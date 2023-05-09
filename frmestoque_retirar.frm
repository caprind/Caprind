VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmestoque_Retirar 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Estoque - Movimentação - Retirada"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15345
   ClipControls    =   0   'False
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
   ForeColor       =   &H00000000&
   Icon            =   "frmestoque_retirar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15345
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
      FormWidthDT     =   15465
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15345
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin TabDlg.SSTab SStab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   45
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
      TabCaption(0)   =   "RM/NF"
      TabPicture(0)   =   "frmestoque_retirar.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameNF"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameRM"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Listamaterial"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "USToolBar1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "FrRastreabilidade"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "RE"
      TabPicture(1)   =   "frmestoque_retirar.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "USToolBar2"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
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
         Height          =   795
         Left            =   60
         TabIndex        =   89
         Top             =   6600
         Width           =   15225
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
            ItemData        =   "frmestoque_retirar.frx":047A
            Left            =   13380
            List            =   "frmestoque_retirar.frx":0484
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   148
            TabStop         =   0   'False
            Top             =   360
            Width           =   1755
         End
         Begin VB.TextBox Txt_cod_ref_RE 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
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
            Left            =   8625
            Locked          =   -1  'True
            TabIndex        =   95
            TabStop         =   0   'False
            ToolTipText     =   "Código de referência."
            Top             =   360
            Width           =   1830
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   94
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   360
            Width           =   1605
         End
         Begin VB.TextBox txtdescricao 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   93
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   360
            Width           =   6825
         End
         Begin VB.TextBox txtPeso 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   11520
            Locked          =   -1  'True
            TabIndex        =   92
            TabStop         =   0   'False
            ToolTipText     =   "Peso unitário."
            Top             =   360
            Width           =   870
         End
         Begin VB.TextBox txtUN 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   12675
            Locked          =   -1  'True
            TabIndex        =   91
            TabStop         =   0   'False
            ToolTipText     =   "Unidade por kilograma."
            Top             =   360
            Width           =   690
         End
         Begin VB.TextBox txtunidade 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   10905
            Locked          =   -1  'True
            TabIndex        =   90
            TabStop         =   0   'False
            ToolTipText     =   "Unidade."
            Top             =   360
            Width           =   600
         End
         Begin DrawSuite2022.USButton Cmd_visualizar_arquivo 
            Height          =   315
            Left            =   10470
            TabIndex        =   150
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            DibPicture      =   "frmestoque_retirar.frx":04A6
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
            BorderColor     =   1154291
            BorderColorDisabled=   13160660
            BorderColorDown =   16576
            BorderColorOver =   8438015
            GradientColor1  =   1154291
            GradientColor2  =   1154291
            GradientColor3  =   1154291
            GradientColor4  =   1154291
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   8438015
            GradientColorOver2=   8438015
            GradientColorOver3=   8438015
            GradientColorOver4=   8438015
            GradientColorDown1=   16576
            GradientColorDown2=   16576
            GradientColorDown3=   16576
            GradientColorDown4=   16576
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   5
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operação da lista"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   49
            Left            =   13620
            TabIndex        =   149
            Top             =   150
            Width           =   1260
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. de referência"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   8910
            TabIndex        =   102
            Top             =   150
            Width           =   1350
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código interno*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   390
            TabIndex        =   101
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição técnica"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4575
            TabIndex        =   100
            Top             =   150
            Width           =   1245
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Kg"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   165
            Left            =   12450
            TabIndex        =   99
            Top             =   420
            Width           =   165
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Peso un."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   11640
            TabIndex        =   98
            Top             =   150
            Width           =   630
         End
         Begin VB.Label Label34 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Un/Kg"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   12810
            TabIndex        =   97
            Top             =   150
            Width           =   435
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Un."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   11078
            TabIndex        =   96
            Top             =   150
            Width           =   255
         End
      End
      Begin VB.Frame Frame5 
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
         Height          =   975
         Left            =   60
         TabIndex        =   103
         Top             =   7320
         Width           =   11445
         Begin VB.TextBox txtLocal_armaz 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5220
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   108
            TabStop         =   0   'False
            ToolTipText     =   "Local de armazenamento."
            Top             =   450
            Width           =   6075
         End
         Begin VB.TextBox txtcorrida 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   107
            TabStop         =   0   'False
            ToolTipText     =   "Número da corrida."
            Top             =   450
            Width           =   1920
         End
         Begin VB.TextBox txtCertificado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   2115
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   106
            TabStop         =   0   'False
            ToolTipText     =   "Número do certificado."
            Top             =   450
            Width           =   1950
         End
         Begin VB.TextBox txtQuant_prevista_PC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   105
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade de peças requsitada."
            Top             =   450
            Width           =   1125
         End
         Begin VB.TextBox txtestoqueatual_PC 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   10890
            Locked          =   -1  'True
            TabIndex        =   104
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Estoque de peças disponível."
            Top             =   1380
            Width           =   2055
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Certificado"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2700
            TabIndex        =   113
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Local de armazenamento"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7365
            TabIndex        =   112
            Top             =   240
            Width           =   1785
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Corrida"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   840
            TabIndex        =   111
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. PÇ"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4305
            TabIndex        =   110
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Disponivel PÇ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   11265
            TabIndex        =   109
            Top             =   1320
            Width           =   1125
         End
      End
      Begin VB.Frame FrRastreabilidade 
         BackColor       =   &H00E0E0E0&
         Height          =   1725
         Left            =   9630
         TabIndex        =   165
         Top             =   8250
         Width           =   1875
         Begin DrawSuite2022.USCheckBox chkindividual 
            Height          =   285
            Left            =   180
            TabIndex        =   166
            Top             =   210
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            Caption         =   "Lote rastreável"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin DrawSuite2022.USButton btnRastreabilidade 
            Height          =   1065
            Left            =   180
            TabIndex        =   167
            ToolTipText     =   "Listar itens da nota fiscal eletrônica"
            Top             =   540
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   1879
            DibPicture      =   "frmestoque_retirar.frx":678A
            Caption         =   "Numeros de série Nota fiscal"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   0
            BorderColorDisabled=   13160660
            BorderColorDown =   4210752
            BorderColorOver =   8421504
            GradientColor1  =   0
            GradientColor2  =   0
            GradientColor3  =   0
            GradientColor4  =   0
            GradientColorDisabled1=   13160660
            GradientColorDisabled2=   13160660
            GradientColorDisabled3=   13160660
            GradientColorDisabled4=   13160660
            GradientColorOver1=   8421504
            GradientColorOver2=   8421504
            GradientColorOver3=   8421504
            GradientColorOver4=   8421504
            GradientColorDown1=   4210752
            GradientColorDown2=   4210752
            GradientColorDown3=   4210752
            GradientColorDown4=   4210752
            PicAlign        =   7
            PicSize         =   3
            PicSizeH        =   32
            PicSizeW        =   32
            ShowFocusRect   =   0   'False
            Theme           =   6
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar lista da..."
         ForeColor       =   &H00000080&
         Height          =   915
         Left            =   60
         TabIndex        =   151
         Top             =   1320
         Width           =   2205
         Begin DrawSuite2022.USButton btnRequisicao 
            Height          =   555
            Left            =   180
            TabIndex        =   152
            ToolTipText     =   "Listar itens de requisição de materiais"
            Top             =   300
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   979
            DibPicture      =   "frmestoque_retirar.frx":1704E
            Caption         =   "Requisição"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   1154291
            BorderColorDisabled=   13160660
            BorderColorDown =   16576
            BorderColorOver =   8438015
            GradientColor1  =   1154291
            GradientColor2  =   1154291
            GradientColor3  =   1154291
            GradientColor4  =   1154291
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   8438015
            GradientColorOver2=   8438015
            GradientColorOver3=   8438015
            GradientColorOver4=   8438015
            GradientColorDown1=   16576
            GradientColorDown2=   16576
            GradientColorDown3=   16576
            GradientColorDown4=   16576
            PicAlign        =   7
            PicSize         =   1
            ShowFocusRect   =   0   'False
            Theme           =   5
         End
         Begin DrawSuite2022.USButton btnNotafiscal 
            Height          =   555
            Left            =   1110
            TabIndex        =   153
            ToolTipText     =   "Listar itens da nota fiscal eletrônica"
            Top             =   300
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   979
            DibPicture      =   "frmestoque_retirar.frx":1D332
            Caption         =   "Nota fiscal"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   0
            BorderColorDisabled=   13160660
            BorderColorDown =   4210752
            BorderColorOver =   8421504
            GradientColor1  =   0
            GradientColor2  =   0
            GradientColor3  =   0
            GradientColor4  =   0
            GradientColorDisabled1=   13160660
            GradientColorDisabled2=   13160660
            GradientColorDisabled3=   13160660
            GradientColorDisabled4=   13160660
            GradientColorOver1=   8421504
            GradientColorOver2=   8421504
            GradientColorOver3=   8421504
            GradientColorOver4=   8421504
            GradientColorDown1=   4210752
            GradientColorDown2=   4210752
            GradientColorDown3=   4210752
            GradientColorDown4=   4210752
            PicAlign        =   7
            PicSize         =   1
            ShowFocusRect   =   0   'False
            Theme           =   6
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   75
         TabIndex        =   79
         Top             =   330
         Width           =   15225
         _ExtentX        =   26855
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
         ButtonCaption1  =   "Retirar"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Retirar (F3)"
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
         ButtonWidth1    =   41
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Retirar selecionados"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Retirar materiais selecionados do estoque (F6)"
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
         ButtonLeft2     =   45
         ButtonTop2      =   2
         ButtonWidth2    =   105
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Status"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Status (F7)"
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
         ButtonLeft3     =   152
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
         ButtonLeft4     =   193
         ButtonTop4      =   4
         ButtonWidth4    =   2
         ButtonHeight4   =   54
         ButtonCaption5  =   "Ajuda"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Ajuda (F1)"
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
         ButtonLeft5     =   197
         ButtonTop5      =   2
         ButtonWidth5    =   36
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Sair"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Sair (Esc)"
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
         ButtonLeft6     =   235
         ButtonTop6      =   2
         ButtonWidth6    =   26
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonKey7      =   "7"
         ButtonAlignment7=   2
         BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState7    =   5
         ButtonLeft7     =   263
         ButtonTop7      =   2
         ButtonWidth7    =   24
         ButtonHeight7   =   24
         ButtonUseMaskColor7=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   6330
            Top             =   150
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmestoque_retirar.frx":2DBF6
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView Listamaterial 
         Height          =   4350
         Left            =   60
         TabIndex        =   141
         Top             =   2250
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   7673
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "Requisitado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Retirado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Retirar"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Qtde. PÇ"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Un."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Código interno"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   9022
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Status"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   915
         Left            =   8340
         TabIndex        =   82
         Top             =   1320
         Width           =   6945
         Begin VB.ComboBox Cmb_empresa 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "frmestoque_retirar.frx":31036
            Left            =   2820
            List            =   "frmestoque_retirar.frx":31038
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Empresa."
            Top             =   450
            Width           =   4035
         End
         Begin VB.TextBox txtIDPedido 
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
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   8430
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.ComboBox Cmb_tipoNF 
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
            ItemData        =   "frmestoque_retirar.frx":3103A
            Left            =   210
            List            =   "frmestoque_retirar.frx":31044
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   1
            ToolTipText     =   "Tipo do formulário."
            Top             =   1380
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtresponsavel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Responsável."
            Top             =   450
            Width           =   2685
         End
         Begin VB.ComboBox cmbN_ref 
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
            ItemData        =   "frmestoque_retirar.frx":3105C
            Left            =   10515
            List            =   "frmestoque_retirar.frx":3105E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            ToolTipText     =   "Código de referência."
            Top             =   4830
            Width           =   1935
         End
         Begin VB.Label Label19 
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
            Left            =   4470
            TabIndex        =   87
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. de referência"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   10800
            TabIndex        =   86
            Top             =   4890
            Width           =   1350
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   615
            TabIndex        =   84
            Top             =   1170
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   1035
            TabIndex        =   85
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         Height          =   915
         Left            =   4650
         TabIndex        =   142
         Top             =   1320
         Width           =   3675
         Begin DrawSuite2022.USButton cmdlote 
            Height          =   555
            Left            =   270
            TabIndex        =   0
            ToolTipText     =   "Filtrar"
            Top             =   270
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   979
            DibPicture      =   "frmestoque_retirar.frx":31060
            Caption         =   "Filtrar documento"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   5263559
            BorderColorDisabled=   13160660
            BorderColorDown =   4013465
            BorderColorOver =   4408288
            GradientColor1  =   5263559
            GradientColor2  =   5263559
            GradientColor3  =   5263559
            GradientColor4  =   5263559
            GradientColorDisabled1=   13160660
            GradientColorDisabled2=   13160660
            GradientColorDisabled3=   13160660
            GradientColorDisabled4=   13160660
            GradientColorOver1=   4408288
            GradientColorOver2=   4408288
            GradientColorOver3=   4408288
            GradientColorOver4=   4408288
            GradientColorDown1=   4013465
            GradientColorDown2=   4013465
            GradientColorDown3=   4013465
            GradientColorDown4=   4013465
            PicAlign        =   7
            PicSize         =   1
            ShowFocusRect   =   0   'False
            Theme           =   4
         End
         Begin DrawSuite2022.USButton btnListaNota 
            Height          =   525
            Left            =   1980
            TabIndex        =   145
            Top             =   300
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   926
            DibPicture      =   "frmestoque_retirar.frx":346B0
            Caption         =   "Listar documentos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   4960354
            BorderColorDisabled=   13160660
            BorderColorDown =   4210752
            BorderColorOver =   49152
            GradientColor1  =   4960354
            GradientColor2  =   4960354
            GradientColor3  =   4960354
            GradientColor4  =   4960354
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   49152
            GradientColorOver2=   49152
            GradientColorOver3=   49152
            GradientColorOver4=   49152
            GradientColorDown1=   32768
            GradientColorDown2=   32768
            GradientColorDown3=   32768
            GradientColorDown4=   32768
            PicAlign        =   8
            PicSize         =   1
            ShowFocusRect   =   0   'False
            Theme           =   3
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Height          =   2655
         Left            =   11520
         TabIndex        =   133
         Top             =   7320
         Width           =   3765
         Begin VB.TextBox txtEmpenhos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   2145
            Locked          =   -1  'True
            TabIndex        =   162
            TabStop         =   0   'False
            Text            =   "0,0000"
            ToolTipText     =   "Estoque disponível."
            Top             =   630
            Width           =   1515
         End
         Begin VB.TextBox txtestoqueatual 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   2145
            Locked          =   -1  'True
            TabIndex        =   139
            TabStop         =   0   'False
            Text            =   "0,0000"
            ToolTipText     =   "Estoque disponível."
            Top             =   240
            Width           =   1515
         End
         Begin VB.TextBox txtquantretirado 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   2145
            TabIndex        =   135
            Text            =   "0,0000"
            ToolTipText     =   "Quantidade de saída."
            Top             =   1380
            Width           =   1515
         End
         Begin VB.TextBox txtRetirar 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2145
            Locked          =   -1  'True
            TabIndex        =   134
            TabStop         =   0   'False
            Text            =   "0,0000"
            ToolTipText     =   "Quantidade requsitada."
            Top             =   1020
            Width           =   1515
         End
         Begin DrawSuite2022.USButton BtnBaixar 
            Height          =   765
            Left            =   2190
            TabIndex        =   136
            ToolTipText     =   "Baixar item do estoque"
            Top             =   1770
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   1349
            DibPicture      =   "frmestoque_retirar.frx":3A994
            Caption         =   "Retirar estoque"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderColor     =   5263559
            BorderColorDisabled=   13160660
            BorderColorDown =   4013465
            BorderColorOver =   4408288
            GradientColor1  =   5263559
            GradientColor2  =   5263559
            GradientColor3  =   5263559
            GradientColor4  =   5263559
            GradientColorDisabled1=   13160660
            GradientColorDisabled2=   13160660
            GradientColorDisabled3=   13160660
            GradientColorDisabled4=   13160660
            GradientColorOver1=   4408288
            GradientColorOver2=   4408288
            GradientColorOver3=   4408288
            GradientColorOver4=   4408288
            GradientColorDown1=   4013465
            GradientColorDown2=   4013465
            GradientColorDown3=   4013465
            GradientColorDown4=   4013465
            PicAlign        =   8
            PicSize         =   3
            PicSizeH        =   32
            PicSizeW        =   32
            ShowFocusRect   =   0   'False
            Theme           =   4
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin MSComCtl2.DTPicker txtdata 
            Height          =   375
            Left            =   390
            TabIndex        =   143
            ToolTipText     =   "Data da movimentação."
            Top             =   2130
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
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
            Format          =   127467521
            CurrentDate     =   39057
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empenhos na RE:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Left            =   630
            TabIndex        =   163
            Top             =   720
            Width           =   1440
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data retirada"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   270
            Left            =   495
            TabIndex        =   144
            Top             =   1860
            Width           =   1275
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Disponível na RE:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   210
            Left            =   705
            TabIndex        =   140
            Top             =   330
            Width           =   1380
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantidade á retirar: "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   210
            Left            =   420
            TabIndex        =   138
            Top             =   1440
            Width           =   1755
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total requisitado:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   210
            Left            =   660
            TabIndex        =   137
            Top             =   1050
            Width           =   1425
         End
      End
      Begin VB.Frame Frame6 
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
         Height          =   1725
         Left            =   60
         TabIndex        =   114
         Top             =   8250
         Width           =   9555
         Begin VB.TextBox txtBaixado 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
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
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   131
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade requsitada."
            Top             =   540
            Width           =   945
         End
         Begin VB.TextBox Txt_cracha 
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
            Left            =   150
            MaxLength       =   8
            TabIndex        =   3
            ToolTipText     =   "Número do cracha."
            Top             =   540
            Width           =   855
         End
         Begin VB.TextBox txtPedidoCompra 
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
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   128
            TabStop         =   0   'False
            ToolTipText     =   "Número do pedido."
            Top             =   1200
            Width           =   1080
         End
         Begin VB.TextBox txtFornecedor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1650
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   127
            TabStop         =   0   'False
            ToolTipText     =   "Fornecedor."
            Top             =   1200
            Width           =   7755
         End
         Begin VB.ComboBox Cmb_RE 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000080&
            Height          =   315
            ItemData        =   "frmestoque_retirar.frx":3E82E
            Left            =   4410
            List            =   "frmestoque_retirar.frx":3E830
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "Número da rastreabilidade de estoque."
            Top             =   540
            Width           =   900
         End
         Begin VB.TextBox txtRequisitante 
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
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   118
            TabStop         =   0   'False
            ToolTipText     =   "Nome do requisitante."
            Top             =   540
            Width           =   2295
         End
         Begin VB.TextBox txtQuant_prevista 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   7170
            Locked          =   -1  'True
            TabIndex        =   117
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade requsitada."
            Top             =   540
            Width           =   855
         End
         Begin VB.ComboBox cmbDestino 
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
            ItemData        =   "frmestoque_retirar.frx":3E832
            Left            =   3330
            List            =   "frmestoque_retirar.frx":3E83C
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "Destino."
            Top             =   540
            Width           =   1065
         End
         Begin VB.TextBox txtDataRE 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   6255
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   115
            TabStop         =   0   'False
            ToolTipText     =   "Data da RE"
            Top             =   540
            Width           =   900
         End
         Begin VB.TextBox Txt_lote 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5310
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   116
            TabStop         =   0   'False
            ToolTipText     =   "Número do lote."
            Top             =   540
            Width           =   930
         End
         Begin VB.ComboBox cmb_Lote 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmestoque_retirar.frx":3E86C
            Left            =   5310
            List            =   "frmestoque_retirar.frx":3E86E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   126
            ToolTipText     =   "Número da rastreabilidade de estoque."
            Top             =   540
            Visible         =   0   'False
            Width           =   930
         End
         Begin DrawSuite2022.USButton cmdListaRetirado 
            Height          =   315
            Left            =   9000
            TabIndex        =   146
            Top             =   540
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            DibPicture      =   "frmestoque_retirar.frx":3E870
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
            BorderColor     =   1154291
            BorderColorDisabled=   13160660
            BorderColorDown =   16576
            BorderColorOver =   8438015
            GradientColor1  =   1154291
            GradientColor2  =   1154291
            GradientColor3  =   1154291
            GradientColor4  =   1154291
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   8438015
            GradientColorOver2=   8438015
            GradientColorOver3=   8438015
            GradientColorOver4=   8438015
            GradientColorDown1=   16576
            GradientColorDown2=   16576
            GradientColorDown3=   16576
            GradientColorDown4=   16576
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   5
         End
         Begin DrawSuite2022.USButton cmdLocalizaPedido 
            Height          =   315
            Left            =   1230
            TabIndex        =   147
            Top             =   1200
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            DibPicture      =   "frmestoque_retirar.frx":44B54
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
            BorderColor     =   4960354
            BorderColorDisabled=   13160660
            BorderColorDown =   4210752
            BorderColorOver =   49152
            GradientColor1  =   4960354
            GradientColor2  =   4960354
            GradientColor3  =   4960354
            GradientColor4  =   4960354
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorOver1=   49152
            GradientColorOver2=   49152
            GradientColorOver3=   49152
            GradientColorOver4=   49152
            GradientColorDown1=   32768
            GradientColorDown2=   32768
            GradientColorDown3=   32768
            GradientColorDown4=   32768
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            Theme           =   3
         End
         Begin DrawSuite2022.USCheckBox chkEmpenhos 
            Height          =   285
            Left            =   150
            TabIndex        =   164
            Top             =   0
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   503
            Caption         =   "Filtrar somente RE(s) empenhada(s)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
            ShowFocusRect   =   0   'False
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Retirado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   8160
            TabIndex        =   132
            Top             =   330
            Width           =   735
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ped. compra"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   255
            TabIndex        =   130
            Top             =   990
            Width           =   900
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedor"
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   5
            Left            =   5040
            TabIndex        =   129
            Top             =   990
            Width           =   945
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° lote"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5535
            TabIndex        =   125
            Top             =   330
            Width           =   495
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Requisitante*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1792
            TabIndex        =   124
            Top             =   330
            Width           =   990
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Requisit."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   7215
            TabIndex        =   123
            Top             =   330
            Width           =   735
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destino*"
            ForeColor       =   &H00000040&
            Height          =   195
            Index           =   0
            Left            =   3540
            TabIndex        =   122
            Top             =   330
            Width           =   630
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº RE*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   4515
            TabIndex        =   121
            Top             =   330
            Width           =   555
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº cracha*"
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   2
            Left            =   195
            TabIndex        =   120
            Top             =   330
            Width           =   795
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data RE"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6420
            TabIndex        =   119
            Top             =   330
            Width           =   585
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   8640
         Left            =   -74925
         TabIndex        =   46
         Top             =   1320
         Width           =   15225
         Begin VB.TextBox Txt_cracha_RE 
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
            MaxLength       =   8
            TabIndex        =   19
            ToolTipText     =   "Número do cracha."
            Top             =   6930
            Width           =   1125
         End
         Begin VB.TextBox Txt_cod_ref_RE_RE 
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
            Left            =   2025
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Código de referência."
            Top             =   6330
            Width           =   2100
         End
         Begin VB.TextBox txtEstoque_Real_RE_PC 
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
            Left            =   4629
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Total estoque peças."
            Top             =   8160
            Width           =   1455
         End
         Begin VB.TextBox txtEstoque_Real_RE 
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
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Total estoque."
            Top             =   8160
            Width           =   1455
         End
         Begin VB.TextBox txtEmpenho_PC 
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
            Left            =   6112
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Empenho de peças."
            Top             =   8160
            Width           =   1455
         End
         Begin VB.TextBox txtEmpenho 
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
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Empenho."
            Top             =   0
            Width           =   1455
         End
         Begin VB.TextBox txtAtualizado_PC_RE 
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
            Left            =   13531
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Estoque de peças disponível atualizado."
            Top             =   8160
            Width           =   1484
         End
         Begin VB.TextBox txtDisponivel_PC_RE 
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
            Left            =   7595
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Estoque de peças disponível."
            Top             =   8160
            Width           =   1455
         End
         Begin VB.TextBox txtQtde_Saida_RE_PC 
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
            Left            =   10561
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Quantidade de saída de peças."
            Top             =   8160
            Width           =   1455
         End
         Begin VB.TextBox txtQtde_PC_RE 
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
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade de peças requsitada."
            Top             =   7500
            Width           =   1125
         End
         Begin VB.CommandButton cmdArquivo_RE 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   10110
            Picture         =   "frmestoque_retirar.frx":4E601
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Visualizar arquivo."
            Top             =   6930
            Width           =   315
         End
         Begin VB.TextBox txtCertificado_RE 
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
            Left            =   12885
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Número do certificado."
            Top             =   6330
            Width           =   2130
         End
         Begin VB.TextBox txtCorrida_RE 
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
            Left            =   10770
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Número da corrida."
            Top             =   6330
            Width           =   2100
         End
         Begin VB.TextBox txtLocal_RE 
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
            Left            =   4135
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Local de armazenamento."
            Top             =   6330
            Width           =   6620
         End
         Begin VB.TextBox txtLote_RE 
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
            MaxLength       =   255
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Número do lote."
            Top             =   6330
            Width           =   1830
         End
         Begin VB.TextBox txtIdPedido_RE 
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
            Left            =   6450
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   7500
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.TextBox txtPedido_RE 
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
            Left            =   6450
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Número do pedido."
            Top             =   7500
            Width           =   990
         End
         Begin VB.TextBox txtFornecedor_RE 
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
            Left            =   7860
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Fornecedor."
            Top             =   7500
            Width           =   4875
         End
         Begin VB.CommandButton cmdPedido_RE 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
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
            Height          =   315
            Left            =   7470
            Picture         =   "frmestoque_retirar.frx":4EBC3
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Localizar pedido de compra."
            Top             =   7500
            Width           =   315
         End
         Begin VB.ComboBox cmbDestino_RE 
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
            ItemData        =   "frmestoque_retirar.frx":4ECC5
            Left            =   5820
            List            =   "frmestoque_retirar.frx":4ECCF
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   21
            ToolTipText     =   "Destino."
            Top             =   6930
            Width           =   2235
         End
         Begin VB.ComboBox cmbEmpresa_RE 
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
            ItemData        =   "frmestoque_retirar.frx":4ECFF
            Left            =   180
            List            =   "frmestoque_retirar.frx":4ED01
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            ToolTipText     =   "Empresa."
            Top             =   390
            Width           =   6375
         End
         Begin VB.TextBox txtUN_RE 
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
            Left            =   12495
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Unidade."
            Top             =   6930
            Width           =   660
         End
         Begin VB.TextBox txtDisponivel_RE 
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
            Left            =   3146
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Estoque disponível."
            Top             =   8160
            Width           =   1455
         End
         Begin VB.TextBox txtAtualizado_RE 
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
            Left            =   12044
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Estoque disponível atualizado."
            Top             =   8160
            Width           =   1455
         End
         Begin VB.TextBox txtQtde_Saida_RE 
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
            Left            =   9060
            TabIndex        =   41
            Text            =   "0,000"
            ToolTipText     =   "Quantidade de saída."
            Top             =   8160
            Width           =   1485
         End
         Begin VB.TextBox txtQtde_RE 
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
            Left            =   12750
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade requsitada."
            Top             =   7500
            Width           =   1125
         End
         Begin VB.TextBox txtRequisitante_RE 
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
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Nome do requisitante."
            Top             =   6930
            Width           =   4500
         End
         Begin VB.TextBox txtRE 
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
            Left            =   13140
            TabIndex        =   11
            ToolTipText     =   "Número da rastreabilidade de estoque."
            Top             =   390
            Width           =   1545
         End
         Begin VB.CommandButton cmdRE 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   14700
            Picture         =   "frmestoque_retirar.frx":4ED03
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Filtrar número da número da rastreabilidade de estoque."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtUnKg_RE 
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
            Left            =   14325
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "Unidade por kilograma."
            Top             =   6930
            Width           =   690
         End
         Begin VB.TextBox txtPeso_RE 
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
            Left            =   13170
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Peso unitário."
            Top             =   6930
            Width           =   870
         End
         Begin VB.ComboBox cmbReferencia_RE 
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
            ItemData        =   "frmestoque_retirar.frx":4F11E
            Left            =   10545
            List            =   "frmestoque_retirar.frx":4F120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   24
            ToolTipText     =   "Código de referência."
            Top             =   6930
            Width           =   1935
         End
         Begin VB.TextBox txtDescricao_RE 
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
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   7500
            Width           =   6255
         End
         Begin VB.TextBox txtCodInterno_RE 
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
            Left            =   8070
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   6930
            Width           =   2025
         End
         Begin VB.TextBox txtResponsavel_RE 
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
            Left            =   7950
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Responsável."
            Top             =   390
            Width           =   5175
         End
         Begin MSComctlLib.ListView ListaMaterial_RE 
            Height          =   5040
            Left            =   180
            TabIndex        =   13
            Top             =   1005
            Width           =   14835
            _ExtentX        =   26167
            _ExtentY        =   8890
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
               Object.Tag             =   "N"
               Text            =   "Ordem"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Object.Tag             =   "N"
               Text            =   "Quantidade"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Qtde. PÇ"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Un."
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Código interno"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   12197
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Object.Tag             =   "T"
               Text            =   "Saída estoque"
               Object.Width           =   2646
            EndProperty
         End
         Begin MSComCtl2.DTPicker txtData_RE 
            Height          =   315
            Left            =   6555
            TabIndex        =   9
            ToolTipText     =   "Data da movimentação."
            Top             =   390
            Width           =   1395
            _ExtentX        =   2461
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
            Format          =   197132289
            CurrentDate     =   39057
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº do cracha*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   232
            TabIndex        =   88
            Top             =   6720
            Width           =   1020
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. de referência"
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
            Left            =   2310
            TabIndex        =   81
            Top             =   6130
            Width           =   1530
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total estoque PÇ"
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
            Left            =   4644
            TabIndex        =   78
            Top             =   7950
            Width           =   1425
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total estoque"
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
            Left            =   322
            TabIndex        =   77
            Top             =   7950
            Width           =   1170
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(-) Empenho PÇ"
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
            Left            =   6187
            TabIndex        =   76
            Top             =   7950
            Width           =   1305
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(-) Empenho"
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
            Left            =   1865
            TabIndex        =   75
            Top             =   7950
            Width           =   1050
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Est. atualizado PÇ"
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
            Left            =   13531
            TabIndex        =   74
            Top             =   7950
            Width           =   1485
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(=) Est. disp. PÇ"
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
            Left            =   7662
            TabIndex        =   73
            Top             =   7950
            Width           =   1320
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtd. saída PÇ*"
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
            Left            =   10688
            TabIndex        =   72
            Top             =   7950
            Width           =   1200
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. PÇ"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   14115
            TabIndex        =   71
            Top             =   7290
            Width           =   660
         End
         Begin VB.Label Label47 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ped. compra"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6495
            TabIndex        =   70
            Top             =   7290
            Width           =   900
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fornecedor"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   9885
            TabIndex        =   69
            Top             =   7290
            Width           =   825
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destino*"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   6622
            TabIndex        =   68
            Top             =   6720
            Width           =   630
         End
         Begin VB.Label Label46 
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
            Left            =   3000
            TabIndex        =   67
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label45 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Un."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   12690
            TabIndex        =   66
            Top             =   6720
            Width           =   255
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Est. atualizado"
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
            Left            =   12156
            TabIndex        =   65
            Top             =   7950
            Width           =   1230
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtd. saída*"
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
            Left            =   9330
            TabIndex        =   64
            Top             =   7950
            Width           =   945
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(=) Est. disp."
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
            Left            =   3341
            TabIndex        =   63
            Top             =   7950
            Width           =   1065
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lista de requisições"
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
            Index           =   9
            Left            =   6765
            TabIndex        =   62
            Top             =   780
            Width           =   1665
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13095
            TabIndex        =   61
            Top             =   7290
            Width           =   420
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Requisitante*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3075
            TabIndex        =   60
            Top             =   6720
            Width           =   990
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Corrida"
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
            Left            =   11520
            TabIndex        =   59
            Top             =   6130
            Width           =   615
         End
         Begin VB.Label Label38 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Un/Kg"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   14460
            TabIndex        =   58
            Top             =   6720
            Width           =   435
         End
         Begin VB.Label Label37 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Peso un."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13290
            TabIndex        =   57
            Top             =   6720
            Width           =   630
         End
         Begin VB.Label Label36 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Local de armazenamento"
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
            Left            =   6375
            TabIndex        =   56
            Top             =   6130
            Width           =   2145
         End
         Begin VB.Label Label35 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Kg"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   165
            Left            =   14070
            TabIndex        =   55
            Top             =   6990
            Width           =   165
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. de referência"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   10837
            TabIndex        =   54
            Top             =   6720
            Width           =   1350
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição técnica"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2685
            TabIndex        =   53
            Top             =   7290
            Width           =   1245
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código interno*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   8512
            TabIndex        =   52
            Top             =   6720
            Width           =   1140
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº RE*"
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
            Left            =   13635
            TabIndex        =   51
            Top             =   180
            Width           =   555
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   10080
            TabIndex        =   50
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Certificado"
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
            Left            =   13500
            TabIndex        =   49
            Top             =   6130
            Width           =   915
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7080
            TabIndex        =   48
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° lote"
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
            Left            =   840
            TabIndex        =   47
            Top             =   6130
            Width           =   570
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   80
         Top             =   330
         Width           =   15225
         _ExtentX        =   26855
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
         ButtonCaption1  =   "Retirar"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Retirar (F3)"
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
         ButtonWidth1    =   41
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Retirar selecionados"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Retirar materiais selecionados do estoque (F6)"
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
         ButtonLeft2     =   45
         ButtonTop2      =   2
         ButtonWidth2    =   105
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
         ButtonLeft3     =   152
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
         ButtonLeft4     =   156
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
         ButtonLeft5     =   194
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
         ButtonLeft6     =   222
         ButtonTop6      =   2
         ButtonWidth6    =   24
         ButtonHeight6   =   24
         ButtonUseMaskColor6=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   6330
            Top             =   150
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmestoque_retirar.frx":4F122
            Count           =   1
         End
      End
      Begin VB.Frame FrameRM 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados do documento"
         Height          =   915
         Left            =   2280
         TabIndex        =   154
         Top             =   1320
         Width           =   2355
         Begin VB.TextBox txt_RM 
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
            Left            =   210
            TabIndex        =   155
            ToolTipText     =   "Numero da requisição de materiais"
            Top             =   420
            Width           =   1905
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº requisição*"
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   630
            TabIndex        =   156
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.Frame FrameNF 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados do documento"
         Height          =   915
         Left            =   2280
         TabIndex        =   157
         Top             =   1320
         Width           =   2355
         Begin VB.TextBox txt_Notafiscal 
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
            Left            =   210
            TabIndex        =   159
            ToolTipText     =   "Numero da nota fiscal"
            Top             =   420
            Width           =   1485
         End
         Begin VB.TextBox Txt_serieNF 
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
            Left            =   1710
            TabIndex        =   158
            ToolTipText     =   "Série da nota fiscal"
            Top             =   420
            Width           =   435
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº nota fiscal*"
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   480
            TabIndex        =   161
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Série"
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   1755
            TabIndex        =   160
            Top             =   240
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "frmestoque_Retirar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Saida As Double 'OK
Dim Requisicao_materiais As Boolean 'OK
Dim Load As Boolean 'OK
Dim Expedir As Boolean

'Corrige formulario
Dim Width_txtresponsavel As Long
Dim Width_Cmb_tipoNF As Long
Dim Width_Txt_RM As Long
Dim Width_Txt_serieNF As Long

Public Sub ProcAtualizaTodas_Listas()
On Error GoTo tratar_erro

Acao = "filtrar"
If Cmb_empresa = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa.SetFocus
    Exit Sub
End If
If txt_Notafiscal = "" And FrameNF.Visible = True Then
    NomeCampo = "0 número da"
    NomeCampo = NomeCampo & " nota fiscal"
    ProcVerificaAcao
    txt_Notafiscal.SetFocus
    Exit Sub
End If

If txt_RM = "" And FrameRM.Visible = True Then
    NomeCampo = "0 número da"
    NomeCampo = NomeCampo & " requisição de materiais"
    ProcVerificaAcao
    txt_RM.SetFocus
    Exit Sub
End If

ProcLimpaCampos
Listamaterial.ListItems.Clear

'===================================================
' Carrega itens da nota fiscal
'===================================================
If Expedir = True Then
    If Txt_serieNF = "" Then
        NomeCampo = "o número de série"
        ProcVerificaAcao
        Txt_serieNF.SetFocus
        Exit Sub
    End If
    If Cmb_tipoNF = "Produtos" Then TipoTexto = "M1" Else TipoTexto = "SA"
    Set TBMateriaprima = CreateObject("adodb.recordset")
    StrSql = "Select NFP.*,NF.* from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Detalhes_Nota NFP ON NFP.ID_nota = NF.ID inner join projProduto SCRE on NFP.codproduto = SCRE.codproduto where NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and NF.int_NotaFiscal = '" & txt_Notafiscal.Text & "' and NF.TipoNF = '" & TipoTexto & "' and Serie = '" & Txt_serieNF & "' and NF.int_status = 1 and NF.int_TipoNota = 1 and NF.DtValidacao IS NOT NULL and SCRE.Estoque = 'true' order by NFP.int_codigo"
    'Debug.print StrSql
    
    TBMateriaprima.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBMateriaprima.EOF = False Then
    txtData.Value = TBMateriaprima!dt_DataEmissao
    ID_nota = TBMateriaprima!ID
        ProcAtualizaLista_NF
    Else
        USMsgBox ("Não foi encontrado nenhum produto com que movimenta estoque com essa nota fiscal do tipo " & Cmb_tipoNF & " com este número e série, ou a mesma está cancelada."), vbExclamation, "CAPRIND v5.0"
    End If
Else
'===================================================
' Carrega itens da requisição de materiais
'===================================================
    If IsNumeric(txt_RM) = True Then
        Requisicao_materiais = False
        Set TBproducao = CreateObject("adodb.recordset")
        TBproducao.Open "Select * from producao where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Ordem = " & txt_RM.Text & " and Status <> 'Cancelada' and DtValidacao IS NOT NULL and DtValidacao_Custo IS NULL order by Desenho", Conexao, adOpenKeyset, adLockOptimistic
        If TBproducao.EOF = False Then
            txt_RM = TBproducao!Ordem
            ProcAtualizalista
        Else
            USMsgBox ("Não foi encontrado nenhuma ordem validada com este número ou o resultado da ordem já foi validado ou está cancelada."), vbExclamation, "CAPRIND v5.0"
        End If
    Else
        Requisicao_materiais = True
        Set TBMateriaprima = CreateObject("adodb.recordset")
        TBMateriaprima.Open "Select ID from Requisicao_materiais where requisicao = '" & txt_RM & "' and DtValidacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
        If TBMateriaprima.EOF = False Then
            Set TBMateriaprima = CreateObject("adodb.recordset")
            TBMateriaprima.Open "Select R.Requisicao, RML.* from Requisicao_materiais R INNER JOIN Requisicao_materiais_lista RML ON R.ID = RML.IDrequisicao where R.requisicao = '" & txt_RM & "' and (RML.ID_CC is not null and RML.ID_CC <> 0 and RML.Data_autorizacao is not null or RML.ID_CC is null or RML.ID_CC = 0) order by RML.Desenho", Conexao, adOpenKeyset, adLockOptimistic
            If TBMateriaprima.EOF = False Then
                txt_RM = TBMateriaprima!requisicao
                ProcAtualizaLista_RM
            Else
                USMsgBox ("É necessário autorizar o centro de custo do(s) produto(s) para baixar a RM."), vbExclamation, "CAPRIND v5.0"
            End If
        Else
            USMsgBox ("Não foi encontrado nenhuma requisição validada com este número."), vbExclamation, "CAPRIND v5.0"
        End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizalista()
On Error GoTo tratar_erro
Dim Retirar As Double
Dim Retirado As Double
Dim Saldo As Double

Listamaterial.ListItems.Clear
Set TBMateriaprima = CreateObject("adodb.recordset")
TBMateriaprima.Open "Select Tipo_item, Idmateriaprima, Requisitado, Total_pc, Unidade, CODIGO, Descricao, Saida from producaomaterial where Ordem = " & txt_RM.Text & "  order by Posicao, Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBMateriaprima.EOF = False Then
    Do While TBMateriaprima.EOF = False
        Retirar = 0
        Retirado = 0
        Saldo = 0
        SaidaPC = 0
        Set TBEstoque = CreateObject("adodb.recordset")
        TBEstoque.Open "Select SUM(EM.Saida) as Saida, SUM(EM.Saida_pc) as SaidaPC, EC.Un from (Estoque_movimentacao EM INNER JOIN Estoque_Controle EC ON EC.IDestoque = EM.IDestoque) where EM.Ordem = " & txt_RM.Text & " and EM.desenho = '" & TBMateriaprima!CODIGO & "' and EM.documento = '" & txt_RM & "' and Left(EM.Operacao, 11) = 'SAIDA_ORDEM' group by EC.Un", Conexao, adOpenKeyset, adLockOptimistic
        If TBEstoque.EOF = False Then
            Retirado = IIf(IsNull(TBEstoque!Saida), 0, TBEstoque!Saida)
            SaidaPC = IIf(IsNull(TBEstoque!SaidaPC), 0, TBEstoque!SaidaPC)
            
            If IsNull(TBMateriaprima!Total_pc) = False And TBMateriaprima!Total_pc > 0 And SaidaPC = 0 Then SaidaPC = FunCalculaQtdePC(TBMateriaprima!CODIGO, Saida, True, TBMateriaprima!Unidade)
        End If
        TBEstoque.Close
        With Listamaterial.ListItems
        Retirar = IIf(IsNull(TBMateriaprima!Requisitado), 0, TBMateriaprima!Requisitado)
             Saldo = Retirar - Retirado
             
            If TBMateriaprima!Saida = "CANCELADO" Then
                status = "CANCELADO"
            End If
            
            If Retirado = 0 Then
                    status = "REQUISITADO"
            End If
            
            If Saldo <= 0 Then
                    status = "RETIRADO"
            End If
            
            If Saldo > 0 And Saldo < Retirar Then
                    status = "PARCIAL"
            End If
             
            .Add , , TBMateriaprima!IdMateriaPrima
            .Item(.Count).SubItems(1) = Format(IIf(Retirar < 0, 0, Retirar), "###,##0.0000")
            .Item(.Count).SubItems(2) = Format(Retirado, "###,##0.0000")
            .Item(.Count).SubItems(3) = Format(IIf(Saldo >= 0, Saldo, 0), "###,##0.0000")
            .Item(.Count).SubItems(5) = IIf(IsNull(TBMateriaprima!Unidade), "", TBMateriaprima!Unidade)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBMateriaprima!CODIGO), "", TBMateriaprima!CODIGO)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBMateriaprima!Descricao), "", TBMateriaprima!Descricao)
             .Item(.Count).SubItems(8) = status
            
        End With
        TBMateriaprima.MoveNext
    Loop
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaRM()
On Error GoTo tratar_erro

Qtde = 0
Qtd = 0
Permitido = False
ListaMaterial_RE.ListItems.Clear
CarregaSemEmpenho:
    If Permitido = False Then
        TextoFiltro = "PNFC.IDestoque = " & TBEstoque!IDEstoque & " and PNFC.Quantidade - PNFC.Qtde_saida > 0"
        Permitido2 = True
    Else
        TextoFiltro = "PM.Codigo = '" & txtCodInterno_RE & "' and (PNFC.ID IS NULL or PNFC.Quantidade - PNFC.Qtde_saida <= 0)"
        Permitido2 = False
    End If
    Set TBMateriaprima = CreateObject("adodb.recordset")
    TBMateriaprima.Open "Select PM.Idmateriaprima, PM.Requisitado, PM.Total_pc, PM.Unidade, PM.CODIGO, PM.Descricao, PM.Saida, P.ordem from (Producaomaterial PM LEFT JOIN Producao_NF_Consignada PNFC ON PNFC.Ordem = PM.Ordem and PNFC.Codinterno = PM.Codigo) INNER JOIN Producao P ON P.Ordem = PM.Ordem where " & TextoFiltro & " and P.Status <> 'Cancelada' and P.Dtvalidacao IS NOT NULL and P.DtValidacao_custo IS NULL and (PM.Saida = 'NÃO' OR PM.Saida = 'PARCIAL') order by PM.Ordem, PM.Codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBMateriaprima.EOF = False Then
        Do While TBMateriaprima.EOF = False
            Saida = 0
            SaidaPC = 0
            Set TBCST = CreateObject("adodb.recordset")
            TBCST.Open "Select SUM(EM.Saida) as Saida, SUM(EM.Saida_pc) as SaidaPC, EC.Un from (Estoque_movimentacao EM INNER JOIN Estoque_Controle EC ON EC.IDestoque = EM.IDestoque) where EM.Ordem = " & TBMateriaprima!Ordem & " and EM.desenho = '" & TBMateriaprima!CODIGO & "' and Left(EM.Operacao, 11) = 'SAIDA_ORDEM' GROUP BY EC.Un", Conexao, adOpenKeyset, adLockOptimistic
            If TBCST.EOF = False Then
                Saida = IIf(IsNull(TBCST!Saida), 0, TBCST!Saida)
                SaidaPC = IIf(IsNull(TBCST!SaidaPC), 0, TBCST!SaidaPC)
                
                If IsNull(TBMateriaprima!Total_pc) = False And TBMateriaprima!Total_pc > 0 And SaidaPC = 0 Then SaidaPC = FunCalculaQtdePC(TBMateriaprima!CODIGO, Saida, True, TBMateriaprima!Unidade)
            End If
            TBCST.Close
            
            With ListaMaterial_RE.ListItems.Add(, , TBMateriaprima!IdMateriaPrima)
                .SubItems(1) = IIf(IsNull(TBMateriaprima!Ordem), "", TBMateriaprima!Ordem)
                qt = TBMateriaprima!Requisitado - Saida
                Qtde = Qtde + qt
                .SubItems(2) = Format(IIf(qt < 0, 0, qt), "###,##0.0000")
                qt = IIf(IsNull(TBMateriaprima!Total_pc), 0, TBMateriaprima!Total_pc) - SaidaPC
                Qtd = Qtd + qt
                .SubItems(3) = Format(IIf(qt < 0, 0, qt), "###,##0.0000")
                .SubItems(4) = IIf(IsNull(TBMateriaprima!Unidade), "", TBMateriaprima!Unidade)
                .SubItems(5) = IIf(IsNull(TBMateriaprima!CODIGO), "", TBMateriaprima!CODIGO)
                .SubItems(6) = IIf(IsNull(TBMateriaprima!Descricao), "", TBMateriaprima!Descricao)
                .SubItems(7) = IIf(IsNull(TBMateriaprima!Saida), "", TBMateriaprima!Saida)
                
                If Permitido2 = True Then
                    .ForeColor = vbBlue
                    .ListSubItems(1).ForeColor = vbBlue
                    .ListSubItems(2).ForeColor = vbBlue
                    .ListSubItems(3).ForeColor = vbBlue
                    .ListSubItems(4).ForeColor = vbBlue
                    .ListSubItems(5).ForeColor = vbBlue
                    .ListSubItems(6).ForeColor = vbBlue
                    .ListSubItems(7).ForeColor = vbBlue
                End If
            End With
            
            TBMateriaprima.MoveNext
        Loop
    End If
    
    If Permitido = False And TBEstoque!Estoque = False Then
        If txtQtde_Saida_RE.Locked = False Then
            If Qtde < txtDisponivel_RE Then
                Permitido = True
                GoTo CarregaSemEmpenho
            End If
        Else
            If Qtd < txtDisponivel_PC_RE Then
                Permitido = True
                GoTo CarregaSemEmpenho
            End If
        End If
    End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizaLista_RM()
On Error GoTo tratar_erro
Dim Retirar As Double
Dim Retirado As Double
Dim Saldo As Double

Do While TBMateriaprima.EOF = False
    With Listamaterial.ListItems
        .Add , , TBMateriaprima!IDlista
        Retirar = IIf(IsNull(TBMateriaprima!Quant), 0, TBMateriaprima!Quant)
StrSql = "Select sum(Saida) as Retirado from Estoque_movimentacao where ID_prod_RM = " & TBMateriaprima!IDlista & ""
'Debug.print StrSql

            Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
'                    Dim Saida
'                    Saida = 0
'                    Do While TBAbrir.EOF = False
'                        Saida = Saida + TBAbrir!Saida
'                        TBAbrir.MoveNext
'                    Loop
                    Retirado = IIf(IsNull(TBAbrir!Retirado), 0, TBAbrir!Retirado)
            TBAbrir.Close
            
           Saldo = Retirar - Retirado
           
          
         .Item(.Count).SubItems(1) = Format(IIf(Retirar < 0, 0, Retirar), "###,##0.0000")
         .Item(.Count).SubItems(2) = Format(IIf(Retirado < 0, 0, Retirado), "###,##0.0000")
         .Item(.Count).SubItems(3) = Format(IIf(Saldo < 0, 0, Saldo), "###,##0.0000")
    
        .Item(.Count).SubItems(4) = ""
        .Item(.Count).SubItems(5) = IIf(IsNull(TBMateriaprima!Un), "", TBMateriaprima!Un)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBMateriaprima!Desenho), "", TBMateriaprima!Desenho)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBMateriaprima!Descricao), "", TBMateriaprima!Descricao)
        If TBMateriaprima!status = "CANCELADO" Then
            status = "CANCELADO"
           .Item(.Count).ListSubItems(3).ForeColor = vbBlack
           .Item(.Count).SubItems(8) = status
           .Item(.Count).ListSubItems(8).ForeColor = vbBlack
        ElseIf .Item(.Count).SubItems(2) = 0 Then
            status = "REQUISITADO"
           .Item(.Count).ListSubItems(3).ForeColor = vbRed
           .Item(.Count).SubItems(8) = status
           .Item(.Count).ListSubItems(8).ForeColor = vbRed
        ElseIf .Item(.Count).SubItems(3) = 0 Then
                status = "RETIRADO"
           .Item(.Count).ListSubItems(3).ForeColor = vbBlue
           .Item(.Count).SubItems(8) = status
           .Item(.Count).ListSubItems(8).ForeColor = vbBlue
        ElseIf .Item(.Count).SubItems(3) > 0 Then 'And .Item(.Count).SubItems(3) > .Item(.Count).SubItems(2) Then
            status = "PARCIAL"
           .Item(.Count).ListSubItems(3).ForeColor = vbRed
           .Item(.Count).SubItems(8) = status
           .Item(.Count).ListSubItems(8).ForeColor = vbRed

        End If
    End With
    TBMateriaprima.MoveNext
Loop

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizaLista_NF()
On Error GoTo tratar_erro
Dim strDoc As String

If Expedir = True And txt_Notafiscal.Text <> "" Then
strDoc = txt_Notafiscal.Text
Else
strDoc = txt_RM.Text
End If

Do While TBMateriaprima.EOF = False
    With Listamaterial.ListItems
        .Add , , TBMateriaprima!Int_codigo
        Saida = 0
        Set TBEstoque = CreateObject("adodb.recordset")
        TBEstoque.Open "Select Sum(Saida) as Qt from estoque_movimentacao where OE = '" & strDoc & "' and ID_prod_NF = " & TBMateriaprima!Int_codigo & " and (operacao = 'SAIDA_NOTA' or operacao = 'SAIDA_NOTA_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
        If TBEstoque.EOF = False Then
            Saida = IIf(IsNull(TBEstoque!qt), 0, TBEstoque!qt)
        End If
        TBEstoque.Close
        
        If TBMateriaprima!txt_Unid <> TBMateriaprima!Unidade_com And (TBMateriaprima!txt_Unid = "KG" Or TBMateriaprima!txt_Unid = "MT" Or TBMateriaprima!txt_Unid = "MM" Or TBMateriaprima!txt_Unid = "BR" Or TBMateriaprima!txt_Unid = "PC" Or TBMateriaprima!txt_Unid = "PÇ") And (TBMateriaprima!Unidade_com = "KG" Or TBMateriaprima!Unidade_com = "MT" Or TBMateriaprima!Unidade_com = "MM" Or TBMateriaprima!Unidade_com = "BR" Or TBMateriaprima!Unidade_com = "PC" Or TBMateriaprima!Unidade_com = "PÇ") And IsNull(TBMateriaprima!Qtde_estoque) = False And TBMateriaprima!Qtde_estoque <> 0 Then
            Qtd = TBMateriaprima!Qtde_estoque
        Else
            Qtd = TBMateriaprima!int_Qtd / FunVerificaTabelaConversaoUnidade(TBMateriaprima!txt_Unid, TBMateriaprima!Unidade_com)
        End If
        
        qt = Format(Qtd - Saida, "###,##0.0000")
        .Item(.Count).SubItems(1) = Format(IIf(Qtd < 0, 0, Qtd), "###,##0.0000")
        .Item(.Count).SubItems(2) = Format(IIf(Saida < 0, 0, Saida), "###,##0.0000")
        .Item(.Count).SubItems(3) = Format(IIf(qt < 0, 0, qt), "###,##0.0000")
        .Item(.Count).SubItems(4) = ""
        .Item(.Count).SubItems(5) = IIf(IsNull(TBMateriaprima!txt_Unid), "", TBMateriaprima!txt_Unid)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBMateriaprima!int_Cod_Produto), "", TBMateriaprima!int_Cod_Produto)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBMateriaprima!Txt_descricao), "", TBMateriaprima!Txt_descricao)
        If Saida > 0 Then
            If qt = 0 Then status = "RETIRADO" Else status = "PARCIAL"
            .Item(.Count).SubItems(8) = status
            If status = "RETIRADO" Then
            .Item(.Count).ListSubItems(8).ForeColor = vbBlue
            End If
            If status = "PARCIAL" Then
            .Item(.Count).ListSubItems(8).ForeColor = vbRed
           End If
        Else
            status = "RETIRAR"
            .Item(.Count).SubItems(8) = status
            .Item(.Count).ListSubItems(8).ForeColor = vbRed

        End If
    End With
    TBMateriaprima.MoveNext
Loop

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ActiveResize1_ResizeComplete()
On Error GoTo tratar_erro

Width_txtresponsavel = txtResponsavel.Width
Width_Cmb_tipoNF = Cmb_tipoNF.Width
Load = True


If TemaCaprind <> "" And TemaINI <> "" Then
frmMDI.SkinFramework1.LoadSkin caminho & TemaCaprind, TemaINI
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub BtnBaixar_Click()
On Error GoTo tratar_erro

ProcRetirar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnListaNota_Click()
On Error GoTo tratar_erro

frmEstoque_Retirar_ListaNF.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub btnNotafiscal_Click()
On Error GoTo tratar_erro
Expedir = True

    With USToolBar1
        .ButtonState(3) = 5
        .Refresh
    End With
    With txtData
        .Value = Date
        .Enabled = True
    End With
    Label7(3).Visible = True
    With Cmb_tipoNF
        .Text = "Produtos"
        .Visible = True
    End With
    With Txt_serieNF
        .Text = "1"
        .Visible = True
    End With
    
If Load = False Then
    ProcLimpaCampos
    Listamaterial.ListItems.Clear
    If txt_Notafiscal.Text <> "" Then cmdlote_Click
End If

FrameRM.Visible = False
FrameNF.Visible = True
Exit Sub

tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnRastreabilidade_Click()
On Error GoTo tratar_erro

frmEstoque_Retirar_NumeroSerie.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnRequisicao_Click()
On Error GoTo tratar_erro
Expedir = False

    With USToolBar1
        .ButtonState(3) = 0
        .Refresh
    End With
    
    txtData.Enabled = True
    Label7(3).Visible = False
    Cmb_tipoNF.Visible = False

If Load = False Then
    ProcLimpaCampos
    Listamaterial.ListItems.Clear
    If txtRM <> "" Then cmdlote_Click
End If

FrameRM.Visible = True
FrameNF.Visible = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcLimpaCampos
Listamaterial.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_lote_Click()
On Error GoTo tratar_erro

If Expedir = True Then
    ProcCarregaRE_NF
Else
    If Requisicao_materiais = False Then ProcCarregaRE Else ProcCarregaRE_RM
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With Listamaterial
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar1
    If Cmb_opcao_lista = "Retirar selecionados" Then
        .ButtonState(2) = 0
        .ButtonState(3) = 5
    Else
        .ButtonState(2) = 5
        .ButtonState(3) = 0
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_RE_Click()
On Error GoTo tratar_erro

If Listamaterial.ListItems.Count = 0 Then
Exit Sub
End If

If Cmb_RE = "" Then Exit Sub
ProcLimpaCamposRE
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_produtos where IDestoque = " & Cmb_RE, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    cmb_Lote.Visible = False
    txtDataRE.Text = IIf(IsNull(TBEstoque!Data), "", TBEstoque!Data)
    Txt_lote = IIf(IsNull(TBEstoque!LOTE), "", TBEstoque!LOTE)
    Txt_cod_ref_RE = IIf(IsNull(TBEstoque!Ref), "", TBEstoque!Ref)
    txtLocal_armaz = IIf(IsNull(TBEstoque!local_armaz), "", TBEstoque!local_armaz)
    txtcorrida = IIf(IsNull(TBEstoque!Corrida), "", TBEstoque!Corrida)
    txtCertificado = IIf(IsNull(TBEstoque!Certificado), "", TBEstoque!Certificado)
    Estoquereal = IIf(IsNull(TBEstoque!Estoque_Disponivel), 0, TBEstoque!Estoque_Disponivel)
    EstoquerealPC = IIf(IsNull(TBEstoque!estoque_real_PC), 0, TBEstoque!estoque_real_PC)
    
    valor = 0
    If Expedir = True Then
        'Verifica se este RE já está empenhado para um pedido interno
        TextoFiltro = ""
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select ID_carteira from tbl_Detalhes_Nota_pedidos where ID_prod_NF = " & Listamaterial.SelectedItem & " and Codinterno = '" & Listamaterial.SelectedItem.ListSubItems(6) & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                If TextoFiltro = "" Then TextoFiltro = "and ID_carteira <> " & TBAbrir!ID_carteira Else TextoFiltro = TextoFiltro & " and ID_carteira <> " & TBAbrir!ID_carteira
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
        
        TextoFiltro1 = ""
    ElseIf Requisicao_materiais = False Then
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Movimentar_estoque_pc from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Movimentar_estoque_pc = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False And EstoquerealPC > 0 Then
                With txtquantretirado
                    .Locked = True
                    .TabStop = False
                End With
'                With txtquantretirado_PC
'                    .Locked = False
'                    .TabStop = True
'                End With
            Else
                With txtquantretirado
                    .Locked = False
                    .TabStop = True
                End With
'                With txtquantretirado_PC
'                    .Locked = True
'                    .TabStop = False
'                End With
            End If
            
            TextoFiltro = ""
            TextoFiltro1 = "and PNFC.Ordem <> " & txt_RM & ""
    End If
    
    'Verifica empenho no estoque
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Sum(Qtde_empenhada - ISNULL(Qtde_saida, 0)) as Valor from Estoque_Controle_Empenho_Vendas where ID_estoque = " & TBEstoque!IDEstoque & " " & TextoFiltro & " and Qtde_empenhada - ISNULL(Qtde_saida, 0) > 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        QTEmpenhado = IIf(IsNull(TBFI!valor), 0, TBFI!valor)
        ''EstoquerealPC = IIf(IsNull(TBEstoque!estoque_real_PC), 0, TBEstoque!estoque_real_PC) - IIf(IsNull(TBFI!valor), 0, TBFI!valor)
    End If

    'Verifica empenho na produção
    Set TBFI = CreateObject("adodb.recordset")
    StrSql = "Select Sum(PNFC.Quantidade - PNFC.Qtde_saida) as Valor, Sum(ISNULL(PNFC.Quantidade_PC, 0) - ISNULL(PNFC.Qtde_saida_PC, 0)) as Valor1 from (Producao_NF_Consignada PNFC INNER JOIN Producaomaterial PM ON PM.Ordem = PNFC.Ordem and PM.Codigo = PNFC.Codinterno) INNER JOIN Producao P ON P.Ordem = PNFC.Ordem where PNFC.IDestoque = " & TBEstoque!IDEstoque & " and PNFC.Quantidade - PNFC.Qtde_saida > 0 and P.Status <> 'Cancelada'  and (PM.Saida = 'NÃO' OR PM.Saida = 'PARCIAL')"
    'Debug.print StrSql
    
    TBFI.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
       QTEmpenhado = QTEmpenhado + IIf(IsNull(TBFI!valor), 0, TBFI!valor)
        EstoquerealPC = EstoquerealPC - IIf(IsNull(TBFI!Valor1), 0, TBFI!Valor1)
    End If
    TBFI.Close
    
    txtestoqueatual_PC = Format(EstoquerealPC, "###,##0.0000")
    txtEmpenhos.Text = Format(QTEmpenhado, "###,##0.0000")
    txtestoqueatual = Format(Estoquereal, "###,##0.0000")
    'If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then txtestoqueatual_PC = Format(EstoquerealPC, "###,##0.0000") Else txtestoqueatual_PC = "0,0000"
End If
TBEstoque.Close

qtdeliberada = IIf(txtestoqueatual <> "", txtestoqueatual, 0)
QTBaixar = IIf(txtRetirar <> "", txtRetirar, 0)

If qtdeliberada < QTBaixar Then
txtquantretirado = txtestoqueatual
Else
txtquantretirado = txtRetirar
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_tipoNF_Click()
On Error GoTo tratar_erro

ProcLimpaCampos
Listamaterial.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbDestino_Click()
On Error GoTo tratar_erro

If cmbDestino <> "Terceiros (Remessa forn.)" Then
    txtIDPedido = ""
    txtFornecedor = ""
    txtPedidoCompra = ""
    cmdLocalizaPedido.Enabled = False
Else
    cmdLocalizaPedido.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbDestino_RE_Click()
On Error GoTo tratar_erro

If cmbDestino_RE <> "Terceiros (Remessa forn.)" Then
    txtIdPedido_RE = ""
    txtFornecedor_RE = ""
    txtPedido_RE = ""
    cmdPedido_RE.Enabled = False
Else
    cmdPedido_RE.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbEmpresa_RE_Click()
On Error GoTo tratar_erro

ProcLimpaCamposRE
ListaMaterial_RE.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If txtCodigo = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select imagem from projproduto where desenho = '" & txtCodigo & "' and imagem is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If TBProduto!imagem <> "" Then ProcAbrirArquivo TBProduto!imagem
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdArquivo_RE_Click()
On Error GoTo tratar_erro

If txtCodInterno_RE = "" Then Exit Sub
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select imagem from projproduto where desenho = '" & txtCodInterno_RE & "' and imagem is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    If TBProduto!imagem <> "" Then ProcAbrirArquivo TBProduto!imagem
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdListaRetirado_Click()
On Error GoTo tratar_erro

If txtBaixado.Text = "" Or txtBaixado.Text = "0,0000" Then
    MsgBox ("Nenhuma baixa realizada para esse item.")
Else
    frmEstoque_Retirar_Lista.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPedido_RE_Click()
On Error GoTo tratar_erro

Sit_REG = 2
If cmbDestino_RE = "Terceiros (Remessa forn.)" Then frmEstoque_retirar_Pedido.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdRE_Click()
On Error GoTo tratar_erro

Acao = "filtrar"
If Cmb_empresa = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa.SetFocus
    Exit Sub
End If
If txtRE = "" Then
    NomeCampo = "a número da RE"
    ProcVerificaAcao
    txtRE.SetFocus
    Exit Sub
End If
ProcLimpaCampos_TabRE
ListaMaterial_RE.ListItems.Clear
txtData_RE = Date

Qtde = 0
Qtd = 0
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select EP.*, EL.Estoque from Estoque_produtos EP INNER JOIN Estoque_Localarmazenamento_criar EL ON EL.descricao = EP.local_armaz where IDestoque = " & txtRE & " and Liberado = 'SIM'", Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    txtCodInterno_RE = IIf(IsNull(TBEstoque!Desenho), "", TBEstoque!Desenho)
    txtLote_RE = IIf(IsNull(TBEstoque!LOTE), "", TBEstoque!LOTE)
    Txt_cod_ref_RE_RE = IIf(IsNull(TBEstoque!Ref), "", TBEstoque!Ref)
    txtLocal_RE = IIf(IsNull(TBEstoque!local_armaz), "", TBEstoque!local_armaz)
    txtCorrida_RE = IIf(IsNull(TBEstoque!Corrida), "", TBEstoque!Corrida)
    txtCertificado_RE = IIf(IsNull(TBEstoque!Certificado), "", TBEstoque!Certificado)
    txtEstoque_Real_RE = Format(TBEstoque!estoque_real, "###,##0.0000")
    txtEstoque_Real_RE_PC = IIf(IsNull(TBEstoque!estoque_real_PC), 0, Format(TBEstoque!estoque_real_PC, "###,##0.0000"))
    'Verifica se no cadastro da empresa esta marcado a opção para movimentar estoque por pç
    If FunVerifMovimentacaoEstPC(cmbEmpresa_RE.ItemData(cmbEmpresa_RE.ListIndex)) = True And TBEstoque!estoque_real_PC > 0 Then
        txtDisponivel_PC_RE = IIf(IsNull(TBEstoque!estoque_real_PC), 0, Format(TBEstoque!estoque_real_PC, "###,##0.0000"))
        With txtQtde_Saida_RE
            .Locked = True
            .TabStop = False
        End With
        With txtQtde_Saida_RE_PC
            .Locked = False
            .TabStop = True
        End With
    Else
        txtDisponivel_PC_RE = "0,0000"
        With txtQtde_Saida_RE
            .Locked = False
            .TabStop = True
        End With
        With txtQtde_Saida_RE_PC
            .Locked = True
            .TabStop = False
        End With
    End If
    
    'Verifica empenho no estoque
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Sum(Qtde_requisitar) as Valor from Qtde_empenhada_produto_venda_detalhado where ID_estoque = " & TBEstoque!IDEstoque & " and Qtde_requisitar > 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Qtde = IIf(IsNull(TBFI!valor), 0, TBFI!valor)
        Qtd = IIf(IsNull(TBFI!valor), 0, TBFI!valor)
    End If
    
    'Verifica empenho na produção
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Sum(Qtde_empenhar) as Valor, Sum(Qtde_empenhar_PC) as Valor1 from Qtde_empenhada_produto_detalhado where IDestoque = " & TBEstoque!IDEstoque & " and (Qtde_empenhar > 0 or Qtde_empenhar_PC > 0)", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Qtde = Qtde + IIf(IsNull(TBFI!valor), 0, TBFI!valor)
        Qtd = Qtd + IIf(IsNull(TBFI!Valor1), 0, TBFI!Valor1)
    End If
    TBFI.Close
    
    txtEmpenhos = Format(Qtde, "###,##0.0000")
    txtEmpenho_PC = Format(Qtd, "###,##0.0000")
    
    txtDisponivel_RE = Format(IIf(IsNull(TBEstoque!estoque_real), 0, TBEstoque!estoque_real) - Qtde, "###,##0.0000")
    txtDisponivel_PC_RE = Format(IIf(IsNull(TBEstoque!estoque_real_PC), 0, TBEstoque!estoque_real_PC) - Qtd, "###,##0.0000")
    
    ProcCarregaListaRM
Else
    USMsgBox ("Não foi encontrado nenhuma RE liberado com este número."), vbExclamation, "CAPRIND v5.0"
End If
TBEstoque.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub




Private Sub Form_Resize()
'btnRequisicao_Click

End Sub

Private Sub Listamaterial_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Listamaterial
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If .ListItems.Item(InitFor).ListSubItems(6) = "RETIRADO" Or .ListItems.Item(InitFor).ListSubItems(6) = "CANCEL." Then
                If .ListItems.Item(InitFor).ListSubItems(6) = "RETIRADO" Then TextoMsg = "já foi baixado" Else TextoMsg = "esta cancelado"
                If Cmb_opcao_lista = "Retirar selecionados" Then
                    If .ListItems.Item(InitFor).ListSubItems(6) = "RETIRADO" Or .ListItems.Item(InitFor).ListSubItems(6) = "CANCEL." Then
                        USMsgBox ("Não é permitido baixar este produto, pois o mesmo " & TextoMsg & "."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                    End If
                Else
                    If .ListItems.Item(InitFor).ListSubItems(6) = "RETIRADO" Then
                        USMsgBox ("Não é permitido alterar o status deste produto, pois o mesmo " & TextoMsg & "."), vbExclamation, "CAPRIND v5.0"
                        .ListItems.Item(InitFor).Checked = False
                    End If
                End If
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaMaterial_RE_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaMaterial_RE
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaMaterial_RE, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaMaterial_RE_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaMaterial_RE
    If .ListItems.Count = 0 Then Exit Sub
    txtQtde_RE = .SelectedItem.ListSubItems(2)
    txtQtde_PC_RE = .SelectedItem.ListSubItems(3)
    txtQtde_Saida_RE = txtQtde_RE
    ProcCalculaEstoqueAtualizadoRE
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        Expedir = True
        txt_RM = ""
        ProcLimpaCampos
        Listamaterial.ListItems.Clear
        
        txt_RM.Visible = True
        Frame1.Visible = True
    Case 1:
        txtRE = ""
        ProcLimpaCampos_TabRE
        ListaMaterial_RE.ListItems.Clear
        
        txt_RM.Visible = False
        Frame1.Visible = False
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_cracha_Change()
On Error GoTo tratar_erro

txtRequisitante = ProcVerifUsuario(Txt_cracha)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function ProcVerifUsuario(Cracha As String) As String
On Error GoTo tratar_erro

ProcVerifUsuario = ""
If Cracha = "" Then Exit Function
Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select Usuario from usuarios where Codigo = '" & Cracha & "' and Bloqueado = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = False Then
    ProcVerifUsuario = TBUsuarios!Usuario
End If
TBUsuarios.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub Txt_cracha_LostFocus()
On Error GoTo tratar_erro

Txt_cracha = FunTamanhoTextoZeroEsq(Txt_cracha, 8)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_cracha_RE_Change()
On Error GoTo tratar_erro

txtRequisitante_RE = ProcVerifUsuario(Txt_cracha_RE)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_cracha_RE_LostFocus()
On Error GoTo tratar_erro

Txt_cracha_RE = FunTamanhoTextoZeroEsq(Txt_cracha_RE, 8)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_lote_Change()
On Error GoTo tratar_erro
If IsNumeric(Txt_lote.Text) Then
Ordem = Txt_lote.Text

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select individual from producao where ordem = " & Ordem, Conexao, adOpenKeyset, adLockOptimistic
    If TBproducao.EOF = False Then
        FrRastreabilidade.Enabled = True
        chkindividual.Value = Checked
    Else
        FrRastreabilidade.Enabled = False
        chkindividual.Value = Unchecked
    End If
        TBproducao.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_RM_Change()
On Error GoTo tratar_erro

'If Txt_RM <> "" Then
'    VerifNumero = Txt_RM
'    ProcVerificaNumero
'    If VerifNumero = False Then
'        Txt_RM = ""
'        Txt_RM.SetFocus
'        Exit Sub
'    End If
'End If
ProcLimpaCampos
Listamaterial.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRetirarSelecionados()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
Permitido1 = False
Permitido2 = False
Desenho = ""
With Listamaterial
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And (.ListItems.Item(InitFor).SubItems(8) = "REQUISITADO" Or .ListItems.Item(InitFor).SubItems(8) = "PARCIAL") Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente baixar este(s) material(ais) do estoque?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
                Acao = "baixar este(s) material(ais) do estoque"
                If txtRequisitante = "" Then
                    NomeCampo = "o requisitante"
                    ProcVerificaAcao
                    Txt_cracha.SetFocus
                    Exit Sub
                End If
                If cmbDestino = "" Then
                    NomeCampo = "o destino"
                    ProcVerificaAcao
                    cmbDestino.SetFocus
                    Exit Sub
                End If
                If cmbDestino.Text = "Terceiros (Remessa forn.)" Then
                    If txtPedidoCompra = "" Then
                        NomeCampo = "o pedido de compra"
                        ProcVerificaAcao
                        cmdLocalizaPedido_Click
                        Exit Sub
                    End If
                End If
                Dataini = txtData
                If Dataini > Date Then
                    USMsgBox ("A data de saída não pode ser maior que a data de hoje."), vbExclamation, "CAPRIND v5.0"
                    'txtdata = Date
                    Exit Sub
                End If
            End If
            
            Modificado = False
            Permitido = True
'======================================================================
' Verifica se é RM ou nota fiscal
'======================================================================
If Expedir = True Or IsNumeric(txt_RM) = False Then
  'frmFaturamento_Estoque.Show 1
  'Exit Sub
  GoTo BaixarRM
End If
'======================================================================
' Verifica se tem empenho da producao
'======================================================================
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select * from Producao_NF_Consignada where Ordem = " & txt_RM & " and Codinterno = '" & .ListItems.Item(InitFor).SubItems(4) & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
 Do While TBproducao.EOF = False
     Set TBEstoque = CreateObject("adodb.recordset")
     TBEstoque.Open "Select * from estoque_controle where IDestoque = " & TBproducao!IDEstoque & " and Estoque_real > 0", Conexao, adOpenKeyset, adLockOptimistic
     If TBEstoque.EOF = False Then
         Permitido2 = True
         ProcRetirarSelecionados1 True
         Modificado = True
     End If
     TBEstoque.Close
     TBproducao.MoveNext
 Loop
End If
            
'=======================================================
' Baixa por requisicao de materiais
'=======================================================
BaixarRM:

Set TBEstoque = CreateObject("adodb.recordset")
StrSql = "Select EC.* from (Estoque_Controle_Empenho_Vendas EE INNER JOIN estoque_controle EC ON EE.ID_estoque = EC.IDEstoque) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = EE.ID_carteira and NFPP.Codinterno = EC.Desenho where NFPP.ID_prod_NF = " & Listamaterial.SelectedItem & " and EC.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and EC.desenho = '" & .ListItems.Item(InitFor).SubItems(6) & "' and EC.Lote is not null and (Left(EC.status, 7) = 'ENTRADA' or EC.status = 'CONSIGNAÇÃO RECEBIDA') and EE.Qtde_empenhada - EE.Qtde_saida > 0"
'Debug.print StrSql

TBEstoque.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    Permitido2 = True
    ProcRetirarSelecionados1 False
    Modificado = True
End If
TBEstoque.Close
'=====================================================================
' Baixar item da requisiscao da ordem de producao
'=====================================================================
TextoFiltro1 = "Consignacao = 'False'"

If IsNumeric(txt_RM) = True Then
  Set TBFI = CreateObject("adodb.recordset")
  If Expedir = True Then
      TBFI.Open "Select IDCliente, Cliente from Producao where ordem = " & txt_RM & " and consignacao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
  Else
      TBFI.Open "Select NF.Id_Int_Cliente AS IDCliente, CFOP.Devolucao, NF.txt_Razao_Nome AS Cliente from (tbl_Detalhes_Nota NFP INNER JOIN tbl_Dados_Nota_Fiscal NF ON NF.ID = NFP.ID_nota) INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.Int_codigo = " & Listamaterial.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
  End If
  If TBFI.EOF = False Then
      TextoFiltro = " and Liberado = 'SIM'"
      If Expedir = True Then
          If TBFI!Devolucao = True Then TextoFiltro = ""
      End If
      TextoFiltro1 = "(Consignacao = 'True' and id_cliente = " & TBFI!IDCliente & " and Cliente = '" & TBFI!Cliente & "' or Consignacao = 'False')"
  End If
  TBFI.Close
End If
'=======================================================================

Set TBFI = CreateObject("adodb.recordset")
StrSql = "Select QPP.Ordem from (Qtde_produzindo_produto QPP INNER JOIN Producao_pedidos PP ON PP.Ordem = QPP.Ordem) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = PP.IDcarteira where NFPP.ID_prod_NF = " & .ListItems.Item(InitFor)
'Debug.print StrSql

TBFI.Open "Select QPP.Ordem from (Qtde_produzindo_produto QPP INNER JOIN Producao_pedidos PP ON PP.Ordem = QPP.Ordem) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = PP.IDcarteira where NFPP.ID_prod_NF = " & .ListItems.Item(InitFor) & "", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = True Then
    Set TBEstoque = CreateObject("adodb.recordset")
    StrSql = "Select * from Estoque_produtos where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Desenho = '" & .ListItems.Item(InitFor).SubItems(6) & "' and Estoque_real > 0 " & TextoFiltro & " and " & TextoFiltro1 & " order by Data, IdEstoque"
    'Debug.print StrSql
    
    TBEstoque.Open "Select * from Estoque_produtos where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Desenho = '" & .ListItems.Item(InitFor).SubItems(6) & "' and Estoque_real > 0 " & TextoFiltro & " and " & TextoFiltro1 & " order by Data, IdEstoque", Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
        Permitido2 = False
        ProcRetirarSelecionados1 False
        Modificado = True
    End If
End If

If Expedir = True And IsNumeric(txt_RM) = True Then TBproducao.Close
End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) material(ais) antes de dar baixa no estoque."), vbExclamation, "CAPRIND v5.0"
Else
    If Permitido1 = True Then
        USMsgBox ("Material(ais) retirado(s) do estoque com sucesso."), vbInformation, "CAPRIND v5.0"
        cmdlote_Click
    Else
        USMsgBox ("O(s) material(ais) não foi(ram) retirado(s) do estoque."), vbExclamation, "CAPRIND v5.0"
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcRetirarSelecionados1(BaixarEmpenhado As Boolean)
On Error GoTo tratar_erro

With Listamaterial
    'Só verifica quando for a primeira vez do produto
    If Modificado = False Then
        If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True And IsNull(TBEstoque!estoque_real_PC) = False And TBEstoque!estoque_real_PC > 0 And .ListItems.Item(InitFor).SubItems(2) <> "" And .ListItems.Item(InitFor).SubItems(2) <> "0,0000" Then
            qtdeliberada = .ListItems.Item(InitFor).SubItems(2)
            CampoFiltro = "Saida_PC"
        Else
            qtdeliberada = .ListItems.Item(InitFor).SubItems(3)
            CampoFiltro = "Saida"
        End If
    End If
    qtdeliberar = qtdeliberada
    If Expedir = True And IsNumeric(txt_RM) = True Then
        If BaixarEmpenhado = True Then
            If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True And IsNull(TBproducao!Quantidade_PC) = False And TBproducao!Quantidade_PC > 0 Then qtdeliberada = TBproducao!Quantidade_PC Else qtdeliberada = TBproducao!quantidade
        End If
    End If
    Qtd = 0
    Do While TBEstoque.EOF = False And qtdeliberada > 0
        valor = 0
        'Verifica se este RE já está empenhado
        If Expedir = True Then
            TextoFiltro = ""
            TextoFiltro1 = ""
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select ID_carteira from tbl_Detalhes_Nota_pedidos where ID_prod_NF = " & .ListItems.Item(InitFor) & " and Codinterno = '" & TBEstoque!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = False Then
                Do While TBFIltro.EOF = False
                    If TextoFiltro = "" Then TextoFiltro = " and ID_carteira <> " & TBFIltro!ID_carteira Else TextoFiltro = TextoFiltro & " and ID_carteira <> " & TBFIltro!ID_carteira
                    TBFIltro.MoveNext
                Loop
            End If
            TBFIltro.Close
        ElseIf Requisicao_materiais = False Then
                TextoFiltro = ""
                TextoFiltro1 = " and Ordem <> " & txt_RM
            Else
                TextoFiltro = ""
                TextoFiltro1 = ""
        End If
                
        If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True And IsNull(TBEstoque!estoque_real_PC) = False And TBEstoque!estoque_real_PC > 0 Then
            Qtd = TBEstoque!estoque_real_PC
            Qtd = Qtd - FunVerificaQtdeEmpenhoREOrdem("IDestoque = " & TBEstoque!IDEstoque & TextoFiltro1, True)
        Else
            Qtd = TBEstoque!estoque_real
            Qtd = Qtd - FunVerificaQtdeEmpenhoREOrdem("IDestoque = " & TBEstoque!IDEstoque & TextoFiltro1, False)
        End If
        Qtd = Qtd - FunVerificaQtdeEmpenhoREPI("ID_estoque = " & TBEstoque!IDEstoque & TextoFiltro)
        If Qtd <= 0 Then GoTo Proximo
        
        If Permitido2 = False Then
            If qtdeliberada > Qtd Then
               ' If USMsgBox("Deseja utilizar mais que um lote para o material " & TBEstoque!Desenho & "?", vbyesno, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
        End If
        Permitido2 = True
        
        If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True And IsNull(TBEstoque!estoque_real_PC) = False And TBEstoque!estoque_real_PC > 0 Then
            If qtdeliberada >= Qtd Then QtdeSaidaPC = Qtd Else QtdeSaidaPC = qtdeliberada
            QtdeSaida = FunCalculaQtdePCKG(TBEstoque!estoque_real, TBEstoque!estoque_real_PC, QtdeSaidaPC, False)
        Else
            If qtdeliberada >= Qtd Then QtdeSaida = Qtd Else QtdeSaida = qtdeliberada
            QtdeSaidaPC = FunCalculaQtdePCKG(TBEstoque!estoque_real, IIf(IsNull(TBEstoque!estoque_real_PC), 0, TBEstoque!estoque_real_PC), QtdeSaida, True)
        End If
        
        Set TBMateriaprima = CreateObject("adodb.recordset")
        If Expedir = True Then
            Set TBMateriaprima = CreateObject("adodb.recordset")
            TBMateriaprima.Open "Select * from tbl_Detalhes_Nota_pedidos where ID_prod_NF = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBMateriaprima.EOF = False Then
            
                SaldoPorc = 0
                Set TBCotacao = CreateObject("adodb.recordset")
                TBCotacao.Open "Select txt_Unid, Unidade_com, int_Cod_Produto, Qtde_estoque from tbl_Detalhes_Nota where Int_codigo = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBCotacao.EOF = False Then
                    If FunVerifUNConversao(TBCotacao!txt_Unid, TBCotacao!Unidade_com) = True Then
                        'Unidade com e un estoque foram invertidas para reverter a conversão
                        SaldoPorc = FunConverteUN(TBCotacao!Unidade_com, TBCotacao!txt_Unid, QtdeSaida, TBCotacao!int_Cod_Produto)
                    End If
                End If
                TBCotacao.Close
            
                ProcAtualizaQtdeExpProdPed TBMateriaprima!ID_prod_NF, TBMateriaprima!Codinterno, IIf(SaldoPorc = 0, QtdeSaida, SaldoPorc), TBEstoque!LOTE, TBEstoque!IDEstoque, txtData
            End If
            If qtdeliberada >= Qtd Then QtdeSaida = Qtd Else QtdeSaida = qtdeliberada
            
            If FunVerifCodRefCliDANFE(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then
                Set TBMateriaprima = CreateObject("adodb.recordset")
                TBMateriaprima.Open "Select N_Referencia from tbl_Detalhes_Nota where Int_codigo = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBMateriaprima.EOF = False Then
                    Conexao.Execute "UPDATE Estoque_controle Set REF = '" & TBMateriaprima!N_referencia & "' where IDestoque = " & TBEstoque!IDEstoque
                End If
                TBMateriaprima.Close
            End If
            
        'Arruma status materia prima
        ElseIf Requisicao_materiais = False Then
                TBMateriaprima.Open "Select * from Producaomaterial where Idmateriaprima = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBMateriaprima.EOF = False Then
                    If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True And IsNull(TBEstoque!estoque_real_PC) = False And TBEstoque!estoque_real_PC > 0 Then
                        If QtdeSaidaPC >= qtdeliberar Then TBMateriaprima!Saida = "SIM" Else TBMateriaprima!Saida = "PARCIAL"
                    Else
                        If QtdeSaida >= qtdeliberar Then TBMateriaprima!Saida = "SIM" Else TBMateriaprima!Saida = "PARCIAL"
                    End If
                    TBMateriaprima!Valor_saida_estoque = TBMateriaprima!Valor_saida_estoque + Format(QtdeSaida * TBEstoque!valor_unitario, "###,##0.00")
                    TBMateriaprima.Update
                    
                    'Atualiza qtde. de saída do empenho da ordem
                    QuantEmpenho = 0
                    QuantEmpenhoPC = 0
                    Set TBFIltro = CreateObject("adodb.recordset")
                    TBFIltro.Open "Select Sum(Saida) as QuantEmpenho, Sum(ISNULL(Saida_PC, 0)) as QuantEmpenhoPC from estoque_movimentacao where IDestoque = " & TBEstoque!IDEstoque & " and oe = '" & txt_RM & "' and desenho = '" & TBEstoque!Desenho & "' and documento = '" & txt_RM & "' and (operacao = 'SAIDA_ORDEM' or operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFIltro.EOF = False Then
                        QuantEmpenho = IIf(IsNull(TBFIltro!QuantEmpenho), 0, Format(TBFIltro!QuantEmpenho, "###,##0.0000"))
                        QuantEmpenhoPC = IIf(IsNull(TBFIltro!QuantEmpenhoPC), 0, TBFIltro!QuantEmpenhoPC)
                        QuantEmpenho = QuantEmpenho + QtdeSaida
                        QuantEmpenhoPC = QuantEmpenhoPC + QtdeSaidaPC
                    End If
                    NovoValor = Replace(QuantEmpenho, ",", ".")
                    NovoValor1 = Replace(QuantEmpenhoPC, ",", ".")
                    Conexao.Execute "UPDATE Producao_NF_Consignada Set Qtde_saida = " & NovoValor & " where IDestoque = " & TBEstoque!IDEstoque & " and Ordem = " & txt_RM & " and Codinterno = '" & TBEstoque!Desenho & "'"
                    Conexao.Execute "UPDATE Producao_NF_Consignada Set Qtde_saida_PC = " & NovoValor1 & " where IDestoque = " & TBEstoque!IDEstoque & " and Ordem = " & txt_RM & " and Codinterno = '" & TBEstoque!Desenho & "' and Quantidade_PC IS NOT NULL and Quantidade_PC > 0"
                   '===================================
                   ' Atualiza tipo item na movimentação
                   '===================================
                   Conexao.Execute "UPDATE Estoque_movimentacao Set IDEstoque = " & TBEstoque!IDEstoque & " where IDoperacao = " & TBFIltro!IDoperacao
                   '===================================

                End If
                TBMateriaprima.Close
            Else
                TBMateriaprima.Open "Select * from Requisicao_materiais_lista where IDlista = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBMateriaprima.EOF = False Then
                    If Format(QtdeSaida, "###,##0.0000") >= Format(qtdeliberar, "###,##0.0000") Then TBMateriaprima!status = "RETIRADO" Else TBMateriaprima!status = "PARCIAL"
                    TBMateriaprima!quant_saida = TBMateriaprima!quant_saida + QtdeSaida
                    TBMateriaprima.Update
                    
                    If IsNull(TBMateriaprima!Ordem) = False And TBMateriaprima!Ordem <> 0 And IsNull(TBMateriaprima!ID_PC) = False And TBMateriaprima!ID_PC <> 0 Then ProcGravarOutrasDespOrdem TBMateriaprima!Ordem, TBMateriaprima!ID_PC, Format(QtdeSaida * TBEstoque!valor_unitario, "###,##0.00")
                End If
                TBMateriaprima.Close
        End If
        
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from projproduto where Desenho = '" & TBEstoque!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            If TBProduto!Estoque = True Then
                QuantEmpenho = Format(TBEstoque!estoque_real - QtdeSaida, "###,##0.0000")
                QuantEmpenhoPC = IIf(IsNull(TBEstoque!estoque_real_PC), 0, TBEstoque!estoque_real_PC) - QtdeSaidaPC
            Else
                QuantEmpenho = 0
                QuantEmpenhoPC = 0
            End If
            NovoValor = Replace(QuantEmpenho, ",", ".")
            NovoValor1 = Replace(QuantEmpenhoPC, ",", ".")
            Conexao.Execute "UPDATE Estoque_controle Set Estoque_real = " & NovoValor & ", Estoque_real_PC = " & NovoValor1 & ", Estoque_venda = " & NovoValor & ", peso_unit = '" & IIf(IsNull(TBProduto!peso_metro), "", TBProduto!peso_metro) & "', Pedido = '" & IIf(IsNull(TBProduto!Un_Kg), "", TBProduto!Un_Kg) & "' where IDestoque = " & TBEstoque!IDEstoque
        End If
        TBProduto.Close
            
        'Atualiza valor do total no estoque
        Conexao.Execute "UPDATE Estoque_controle Set Valor_Total = ROUND(valor_unitario * Estoque_real, 2) where IDestoque = " & TBEstoque!IDEstoque
                
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
        TBProduto.AddNew
        TBProduto!Destino = IIf(cmbDestino = "Interno/Cliente", "Interno", "Terceiros")
        TBProduto!Terceiros = False
        TBProduto!Documento = IIf(txt_RM.Text = "", txt_Notafiscal, txt_RM)
        TBProduto!LOTE = IIf(IsNull(TBEstoque!LOTE), "", TBEstoque!LOTE)
        TBProduto!Desenho = IIf(IsNull(TBEstoque!Desenho), "", TBEstoque!Desenho)
        TBProduto!Data = txtData
        TBProduto!Descricao = IIf(IsNull(TBEstoque!Descricao), "", TBEstoque!Descricao)
        TBProduto!Familia = IIf(IsNull(TBEstoque!Classe), "", TBEstoque!Classe)
        TBProduto!Requisitante = txtRequisitante
        TBProduto!Responsavel = pubUsuario
        TBProduto!IDEstoque = TBEstoque!IDEstoque
        TBProduto!OE = IIf(txt_RM.Text = "", txt_Notafiscal, txt_RM)
        
        If Expedir = True Then
            TBProduto!ID_prod_NF = .ListItems.Item(InitFor)
            If QtdeSaida >= qtdeliberada Then TBProduto!Operacao = "SAIDA_NOTA" Else TBProduto!Operacao = "SAIDA_NOTA_PARCIAL"
        ElseIf Requisicao_materiais = False Then
                TBProduto!Ordem = txt_RM
                If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True And IsNull(TBEstoque!estoque_real_PC) = False And TBEstoque!estoque_real_PC > 0 Then
                    If QtdeSaidaPC >= qtdeliberar Then TBProduto!Operacao = "SAIDA_ORDEM" Else TBProduto!Operacao = "SAIDA_ORDEM_PARCIAL"
                Else
                    If QtdeSaida >= qtdeliberar Then TBProduto!Operacao = "SAIDA_ORDEM" Else TBProduto!Operacao = "SAIDA_ORDEM_PARCIAL"
                End If
            Else
                TBProduto!ID_prod_RM = .ListItems.Item(InitFor)
                If QtdeSaida >= qtdeliberada Then TBProduto!Operacao = "SAIDA_REQUISICAO" Else TBProduto!Operacao = "SAIDA_REQUISICAO_PARCIAL"
        End If
        
        If qtdeliberar >= Qtd Then
            qtdeliberada = qtdeliberada - Qtd
            qtdeliberar = qtdeliberar - Qtd
        Else
            qtdeliberada = 0
            qtdeliberar = 0
        End If
        
        TBProduto!Saida = QtdeSaida
        TBProduto!Saida_PC = QtdeSaidaPC
        TBProduto!estoque_venda = QtdeSaida
    
        'Atualiza valor do material no estoque
        TBProduto!VlrUnit = IIf(IsNull(TBEstoque!valor_unitario), 0, Format(TBEstoque!valor_unitario, "###,##0.0000000000"))
        TBProduto!vlrTotal = Format(QtdeSaida * TBProduto!VlrUnit, "###,##0.00")
            
        '==================================
        Modulo = "Estoque/Movimentação/Retirada"
        Evento = "Retirar"
        ID_documento = .ListItems.Item(InitFor)
        Documento = "Cód. interno: " & TBEstoque!Desenho & " - RE: " & TBEstoque!IDEstoque
        Documento1 = ""
        ProcGravaEvento
        '==================================
        Permitido1 = True
        
        TBProduto.Update
        '===================================
        ' Atualiza tipo item na movimentação
        '===================================
        Conexao.Execute "UPDATE Estoque_movimentacao Set IDEstoque = " & TBEstoque!IDEstoque & " where IDoperacao = " & TBProduto!IDoperacao
        '===================================
        
        valor = TBProduto!vlrTotal
        
        'Centro de custo
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select Codproduto, ID_PC from projproduto where desenho = '" & TBEstoque!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            Codproduto = TBItem!Codproduto
            IDAntigo = IIf(IsNull(TBItem!ID_PC), 0, TBItem!ID_PC)
            
            ProcCriaCreditoCCProdutoItemSelecionada TBItem!Codproduto, txtData, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBProduto!IDoperacao, valor
        End If
        TBItem.Close
                
        If Requisicao_materiais = True Then
            Set TBMateriaprima = CreateObject("adodb.recordset")
            TBMateriaprima.Open "Select * from Requisicao_materiais_lista where idlista = " & .ListItems.Item(InitFor) & " and ID_CC is not null", Conexao, adOpenKeyset, adLockOptimistic
            If TBMateriaprima.EOF = False Then
                If TBMateriaprima!ID_CC <> "" Then
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                    TBFI.AddNew
                    ProcEnviaDadosCCRealizadoSel TBMateriaprima!ID_CC
                    TBFI.Update
                    
                    'Grava movimentação no centro consolidado
                    Set TBAfericao = CreateObject("adodb.recordset")
                    TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBMateriaprima!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAfericao.EOF = False Then
                        Do While TBAfericao.EOF = False
                            Set TBFI = CreateObject("adodb.recordset")
                            TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                            TBFI.AddNew
                            ProcEnviaDadosCCRealizadoSel TBAfericao!ID_CC
                            TBFI.Update
                            
                            Set TBCiclo = CreateObject("adodb.recordset")
                            TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                            If TBCiclo.EOF = False Then
                                Do While TBCiclo.EOF = False
                                    Set TBFI = CreateObject("adodb.recordset")
                                    TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                                    TBFI.AddNew
                                    ProcEnviaDadosCCRealizadoSel TBCiclo!ID_CC
                                    TBFI.Update
                                    TBCiclo.MoveNext
                                Loop
                            End If
                            TBCiclo.Close
                            
                            TBAfericao.MoveNext
                        Loop
                    End If
                    TBAfericao.Close
                End If
            End If
            TBMateriaprima.Close
        End If
        
        If Expedir = True And IsNumeric(txt_RM) = True Then ProcAtualizaCTMaterialOrdem Cmb_empresa.ItemData(Cmb_empresa.ListIndex), txt_RM
        
        TBProduto.Close
Proximo:
        TBEstoque.MoveNext
    Loop
End With

If Requisicao_materiais = True And Permitido2 = True Then ProcAtualizaStatus_RM

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRetirarSelecionadosRE()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
Permitido1 = False
Permitido2 = False
Desenho = ""
With ListaMaterial_RE
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And (.ListItems.Item(InitFor).SubItems(7) = "NÃO" Or .ListItems.Item(InitFor).SubItems(7) = "PARCIAL") Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente baixar este(s) material(ais) do estoque?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
                Acao = "baixar este(s) material(ais) do estoque"
                If txtRequisitante_RE = "" Then
                    NomeCampo = "o requisitante"
                    ProcVerificaAcao
                    Txt_cracha_RE.SetFocus
                    Exit Sub
                End If
                If cmbDestino_RE = "" Then
                    NomeCampo = "o destino"
                    ProcVerificaAcao
                    cmbDestino_RE.SetFocus
                    Exit Sub
                End If
                If cmbDestino_RE.Text = "Terceiros (Remessa forn.)" Then
                    If txtPedidoCompra = "" Then
                        NomeCampo = "o pedido de compra"
                        ProcVerificaAcao
                        cmdPedido_RE_Click
                        Exit Sub
                    End If
                End If
                Dataini = txtData_RE
                If Dataini > Date Then
                    USMsgBox ("A data de saída não pode ser maior que a data de hoje."), vbExclamation, "CAPRIND v5.0"
                    txtData_RE = Date
                    Exit Sub
                End If
            End If
            
            Modificado = False
            Permitido = True
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select * from Producao_NF_Consignada where Ordem = " & .ListItems.Item(InitFor).SubItems(1) & " and Codinterno = '" & .ListItems.Item(InitFor).SubItems(5) & "' and IDestoque = " & txtRE, Conexao, adOpenKeyset, adLockOptimistic
            If TBproducao.EOF = False Then
                Do While TBproducao.EOF = False
                    Set TBEstoque = CreateObject("adodb.recordset")
                    TBEstoque.Open "Select * from estoque_controle where IDestoque = " & TBproducao!IDEstoque & " and Estoque_real > 0", Conexao, adOpenKeyset, adLockOptimistic
                    If TBEstoque.EOF = False Then
                        Permitido2 = True
                        ProcRetirarSelecionadosRE1 True
                        Modificado = True
                    End If
                    TBEstoque.Close
                    TBproducao.MoveNext
                Loop
            Else
                Set TBEstoque = CreateObject("adodb.recordset")
                TBEstoque.Open "Select EC.* from (Estoque_Controle_Empenho_Vendas EE INNER JOIN estoque_controle EC ON EE.ID_estoque = EC.IDEstoque) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = EE.ID_carteira and NFPP.Codinterno = EC.Desenho where EC.IDestoque = " & txtRE & " and NFPP.ID_prod_NF = " & .SelectedItem & " and EC.ID_empresa = " & cmbEmpresa_RE.ItemData(cmbEmpresa_RE.ListIndex) & " and EC.desenho = '" & .ListItems.Item(InitFor).SubItems(4) & "' and EC.Lote IS NOT NULL and (Left(EC.status, 7) = 'ENTRADA' or EC.status = 'CONSIGNAÇÃO RECEBIDA') and EE.Qtde_empenhada - EE.Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
                If TBEstoque.EOF = False Then
                    Permitido2 = True
                    ProcRetirarSelecionadosRE1 False
                    Modificado = True
                Else
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select QPP.Ordem from (Qtde_produzindo_produto QPP INNER JOIN Producao_pedidos PP ON PP.Ordem = QPP.Ordem) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = PP.IDcarteira where NFPP.ID_prod_NF = " & .ListItems.Item(InitFor) & "", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = True Then
                        Set TBEstoque = CreateObject("adodb.recordset")
                        TBEstoque.Open "Select * from Estoque_produtos where IDestoque = " & txtRE, Conexao, adOpenKeyset, adLockOptimistic
                        If TBEstoque.EOF = False Then
                            Permitido2 = False
                            ProcRetirarSelecionadosRE1 False
                            Modificado = True
                        End If
                    End If
                End If
                TBEstoque.Close
            End If
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) material(ais) antes de dar baixa no estoque."), vbExclamation, "CAPRIND v5.0"
Else
    If Permitido1 = True Then
        USMsgBox ("Material(ais) retirado(s) do estoque com sucesso."), vbInformation, "CAPRIND v5.0"
        cmdRE_Click
    Else
        USMsgBox ("O(s) material(ais) não foi(ram) retirado(s) do estoque."), vbExclamation, "CAPRIND v5.0"
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcRetirarSelecionadosRE1(BaixarEmpenhado As Boolean)
On Error GoTo tratar_erro

With ListaMaterial_RE
    'Só verifica quando for a primeira vez do produto
    If Modificado = False Then
        If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True And IsNull(TBEstoque!estoque_real_PC) = False And TBEstoque!estoque_real_PC > 0 And .ListItems.Item(InitFor).SubItems(3) <> "" And .ListItems.Item(InitFor).SubItems(3) <> "0,0000" Then
            qtdeliberada = .ListItems.Item(InitFor).SubItems(3)
            CampoFiltro = "Saida_PC"
        Else
            qtdeliberada = .ListItems.Item(InitFor).SubItems(2)
            CampoFiltro = "Saida"
        End If
    End If
    
    qtdeliberar = qtdeliberada
    If BaixarEmpenhado = True Then
        If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True And IsNull(TBproducao!Quantidade_PC) = False And TBproducao!Quantidade_PC > 0 Then qtdeliberada = TBproducao!Quantidade_PC Else qtdeliberada = TBproducao!quantidade
    End If
    Qtd = 0
    Do While TBEstoque.EOF = False And qtdeliberada > 0
        valor = 0
        'Verifica se este RE já está empenhado
        If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True And IsNull(TBEstoque!estoque_real_PC) = False And TBEstoque!estoque_real_PC > 0 Then
            Qtd = TBEstoque!estoque_real_PC
            'Verifica se este RE já está empenhado
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Sum(PNFC.Quantidade_PC - ISNULL(PNFC.Qtde_saida_PC, 0)) as Valor from (Producao_NF_Consignada PNFC INNER JOIN Producaomaterial PM ON PM.Ordem = PNFC.Ordem and PM.Codigo = PNFC.Codinterno) INNER JOIN Producao P ON P.Ordem = PNFC.Ordem where PNFC.IDestoque = " & TBEstoque!IDEstoque & " and PNFC.Ordem <> " & .ListItems.Item(InitFor).SubItems(1) & " and PNFC.Quantidade_PC - ISNULL(PNFC.Qtde_saida_PC, 0) > 0 and P.Status <> 'Cancelada' and P.Concluida = 0 and (PM.Saida = 'NÃO' OR PM.Saida = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Qtd = Qtd - IIf(IsNull(TBFI!valor), 0, TBFI!valor)
            End If
        Else
            Qtd = TBEstoque!estoque_real
            'Verifica se este RE já está empenhado
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Sum(PNFC.Quantidade - ISNULL(PNFC.Qtde_saida, 0)) as Valor from (Producao_NF_Consignada PNFC INNER JOIN Producaomaterial PM ON PM.Ordem = PNFC.Ordem and PM.Codigo = PNFC.Codinterno) INNER JOIN Producao P ON P.Ordem = PNFC.Ordem where PNFC.IDestoque = " & TBEstoque!IDEstoque & " and PNFC.Ordem <> " & .ListItems.Item(InitFor).SubItems(1) & " and PNFC.Quantidade - ISNULL(PNFC.Qtde_saida, 0) > 0 and P.Status <> 'Cancelada' and P.Concluida = 0 and (PM.Saida = 'NÃO' OR PM.Saida = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Qtd = Qtd - IIf(IsNull(TBFI!valor), 0, TBFI!valor)
            End If
        End If
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select Sum(Qtde_empenhada - ISNULL(Qtde_saida, 0)) as Valor from Estoque_Controle_Empenho_Vendas where ID_estoque = " & TBEstoque!IDEstoque & " and Qtde_empenhada - ISNULL(Qtde_saida, 0) > 0", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            Qtd = Qtd - IIf(IsNull(TBFI!valor), 0, TBFI!valor)
        End If
        TBFI.Close
        
        If Qtd <= 0 Then GoTo Proximo
        
        If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True And IsNull(TBEstoque!estoque_real_PC) = False And TBEstoque!estoque_real_PC > 0 Then
            If qtdeliberada >= Qtd Then QtdeSaidaPC = Qtd Else QtdeSaidaPC = qtdeliberada
            QtdeSaida = FunCalculaQtdePCKG(TBEstoque!estoque_real, TBEstoque!estoque_real_PC, QtdeSaidaPC, False)
        Else
            If qtdeliberada >= Qtd Then QtdeSaida = Qtd Else QtdeSaida = qtdeliberada
            QtdeSaidaPC = FunCalculaQtdePCKG(TBEstoque!estoque_real, IIf(IsNull(TBEstoque!estoque_real_PC), 0, TBEstoque!estoque_real_PC), QtdeSaida, True)
        End If
        
        'Arruma status materia prima
        Set TBMateriaprima = CreateObject("adodb.recordset")
        TBMateriaprima.Open "Select * from Producaomaterial where Idmateriaprima = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
        If TBMateriaprima.EOF = False Then
            If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True And IsNull(TBEstoque!estoque_real_PC) = False And TBEstoque!estoque_real_PC > 0 Then
                If QtdeSaidaPC >= qtdeliberar Then TBMateriaprima!Saida = "SIM" Else TBMateriaprima!Saida = "PARCIAL"
            Else
                If QtdeSaida >= qtdeliberar Then TBMateriaprima!Saida = "SIM" Else TBMateriaprima!Saida = "PARCIAL"
            End If
            TBMateriaprima!Valor_saida_estoque = TBMateriaprima!Valor_saida_estoque + Format(QtdeSaida * TBEstoque!valor_unitario, "###,##0.00")
            TBMateriaprima.Update
            
            'Atualiza qtde. de saída do empenho da ordem
            QuantEmpenho = 0
            QuantEmpenhoPC = 0
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select Sum(Saida) as QuantEmpenho, Sum(ISNULL(Saida_PC, 0)) as QuantEmpenhoPC from estoque_movimentacao where IDestoque = " & TBEstoque!IDEstoque & " and oe = '" & .ListItems.Item(InitFor).SubItems(1) & "' and desenho = '" & TBEstoque!Desenho & "' and documento = '" & .ListItems.Item(InitFor).SubItems(2) & "' and (operacao = 'SAIDA_ORDEM' or operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = False Then
                QuantEmpenho = IIf(IsNull(TBFIltro!QuantEmpenho), 0, Format(TBFIltro!QuantEmpenho, "###,##0.0000"))
                QuantEmpenhoPC = IIf(IsNull(TBFIltro!QuantEmpenhoPC), 0, TBFIltro!QuantEmpenhoPC)
                QuantEmpenho = QuantEmpenho + QtdeSaida
                QuantEmpenhoPC = QuantEmpenhoPC + QtdeSaidaPC
            End If
            NovoValor = Replace(QuantEmpenho, ",", ".")
            NovoValor1 = Replace(QuantEmpenhoPC, ",", ".")
            Conexao.Execute "UPDATE Producao_NF_Consignada Set Qtde_saida = " & NovoValor & " where IDestoque = " & TBEstoque!IDEstoque & " and Ordem = " & .ListItems.Item(InitFor).SubItems(1) & " and Codinterno = '" & TBEstoque!Desenho & "'"
            Conexao.Execute "UPDATE Producao_NF_Consignada Set Qtde_saida_PC = " & NovoValor1 & " where IDestoque = " & TBEstoque!IDEstoque & " and Ordem = " & .ListItems.Item(InitFor).SubItems(1) & " and Codinterno = '" & TBEstoque!Desenho & "' and Quantidade_PC IS NOT NULL and Quantidade_PC > 0"
            '===================================
            ' Atualiza tipo item na movimentação
            '===================================
            Conexao.Execute "UPDATE Estoque_movimentacao Set IDEstoque = " & TBEstoque!IDEstoque & " where IDoperacao = " & TBFIltro!IDoperacao
            '===================================
       
        End If
        TBMateriaprima.Close
        
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from projproduto where Desenho = '" & TBEstoque!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = False Then
            If TBProduto!Estoque = True Then
                QuantEmpenho = Format(TBEstoque!estoque_real - QtdeSaida, "###,##0.0000")
                QuantEmpenhoPC = IIf(IsNull(TBEstoque!estoque_real_PC), 0, TBEstoque!estoque_real_PC) - QtdeSaidaPC
            Else
                QuantEmpenho = 0
                QuantEmpenhoPC = 0
            End If
            NovoValor = Replace(QuantEmpenho, ",", ".")
            NovoValor1 = Replace(QuantEmpenhoPC, ",", ".")
            Conexao.Execute "UPDATE Estoque_controle Set Estoque_real = " & NovoValor & ", Estoque_real_PC = " & NovoValor1 & ", Estoque_venda = " & NovoValor & ", peso_unit = '" & IIf(IsNull(TBProduto!peso_metro), "", TBProduto!peso_metro) & "', Pedido = '" & IIf(IsNull(TBProduto!Un_Kg), "", TBProduto!Un_Kg) & "' where IDestoque = " & TBEstoque!IDEstoque
        End If
        TBProduto.Close
            
        'Atualiza valor do total no estoque
        Conexao.Execute "UPDATE Estoque_controle Set Valor_Total = ROUND(valor_unitario * Estoque_real, 2) where IDestoque = " & TBEstoque!IDEstoque
                
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
        TBProduto.AddNew
        TBProduto!Destino = IIf(cmbDestino_RE = "Interno/Cliente", "Interno", "Terceiros")
        TBProduto!Terceiros = False
        TBProduto!Documento = .ListItems.Item(InitFor).SubItems(1)
        TBProduto!LOTE = IIf(IsNull(TBEstoque!LOTE), "", TBEstoque!LOTE)
        TBProduto!Desenho = IIf(IsNull(TBEstoque!Desenho), "", TBEstoque!Desenho)
        TBProduto!Data = Date
        TBProduto!Descricao = IIf(IsNull(TBEstoque!Descricao), "", TBEstoque!Descricao)
        TBProduto!Familia = IIf(IsNull(TBEstoque!Classe), "", TBEstoque!Classe)
        TBProduto!Requisitante = txtRequisitante_RE
        TBProduto!Responsavel = pubUsuario
        TBProduto!IDEstoque = TBEstoque!IDEstoque
        TBProduto!OE = .ListItems.Item(InitFor).SubItems(1)
        
        TBProduto!Ordem = .ListItems.Item(InitFor).SubItems(1)
        If FunVerifMovimentacaoEstPC(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True And IsNull(TBEstoque!estoque_real_PC) = False And TBEstoque!estoque_real_PC > 0 Then
            If QtdeSaidaPC >= qtdeliberar Then TBProduto!Operacao = "SAIDA_ORDEM" Else TBProduto!Operacao = "SAIDA_ORDEM_PARCIAL"
        Else
            If QtdeSaida >= qtdeliberar Then TBProduto!Operacao = "SAIDA_ORDEM" Else TBProduto!Operacao = "SAIDA_ORDEM_PARCIAL"
        End If
        
        If qtdeliberar >= Qtd Then
            qtdeliberada = qtdeliberada - Qtd
            qtdeliberar = qtdeliberar - Qtd
        Else
            qtdeliberada = 0
            qtdeliberar = 0
        End If
        
        TBProduto!Saida = QtdeSaida
        TBProduto!Saida_PC = QtdeSaidaPC
        TBProduto!estoque_venda = QtdeSaida
    
        'Atualiza valor do material no estoque
        TBProduto!VlrUnit = IIf(IsNull(TBEstoque!valor_unitario), 0, Format(TBEstoque!valor_unitario, "###,##0.0000000000"))
        TBProduto!vlrTotal = Format(QtdeSaida * TBProduto!VlrUnit, "###,##0.00")
            
        '==================================
        Modulo = "Estoque/Movimentação/Retirada"
        Evento = "Retirar"
        ID_documento = .ListItems.Item(InitFor)
        Documento = "Cód. interno: " & TBEstoque!Desenho & " - RE: " & TBEstoque!IDEstoque
        Documento1 = ""
        ProcGravaEvento
        '==================================
        Permitido1 = True
        
        TBProduto.Update
        '===================================
        ' Atualiza tipo item na movimentação
        '===================================
        Conexao.Execute "UPDATE Estoque_movimentacao Set IDEstoque = " & TBEstoque!IDEstoque & " where IDoperacao = " & TBProduto!IDoperacao
        '===================================
        
        valor = TBProduto!vlrTotal
        
        'Centro de custo
        Set TBItem = CreateObject("adodb.recordset")
        TBItem.Open "Select * from projproduto where desenho = '" & TBEstoque!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBItem.EOF = False Then
            Codproduto = TBItem!Codproduto
            IDAntigo = IIf(IsNull(TBItem!ID_PC), 0, TBItem!ID_PC)
            
            ProcCriaCreditoCCProdutoItemSelecionada TBItem!Codproduto, txtData, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), TBProduto!IDoperacao, valor
        End If
        TBItem.Close
        
        If Requisicao_materiais = True Then
            Set TBMateriaprima = CreateObject("adodb.recordset")
            TBMateriaprima.Open "Select * from Requisicao_materiais_lista where idlista = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBMateriaprima.EOF = False Then
                If IsNull(TBMateriaprima!ID_CC) = False And TBMateriaprima!ID_CC <> "" And TBMateriaprima!ID_CC <> "0" Then
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                    TBFI.AddNew
                    ProcEnviaDadosCCRealizadoSel TBMateriaprima!ID_CC
                    TBFI.Update
                    
                    'Grava movimentação no centro consolidado
                    Set TBAfericao = CreateObject("adodb.recordset")
                    TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBMateriaprima!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAfericao.EOF = False Then
                        Do While TBAfericao.EOF = False
                            Set TBFI = CreateObject("adodb.recordset")
                            TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                            TBFI.AddNew
                            ProcEnviaDadosCCRealizadoSel TBAfericao!ID_CC
                            TBFI.Update
                            
                            Set TBCiclo = CreateObject("adodb.recordset")
                            TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                            If TBCiclo.EOF = False Then
                                Do While TBCiclo.EOF = False
                                    Set TBFI = CreateObject("adodb.recordset")
                                    TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                                    TBFI.AddNew
                                    ProcEnviaDadosCCRealizadoSel TBCiclo!ID_CC
                                    TBFI.Update
                                    TBCiclo.MoveNext
                                Loop
                            End If
                            TBCiclo.Close
                            
                            TBAfericao.MoveNext
                        Loop
                    End If
                    TBAfericao.Close
                End If
                
                If IsNull(TBMateriaprima!Ordem) = False And TBMateriaprima!Ordem <> 0 And IsNull(TBMateriaprima!ID_PC) = False And TBMateriaprima!ID_PC <> 0 Then ProcGravarOutrasDespOrdem TBMateriaprima!Ordem, TBMateriaprima!ID_PC, Format(QtdeSaida * TBProduto!VlrUnit, "###,##0.00")
            End If
            TBMateriaprima.Close
        End If
        
        ProcAtualizaCTMaterialOrdem cmbEmpresa_RE.ItemData(cmbEmpresa_RE.ListIndex), .ListItems.Item(InitFor).SubItems(1)
        
        TBProduto.Close
        TBEstoque.MoveNext
    Loop
End With
Proximo:
    If Requisicao_materiais = True And Permitido2 = True Then ProcAtualizaStatus_RM

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosCCRealizadoSel(ID_CC As Long)
On Error GoTo tratar_erro

Valor3 = IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario)
valor = Format(Valor3 * QtdeSaida, "###,##0.00")

TBFI!valor = valor
TBFI!Data = txtData
TBFI!Responsavel = pubUsuario
TBFI!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBFI!Operacao = "Débito"
TBFI!ID_estoque = TBProduto!IDoperacao
TBFI!ID_CC = ID_CC
TBFI!Cod_produto = Codproduto
TBFI!ID_PC = IDAntigo
TBFI!Bloqueado = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCancelarReq()
On Error GoTo tratar_erro

If Listamaterial.ListItems.Count = 0 Then Exit Sub
If Listamaterial.SelectedItem.ListSubItems(6) <> "NÃO" And Listamaterial.SelectedItem.ListSubItems(6) <> "CANCEL." Then
    USMsgBox ("Não é permitido alterar o status deste produto requisitado, pois o mesmo já sofreu uma alteração."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
If Listamaterial.SelectedItem.ListSubItems(6) = "CANCEL." Then
    If USMsgBox("Deseja realmente alterar o status deste produto requisitado para NÃO?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        If Requisicao_materiais = False Then
            Conexao.Execute "Update Producaomaterial Set Saida = 'NÃO' where Idmateriaprima = " & Listamaterial.SelectedItem
        Else
            Conexao.Execute "Update Requisicao_materiais_lista Set status = 'REQUISIT.' where Idlista = " & Listamaterial.SelectedItem
        End If
        Permitido = True
    End If
Else
    If USMsgBox("Deseja realmente alterar o status deste produto requisitado para CANCELADO?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        If Requisicao_materiais = False Then
            Conexao.Execute "Update Producaomaterial Set Saida = 'CANCEL.' where Idmateriaprima = " & Listamaterial.SelectedItem
        Else
            Conexao.Execute "Update Requisicao_materiais_lista Set status = 'CANCELADO' where Idlista = " & Listamaterial.SelectedItem
        End If
        Permitido = True
    End If
End If
If Permitido = True Then
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Estoque/Movimentação/Retirada"
    Evento = "Alterar status do produto requisitado"
    ID_documento = Listamaterial.SelectedItem
    Documento = "Cód. interno: " & txtCodigo & " - Descrição: " & txtdescricao
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Listamaterial.ListItems.Clear
    If Requisicao_materiais = False Then
        ProcAtualizalista
    Else
        Set TBMateriaprima = CreateObject("adodb.recordset")
        TBMateriaprima.Open "Select R.Requisicao, RML.* from Requisicao_materiais R INNER JOIN Requisicao_materiais_lista RML ON R.ID = RML.IDrequisicao where R.requisicao = '" & txt_RM & "' and (RML.ID_CC is not null and RML.ID_CC <> 0 and RML.Data_autorizacao is not null or RML.ID_CC is null or RML.ID_CC = 0) order by RML.idlista", Conexao, adOpenKeyset, adLockOptimistic
        If TBMateriaprima.EOF = False Then
            ProcAtualizaLista_RM
        End If
        TBMateriaprima.Close
        ProcAtualizaStatus_RM
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdLocalizaPedido_Click()
On Error GoTo tratar_erro

Sit_REG = 1
If cmbDestino = "Terceiros (Remessa forn.)" Then frmEstoque_retirar_Pedido.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdlote_Click()
On Error GoTo tratar_erro

ProcAtualizaTodas_Listas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

If SSTab1.Tab = 0 Then
    Select Case KeyCode
        Case vbKeyEscape: Unload Me
        Case vbKeyF3: ProcRetirar
        Case vbKeyF4: If Cmb_opcao_lista = "Status" Then ProcCancelarReq
        Case vbKeyF7: If Cmb_opcao_lista = "Retirar selecionados" Then ProcRetirarSelecionados
    End Select
Else
    Select Case KeyCode
        Case vbKeyEscape: Unload Me
        Case vbKeyF3: ProcRetirarRE
        Case vbKeyF7: ProcRetirarSelecionadosRE
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRetirar()
On Error GoTo tratar_erro

Dim qtderetirar As Double
Dim qtderetirado As Double
Dim qtdedisponivel As Double

EstoqueAtual = 0
QtdeSaida = 0
Qtde = 0

Acao = "salvar"

If txt_RM.Text = "" And FrameRM.Visible = True Then
    NomeCampo = "o número da requisição de material"
    ProcVerificaAcao
    txt_RM.SetFocus
    Exit Sub
End If

If txt_Notafiscal.Text = "" And FrameNF.Visible = True Then
    NomeCampo = "o número da nota fiscal"
    ProcVerificaAcao
    txt_Notafiscal.SetFocus
    Exit Sub
End If


If txtRetirar = 0 Then
   If USMsgBox("Já foram retirados todos os produtos desta requisição de material, deseja continuar?", vbYesNo, "CAPRIND v5.0") = vbNo Then
    
    Exit Sub
End If
End If

qtderetirar = txtRetirar
qtderetirado = txtquantretirado
qtdedisponivel = txtestoqueatual

If qtderetirado > qtderetirar Then

    If USMsgBox("A quantidade de saida é maior que o saldo a retirar, deseja realmente baixar a requisição?", vbYesNo, "CAPRIND v5.0") = vbNo Then
    Exit Sub
    End If

End If

If qtderetirado > qtdedisponivel Then
    USMsgBox ("Não existe saldo disponível para essa retirada."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If Listamaterial.SelectedItem.ListSubItems(8) = "CANCEL." Then
    USMsgBox ("Não é permitido retirar este produto do estoque, pois o mesmo está CANCELADO."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Dataini = Format(txtData, "DD/MM/YYYY")
If Dataini > Date Then
    USMsgBox ("A data de saída não pode ser maior que a data de hoje."), vbExclamation, "CAPRIND v5.0"
    'txtdata = Date
    Exit Sub
End If
If txtRequisitante = "" Then
    NomeCampo = "o requisitante"
    ProcVerificaAcao
    Txt_cracha.SetFocus
    Exit Sub
End If
If cmbDestino.Text = "" Then
    NomeCampo = "o destino"
    ProcVerificaAcao
    cmbDestino.SetFocus
    Exit Sub
End If
If txtCodigo.Text = "" Then
    NomeCampo = "o material na lista"
    ProcVerificaAcao
    Exit Sub
End If
If cmbDestino.Text = "Terceiros (Remessa forn.)" Then
    If txtPedidoCompra = "" Then
        NomeCampo = "o pedido de compra"
        ProcVerificaAcao
        cmdLocalizaPedido_Click
        Exit Sub
    End If
End If
If Cmb_RE.Text = "" Then
    NomeCampo = "o número de rastreabilidade do estoque"
    ProcVerificaAcao
    Cmb_RE.SetFocus
    Exit Sub
End If
If txtquantretirado.Locked = False Then NomeCampo = "a quantidade de saída" Else NomeCampo = "a quantidade de saída de peças"
QtdeSaida = IIf(txtquantretirado = "", 0, txtquantretirado)
If QtdeSaida <= 0 Then
    ProcVerificaAcao
    If txtquantretirado.Locked = False Then txtquantretirado.SetFocus 'Else txtquantretirado_PC.SetFocus
    Exit Sub
End If

'Verifica se a quantidade do lote é menor q a quant. retirada
QtdeSaida = txtquantretirado
'QtdeSaidaPC = IIf(txtquantretirado_PC = "", 0, txtquantretirado_PC)
Permitido = True

EstoqueAtual = txtestoqueatual
EstoqueAtualPC = txtestoqueatual_PC

'If EstoqueAtual < QtdeSaida And EstoqueAtualPC = 0 Or EstoqueAtualPC <> 0 And EstoqueAtualPC < QtdeSaidaPC Then
If EstoqueAtual < QtdeSaida Then
    USMsgBox ("Não é permitido retirar, pois a quantidade disponível no estoque é menor que a quantidade de saída."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

''Se for expedição ou RM verifica se a quantidade baixada é maior que a requisitada
'If Expedir = true Or Requisicao_materiais = True Then
'    With Listamaterial.SelectedItem
'        qt = .ListSubItems(1)
'        Qtd = IIf(.ListSubItems(2) = "", 0, .ListSubItems(2))
'        If QtdeSaida > qt Or Qtd <> 0 And QtdeSaidaPC > EstoqueAtualPC Then
'            USMsgBox ("Não é permitido retirar, pois a quantidade requisitada é menor que a quantidade de saída."), vbExclamation, "CAPRIND v5.0"
'            Exit Sub
'        End If
'    End With
'End If

'Aqui começa a retirar do estoque o produto final da nota fiscal
If Expedir = True Then

'======================================================================
' Se for item com controle de rastreabilidade, abre formulário de numero de série
'======================================================================
If chkindividual = Checked Then
frmEstoque_Retirar_NumeroSerie.Show 1

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Count(ID_Nota) AS TotalAnexado from Producao_Etiquetas where ID_Nota = '" & ID_nota & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
    QTSaida = txtRetirar.Text
            If TBAbrir!TotalAnexado < QTSaida Then
                USMsgBox "Atenção o Lote " & Txt_lote.Text & " da RE " & Cmb_RE.Text & " é rastreável, portanto é necessário anexar todos os numeros de série a nota fiscal!", vbCritical, "CAPRIND v5.0"
            Exit Sub
        End If
    End If
    TBAbrir.Close
End If

'Verifica se existe ordem produzindo que esteja empenhada para o produto
If IsNumeric(Txt_lote) = True Then
        OPTexto = ""
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select ECEV.ID_estoque from Estoque_Controle_Empenho_Vendas ECEV INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = ECEV.ID_carteira where NFPP.ID_prod_NF = " & Listamaterial.SelectedItem & " And ECEV.ID_estoque = " & Cmb_RE, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = True Then
            'Verifica se a ordem vinculada ao RE está empenhada para o mesmo pedido da nota
            Set TBproducao = CreateObject("adodb.recordset")
            TBproducao.Open "Select PP.ID from Producao_pedidos PP INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = PP.IDcarteira where PP.Ordem = " & Txt_lote & " and NFPP.ID_prod_NF = " & Listamaterial.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
            If TBproducao.EOF = True Then
                Set TBproducao = CreateObject("adodb.recordset")
                TBproducao.Open "Select QPP.Ordem from (Qtde_produzindo_produto QPP INNER JOIN Producao_pedidos PP ON PP.Ordem = QPP.Ordem) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = PP.IDcarteira where NFPP.ID_prod_NF = " & Listamaterial.SelectedItem & " and QPP.Desenho = '" & txtCodigo & "' and QPP.Ordem <> " & Txt_lote & " and PP.Qtde_empenho > ISNULL(PP.Qtde_entrada, 0)", Conexao, adOpenKeyset, adLockOptimistic
                If TBproducao.EOF = False Then
                    Do While TBproducao.EOF = False
                        If OPTexto = "" Then OPTexto = TBproducao!Ordem Else OPTexto = OPTexto & "|" & TBproducao!Ordem
                        TBproducao.MoveNext
                    Loop
                    USMsgBox ("Não é permitido baixar este RE, pois existe(m) ordem(ns) empenhada(s) na produção para este produto." & vbCrLf & "Ordem(ns): " & OPTexto), vbExclamation, "CAPRIND v5.0"
                    TBAbrir.Close
                    TBproducao.Close
                    Exit Sub
                End If
            End If
            TBproducao.Close
        End If
        TBAbrir.Close
    End If
    
    'Verifica se existe empenho de RE anterior que não foi baixado
    RETexto = ""
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select ECEV.ID_estoque from (Estoque_Controle_Empenho_Vendas ECEV INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = ECEV.ID_carteira) INNER JOIN Estoque_controle EC ON EC.IDestoque = ECEV.ID_estoque where NFPP.ID_prod_NF = " & Listamaterial.SelectedItem & " And ECEV.ID_estoque < " & Cmb_RE & " and ECEV.Qtde_empenhada > ISNULL(ECEV.Qtde_saida, 0) and EC.Estoque_real > 0 group by ECEV.ID_estoque", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            If RETexto = "" Then RETexto = TBAbrir!ID_estoque Else RETexto = RETexto & "|" & TBAbrir!ID_estoque
            TBAbrir.MoveNext
        Loop
        USMsgBox ("Não é permitido baixar este RE, pois existe(m) RE('s) anterior(es) a este que não foi(ram) baixado(s)." & vbCrLf & "RE('s): " & RETexto), vbExclamation, "CAPRIND v5.0"
        TBAbrir.Close
        Exit Sub
    End If
    TBAbrir.Close
End If

'Verifica se a ordem é controlada e trava se a qtde for menor que a requisitada
'If Expedir = true And IsNumeric(Txt_RM) = True Then
'    Set TBproducao = CreateObject("adodb.recordset")
'    TBproducao.Open "Select * from producao where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Ordem = " & Txt_RM & " and (OSControlada = 'True' or Processo_controlado = 'True')", Conexao, adOpenKeyset, adLockOptimistic
'    If TBproducao.EOF = False Then
'        If QtdeSaida > Qtde Then
'            usMsgbox ("Só é permitido retirar a quantidade menor ou igual a requisitada, pois a ordem é controlada."), vbExclamation, "CAPRIND v5.0"
'            TBproducao.Close
'            Exit Sub
'        End If
'    End If
'    TBproducao.Close
'End If

'Aqui retira do estoque controle o item empenhado
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_controle where IDestoque = " & Cmb_RE, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from projproduto where Desenho = '" & txtCodigo & "' and Estoque = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        qtdeliberada = 0
        qtdeliberadaPC = 0
        qtdeliberar = 0
        qtdeliberarPC = 0
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select Sum(Entrada) as qtdeliberada, Sum(ISNULL(Entrada_PC, 0)) as qtdeliberadaPC, Sum(Saida) as qtdeliberar, Sum(ISNULL(Saida_PC, 0)) as qtdeliberarPC from Estoque_movimentacao where IDestoque = " & Cmb_RE, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            qtdeliberada = IIf(IsNull(TBFI!qtdeliberada), 0, TBFI!qtdeliberada)
            qtdeliberadaPC = IIf(IsNull(TBFI!qtdeliberadaPC), 0, TBFI!qtdeliberadaPC)
            qtdeliberar = IIf(IsNull(TBFI!qtdeliberar), 0, TBFI!qtdeliberar)
            qtdeliberarPC = IIf(IsNull(TBFI!qtdeliberarPC), 0, TBFI!qtdeliberarPC)
            QtdeEstoque = Format(qtdeliberada - (qtdeliberar + QtdeSaida), "###,##0.0000")
            QtdeEstoquePC = Format(qtdeliberadaPC - (qtdeliberarPC + QtdeSaidaPC), "###,##0.0000")
        End If
        TBFI.Close
    Else
        QtdeEstoque = 0
    End If
    TBAbrir.Close
    
    TBEstoque!peso_unit = IIf(txtpeso = "", Null, txtpeso)
    TBEstoque!Pedido = IIf(txtPedidoCompra = "", Null, txtPedidoCompra)
   
    TBEstoque!estoque_real = QtdeEstoque
    TBEstoque!estoque_real_PC = QtdeEstoquePC
    TBEstoque!estoque_venda = QtdeEstoque
    TBEstoque!Valor_total = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * QtdeEstoque, "###,##0.00")
           
    Set TBProduto = CreateObject("adodb.recordset")
'Aqui cria a movimentação no estoque da saida por nota fiscal
    TBProduto.Open "Select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
    TBProduto.AddNew
    TBProduto!Documento = IIf(Expedir = True, txt_Notafiscal, txt_RM.Text)
    TBProduto!LOTE = Txt_lote.Text
    TBProduto!Ordem = IIf(IsNumeric(txt_RM.Text), txt_RM.Text, 0)
    TBProduto!Desenho = txtCodigo.Text
    TBProduto!Data = txtData
    TBProduto!Descricao = txtdescricao.Text
    TBProduto!Familia = TBEstoque!Classe
    TBProduto!Requisitante = txtRequisitante.Text
    TBProduto!Responsavel = txtResponsavel.Text
    TBProduto!IDEstoque = TBEstoque!IDEstoque
    'If Expedir = True And Requisicao_materiais = False Then TBProduto!Ordem = txt_RM
    TBProduto!OE = IIf(Expedir = True, txt_Notafiscal, txt_RM.Text)
    TBProduto!Destino = IIf(cmbDestino = "Interno/Cliente", "Interno", "Terceiros")
    If cmbDestino = "Terceiros (Remessa forn.)" Then
        TBProduto!IDpedido = txtIDPedido
        TBProduto!Pedidocompra = txtPedidoCompra
        TBProduto!Terceiros = True
    Else
        TBProduto!Terceiros = False
    End If
    
    TBProduto!Saida = txtquantretirado.Text
    'TBProduto!Saida_PC = IIf(txtquantretirado_PC = "", 0, txtquantretirado_PC)
    TBProduto!estoque_venda = QtdeEstoque

    'Atualiza valor do material no estoque
    TBProduto!VlrUnit = IIf(IsNull(TBEstoque!valor_unitario), 0, Format(TBEstoque!valor_unitario, "###,##0.0000000000"))
    TBProduto!vlrTotal = Format(QtdeSaida * TBProduto!VlrUnit, "###,##0.00")
    
    'verifica se a quantidade retirada e menor q a quant. solicitada
    If Expedir = True Then
        Set TBMateriaprima = CreateObject("adodb.recordset")
        TBMateriaprima.Open "Select NF.ID, NF.int_NotaFiscal, NF.dt_DataEmissao, NFP.* from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Detalhes_Nota NFP ON NFP.ID_nota = NF.ID where NFP.Int_codigo = " & Listamaterial.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBMateriaprima.EOF = False Then
            TBProduto!ID_prod_NF = TBMateriaprima!Int_codigo
            
            SaldoPorc = 0
            If FunVerifUNConversao(TBMateriaprima!txt_Unid, TBMateriaprima!Unidade_com) = True Then
                Qtde = TBMateriaprima!Qtde_estoque
                'Unidade com e un estoque foram invertidas para reverter a conversão
                SaldoPorc = FunConverteUN(TBMateriaprima!Unidade_com, TBMateriaprima!txt_Unid, txtquantretirado, TBMateriaprima!int_Cod_Produto)
            Else
                Qtde = IIf(IsNull(TBMateriaprima!int_Qtd), 0, TBMateriaprima!int_Qtd / FunVerificaTabelaConversaoUnidade(TBMateriaprima!txt_Unid, TBMateriaprima!Unidade_com))
            End If

            'Verifica qtde. de saida
            qtdeliberada = 0
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(Saida) as qtdeliberada from estoque_movimentacao where oe = '" & txt_RM & "' and desenho = '" & Listamaterial.SelectedItem.ListSubItems(4) & "' and documento = '" & txt_RM & "' and (operacao = 'SAIDA_NOTA' or operacao = 'SAIDA_NOTA_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                qtdeliberada = IIf(IsNull(TBAbrir!qtdeliberada), 0, TBAbrir!qtdeliberada)
            End If
            If (QtdeSaida + qtdeliberada) < Qtde Then TBProduto!Operacao = "SAIDA_NOTA_PARCIAL" Else TBProduto!Operacao = "SAIDA_NOTA"

            ProcAtualizaQtdeExpProdPed TBMateriaprima!Int_codigo, TBMateriaprima!int_Cod_Produto, IIf(SaldoPorc = 0, txtquantretirado, SaldoPorc), Txt_lote, TBEstoque!IDEstoque, txtData
            
            If FunVerifCodRefCliDANFE(Cmb_empresa.ItemData(Cmb_empresa.ListIndex)) = True Then TBEstoque!Ref = TBMateriaprima!N_referencia
        End If
        TBMateriaprima.Close
    ElseIf Requisicao_materiais = False Then
            Set TBMateriaprima = CreateObject("adodb.recordset")
            TBMateriaprima.Open "Select * from Producaomaterial where Idmateriaprima = " & Listamaterial.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
            If TBMateriaprima.EOF = False Then
                TBMateriaprima!Valor_saida_estoque = IIf(IsNull(TBMateriaprima!Valor_saida_estoque), 0, TBMateriaprima!Valor_saida_estoque) + TBProduto!vlrTotal
                Qtde = IIf(IsNull(TBMateriaprima!Requisitado), 0, Format(TBMateriaprima!Requisitado, "###,##0.0000"))
                QtdePC = IIf(IsNull(TBMateriaprima!Total_pc), 0, TBMateriaprima!Total_pc)
                
                'Verifica qtde. de saida
                qtdeliberada = 0
                qtdeliberadaPC = 0
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select SUM(Saida) as qtdeliberada, SUM(Saida_pc) as qtdeliberadaPC from Estoque_movimentacao where Ordem = " & txt_RM.Text & " and desenho = '" & Listamaterial.SelectedItem.ListSubItems(4) & "' and documento = '" & txt_RM & "' and Left(Operacao, 11) = 'SAIDA_ORDEM'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    qtdeliberada = IIf(IsNull(TBAbrir!qtdeliberada), 0, Format(TBAbrir!qtdeliberada, "###,##0.0000"))
                    qtdeliberadaPC = IIf(IsNull(TBAbrir!qtdeliberadaPC), 0, Format(TBAbrir!qtdeliberadaPC, "###,##0.0000"))
                End If
                
                If (QtdeSaida + qtdeliberada) >= Qtde Or QtdePC > 0 And (QtdeSaidaPC + qtdeliberadaPC) >= QtdePC Then
                    TBProduto!Operacao = "SAIDA_ORDEM"
                    TBMateriaprima!Saida = "SIM"
                Else
                    TBProduto!Operacao = "SAIDA_ORDEM_PARCIAL"
                    TBMateriaprima!Saida = "PARCIAL"
                End If
                TBMateriaprima.Update
                
                'Atualiza qtde. de saída do empenho da ordem
                QuantEmpenho = 0
                QuantEmpenhoPC = 0
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Sum(Saida) as QuantEmpenho, Sum(ISNULL(Saida, 0)) as QuantEmpenhoPC from estoque_movimentacao where IDestoque = " & Cmb_RE & " and oe = '" & txt_RM & "' and desenho = '" & Listamaterial.SelectedItem.ListSubItems(4) & "' and documento = '" & txt_RM & "' and (operacao = 'SAIDA_ORDEM' or operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    QuantEmpenho = IIf(IsNull(TBAbrir!QuantEmpenho), 0, Format(TBAbrir!QuantEmpenho, "###,##0.0000"))
                    QuantEmpenhoPC = IIf(IsNull(TBAbrir!QuantEmpenhoPC), 0, Format(TBAbrir!QuantEmpenhoPC, "###,##0.0000"))
                    QuantEmpenho = QuantEmpenho + QtdeSaida
                    QuantEmpenhoPC = QuantEmpenhoPC + QtdeSaidaPC
                End If
                NovoValor = Replace(QuantEmpenho, ",", ".")
                NovoValor1 = Replace(QuantEmpenhoPC, ",", ".")
                Conexao.Execute "UPDATE Producao_NF_Consignada Set Qtde_saida = " & NovoValor & " where IDestoque = " & Cmb_RE & " and Ordem = " & txt_RM & " and Codinterno = '" & Listamaterial.SelectedItem.ListSubItems(6) & "'"
                Conexao.Execute "UPDATE Producao_NF_Consignada Set Qtde_saida_PC = " & NovoValor1 & " where IDestoque = " & Cmb_RE & " and Ordem = " & txt_RM & " and Codinterno = '" & Listamaterial.SelectedItem.ListSubItems(6) & "' and Quantidade_PC IS NOT NULL and Quantidade_PC > 0"
                
            End If
            TBMateriaprima.Close
        Else
            Set TBMateriaprima = CreateObject("adodb.recordset")
            TBMateriaprima.Open "Select * from Requisicao_materiais_lista where idlista = " & Listamaterial.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
            If TBMateriaprima.EOF = False Then
                TBProduto!ID_prod_RM = TBMateriaprima!IDlista
                
                Qtde = IIf(IsNull(TBMateriaprima!Quant), 0, TBMateriaprima!Quant)
                If (TBMateriaprima!quant_saida + QtdeSaida) < Qtde Then
                    TBMateriaprima!status = "PARCIAL"
                    TBProduto!Operacao = "SAIDA_REQUISICAO_PARCIAL"
                Else
                    TBMateriaprima!status = "RETIRADO"
                    TBProduto!Operacao = "SAIDA_REQUISICAO"
                End If
                TBMateriaprima!quant_saida = Format(IIf(IsNull(TBMateriaprima!quant_saida), 0, TBMateriaprima!quant_saida) + txtquantretirado, "###,##0.0000")
                TBMateriaprima.Update
                
                If IsNull(TBMateriaprima!Ordem) = False And TBMateriaprima!Ordem <> 0 And IsNull(TBMateriaprima!ID_PC) = False And TBMateriaprima!ID_PC <> 0 Then ProcGravarOutrasDespOrdem TBMateriaprima!Ordem, TBMateriaprima!ID_PC, Format(QtdeSaida * TBProduto!VlrUnit, "###,##0.00")
            End If
            TBMateriaprima.Close
            'ProcAtualizaStatus_RM
    End If
    
    TBEstoque.Update
    TBProduto.Update
    
   '===================================
   ' Atualiza tipo item na movimentação
   '===================================
   Conexao.Execute "UPDATE Estoque_movimentacao Set IDEstoque = " & TBEstoque!IDEstoque & " where IDoperacao = " & TBProduto!IDoperacao
   '===================================
    
    
    'Centro de custo
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select * from projproduto where desenho = '" & txtCodigo & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        Codproduto = TBItem!Codproduto
        IDAntigo = IIf(IsNull(TBItem!ID_PC), 0, TBItem!ID_PC)
    End If
    TBItem.Close
    
    valor = TBProduto!vlrTotal
    If Requisicao_materiais = True Then
        ProcCriaCreditoCCProdutoItem
        Set TBMateriaprima = CreateObject("adodb.recordset")
        TBMateriaprima.Open "Select * from Requisicao_materiais_lista where idlista = " & Listamaterial.SelectedItem & " and ID_CC is not null", Conexao, adOpenKeyset, adLockOptimistic
        If TBMateriaprima.EOF = False Then
            If TBMateriaprima!ID_CC <> "" Then
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                TBFI.AddNew
                ProcEnviaDadosCCRealizado TBMateriaprima!ID_CC
                TBFI.Update
                
                'Grava movimentação no centro consolidado
                Set TBAfericao = CreateObject("adodb.recordset")
                TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBMateriaprima!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                If TBAfericao.EOF = False Then
                    Do While TBAfericao.EOF = False
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                        TBFI.AddNew
                        ProcEnviaDadosCCRealizado TBAfericao!ID_CC
                        TBFI.Update
                        
                        Set TBCiclo = CreateObject("adodb.recordset")
                        TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                        If TBCiclo.EOF = False Then
                            Do While TBCiclo.EOF = False
                                Set TBFI = CreateObject("adodb.recordset")
                                TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                                TBFI.AddNew
                                ProcEnviaDadosCCRealizado TBCiclo!ID_CC
                                TBFI.Update
                                TBCiclo.MoveNext
                            Loop
                        End If
                        TBCiclo.Close
                        
                        TBAfericao.MoveNext
                    Loop
                End If
                TBAfericao.Close
            End If
        End If
        TBMateriaprima.Close
    Else
        ProcCriaCreditoCCProdutoItem
    End If
    TBProduto.Close
End If

If Expedir = True And IsNumeric(txt_RM) = True Then
ProcAtualizaCTMaterialOrdem Cmb_empresa.ItemData(Cmb_empresa.ListIndex), txt_RM
End If

USMsgBox ("Produto retirado do estoque com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Estoque/Movimentação/Retirada"
Evento = "Retirar"
ID_documento = Listamaterial.SelectedItem
Documento = "Cód. interno: " & txtCodigo & " - RE: " & txt_RE
Documento1 = ""
ProcGravaEvento
'==================================
cmdlote_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarOutrasDespOrdem(Ordem As Long, ID_PC As Long, valor As Double)
On Error GoTo tratar_erro

Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select * from Producao_outras_despesas where Ordem = " & Ordem & " and ID_PC = " & ID_PC, Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = True Then TBOrdem.AddNew
TBOrdem!Ordem = Ordem
TBOrdem!ID_PC = ID_PC
TBOrdem!valor = valor + IIf(IsNull(TBOrdem!valor), 0, TBOrdem!valor)
TBOrdem.Update

Valor1 = 0
Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select Sum(Valor) as Valor1 from Producao_outras_despesas where Ordem = " & Ordem, Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
    Valor1 = IIf(IsNull(TBOrdem!Valor1), 0, TBOrdem!Valor1)
End If
NovoValor = Replace(Valor1, ",", ".")
Conexao.Execute "Update Producao Set CTOutras = " & NovoValor & " where Ordem = " & Ordem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcRetirarRE()
On Error GoTo tratar_erro

EstoqueAtual = 0
QtdeSaida = 0
Qtde = 0

Acao = "salvar"
If txtRE.Text = "" Then
    NomeCampo = "o número de rastreabilidade do estoque"
    ProcVerificaAcao
    Exit Sub
End If
If ListaMaterial_RE.ListItems.Count = 0 Then
    USMsgBox ("Não existe nenhuma requisição de material para baixar essa RE."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If ListaMaterial_RE.SelectedItem.ListSubItems(7) = "CANCEL." Then
    USMsgBox ("Não é permitido retirar este produto do estoque, pois o mesmo está CANCELADO."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Dataini = txtData
If Dataini > Date Then
    USMsgBox ("A data de saída não pode ser maior que a data de hoje."), vbExclamation, "CAPRIND v5.0"
    txtData_RE = Date
    Exit Sub
End If
If txtRequisitante_RE = "" Then
    NomeCampo = "o requisitante"
    ProcVerificaAcao
    Txt_cracha_RE.SetFocus
    Exit Sub
End If
If cmbDestino_RE.Text = "" Then
    NomeCampo = "o destino"
    ProcVerificaAcao
    cmbDestino_RE.SetFocus
    Exit Sub
End If
If txtQtde_RE = "" Then
    NomeCampo = "o material na lista"
    ProcVerificaAcao
    Exit Sub
End If
If cmbDestino_RE.Text = "Terceiros (Remessa forn.)" Then
    If txtPedidoCompra = "" Then
        NomeCampo = "o pedido de compra"
        ProcVerificaAcao
        cmdPedido_RE_Click
        Exit Sub
    End If
End If
If txtQtde_Saida_RE.Locked = False Then NomeCampo = "a quantidade de saída" Else NomeCampo = "a quantidade de saída de peças"
QtdeSaida = IIf(txtQtde_Saida_RE = "", 0, txtQtde_Saida_RE)
If QtdeSaida <= 0 Then
    ProcVerificaAcao
    If txtQtde_Saida_RE.Locked = False Then txtQtde_Saida_RE.SetFocus Else txtQtde_Saida_RE_PC.SetFocus
    Exit Sub
End If

'Verifica se a quantidade do lote é menor q a quant. retirada
QtdeSaida = txtQtde_Saida_RE
QtdeSaidaPC = IIf(txtQtde_Saida_RE_PC = "", 0, txtQtde_Saida_RE_PC)
Permitido = True

EstoqueAtual = txtEstoque_Real_RE
EstoqueAtualPC = txtEstoque_Real_RE_PC

If EstoqueAtual < QtdeSaida And EstoqueAtualPC = 0 Or EstoqueAtualPC <> 0 And EstoqueAtualPC < QtdeSaidaPC Then
    USMsgBox ("Não é permitido retirar, pois a quantidade disponível no estoque é menor que a quantidade de saída."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

'Se for Nota fiscal ou RM verifica se a quantidade baixada é maior que a requisitada
If Expedir = True Or Requisicao_materiais = True Then
    With Listamaterial.SelectedItem
        qt = .ListSubItems(2)
        Qtd = IIf(.ListSubItems(3) = "", 0, .ListSubItems(3))
        If QtdeSaida > qt Or Qtd <> 0 And QtdeSaidaPC > EstoqueAtualPC Then
            USMsgBox ("Não é permitido retirar, pois a quantidade requisitada é menor que a quantidade de saída."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
    End With
End If

Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_controle where IDestoque = " & txtRE, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from projproduto where Desenho = '" & txtCodInterno_RE & "' and Estoque = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        qtdeliberada = 0
        qtdeliberadaPC = 0
        qtdeliberar = 0
        qtdeliberarPC = 0
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select Sum(Entrada) as qtdeliberada, Sum(ISNULL(Entrada_PC, 0)) as qtdeliberadaPC, Sum(Saida) as qtdeliberar, Sum(ISNULL(Saida_PC, 0)) as qtdeliberarPC from Estoque_movimentacao where IDestoque = " & txtRE, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            qtdeliberada = IIf(IsNull(TBFI!qtdeliberada), 0, TBFI!qtdeliberada)
            qtdeliberadaPC = IIf(IsNull(TBFI!qtdeliberadaPC), 0, TBFI!qtdeliberadaPC)
            qtdeliberar = IIf(IsNull(TBFI!qtdeliberar), 0, TBFI!qtdeliberar)
            qtdeliberarPC = IIf(IsNull(TBFI!qtdeliberarPC), 0, TBFI!qtdeliberarPC)
            QtdeEstoque = Format(qtdeliberada - (qtdeliberar + QtdeSaida), "###,##0.0000")
            QtdeEstoquePC = Format(qtdeliberadaPC - (qtdeliberarPC + QtdeSaidaPC), "###,##0.0000")
        End If
        TBFI.Close
    Else
        QtdeEstoque = 0
    End If
    TBAbrir.Close
    
    TBEstoque!peso_unit = IIf(txtPeso_RE = "", Null, txtPeso_RE)
    TBEstoque!Pedido = IIf(txtPedido_RE = "", Null, txtPedido_RE)
   
    TBEstoque!estoque_real = QtdeEstoque
    TBEstoque!estoque_real_PC = QtdeEstoquePC
    TBEstoque!estoque_venda = QtdeEstoque
    TBEstoque!Valor_total = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * QtdeEstoque, "###,##0.00")
           
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
    TBProduto.AddNew
    TBProduto!Documento = ListaMaterial_RE.SelectedItem.ListSubItems(1)
    TBProduto!LOTE = txtLote_RE
    TBProduto!Desenho = txtCodInterno_RE
    TBProduto!Data = txtData_RE
    TBProduto!Descricao = txtDescricao_RE
    TBProduto!Familia = TBEstoque!Classe
    TBProduto!Requisitante = txtRequisitante_RE
    TBProduto!Responsavel = txtResponsavel_RE
    TBProduto!IDEstoque = TBEstoque!IDEstoque
    TBProduto!Ordem = ListaMaterial_RE.SelectedItem.ListSubItems(1)
    TBProduto!OE = ListaMaterial_RE.SelectedItem.ListSubItems(1)
    TBProduto!Destino = IIf(cmbDestino_RE = "Interno/Cliente", "Interno", "Terceiros")
    If cmbDestino_RE = "Terceiros (Remessa forn.)" Then
        TBProduto!IDpedido = txtIdPedido_RE
        TBProduto!Pedidocompra = txtPedido_RE
        TBProduto!Terceiros = True
    Else
        TBProduto!Terceiros = False
    End If
    
    TBProduto!Saida = txtQtde_Saida_RE.Text
    TBProduto!Saida_PC = IIf(txtQtde_Saida_RE_PC = "", 0, txtQtde_Saida_RE_PC)
    TBProduto!estoque_venda = QtdeEstoque

    'Atualiza valor do material no estoque
    TBProduto!VlrUnit = IIf(IsNull(TBEstoque!valor_unitario), 0, Format(TBEstoque!valor_unitario, "###,##0.0000000000"))
    TBProduto!vlrTotal = Format(QtdeSaida * TBProduto!VlrUnit, "###,##0.00")

    'verifica se a quantidade retirada e menor q a quant. solicitada
    Set TBMateriaprima = CreateObject("adodb.recordset")
    TBMateriaprima.Open "Select * from Producaomaterial where Idmateriaprima = " & ListaMaterial_RE.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBMateriaprima.EOF = False Then
        TBMateriaprima!Valor_saida_estoque = IIf(IsNull(TBMateriaprima!Valor_saida_estoque), 0, TBMateriaprima!Valor_saida_estoque) + TBProduto!vlrTotal
        Qtde = IIf(IsNull(TBMateriaprima!Requisitado), 0, Format(TBMateriaprima!Requisitado, "###,##0.0000"))
        QtdePC = IIf(IsNull(TBMateriaprima!Total_pc), 0, TBMateriaprima!Total_pc)
        
        'Verifica qtde. de saida
        qtdeliberada = 0
        qtdeliberadaPC = 0
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select SUM(Saida) as qtdeliberada, SUM(Saida_pc) as qtdeliberadaPC from Estoque_movimentacao where Ordem = " & ListaMaterial_RE.SelectedItem.ListSubItems(1) & " and desenho = '" & ListaMaterial_RE.SelectedItem.ListSubItems(5) & "' and documento = '" & ListaMaterial_RE.SelectedItem.ListSubItems(1) & "' and Left(Operacao, 11) = 'SAIDA_ORDEM'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            qtdeliberada = IIf(IsNull(TBAbrir!qtdeliberada), 0, Format(TBAbrir!qtdeliberada, "###,##0.0000"))
            qtdeliberadaPC = IIf(IsNull(TBAbrir!qtdeliberadaPC), 0, Format(TBAbrir!qtdeliberadaPC, "###,##0.0000"))
        End If
        
        If (QtdeSaida + qtdeliberada) >= Qtde Or QtdePC > 0 And (QtdeSaidaPC + qtdeliberadaPC) >= QtdePC Then
            TBProduto!Operacao = "SAIDA_ORDEM"
            TBMateriaprima!Saida = "SIM"
        Else
            TBProduto!Operacao = "SAIDA_ORDEM_PARCIAL"
            TBMateriaprima!Saida = "PARCIAL"
        End If
        TBMateriaprima.Update
        
        'Atualiza qtde. de saída do empenho da ordem
        QuantEmpenho = 0
        QuantEmpenhoPC = 0
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Sum(Saida) as QuantEmpenho, Sum(ISNULL(Saida, 0)) as QuantEmpenhoPC from estoque_movimentacao where IDestoque = " & txtRE & " and OE = '" & ListaMaterial_RE.SelectedItem.ListSubItems(1) & "' and desenho = '" & ListaMaterial_RE.SelectedItem.ListSubItems(5) & "' and documento = '" & ListaMaterial_RE.SelectedItem.ListSubItems(1) & "' and (operacao = 'SAIDA_ORDEM' or operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            QuantEmpenho = IIf(IsNull(TBAbrir!QuantEmpenho), 0, Format(TBAbrir!QuantEmpenho, "###,##0.0000"))
            QuantEmpenhoPC = IIf(IsNull(TBAbrir!QuantEmpenhoPC), 0, Format(TBAbrir!QuantEmpenhoPC, "###,##0.0000"))
            QuantEmpenho = QuantEmpenho + QtdeSaida
            QuantEmpenhoPC = QuantEmpenhoPC + QtdeSaidaPC
        End If
        NovoValor = Replace(QuantEmpenho, ",", ".")
        NovoValor1 = Replace(QuantEmpenhoPC, ",", ".")
        Conexao.Execute "UPDATE Producao_NF_Consignada Set Qtde_saida = " & NovoValor & " where IDestoque = " & txtRE & " and Ordem = " & ListaMaterial_RE.SelectedItem.ListSubItems(1) & " and Codinterno = '" & ListaMaterial_RE.SelectedItem.ListSubItems(5) & "'"
        Conexao.Execute "UPDATE Producao_NF_Consignada Set Qtde_saida_PC = " & NovoValor1 & " where IDestoque = " & txtRE & " and Ordem = " & ListaMaterial_RE.SelectedItem.ListSubItems(1) & " and Codinterno = '" & ListaMaterial_RE.SelectedItem.ListSubItems(5) & "' and Quantidade_PC IS NOT NULL and Quantidade_PC > 0"
    End If
    TBMateriaprima.Close
    
    TBEstoque.Update
    TBProduto.Update
    '===================================
    ' Atualiza tipo item na movimentação
    '===================================
    Conexao.Execute "UPDATE Estoque_movimentacao Set IDEstoque = " & TBEstoque!IDEstoque & " where IDoperacao = " & TBProduto!IDoperacao
    '===================================
    
    
    'Centro de custo
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select * from projproduto where desenho = '" & txtCodInterno_RE & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        Codproduto = TBItem!Codproduto
        IDAntigo = IIf(IsNull(TBItem!ID_PC), 0, TBItem!ID_PC)
    End If
    TBItem.Close
    
    valor = TBProduto!vlrTotal
    If Requisicao_materiais = True Then
        ProcCriaCreditoCCProdutoItem
        Set TBMateriaprima = CreateObject("adodb.recordset")
        TBMateriaprima.Open "Select * from Requisicao_materiais_lista where idlista = " & ListaMaterial_RE.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBMateriaprima.EOF = False Then
            If IsNull(TBMateriaprima!ID_CC) = False And TBMateriaprima!ID_CC <> "" And TBMateriaprima!ID_CC <> "0" Then
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                TBFI.AddNew
                ProcEnviaDadosCCRealizado TBMateriaprima!ID_CC
                TBFI.Update
                
                'Grava movimentação no centro consolidado
                Set TBAfericao = CreateObject("adodb.recordset")
                TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBMateriaprima!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                If TBAfericao.EOF = False Then
                    Do While TBAfericao.EOF = False
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                        TBFI.AddNew
                        ProcEnviaDadosCCRealizado TBAfericao!ID_CC
                        TBFI.Update
                        
                        Set TBCiclo = CreateObject("adodb.recordset")
                        TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                        If TBCiclo.EOF = False Then
                            Do While TBCiclo.EOF = False
                                Set TBFI = CreateObject("adodb.recordset")
                                TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                                TBFI.AddNew
                                ProcEnviaDadosCCRealizado TBCiclo!ID_CC
                                TBFI.Update
                                TBCiclo.MoveNext
                            Loop
                        End If
                        TBCiclo.Close
                        
                        TBAfericao.MoveNext
                    Loop
                End If
                TBAfericao.Close
                
                If IsNull(TBMateriaprima!Ordem) = False And TBMateriaprima!Ordem <> 0 And IsNull(TBMateriaprima!ID_PC) = False And TBMateriaprima!ID_PC <> 0 Then ProcGravarOutrasDespOrdem TBMateriaprima!Ordem, TBMateriaprima!ID_PC, Format(QtdeSaida * TBProduto!VlrUnit, "###,##0.00")
            End If
        End If
        TBMateriaprima.Close
    Else
        ProcCriaCreditoCCProdutoItem
    End If
    TBProduto.Close
End If

ProcAtualizaCTMaterialOrdem cmbEmpresa_RE.ItemData(cmbEmpresa_RE.ListIndex), ListaMaterial_RE.SelectedItem.ListSubItems(1).Text

USMsgBox ("Produto retirado do estoque com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Estoque/Movimentação/Retirada"
Evento = "Retirar"
ID_documento = ListaMaterial_RE.SelectedItem
Documento = "Cód. interno: " & txtCodInterno_RE & " - RE: " & txtRE
Documento1 = ""
ProcGravaEvento
'==================================
cmdRE_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosCCRealizado(ID_CC As Long)
On Error GoTo tratar_erro

Valor3 = IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario)
Qtde = txtquantretirado
valor = Format(Valor3 * Qtde, "###,##0.00")

TBFI!valor = valor
TBFI!Data = txtData
TBFI!Responsavel = pubUsuario
TBFI!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBFI!Operacao = "Débito"
TBFI!ID_estoque = TBProduto!IDoperacao
TBFI!ID_CC = ID_CC
TBFI!Cod_produto = Codproduto
TBFI!ID_PC = IDAntigo
TBFI!Bloqueado = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriaCreditoCCProdutoItem()
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from projproduto where Codproduto = " & Codproduto & " and ID_CC is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    If TBFIltro!ID_CC <> "" Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
        TBFI.AddNew
        ProcEnviaDadosCCProdutoItem TBFIltro!ID_CC
        TBFI.Update
        
        'Grava movimentação no centro consolidado
        Set TBAfericao = CreateObject("adodb.recordset")
        TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBFIltro!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
        If TBAfericao.EOF = False Then
            Do While TBAfericao.EOF = False
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                TBFI.AddNew
                ProcEnviaDadosCCProdutoItem TBAfericao!ID_CC
                TBFI.Update
                
                Set TBCiclo = CreateObject("adodb.recordset")
                TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                If TBCiclo.EOF = False Then
                    Do While TBCiclo.EOF = False
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
                        TBFI.AddNew
                        ProcEnviaDadosCCProdutoItem TBCiclo!ID_CC
                        TBFI.Update
                        TBCiclo.MoveNext
                    Loop
                End If
                TBCiclo.Close
                
                TBAfericao.MoveNext
            Loop
        End If
        TBAfericao.Close
    End If
End If
TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosCCProdutoItem(ID_CC As Long)
On Error GoTo tratar_erro

TBFI!valor = valor
TBFI!Data = txtData
TBFI!Responsavel = pubUsuario
TBFI!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBFI!Operacao = "Crédito"
TBFI!ID_estoque = TBProduto!IDoperacao
TBFI!ID_CC = ID_CC
TBFI!Cod_produto = Codproduto
TBFI!ID_PC = IDAntigo
TBFI!Bloqueado = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15225, 7, True
ProcCarregaToolBar2 Me, 15225, 5, True
SSTab1.Tab = 0
txtData = Date
txtData_RE = Date
txtResponsavel.Text = pubUsuario
Cmb_opcao_lista = "Retirar selecionados"
txtResponsavel_RE.Text = pubUsuario
ProcCarregaComboEmpresa Cmb_empresa, False
ProcCarregaComboEmpresa cmbEmpresa_RE, False
ProcLimpaCampos
ProcLimpaCampos_TabRE

ProcRemoveObjetosResize Me
btnNotafiscal_Click

Load = True
Load = False

FrRastreabilidade.Enabled = False
chkindividual.Value = Unchecked

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listamaterial_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Listamaterial
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If Cmb_opcao_lista = "Retirar selecionados" Then
                    If .ListItems.Item(InitFor).ListSubItems(6) <> "RETIRADO" And .ListItems.Item(InitFor).ListSubItems(6) <> "CANCEL." Then .ListItems.Item(InitFor).Checked = True
                Else
                    If .ListItems.Item(InitFor).ListSubItems(6) <> "RETIRADO" Then .ListItems.Item(InitFor).Checked = True
                End If
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Listamaterial, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listamaterial_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

ProcCarregaComboRE
ID_produto_nota = Listamaterial.SelectedItem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboRE()
On Error GoTo tratar_erro

If Listamaterial.ListItems.Count = 0 Then Exit Sub
ProcLimpaCampos
With Listamaterial.SelectedItem
    txtCodigo.Text = .ListSubItems(6)
    txtQuant_prevista = .ListSubItems(1)
    txtQuant_prevista_PC = .ListSubItems(4)
    txtBaixado = .ListSubItems(2)
    txtRetirar = .ListSubItems(3)
    
   ' If Cmb_RE <> "" Then txtquantretirado = txtQuant_prevista
End With

ProcCarregaProduto
If Expedir = True Then 'Retirar item de nota fiscal
    ProcCarregaRE_NF
    ProcCarregaLote_NF
Else 'Retirar item de requisicao de materiais
    If Requisicao_materiais = True Then ProcCarregaRE_RM Else ProcCarregaRE
    If Requisicao_materiais = False Then ProcCarregaLote
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtCodigo.Text = ""
'txtdata = Date
txtdescricao.Text = ""
txtunidade.Text = ""
cmbN_ref.Clear
txtQuant_prevista = ""
txtpeso = ""
txtUN = ""
txtPedidoCompra = ""
txtIDPedido = ""
Cmb_RE.ListIndex = -1
ProcLimpaCamposRE

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos_TabRE()
On Error GoTo tratar_erro

txtCodInterno_RE = ""
txtEstoque_Real_RE = "0,0000"
txtEmpenhos = "0,0000"
txtDisponivel_RE = "0,0000"
txtEstoque_Real_RE_PC = "0,0000"
txtEmpenho_PC = "0,0000"
txtDisponivel_PC_RE = "0,0000"
txtQtde_Saida_RE = "0,0000"
txtQtde_Saida_RE_PC = 0
txtAtualizado_RE = "0,0000"
txtAtualizado_PC_RE = "0,0000"
txtData_RE = Date
txtDescricao_RE = ""
txtUN_RE = ""
cmbReferencia_RE.Clear
txtQtde_RE = ""
txtQtde_PC_RE = ""
txtPeso_RE = ""
txtUnKg_RE = ""
txtPedido_RE = ""
txtIdPedido_RE = ""
txtLote_RE = ""
Txt_cod_ref_RE_RE = ""
txtLocal_RE = ""
txtCorrida_RE = ""
txtCertificado_RE = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposRE()
On Error GoTo tratar_erro

Txt_lote = ""
Txt_cod_ref_RE = ""
txtLocal_armaz = ""
txtCertificado = ""
txtcorrida = ""
txtestoqueatual = "0,0000"
txtestoqueatual_PC = "0,0000"
txtquantretirado.Text = "0,0000"
'txtquantretirado_PC = 0
txtestoqueatual = "0,0000"
txtestoqueatual_PC = "0,0000"
txtBaixado = "0,0000"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaProduto()
On Error GoTo tratar_erro

'txtdata = Date
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from projproduto where desenho = '" & txtCodigo & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcCarregaComboCodRef cmbN_ref, "P.codproduto = " & TBProduto!Codproduto & " and IA.N_Referencia <> 'Null'", 0, "", False, False
    txtdescricao = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
    txtunidade.Text = TBProduto!Unidade
    txtpeso.Text = IIf(IsNull(TBProduto!peso_metro), "", TBProduto!peso_metro)
    txtUN.Text = IIf(IsNull(TBProduto!Un_Kg), "", TBProduto!Un_Kg)
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaProduto_RE()
On Error GoTo tratar_erro

txtData_RE = Date
cmbReferencia_RE.Clear
txtDescricao_RE = ""
txtUN_RE = ""
txtPeso_RE = ""
txtUnKg_RE = ""
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select Codproduto, Descricao, Unidade, peso_metro, un_kg from projproduto where desenho = '" & txtCodInterno_RE & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    ProcCarregaComboCodRef cmbReferencia_RE, "P.codproduto = " & TBProduto!Codproduto & " and IA.N_Referencia <> 'Null'", 0, "", False, False
    txtDescricao_RE = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
    txtUN_RE = TBProduto!Unidade
    txtPeso_RE = IIf(IsNull(TBProduto!peso_metro), "", TBProduto!peso_metro)
    txtUnKg_RE = IIf(IsNull(TBProduto!Un_Kg), "", TBProduto!Un_Kg)
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaRE()
On Error GoTo tratar_erro

With Cmb_RE
    .Clear
    Set TBFIltro = CreateObject("adodb.recordset")
    StrSql = "Select PNC.IDestoque from (Producao P INNER JOIN Producao_NF_Consignada PNC ON P.Ordem = PNC.Ordem) INNER JOIN Estoque_controle EC ON EC.IDestoque = PNC.IDestoque where P.ordem = " & txt_RM & " and PNC.Codinterno = '" & txtCodigo & "' and EC.Estoque_real > 0 group by PNC.IDestoque"
    'Debug.print StrSql
   ' StrSql = "Select PNC.IDestoque from (Producao P INNER JOIN Producao_NF_Consignada PNC ON P.Ordem = PNC.Ordem) INNER JOIN Estoque_controle EC ON EC.IDestoque = PNC.IDestoque where P.ordem = " & txt_RM & " and PNC.Codinterno = '" & txtCodigo & "' and EC.Estoque_real > 0 group by PNC.IDestoque"
    'Debug.print StrSql
    
    TBFIltro.Open "Select PNC.IDestoque from (Producao P INNER JOIN Producao_NF_Consignada PNC ON P.Ordem = PNC.Ordem) INNER JOIN Estoque_controle EC ON EC.IDestoque = PNC.IDestoque where P.ordem = " & txt_RM & " and PNC.Codinterno = '" & txtCodigo & "' and EC.Estoque_real > 0 group by PNC.IDestoque", Conexao, adOpenKeyset, adLockOptimistic
    
    If TBFIltro.EOF = False Then
        Do While TBFIltro.EOF = False
            .AddItem TBFIltro!IDEstoque
            TBFIltro.MoveNext
        Loop
    Else
        TextoFiltro = "Consignacao = 'False'"
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select IDCliente, Cliente from Producao where ordem = " & txt_RM & " and consignacao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            TextoFiltro = "(EP.Consignacao = 'True' and EP.id_cliente = " & TBFI!IDCliente & " and EP.Cliente = '" & TBFI!Cliente & "' or EP.Consignacao = 'False')"
            'TextoFiltro = "(EP.Consignacao = 'True' and EP.Cliente = '" & TBFI!Cliente & "' or EP.Consignacao = 'False')"
        End If
        TBFI.Close
        
        Set TBFIltro = CreateObject("adodb.recordset")
        StrSql = "Select EP.IDestoque from (Estoque_produtos EP LEFT JOIN Qtde_empenhada_produto_venda_detalhado EE ON EE.ID_estoque = EP.IDestoque) INNER JOIN Estoque_Localarmazenamento_criar EL ON EL.descricao = EP.local_armaz LEFT JOIN Qtde_empenhada_produto_detalhado PNFC ON PNFC.IDestoque = EP.IDestoque WHERE EP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and EP.Desenho = '" & txtCodigo & "' and " & TextoFiltro & " and EP.Lote is not null and EP.Liberado = 'SIM' and (Left(EP.status, 7) = 'ENTRADA' or EP.status = 'CONSIGNAÇÃO RECEBIDA') AND EL.Estoque = 'False' Group by EP.IDestoque, EP.estoque_real HAVING EP.estoque_real - SUM(ISNULL(EE.Qtde_requisitar, 0) + ISNULL(PNFC.Qtde_empenhar, 0)) > 0"
        'Debug.print StrSql
        
        TBFIltro.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            Do While TBFIltro.EOF = False
                .AddItem TBFIltro!IDEstoque
                TBFIltro.MoveNext
            Loop
        End If
    End If
    TBFIltro.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLote()
On Error GoTo tratar_erro

With cmb_Lote
    .Clear
'=================================================================
' Materia prima consignada na nota fiscal
'=================================================================
Set TBFIltro = CreateObject("adodb.recordset")
StrSql = "Select PNC.IDestoque from (Producao P INNER JOIN Producao_NF_Consignada PNC ON P.Ordem = PNC.Ordem) INNER JOIN Estoque_controle EC ON EC.IDestoque = PNC.IDestoque where P.ordem = " & txt_RM & " and PNC.Codinterno = '" & txtCodigo & "' and EC.Estoque_real > 0 group by PNC.IDestoque"
'Debug.print StrSql

TBFIltro.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
Do While TBFIltro.EOF = False
   .AddItem TBFIltro!IDEstoque
   TBFIltro.MoveNext
Loop
Else
'==================================================================
' Produto consignado
'==================================================================
 TextoFiltro = "Consignacao = 'False'"
 Set TBFI = CreateObject("adodb.recordset")
        
StrSql = "Select IDCliente, Cliente from Producao where ordem = " & txt_RM & " and consignacao = 'True'"
'Debug.print StrSql

TBFI.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
   TextoFiltro = "(EP.Consignacao = 'True' and EP.id_cliente = " & TBFI!IDCliente & " and EP.Cliente = '" & TBFI!Cliente & "' or EP.Consignacao = 'False')"
End If
TBFI.Close

StrSql = "Select EP.IDEstoque,EP.Lote from (Estoque_produtos EP LEFT JOIN Qtde_empenhada_produto_venda_detalhado EE ON EE.ID_estoque = EP.IDestoque) INNER JOIN Estoque_Localarmazenamento_criar EL ON EL.descricao = EP.local_armaz LEFT JOIN Qtde_empenhada_produto_detalhado PNFC ON PNFC.IDestoque = EP.IDestoque WHERE EP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and EP.Desenho = '" & txtCodigo & "' and " & TextoFiltro & " and EP.Lote is not null and EP.Liberado = 'SIM' and (Left(EP.status, 7) = 'ENTRADA' or EP.status = 'CONSIGNAÇÃO RECEBIDA') AND EL.Estoque = 'False' Group by EP.Lote, EP.IDestoque, EP.estoque_real HAVING EP.estoque_real - SUM(ISNULL(EE.Qtde_requisitar, 0) + ISNULL(PNFC.Qtde_empenhar, 0)) > 0"
'Debug.print StrSql

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
   Do While TBFIltro.EOF = False
       .AddItem TBFIltro!LOTE
       TBFIltro.MoveNext
   Loop
End If
End If
TBFIltro.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaRE_RM()
On Error GoTo tratar_erro

StrSql = "Select IDestoque,Estoque_disponivel, Data from Estoque_produtos Where Estoque_disponivel > 0  and Data < = '" & txtData & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Desenho = '" & txtCodigo & "' and Lote is not null and Liberado = 'SIM' and (Left(status, 7) = 'ENTRADA' or status = 'CONSIGNAÇÃO RECEBIDA') Group by data, IDestoque,Estoque_disponivel order by Data desc"
'Debug.print StrSql

With Cmb_RE
    .Clear
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        Do While TBFIltro.EOF = False
            .AddItem TBFIltro!IDEstoque
            TBFIltro.MoveNext
        Loop
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaRE_NF_Lote()
On Error GoTo tratar_erro

With Cmb_RE
    .Clear
    .Enabled = True
    .Locked = False
    .TabStop = True
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select NFP.IDestoque, CFOP.Devolucao from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.Int_codigo = " & Listamaterial.SelectedItem, Conexao, adOpenKeyset, adLockReadOnly
    If TBAbrir.EOF = False Then
        Set TBFI = CreateObject("adodb.recordset")
        
        If chkEmpenhos.Value = False Then
            TBFI.Open "Select EC.IDestoque from Estoque_Controle EC where EC.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and EC.desenho = '" & txtCodigo.Text & "' and EC.Lote is not null and EC.estoque_real > 0 and (Left(EC.status, 7) = 'ENTRADA' or EC.status = 'CONSIGNAÇÃO RECEBIDA') Group by EC.IDestoque", Conexao, adOpenKeyset, adLockReadOnly
        Else
            TBFI.Open "Select EC.IDestoque from (Estoque_Controle_Empenho_Vendas EE INNER JOIN estoque_controle EC ON EE.ID_estoque = EC.IDEstoque) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = EE.ID_carteira and NFPP.Codinterno = EC.Desenho where NFPP.ID_prod_NF = " & Listamaterial.SelectedItem & " and EC.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and EC.desenho = '" & txtCodigo.Text & "' and EC.Lote is not null and EC.estoque_real > 0 and (Left(EC.status, 7) = 'ENTRADA' or EC.status = 'CONSIGNAÇÃO RECEBIDA') Group by EC.IDestoque", Conexao, adOpenKeyset, adLockReadOnly
        End If

        If TBFI.EOF = False Then
            Do While TBFI.EOF = False
                .AddItem TBFI!IDEstoque
                TBFI.MoveNext
            Loop
        Else
            If TBAbrir!Devolucao = True Then TextoFiltro = "" Else TextoFiltro = " and EP.Liberado = 'SIM'"
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select EP.IDestoque from Estoque_produtos EP INNER JOIN Estoque_Localarmazenamento_criar EL ON EL.descricao = EP.local_armaz where EP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and EP.Desenho = '" & txtCodigo & "' and EP.Lote is not null " & TextoFiltro & " and EP.Estoque_disponivel > 0 and (Left(EP.status, 7) = 'ENTRADA' or EP.status = 'CONSIGNAÇÃO RECEBIDA') and EL.Estoque = 'False' Group by EP.IDestoque", Conexao, adOpenKeyset, adLockReadOnly
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    .AddItem TBFI!IDEstoque
                    TBFI.MoveNext
                Loop
            End If
        End If
    End If
    TBAbrir.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaRE_NF()
On Error GoTo tratar_erro

With Cmb_RE
    .Clear
    .Enabled = True
    .Locked = False
    .TabStop = True
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select NFP.IDestoque, CFOP.Devolucao from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.Int_codigo = " & Listamaterial.SelectedItem, Conexao, adOpenKeyset, adLockReadOnly
    If TBAbrir.EOF = False Then
        Set TBFI = CreateObject("adodb.recordset")
        
        If chkEmpenhos.Value = False Then
            TBFI.Open "Select EC.IDestoque from Estoque_Controle_Saldo_RE EC where EC.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and EC.Codigo = '" & txtCodigo.Text & "' and EC.Lote is not null and EC.Saldo > 0 and (Left(EC.status, 7) = 'ENTRADA' or EC.status = 'CONSIGNAÇÃO RECEBIDA') Group by EC.IDestoque", Conexao, adOpenKeyset, adLockReadOnly
        Else
        StrSql = "Select EC.IDestoque from (Estoque_Controle_Empenho_Vendas EE INNER JOIN Estoque_Controle_Saldo_RE EC ON EE.ID_estoque = EC.IDEstoque) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = EE.ID_carteira and NFPP.Codinterno = EC.Codigo where NFPP.ID_prod_NF = " & Listamaterial.SelectedItem & " and EC.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and EC.Codigo = '" & txtCodigo.Text & "' and EC.Lote is not null and EC.Saldo > 0 and (Left(EC.status, 7) = 'ENTRADA' or EC.status = 'CONSIGNAÇÃO RECEBIDA') Group by EC.IDestoque"
        'Debug.print StrSql
        
            TBFI.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
        End If

        If TBFI.EOF = False Then
            Do While TBFI.EOF = False
                .AddItem TBFI!IDEstoque
                TBFI.MoveNext
            Loop
        Else
            If TBAbrir!Devolucao = True Then TextoFiltro = "" Else TextoFiltro = " and EP.Liberado = 'SIM'"
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select EP.IDestoque from Estoque_produtos EP INNER JOIN Estoque_Localarmazenamento_criar EL ON EL.descricao = EP.local_armaz where EP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and EP.Desenho = '" & txtCodigo & "' and EP.Lote is not null " & TextoFiltro & " and EP.Estoque_disponivel > 0 and (Left(EP.status, 7) = 'ENTRADA' or EP.status = 'CONSIGNAÇÃO RECEBIDA') and EL.Estoque = 'False' Group by EP.IDestoque", Conexao, adOpenKeyset, adLockReadOnly
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    .AddItem TBFI!IDEstoque
                    TBFI.MoveNext
                Loop
            End If
        End If
    End If
    TBAbrir.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLote_NF()
On Error GoTo tratar_erro

With cmb_Lote
    .Clear
    .Locked = False
    .TabStop = True
    Set TBAbrir = CreateObject("adodb.recordset")
    StrSql = "Select NFP.IDestoque, CFOP.Devolucao from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.Int_codigo = " & Listamaterial.SelectedItem & ""
    'Debug.print StrSql
    
    TBAbrir.Open "Select NFP.IDestoque, CFOP.Devolucao from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.Int_codigo = " & Listamaterial.SelectedItem, Conexao, adOpenKeyset, adLockReadOnly
    If TBAbrir.EOF = False Then
        Set TBFI = CreateObject("adodb.recordset")
        
        If chkEmpenhos.Value = False Then
            TBFI.Open "Select EC.IDestoque,EC.Lote from Estoque_Controle EC where EC.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and EC.desenho = '" & txtCodigo.Text & "' and EC.Lote is not null and EC.estoque_real > 0 and (Left(EC.status, 7) = 'ENTRADA' or EC.status = 'CONSIGNAÇÃO RECEBIDA') Group by EC.IDestoque, EC.Lote", Conexao, adOpenKeyset, adLockReadOnly
        Else
            StrSql = "Select EC.IDestoque,EC.lote from (Estoque_Controle_Empenho_Vendas EE INNER JOIN estoque_controle EC ON EE.ID_estoque = EC.IDEstoque) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = EE.ID_carteira and NFPP.Codinterno = EC.Desenho where NFPP.ID_prod_NF = " & Listamaterial.SelectedItem & " and EC.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and EC.desenho = '" & txtCodigo.Text & "' and EC.Lote is not null and EC.estoque_real > 0 and (Left(EC.status, 7) = 'ENTRADA' or EC.status = 'CONSIGNAÇÃO RECEBIDA') Group by EC.IDestoque, EC.Lote"
            'Debug.print StrSql
            TBFI.Open "Select EC.IDestoque,EC.lote from (Estoque_Controle_Empenho_Vendas EE INNER JOIN estoque_controle EC ON EE.ID_estoque = EC.IDEstoque) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = EE.ID_carteira and NFPP.Codinterno = EC.Desenho where NFPP.ID_prod_NF = " & Listamaterial.SelectedItem & " and EC.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and EC.desenho = '" & txtCodigo.Text & "' and EC.Lote is not null and EC.estoque_real > 0 and (Left(EC.status, 7) = 'ENTRADA' or EC.status = 'CONSIGNAÇÃO RECEBIDA') Group by EC.IDestoque, EC.Lote", Conexao, adOpenKeyset, adLockReadOnly
        End If

        If TBFI.EOF = False Then
            Do While TBFI.EOF = False
                .AddItem TBFI!IDEstoque
                TBFI.MoveNext
            Loop
        Else
            If TBAbrir!Devolucao = True Then TextoFiltro = "" Else TextoFiltro = " and EP.Liberado = 'SIM'"
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select EP.IDestoque, EP.Lote from Estoque_produtos EP INNER JOIN Estoque_Localarmazenamento_criar EL ON EL.descricao = EP.local_armaz where EP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and EP.Desenho = '" & txtCodigo & "' and EP.Lote is not null " & TextoFiltro & " and EP.Estoque_disponivel > 0 and (Left(EP.status, 7) = 'ENTRADA' or EP.status = 'CONSIGNAÇÃO RECEBIDA') and EL.Estoque = 'False' Group by EP.IDestoque, EP.Lote Order By EP.IDestoque", Conexao, adOpenKeyset, adLockReadOnly
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    .AddItem TBFI!LOTE
                    TBFI.MoveNext
                Loop
            End If
        End If
    End If
    TBAbrir.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_serieNF_Change()
On Error GoTo tratar_erro

'If Txt_serieNF <> "" Then
'    VerifNumero = Txt_serieNF
'    ProcVerificaNumero
'    If VerifNumero = False Then
'        Txt_serieNF = ""
'        Txt_serieNF.SetFocus
'        Exit Sub
'    End If
'End If
ProcLimpaCampos
Listamaterial.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCodInterno_RE_Change()
On Error GoTo tratar_erro

ProcCarregaProduto_RE

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub txtdata_Change()
On Error GoTo tratar_erro

ProcCarregaComboRE

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDisponivel_RE_Change()
On Error GoTo tratar_erro

ProcCalculaEstoqueAtualizadoRE

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub




Private Sub txtestoqueatual_Change()
On Error GoTo tratar_erro

'ProcCalculaEstoqueAtualizado

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txt_Notafiscal_LostFocus()
On Error GoTo tratar_erro

If IsNumeric(txt_Notafiscal.Text) = True Then txt_Notafiscal = FunTamanhoTextoZeroEsq(txt_Notafiscal, 9)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Saida_RE_PC_Change()
On Error GoTo tratar_erro

If txtQtde_Saida_RE_PC.Text <> "" Then
    VerifNumero = txtQtde_Saida_RE_PC.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde_Saida_RE_PC.Text = ""
        txtQtde_Saida_RE_PC.SetFocus
        Exit Sub
    End If
    If txtQtde_Saida_RE_PC.Locked = False Then txtQtde_Saida_RE = Format(FunCalculaQtdePCKG(txtDisponivel_RE, txtDisponivel_PC_RE, txtQtde_Saida_RE_PC, False), "###,##0.0000")
    ProcCalculaEstoqueAtualizadoRE
Else
    txtQtde_Saida = ""
    
    txtAtualizado_RE = Format(txtDisponivel_RE, "###,##0.0000")
    txtAtualizado_PC_RE = Format(txtDisponivel_PC_RE, "###,##0.0000")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Saida_RE_PC_GotFocus()
On Error GoTo tratar_erro

FunGotFocus txtQtde_Saida_RE_PC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Saida_RE_PC_LostFocus()
On Error GoTo tratar_erro

txtQtde_Saida_RE_PC = IIf(txtQtde_Saida_RE_PC = "", "0,0000", Format(txtQtde_Saida_RE_PC, "###,##0.0000"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Saida_RE_Change()
On Error GoTo tratar_erro

If txtQtde_Saida_RE.Text <> "" Then
    VerifNumero = txtQtde_Saida_RE.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde_Saida_RE.Text = ""
        txtQtde_Saida_RE.SetFocus
        Exit Sub
    End If
    If txtQtde_Saida_RE.Locked = False Then txtQtde_Saida_RE_PC = Format(FunCalculaQtdePCKG(txtEstoque_Real_RE, txtEstoque_Real_RE_PC, txtQtde_Saida_RE, True), "###,##0.0000")
    ProcCalculaEstoqueAtualizadoRE
Else
    txtQtde_Saida_RE_PC = ""
    
    txtAtualizado_RE = Format(txtDisponivel_RE, "###,##0.0000")
    txtAtualizado_PC_RE = Format(txtDisponivel_PC_RE, "###,##0.0000")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Saida_RE_GotFocus()
On Error GoTo tratar_erro

FunGotFocus txtQtde_Saida_RE

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Saida_RE_LostFocus()
On Error GoTo tratar_erro

txtQtde_Saida_RE = IIf(txtQtde_Saida_RE = "", "0,0000", Format(txtQtde_Saida_RE, "###,##0.0000"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtquantretirado_Change()
On Error GoTo tratar_erro

If txtquantretirado.Text <> "" Then
    VerifNumero = txtquantretirado.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtquantretirado.Text = ""
        txtquantretirado.SetFocus
        Exit Sub
    End If
'    If txtquantretirado.Locked = False Then txtquantretirado_PC = Format(FunCalculaQtdePCKG(txtestoqueatual, txtestoqueatual_PC, txtquantretirado, True), "###,##0.0000")
   ' ProcCalculaEstoqueAtualizado
Else
    'txtquantretirado_PC = ""
    
    txtestoquereal.Text = Format(txtestoqueatual, "###,##0.0000")
    txtestoquereal_PC = Format(txtestoqueatual_PC, "###,##0.0000")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcCalculaEstoqueAtualizadoRE()
On Error GoTo tratar_erro

QuantSolicitado = txtQtde_Saida_RE
QuantsolicitadoN1 = txtQtde_Saida_RE_PC
EstoqueAtual = IIf(txtDisponivel_RE = "", 0, txtDisponivel_RE)
EstoqueAtualPC = IIf(txtDisponivel_PC_RE = "", 0, txtDisponivel_PC_RE)
Qtde = IIf(txtEmpenhos = "", 0, txtEmpenhos)
Qtd = IIf(txtEmpenho_PC = "", 0, txtEmpenho_PC)

'Verifica se a ordem esta empenhada
If ListaMaterial_RE.ListItems.Count <> 0 And txtRE <> "" Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Sum(PNFC.Quantidade - PNFC.Qtde_saida) as Valor, Sum(ISNULL(PNFC.Quantidade_PC, 0) - ISNULL(PNFC.Qtde_saida_PC, 0)) as Valor1 from (Producao_NF_Consignada PNFC INNER JOIN Producaomaterial PM ON PM.Ordem = PNFC.Ordem and PM.Codigo = PNFC.Codinterno) INNER JOIN Producao P ON P.Ordem = PNFC.Ordem where PNFC.IDestoque = " & txtRE & " and PNFC.Ordem = " & ListaMaterial_RE.SelectedItem.ListSubItems(1) & " and PNFC.Quantidade - PNFC.Qtde_saida > 0 and P.Status <> 'Cancelada' and P.Concluida = 0 and (PM.Saida = 'NÃO' OR PM.Saida = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        If QuantSolicitado <> 0 Then
            qt = IIf(IsNull(TBFI!valor), 0, TBFI!valor)
            QuantSolicitado = IIf(QuantSolicitado - qt < 0, 0, QuantSolicitado - qt)
        End If
        If QuantsolicitadoN1 <> 0 Then
            qt = IIf(IsNull(TBFI!Valor1), 0, TBFI!Valor1)
            QuantsolicitadoN1 = IIf(QuantsolicitadoN1 - qt < 0, 0, QuantsolicitadoN1 - qt)
        End If
    End If
    TBFI.Close
End If

txtAtualizado_RE.Text = Format(EstoqueAtual - QuantSolicitado, "###,##0.0000")
txtAtualizado_PC_RE = Format(EstoqueAtualPC - QuantsolicitadoN1, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAtualizaStatus_RM()
On Error GoTo tratar_erro

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select * from Requisicao_materiais where requisicao = '" & txt_RM & "' ", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    Set TBMateriaprima = CreateObject("adodb.recordset")
    TBMateriaprima.Open "Select * from Requisicao_materiais_lista where idrequisicao = " & TBproducao!ID & " and status <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBMateriaprima.EOF = False Then
        Set TBMateriaprima = CreateObject("adodb.recordset")
        TBMateriaprima.Open "Select * from Requisicao_materiais_lista where idrequisicao = " & TBproducao!ID & " and status <> 'REQUISIT.'", Conexao, adOpenKeyset, adLockOptimistic
        If TBMateriaprima.EOF = True Then
            TBproducao!status = "ABERTA"
        Else
            Set TBMateriaprima = CreateObject("adodb.recordset")
            TBMateriaprima.Open "Select * from Requisicao_materiais_lista where idrequisicao = " & TBproducao!ID & " and status <> 'RETIRADO'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMateriaprima.EOF = True Then
                TBproducao!status = "RETIRADA"
            Else
                TBproducao!status = "PARCIAL"
            End If
        End If
    End If
    TBMateriaprima.Close
    TBproducao.Update
End If
TBproducao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtquantretirado_GotFocus()
On Error GoTo tratar_erro

FunGotFocus txtquantretirado

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtquantretirado_LostFocus()
On Error GoTo tratar_erro

txtquantretirado = IIf(txtquantretirado = "", "0,0000", Format(txtquantretirado, "###,##0.0000"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

'Private Sub txtquantretirado_PC_Change()
'On Error GoTo tratar_erro
'
'If txtquantretirado_PC.Text <> "" Then
'    VerifNumero = txtquantretirado_PC.Text
'    ProcVerificaNumero
'    If VerifNumero = False Then
'        txtquantretirado_PC.Text = ""
'        txtquantretirado_PC.SetFocus
'        Exit Sub
'    End If
'    If txtquantretirado_PC.Locked = False Then txtquantretirado = Format(FunCalculaQtdePCKG(txtestoqueatual, txtestoqueatual_PC, txtquantretirado_PC, False), "###,##0.0000")
'    ProcCalculaEstoqueAtualizado
'Else
'    txtquantretirado = ""
'
'    txtestoquereal.Text = Format(txtestoqueatual, "###,##0.0000")
'    txtestoquereal_PC = Format(txtestoqueatual_PC, "###,##0.0000")
'End If
'
'Exit Sub
'tratar_erro:
'    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub
'
'Private Sub txtquantretirado_PC_GotFocus()
'On Error GoTo tratar_erro
'
'FunGotFocus txtquantretirado_PC
'
'Exit Sub
'tratar_erro:
'    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub
'
'Private Sub txtquantretirado_PC_LostFocus()
'On Error GoTo tratar_erro
'
'txtquantretirado_PC = IIf(txtquantretirado_PC = "", "0,0000", Format(txtquantretirado_PC, "###,##0.0000"))
'
'Exit Sub
'tratar_erro:
'    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub

Private Sub txtRE_Change()
On Error GoTo tratar_erro

ProcLimpaCampos_TabRE
ListaMaterial_RE.ListItems.Clear

If txtRE <> "" Then
    VerifNumero = txtRE
    ProcVerificaNumero
    If VerifNumero = False Then
        txtRE = ""
        txtRE.SetFocus
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
ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
Select Case ButtonIndex
    Case 1: ProcRetirar
    Case 2:
    If txt_Notafiscal <> "" Then
    frmEstoque_Saldos.Show 1
    Else
    ProcRetirarSelecionados
    End If
    Case 3: ProcCancelarReq
    'Case 5: ProcAjuda
    Case 6: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcRetirarRE
    Case 2: ProcRetirarSelecionadosRE
    'Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
