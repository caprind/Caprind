VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMDI_AvisoDiario 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Avisos diário (Lembrete)"
   ClientHeight    =   10005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   4  'Mask Not Pen
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMDI_AvisoDiario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10005
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
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
      FormHeightDT    =   10470
      FormWidthDT     =   15240
      FormScaleHeightDT=   10005
      FormScaleWidthDT=   15120
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9315
      Left            =   60
      TabIndex        =   76
      Top             =   1020
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   16431
      _Version        =   393216
      Tabs            =   18
      Tab             =   13
      TabsPerRow      =   18
      TabHeight       =   706
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Contas a pagar"
      TabPicture(0)   =   "FrmMDI_AvisoDiario.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "PBLista"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lista"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Contas a receber"
      TabPicture(1)   =   "FrmMDI_AvisoDiario.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Lista1"
      Tab(1).Control(2)=   "PBLista1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Solicitação para aprovar"
      TabPicture(2)   =   "FrmMDI_AvisoDiario.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Lista2"
      Tab(2).Control(2)=   "PBLista2"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Manutenção"
      TabPicture(3)   =   "FrmMDI_AvisoDiario.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(1)=   "Lista3"
      Tab(3).Control(2)=   "PBLista3"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "RM"
      TabPicture(4)   =   "FrmMDI_AvisoDiario.frx":0D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).Control(1)=   "Lista4"
      Tab(4).Control(2)=   "PBLista4"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Necessidade de compras"
      TabPicture(5)   =   "FrmMDI_AvisoDiario.frx":0D56
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).Control(1)=   "Lista5"
      Tab(5).Control(2)=   "PBLista5"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Necessidade de produção"
      TabPicture(6)   =   "FrmMDI_AvisoDiario.frx":0D72
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame7"
      Tab(6).Control(1)=   "Lista6"
      Tab(6).Control(2)=   "PBLista6"
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Necessidade de estoque"
      TabPicture(7)   =   "FrmMDI_AvisoDiario.frx":0D8E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame8"
      Tab(7).Control(1)=   "Lista7"
      Tab(7).Control(2)=   "PBLista7"
      Tab(7).ControlCount=   3
      TabCaption(8)   =   "Faturar"
      TabPicture(8)   =   "FrmMDI_AvisoDiario.frx":0DAA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame10"
      Tab(8).Control(1)=   "FrameObs"
      Tab(8).Control(2)=   "PBLista8"
      Tab(8).Control(3)=   "Lista8"
      Tab(8).Control(4)=   "Frame9"
      Tab(8).ControlCount=   5
      TabCaption(9)   =   "OS's em atraso"
      TabPicture(9)   =   "FrmMDI_AvisoDiario.frx":0DC6
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "PBLista9"
      Tab(9).Control(1)=   "ImageList1"
      Tab(9).Control(2)=   "FlexGridOS"
      Tab(9).ControlCount=   3
      TabCaption(10)  =   "Centro de custo"
      TabPicture(10)  =   "FrmMDI_AvisoDiario.frx":0DE2
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "ImageList2"
      Tab(10).Control(1)=   "FlexGridCC"
      Tab(10).Control(2)=   "PBLista10"
      Tab(10).ControlCount=   3
      TabCaption(11)  =   "Análise crítica"
      TabPicture(11)  =   "FrmMDI_AvisoDiario.frx":0DFE
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Frame12"
      Tab(11).Control(1)=   "Tab_Analise"
      Tab(11).Control(2)=   "PBLista11"
      Tab(11).ControlCount=   3
      TabCaption(12)  =   "Terceiros"
      TabPicture(12)  =   "FrmMDI_AvisoDiario.frx":0E1A
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "frameOBS_Terceiros"
      Tab(12).Control(1)=   "Lista_Terceiros"
      Tab(12).Control(2)=   "PBLista12"
      Tab(12).Control(3)=   "Frame13"
      Tab(12).ControlCount=   4
      TabCaption(13)  =   "Pedidos em atraso"
      TabPicture(13)  =   "FrmMDI_AvisoDiario.frx":0E36
      Tab(13).ControlEnabled=   -1  'True
      Tab(13).Control(0)=   "PBLista13"
      Tab(13).Control(0).Enabled=   0   'False
      Tab(13).Control(1)=   "Lista_pedidosAtraso"
      Tab(13).Control(1).Enabled=   0   'False
      Tab(13).Control(2)=   "Frame14"
      Tab(13).Control(2).Enabled=   0   'False
      Tab(13).ControlCount=   3
      TabCaption(14)  =   "Calibração á vencer"
      TabPicture(14)  =   "FrmMDI_AvisoDiario.frx":0E52
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "Lista_instrumentos"
      Tab(14).Control(0).Enabled=   0   'False
      Tab(14).Control(1)=   "PBLista_instrumentos"
      Tab(14).Control(1).Enabled=   0   'False
      Tab(14).Control(2)=   "Frame15"
      Tab(14).Control(2).Enabled=   0   'False
      Tab(14).ControlCount=   3
      TabCaption(15)  =   "Não conforme"
      TabPicture(15)  =   "FrmMDI_AvisoDiario.frx":0E6E
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "Lista_NaoConforme"
      Tab(15).Control(0).Enabled=   0   'False
      Tab(15).Control(1)=   "PBLista_NaoConforme"
      Tab(15).Control(1).Enabled=   0   'False
      Tab(15).Control(2)=   "Frame16"
      Tab(15).Control(2).Enabled=   0   'False
      Tab(15).ControlCount=   3
      TabCaption(16)  =   "Produtos a vencer"
      TabPicture(16)  =   "FrmMDI_AvisoDiario.frx":0E8A
      Tab(16).ControlEnabled=   0   'False
      Tab(16).Control(0)=   "ListaProdVencer"
      Tab(16).Control(0).Enabled=   0   'False
      Tab(16).ControlCount=   1
      TabCaption(17)  =   "Processos Sugestões"
      TabPicture(17)  =   "FrmMDI_AvisoDiario.frx":0EA6
      Tab(17).ControlEnabled=   0   'False
      Tab(17).Control(0)=   "ListaSugestoes"
      Tab(17).Control(0).Enabled=   0   'False
      Tab(17).ControlCount=   1
      Begin VB.Frame Frame16 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74970
         TabIndex        =   195
         Top             =   8370
         Width           =   14985
         Begin VB.TextBox txtPagIr_NaoConforme 
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
            TabIndex        =   197
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtNreg_NaoConforme 
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
            Left            =   3750
            TabIndex        =   196
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx_NaoConforme 
            Height          =   315
            Left            =   11760
            TabIndex        =   198
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":0EC2
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
         Begin DrawSuite2022.USButton cmdPagAnt_NaoConforme 
            Height          =   315
            Left            =   11220
            TabIndex        =   199
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":4666
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
         Begin DrawSuite2022.USButton cmdPagIr_NaoConforme 
            Height          =   315
            Left            =   10110
            TabIndex        =   200
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
         Begin DrawSuite2022.USButton cmdPagPrim_NaoConforme 
            Height          =   315
            Left            =   10680
            TabIndex        =   201
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":816F
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
         Begin DrawSuite2022.USButton cmdPagUlt_NaoConforme 
            Height          =   315
            Left            =   12300
            TabIndex        =   202
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":C25E
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
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3060
            TabIndex        =   210
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblPaginas_NaoConforme 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   205
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros_NaoConforme 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   204
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4380
            TabIndex        =   203
            Top             =   240
            Width           =   1440
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74970
         TabIndex        =   183
         Top             =   8370
         Width           =   14985
         Begin VB.TextBox txtPagIr_Instrumentos 
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
            TabIndex        =   185
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtNreg_Instrumentos 
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
            Left            =   3750
            TabIndex        =   184
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx_Instrumentos 
            Height          =   315
            Left            =   11760
            TabIndex        =   186
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":FAEA
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
         Begin DrawSuite2022.USButton cmdPagAnt_Instrumentos 
            Height          =   315
            Left            =   11220
            TabIndex        =   187
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":1328E
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
         Begin DrawSuite2022.USButton cmdPagIr_Instrumentos 
            Height          =   315
            Left            =   10110
            TabIndex        =   188
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
         Begin DrawSuite2022.USButton cmdPagPrim_Instrumentos 
            Height          =   315
            Left            =   10680
            TabIndex        =   189
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":16D97
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
         Begin DrawSuite2022.USButton cmdPagUlt_Instrumentos 
            Height          =   315
            Left            =   12300
            TabIndex        =   190
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":1AE86
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
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4380
            TabIndex        =   209
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lblPaginas_Instrumentos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   193
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros_instrumentos 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   192
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3060
            TabIndex        =   191
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   30
         TabIndex        =   172
         Top             =   8370
         Width           =   14985
         Begin VB.TextBox txtPagIr11 
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
            TabIndex        =   174
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtNreg11 
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
            Left            =   3750
            TabIndex        =   173
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx11 
            Height          =   315
            Left            =   11760
            TabIndex        =   175
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":1E712
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
         Begin DrawSuite2022.USButton cmdPagAnt11 
            Height          =   315
            Left            =   11220
            TabIndex        =   176
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":21EB6
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
         Begin DrawSuite2022.USButton cmdPagIr11 
            Height          =   315
            Left            =   10110
            TabIndex        =   177
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
         Begin DrawSuite2022.USButton cmdPagPrim11 
            Height          =   315
            Left            =   10680
            TabIndex        =   178
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":259BF
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
         Begin DrawSuite2022.USButton cmdPagUlt11 
            Height          =   315
            Left            =   12300
            TabIndex        =   179
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":29AAE
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
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4380
            TabIndex        =   222
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lblPaginas11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   182
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   181
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3060
            TabIndex        =   180
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Frame frameOBS_Terceiros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Observações"
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
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   -74970
         TabIndex        =   169
         Top             =   7890
         Width           =   14985
         Begin VB.TextBox txtObs_Terceiros 
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
            Height          =   675
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   156
            ToolTipText     =   "Observação."
            Top             =   300
            Width           =   13905
         End
         Begin DrawSuite2022.USButton cmdSalvar_Terceiros 
            Height          =   675
            Left            =   14130
            TabIndex        =   157
            ToolTipText     =   "Salvar (F3)"
            Top             =   300
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   1191
            DibPicture      =   "FrmMDI_AvisoDiario.frx":2D33A
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
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74970
         TabIndex        =   138
         Top             =   8370
         Width           =   14985
         Begin VB.TextBox txtPagIr9 
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
            TabIndex        =   140
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtNreg9 
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
            Left            =   3750
            TabIndex        =   139
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx9 
            Height          =   315
            Left            =   11760
            TabIndex        =   141
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":2DACC
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
         Begin DrawSuite2022.USButton cmdPagAnt9 
            Height          =   315
            Left            =   11220
            TabIndex        =   142
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":31270
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
         Begin DrawSuite2022.USButton cmdPagIr9 
            Height          =   315
            Left            =   10110
            TabIndex        =   143
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
         Begin DrawSuite2022.USButton cmdPagPrim9 
            Height          =   315
            Left            =   10680
            TabIndex        =   144
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":34D79
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
         Begin DrawSuite2022.USButton cmdPagUlt9 
            Height          =   315
            Left            =   12300
            TabIndex        =   145
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":38E68
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
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4380
            TabIndex        =   220
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lblPaginas9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   158
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   147
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3060
            TabIndex        =   146
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produtos/Serviços"
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
         Height          =   2175
         Left            =   -74970
         TabIndex        =   127
         Top             =   4800
         Width           =   14985
         Begin MSComctlLib.ListView ListaItensNota 
            Height          =   1755
            Left            =   180
            TabIndex        =   65
            Top             =   300
            Width           =   14625
            _ExtentX        =   25797
            _ExtentY        =   3096
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
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
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Código interno"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   13944
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Un."
               Object.Width           =   794
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Quantidade"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Pedido"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Rev."
               Object.Width           =   970
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Object.Tag             =   "D"
               Text            =   "Prazo final"
               Object.Width           =   2117
            EndProperty
         End
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   -68730
         Top             =   3930
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   21
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMDI_AvisoDiario.frx":3C6F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMDI_AvisoDiario.frx":3CD46
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista9 
         Height          =   255
         Left            =   -74970
         TabIndex        =   124
         Top             =   8730
         Width           =   14985
         _ExtentX        =   26432
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -67620
         Top             =   2220
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   21
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMDI_AvisoDiario.frx":3D608
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMDI_AvisoDiario.frx":3DC5A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin DrawSuite2022.USFlexGrid FlexGridOS 
         Height          =   8265
         Left            =   -74970
         TabIndex        =   123
         Top             =   450
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   14579
         BackColorEvenRows=   14737632
         BackColorSelected1=   16643298
         BackColorSelected2=   16643298
         FocusRectColor  =   15181413
         GridColor       =   16247519
         HeaderGradientColor2=   12632256
         ProgressBarColor2=   2277891
         ForeColorSelected=   0
         AllowColumnResizing=   -1  'True
         CaptionHeight   =   28
         ColumnHeaderSmall=   -1  'True
         FocusRowHighlightKeepTextForeColor=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderFormatString=   $"FrmMDI_AvisoDiario.frx":3E51C
         MinRowHeight    =   14
         ScrollBars      =   1
         TotalLineShow   =   0   'False
      End
      Begin VB.Frame FrameObs 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Observações"
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
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   -74970
         TabIndex        =   122
         Top             =   7890
         Width           =   14985
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
            Height          =   675
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   73
            ToolTipText     =   "Observação."
            Top             =   300
            Width           =   13905
         End
         Begin DrawSuite2022.USButton cmdSalvar 
            Height          =   675
            Left            =   14130
            TabIndex        =   74
            ToolTipText     =   "Salvar (F3)"
            Top             =   300
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   1191
            DibPicture      =   "FrmMDI_AvisoDiario.frx":3E647
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
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74970
         TabIndex        =   112
         Top             =   8370
         Width           =   14985
         Begin VB.TextBox txtNreg7 
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
            Left            =   3750
            TabIndex        =   57
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtPagIr7 
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
            TabIndex        =   58
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx7 
            Height          =   315
            Left            =   11760
            TabIndex        =   62
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":3EDD9
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
         Begin DrawSuite2022.USButton cmdPagAnt7 
            Height          =   315
            Left            =   11220
            TabIndex        =   61
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":4257D
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
         Begin DrawSuite2022.USButton cmdPagIr7 
            Height          =   315
            Left            =   10110
            TabIndex        =   59
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
         Begin DrawSuite2022.USButton cmdPagPrim7 
            Height          =   315
            Left            =   10680
            TabIndex        =   60
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":46086
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
         Begin DrawSuite2022.USButton cmdPagUlt7 
            Height          =   315
            Left            =   12300
            TabIndex        =   63
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":4A175
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
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4380
            TabIndex        =   218
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3060
            TabIndex        =   115
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblRegistros7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   114
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblPaginas7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   113
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74970
         TabIndex        =   107
         Top             =   8370
         Width           =   14985
         Begin VB.TextBox txtPagIr6 
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
            TabIndex        =   50
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtNreg6 
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
            Left            =   3750
            TabIndex        =   49
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx6 
            Height          =   315
            Left            =   11760
            TabIndex        =   54
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":4DA01
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
         Begin DrawSuite2022.USButton cmdPagAnt6 
            Height          =   315
            Left            =   11220
            TabIndex        =   53
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":511A5
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
         Begin DrawSuite2022.USButton cmdPagIr6 
            Height          =   315
            Left            =   10110
            TabIndex        =   51
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
         Begin DrawSuite2022.USButton cmdPagPrim6 
            Height          =   315
            Left            =   10680
            TabIndex        =   52
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":54CAE
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
         Begin DrawSuite2022.USButton cmdPagUlt6 
            Height          =   315
            Left            =   12300
            TabIndex        =   55
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":58D9D
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
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4380
            TabIndex        =   217
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lblPaginas6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   110
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   109
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3060
            TabIndex        =   108
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74970
         TabIndex        =   102
         Top             =   8370
         Width           =   14985
         Begin VB.TextBox txtNreg5 
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
            Left            =   3750
            TabIndex        =   41
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtPagIr5 
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
            TabIndex        =   42
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx5 
            Height          =   315
            Left            =   11760
            TabIndex        =   46
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":5C629
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
         Begin DrawSuite2022.USButton cmdPagAnt5 
            Height          =   315
            Left            =   11220
            TabIndex        =   45
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":5FDCD
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
         Begin DrawSuite2022.USButton cmdPagIr5 
            Height          =   315
            Left            =   10110
            TabIndex        =   43
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
         Begin DrawSuite2022.USButton cmdPagPrim5 
            Height          =   315
            Left            =   10680
            TabIndex        =   44
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":638D6
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
         Begin DrawSuite2022.USButton cmdPagUlt5 
            Height          =   315
            Left            =   12300
            TabIndex        =   47
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":679C5
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
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4380
            TabIndex        =   216
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3060
            TabIndex        =   105
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblRegistros5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   104
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblPaginas5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   103
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74970
         TabIndex        =   97
         Top             =   8370
         Width           =   14985
         Begin VB.TextBox txtPagIr4 
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
            TabIndex        =   34
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtNreg4 
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
            Left            =   3750
            TabIndex        =   33
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx4 
            Height          =   315
            Left            =   11760
            TabIndex        =   38
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":6B251
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
         Begin DrawSuite2022.USButton cmdPagAnt4 
            Height          =   315
            Left            =   11220
            TabIndex        =   37
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":6E9F5
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
         Begin DrawSuite2022.USButton cmdPagIr4 
            Height          =   315
            Left            =   10110
            TabIndex        =   35
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
         Begin DrawSuite2022.USButton cmdPagPrim4 
            Height          =   315
            Left            =   10680
            TabIndex        =   36
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":724FE
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
         Begin DrawSuite2022.USButton cmdPagUlt4 
            Height          =   315
            Left            =   12300
            TabIndex        =   39
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":765ED
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
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4380
            TabIndex        =   215
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lblPaginas4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   100
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   99
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3060
            TabIndex        =   98
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74970
         TabIndex        =   92
         Top             =   8370
         Width           =   14985
         Begin VB.TextBox txtNreg3 
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
            Left            =   3750
            TabIndex        =   25
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtPagIr3 
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
            TabIndex        =   26
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx3 
            Height          =   315
            Left            =   11760
            TabIndex        =   30
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":79E79
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
         Begin DrawSuite2022.USButton cmdPagAnt3 
            Height          =   315
            Left            =   11220
            TabIndex        =   29
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":7D61D
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
         Begin DrawSuite2022.USButton cmdPagIr3 
            Height          =   315
            Left            =   10110
            TabIndex        =   27
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
         Begin DrawSuite2022.USButton cmdPagPrim3 
            Height          =   315
            Left            =   10680
            TabIndex        =   28
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":81126
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
         Begin DrawSuite2022.USButton cmdPagUlt3 
            Height          =   315
            Left            =   12300
            TabIndex        =   31
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":85215
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
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4380
            TabIndex        =   214
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   3060
            TabIndex        =   95
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblRegistros3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   94
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblPaginas3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   93
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74970
         TabIndex        =   87
         Top             =   8370
         Width           =   14985
         Begin VB.TextBox txtPagIr2 
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
            TabIndex        =   18
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtNreg2 
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
            Left            =   3750
            TabIndex        =   17
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx2 
            Height          =   315
            Left            =   11760
            TabIndex        =   22
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":88AA1
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
         Begin DrawSuite2022.USButton cmdPagAnt2 
            Height          =   315
            Left            =   11220
            TabIndex        =   21
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":8C245
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
         Begin DrawSuite2022.USButton cmdPagIr2 
            Height          =   315
            Left            =   10110
            TabIndex        =   19
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
         Begin DrawSuite2022.USButton cmdPagPrim2 
            Height          =   315
            Left            =   10680
            TabIndex        =   20
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":8FD4E
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
         Begin DrawSuite2022.USButton cmdPagUlt2 
            Height          =   315
            Left            =   12300
            TabIndex        =   23
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":93E3D
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
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4380
            TabIndex        =   213
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lblPaginas2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   90
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   89
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   3060
            TabIndex        =   88
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74970
         TabIndex        =   82
         Top             =   8370
         Width           =   14985
         Begin VB.TextBox txtNreg1 
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
            Left            =   3750
            TabIndex        =   9
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtPagIr1 
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
            TabIndex        =   10
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx1 
            Height          =   315
            Left            =   11760
            TabIndex        =   14
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":976C9
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
         Begin DrawSuite2022.USButton cmdPagAnt1 
            Height          =   315
            Left            =   11220
            TabIndex        =   13
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":9AE6D
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
         Begin DrawSuite2022.USButton cmdPagIr1 
            Height          =   315
            Left            =   10110
            TabIndex        =   11
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
         Begin DrawSuite2022.USButton cmdPagPrim1 
            Height          =   315
            Left            =   10680
            TabIndex        =   12
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":9E976
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
         Begin DrawSuite2022.USButton cmdPagUlt1 
            Height          =   315
            Left            =   12300
            TabIndex        =   15
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":A2A65
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
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4380
            TabIndex        =   212
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3060
            TabIndex        =   85
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblRegistros1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   84
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblPaginas1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   83
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74970
         TabIndex        =   77
         Top             =   8370
         Width           =   14985
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
            TabIndex        =   2
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
            Left            =   3750
            TabIndex        =   1
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   6
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":A62F1
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
            TabIndex        =   5
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":A9A95
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
            TabIndex        =   3
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
            TabIndex        =   4
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":AD59E
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
            TabIndex        =   7
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":B168D
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
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4380
            TabIndex        =   211
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lblPaginas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   80
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   79
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3060
            TabIndex        =   78
            Top             =   240
            Width           =   645
         End
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   7635
         Left            =   -74970
         TabIndex        =   0
         Top             =   450
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   13467
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "D"
            Text            =   "Vencimento"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Fornecedor"
            Object.Width           =   21511
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   -74970
         TabIndex        =   81
         Top             =   8100
         Width           =   14985
         _ExtentX        =   26432
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
      Begin MSComctlLib.ListView Lista1 
         Height          =   7635
         Left            =   -74970
         TabIndex        =   8
         Top             =   450
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   13467
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "D"
            Text            =   "Vencimento"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Cliente"
            Object.Width           =   21511
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista1 
         Height          =   255
         Left            =   -74970
         TabIndex        =   86
         Top             =   8100
         Width           =   14985
         _ExtentX        =   26432
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
      Begin MSComctlLib.ListView Lista2 
         Height          =   7635
         Left            =   -74970
         TabIndex        =   16
         Top             =   450
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   13467
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Nº solicitação"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Responsável"
            Object.Width           =   20452
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista2 
         Height          =   255
         Left            =   -74970
         TabIndex        =   91
         Top             =   8100
         Width           =   14985
         _ExtentX        =   26432
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
      Begin MSComctlLib.ListView Lista3 
         Height          =   7635
         Left            =   -74970
         TabIndex        =   24
         Top             =   450
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   13467
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Tipo"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "D"
            Text            =   "Posto de trab."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   15513
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "D"
            Text            =   "Data próx."
            Object.Width           =   2293
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista3 
         Height          =   255
         Left            =   -74970
         TabIndex        =   96
         Top             =   8100
         Width           =   14985
         _ExtentX        =   26432
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
      Begin MSComctlLib.ListView Lista4 
         Height          =   7635
         Left            =   -74970
         TabIndex        =   32
         Top             =   450
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   13467
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Nº requisição"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Cód. int."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   20452
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista4 
         Height          =   255
         Left            =   -74970
         TabIndex        =   101
         Top             =   8100
         Width           =   14985
         _ExtentX        =   26432
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
      Begin MSComctlLib.ListView Lista5 
         Height          =   7635
         Left            =   -74970
         TabIndex        =   40
         Top             =   450
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   13467
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Cód. int."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   20364
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Necessidade"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Un."
            Object.Width           =   970
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista5 
         Height          =   255
         Left            =   -74970
         TabIndex        =   106
         Top             =   8100
         Width           =   14985
         _ExtentX        =   26432
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
      Begin MSComctlLib.ListView Lista6 
         Height          =   7635
         Left            =   -74970
         TabIndex        =   48
         Top             =   450
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   13467
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Cód. int."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   20364
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Necessidade"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Un."
            Object.Width           =   970
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista6 
         Height          =   255
         Left            =   -74970
         TabIndex        =   111
         Top             =   8100
         Width           =   14985
         _ExtentX        =   26432
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
      Begin MSComctlLib.ListView Lista7 
         Height          =   7635
         Left            =   -74970
         TabIndex        =   56
         Top             =   450
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   13467
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Cód. int."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   20364
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Necessidade"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Un."
            Object.Width           =   970
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista7 
         Height          =   255
         Left            =   -74970
         TabIndex        =   116
         Top             =   8100
         Width           =   14985
         _ExtentX        =   26432
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
      Begin DrawSuite2022.USProgressBar PBLista8 
         Height          =   255
         Left            =   -74970
         TabIndex        =   121
         Top             =   6990
         Width           =   14985
         _ExtentX        =   26432
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
      Begin MSComctlLib.ListView Lista8 
         Height          =   4335
         Left            =   -74970
         TabIndex        =   64
         Top             =   450
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   7646
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
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
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Empresa"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "D"
            Text            =   "Dt. emissão"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Ordem fat."
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
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
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "ID"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Destinatário"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Observação"
            Object.Width           =   5900
         EndProperty
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74970
         TabIndex        =   117
         Top             =   7260
         Width           =   14985
         Begin VB.TextBox txtPagIr8 
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
            TabIndex        =   67
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtNreg8 
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
            Left            =   3750
            TabIndex        =   66
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx8 
            Height          =   315
            Left            =   11760
            TabIndex        =   71
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":B4F19
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
         Begin DrawSuite2022.USButton cmdPagAnt8 
            Height          =   315
            Left            =   11220
            TabIndex        =   70
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":B86BD
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
         Begin DrawSuite2022.USButton cmdPagIr8 
            Height          =   315
            Left            =   10110
            TabIndex        =   68
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
         Begin DrawSuite2022.USButton cmdPagPrim8 
            Height          =   315
            Left            =   10680
            TabIndex        =   69
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":BC1C6
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
         Begin DrawSuite2022.USButton cmdPagUlt8 
            Height          =   315
            Left            =   12300
            TabIndex        =   72
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":C02B5
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
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4380
            TabIndex        =   219
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label lblPaginas8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   120
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblRegistros8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   119
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3060
            TabIndex        =   118
            Top             =   240
            Width           =   645
         End
      End
      Begin DrawSuite2022.USFlexGrid FlexGridCC 
         Height          =   8265
         Left            =   -74970
         TabIndex        =   125
         Top             =   450
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   14579
         BackColorEvenRows=   14737632
         BackColorSelected1=   16643298
         BackColorSelected2=   16643298
         FocusRectColor  =   15181413
         GridColor       =   16247519
         HeaderGradientColor2=   12632256
         ProgressBarColor2=   2277891
         ForeColorSelected=   0
         AllowColumnResizing=   -1  'True
         CaptionHeight   =   28
         ColumnHeaderSmall=   -1  'True
         FocusRowHighlightKeepTextForeColor=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderFormatString=   $"FrmMDI_AvisoDiario.frx":C3B41
         MinRowHeight    =   14
         ScrollBars      =   1
         TotalLineShow   =   0   'False
      End
      Begin DrawSuite2022.USProgressBar PBLista10 
         Height          =   255
         Left            =   -74970
         TabIndex        =   126
         Top             =   8730
         Width           =   14985
         _ExtentX        =   26432
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
      Begin TabDlg.SSTab Tab_Analise 
         Height          =   7635
         Left            =   -74970
         TabIndex        =   128
         Top             =   435
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   13467
         _Version        =   393216
         Tabs            =   5
         Tab             =   2
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
         TabCaption(0)   =   "Engenharia"
         TabPicture(0)   =   "FrmMDI_AvisoDiario.frx":C3BF1
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Lista_Analise_Engenharia"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Processo"
         TabPicture(1)   =   "FrmMDI_AvisoDiario.frx":C3C0D
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Lista_Analise_Processo"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "PCP"
         TabPicture(2)   =   "FrmMDI_AvisoDiario.frx":C3C29
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Lista_Analise_PCP"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Qualidade"
         TabPicture(3)   =   "FrmMDI_AvisoDiario.frx":C3C45
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Lista_Analise_Qualidade"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Compras"
         TabPicture(4)   =   "FrmMDI_AvisoDiario.frx":C3C61
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Lista_Analise_Compras"
         Tab(4).ControlCount=   1
         Begin VB.Frame Frame11 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   600
            Left            =   -74940
            TabIndex        =   129
            Top             =   330
            Width           =   15135
            Begin VB.TextBox txtValorCentro 
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
               Left            =   11295
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   134
               TabStop         =   0   'False
               ToolTipText     =   "Valor."
               Top             =   180
               Width           =   1155
            End
            Begin VB.TextBox txtPercentualCentro 
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
               Left            =   13785
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   133
               TabStop         =   0   'False
               ToolTipText     =   "Percentual."
               Top             =   180
               Width           =   1155
            End
            Begin VB.CheckBox chkValor 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Valor :"
               Height          =   255
               Left            =   10530
               TabIndex        =   132
               Top             =   180
               Width           =   765
            End
            Begin VB.CheckBox chkPercentual 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Percentual :"
               Height          =   255
               Left            =   12600
               TabIndex        =   131
               Top             =   180
               Width           =   1185
            End
            Begin VB.ComboBox Cmb_centro 
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
               Left            =   1500
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   130
               ToolTipText     =   "Centro de custo."
               Top             =   180
               Width           =   8910
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Centro de custo :"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   180
               TabIndex        =   135
               Top             =   180
               Width           =   1260
            End
         End
         Begin MSComctlLib.ListView Lista_custo 
            Height          =   6240
            Left            =   -74940
            TabIndex        =   136
            Top             =   945
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   11007
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Código"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   18600
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Valor"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Percentual"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "ID_CC"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_Analise_Engenharia 
            Height          =   7245
            Left            =   -74970
            TabIndex        =   137
            Top             =   330
            Width           =   14865
            _ExtentX        =   26220
            _ExtentY        =   12779
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
               Object.Tag             =   "N"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Nº análise"
               Object.Width           =   2117
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
               Text            =   "Cód. interno"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Cliente"
               Object.Width           =   12656
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_Analise_Processo 
            Height          =   7245
            Left            =   -74970
            TabIndex        =   160
            Top             =   330
            Width           =   14865
            _ExtentX        =   26220
            _ExtentY        =   12779
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
               Object.Tag             =   "N"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Nº análise"
               Object.Width           =   2117
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
               Text            =   "Cód. interno"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Cliente"
               Object.Width           =   12656
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_Analise_PCP 
            Height          =   7245
            Left            =   30
            TabIndex        =   161
            Top             =   330
            Width           =   14865
            _ExtentX        =   26220
            _ExtentY        =   12779
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
               Object.Tag             =   "N"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Nº análise"
               Object.Width           =   2117
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
               Text            =   "Cód. interno"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Cliente"
               Object.Width           =   12656
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_Analise_Qualidade 
            Height          =   7245
            Left            =   -74970
            TabIndex        =   162
            Top             =   330
            Width           =   14865
            _ExtentX        =   26220
            _ExtentY        =   12779
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
               Object.Tag             =   "N"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Nº análise"
               Object.Width           =   2117
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
               Text            =   "Cód. interno"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Cliente"
               Object.Width           =   12656
            EndProperty
         End
         Begin MSComctlLib.ListView Lista_Analise_Compras 
            Height          =   7245
            Left            =   -74970
            TabIndex        =   163
            Top             =   330
            Width           =   14865
            _ExtentX        =   26220
            _ExtentY        =   12779
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
               Object.Tag             =   "N"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Nº análise"
               Object.Width           =   2117
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
               Text            =   "Cód. interno"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Cliente"
               Object.Width           =   12656
            EndProperty
         End
      End
      Begin DrawSuite2022.USProgressBar PBLista11 
         Height          =   255
         Left            =   -74970
         TabIndex        =   159
         Top             =   8100
         Width           =   14985
         _ExtentX        =   26432
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
      Begin MSComctlLib.ListView Lista_Terceiros 
         Height          =   6525
         Left            =   -74970
         TabIndex        =   148
         Top             =   450
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   11509
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "ID item"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "ID pedido"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Pedido"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Fornecedor"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Data/Qtde exped."
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "D"
            Text            =   "Prazo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Ordem"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "OS"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Tag             =   "T"
            Text            =   "Obs"
            Object.Width           =   5106
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista12 
         Height          =   255
         Left            =   -74970
         TabIndex        =   168
         Top             =   6990
         Width           =   14985
         _ExtentX        =   26432
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
      Begin VB.Frame Frame13 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74970
         TabIndex        =   164
         Top             =   7260
         Width           =   14985
         Begin VB.TextBox txtNreg10 
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
            Left            =   3750
            TabIndex        =   149
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtPagIr10 
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
            TabIndex        =   150
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx10 
            Height          =   315
            Left            =   11760
            TabIndex        =   154
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":C3C7D
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
         Begin DrawSuite2022.USButton cmdPagAnt10 
            Height          =   315
            Left            =   11220
            TabIndex        =   153
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":C7421
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
         Begin DrawSuite2022.USButton cmdPagIr10 
            Height          =   315
            Left            =   10110
            TabIndex        =   151
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
         Begin DrawSuite2022.USButton cmdPagPrim10 
            Height          =   315
            Left            =   10680
            TabIndex        =   152
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":CAF2A
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
         Begin DrawSuite2022.USButton cmdPagUlt10 
            Height          =   315
            Left            =   12300
            TabIndex        =   155
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "FrmMDI_AvisoDiario.frx":CF019
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
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4380
            TabIndex        =   221
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3060
            TabIndex        =   167
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblRegistros10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   166
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblPaginas10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   165
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView Lista_pedidosAtraso 
         Height          =   7635
         Left            =   30
         TabIndex        =   170
         Top             =   450
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   13467
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "ID item"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "ID pedido"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Pedido"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Fornecedor"
            Object.Width           =   13221
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   2117
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
            Text            =   "Qtde. recebida"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Qtde. a receber"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Object.Tag             =   "D"
            Text            =   "Prazo"
            Object.Width           =   1587
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista13 
         Height          =   255
         Left            =   30
         TabIndex        =   171
         Top             =   8100
         Width           =   14985
         _ExtentX        =   26432
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
      Begin DrawSuite2022.USProgressBar PBLista_instrumentos 
         Height          =   255
         Left            =   -74970
         TabIndex        =   194
         Top             =   8100
         Width           =   14985
         _ExtentX        =   26432
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
      Begin DrawSuite2022.USProgressBar PBLista_NaoConforme 
         Height          =   255
         Left            =   -74970
         TabIndex        =   206
         Top             =   8100
         Width           =   14985
         _ExtentX        =   26432
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
      Begin MSComctlLib.ListView Lista_instrumentos 
         Height          =   7635
         Left            =   -74970
         TabIndex        =   207
         Top             =   450
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   13467
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
            Object.Width           =   0
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
            Text            =   "N. série"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   7927
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "D"
            Text            =   "Dt. aquis."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Fabricante"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "D"
            Text            =   "Dt. calib."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Órgão"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Object.Tag             =   "D"
            Text            =   "Próx. calib."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Tag             =   "T"
            Text            =   "Certificado"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ListView Lista_NaoConforme 
         Height          =   7635
         Left            =   -74970
         TabIndex        =   208
         Top             =   450
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   13467
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "Ordem"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "N"
            Text            =   "OS"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Fase"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Qtde."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Qtde. NC"
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
            Text            =   "Operador"
            Object.Width           =   6697
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "Disposição"
            Object.Width           =   6697
         EndProperty
      End
      Begin MSComctlLib.ListView ListaProdVencer 
         Height          =   8505
         Left            =   -74970
         TabIndex        =   223
         Top             =   450
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   15002
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "RE"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Lote"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Codigo"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   13219
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "D"
            Text            =   "Vencimento"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Saldo"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListaSugestoes 
         Height          =   8535
         Left            =   -74970
         TabIndex        =   224
         Top             =   450
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   15055
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Processo"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Código"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   9691
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Fase"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Sugestão"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Responsável"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   75
      Top             =   30
      Width           =   15195
      _ExtentX        =   26802
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
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   3090
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "FrmMDI_AvisoDiario.frx":D28A5
         Count           =   1
      End
   End
End
Attribute VB_Name = "FrmMDI_AvisoDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Aviso_diario_utiliza_formularel          As Boolean 'OK
Dim TBLista_aviso_diario                        As ADODB.Recordset 'OK
Dim TBLista_aviso_diario1                       As ADODB.Recordset 'OK
Dim TBLista_aviso_diario2                       As ADODB.Recordset 'OK
Dim TBLista_aviso_diario3                       As ADODB.Recordset 'OK
Dim TBLista_aviso_diario4                       As ADODB.Recordset 'OK
Dim TBLista_aviso_diario5                       As ADODB.Recordset 'OK
Dim TBLista_aviso_diario6                       As ADODB.Recordset 'OK
Dim TBLista_aviso_diario7                       As ADODB.Recordset 'OK
Dim TBLista_aviso_diario8                       As ADODB.Recordset 'OK
Dim TBLista_aviso_diario9                       As ADODB.Recordset 'OK
Dim TBLista_aviso_diario10                      As ADODB.Recordset 'OK
Dim TBLista_aviso_diario11                      As ADODB.Recordset 'OK
Dim TBLista_aviso_diario12                      As ADODB.Recordset 'OK
Dim TBLista_aviso_diario13                      As ADODB.Recordset 'OK
Dim TBLista_aviso_diario14                      As ADODB.Recordset 'OK
Dim TBLista_aviso_diarioAnalise                 As ADODB.Recordset 'OK
Dim TBLista_aviso_diarioCC                      As ADODB.Recordset 'OK
Dim TBLista_aviso_diarioItensNota               As ADODB.Recordset 'OK
Dim TBLista_aviso_diarioItensNota_pedido        As ADODB.Recordset 'OK
Dim TBLista_aviso_diarioAcesso                  As ADODB.Recordset 'OK
Dim TBLista_aviso_diarioTerceiros               As ADODB.Recordset 'OK
Dim TBLista_aviso_diarioTerceirosData           As ADODB.Recordset 'OK
Dim TBLista_aviso_diarioPedidoAtraso            As ADODB.Recordset 'OK
Dim TBLista_aviso_diario_Instrumentos           As ADODB.Recordset 'OK
Dim TBLista_aviso_diario_NaoConformidade        As ADODB.Recordset 'OK
Dim ContadorTAB                                 As Integer 'OK

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

ContadorTAB = 18

With SSTab1
    If .TabVisible(0) = True Then ProcCarregaListaContasPagar Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(1) = True Then ProcCarregaListaContasReceber Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(2) = True Then ProcCarregaListaSolicitacao Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(3) = True Then ProcCarregaListaManutencao Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(4) = True Then ProcCarregaListaRM Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(5) = True Then ProcCarregaListaNecessidadeCompras Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(6) = True Then ProcCarregaListaNecessidadePCP Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(7) = True Then ProcCarregaListaNecessidadeEstoque Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(8) = True Then ProcCarregaListaFaturar Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(9) = True Then ProcCarregaOS Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(10) = True Then ProcCarregaCC Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(11) = True Then ProcCarregaAnalise Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(12) = True Then ProcCarregaListaTerceiros Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(13) = True Then ProcCarregaListaPedidoAtraso Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(14) = True Then ProcCarregaListaInstrumentos Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(15) = True Then ProcCarregaListaNaoConforme Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(16) = True Then ProcCarregaListaProdVencer Else ContadorTAB = ContadorTAB - 1
    If .TabVisible(17) = True Then ProcCarregaListaSugestoes Else ContadorTAB = ContadorTAB - 1
End With
If ContadorTAB < 1 Then
    Unload Me
    USMsgBox ("Não foi encontrado nenhum aviso diário."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
SSTab1.TabsPerRow = ContadorTAB

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

frmMDI.Timer_avisodiario.Enabled = True
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario.AbsolutePage <> 2 Then
    If TBLista_aviso_diario.AbsolutePage = -3 Then
        ProcExibePagina (TBLista_aviso_diario.PageCount - 1)
    Else
        TBLista_aviso_diario.AbsolutePage = TBLista_aviso_diario.AbsolutePage - 2
        ProcExibePagina (TBLista_aviso_diario.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Instrumentos_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas_Instrumentos.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario_Instrumentos.AbsolutePage <> 2 Then
    If TBLista_aviso_diario_Instrumentos.AbsolutePage = -3 Then
        ProcExibePaginaInstrumento (TBLista_aviso_diario_Instrumentos.PageCount - 1)
    Else
        TBLista_aviso_diario_Instrumentos.AbsolutePage = TBLista_aviso_diario_Instrumentos.AbsolutePage - 2
        ProcExibePaginaInstrumento (TBLista_aviso_diario_Instrumentos.AbsolutePage)
    End If
Else
    ProcExibePaginaInstrumento (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_NaoConforme_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas_NaoConforme.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario_NaoConformidade.AbsolutePage <> 2 Then
    If TBLista_aviso_diario_NaoConformidade.AbsolutePage = -3 Then
        ProcExibePaginaNaoConforme (TBLista_aviso_diario_NaoConformidade.PageCount - 1)
    Else
        TBLista_aviso_diario_NaoConformidade.AbsolutePage = TBLista_aviso_diario_NaoConformidade.AbsolutePage - 2
        ProcExibePaginaNaoConforme (TBLista_aviso_diario_NaoConformidade.AbsolutePage)
    End If
Else
    ProcExibePaginaNaoConforme (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt10_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas10.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diarioTerceiros.AbsolutePage <> 2 Then
    If TBLista_aviso_diarioTerceiros.AbsolutePage = -3 Then
        ProcExibePaginaTerceiros (TBLista_aviso_diarioTerceiros.PageCount - 1)
    Else
        TBLista_aviso_diarioTerceiros.AbsolutePage = TBLista_aviso_diarioTerceiros.AbsolutePage - 2
        ProcExibePaginaTerceiros (TBLista_aviso_diarioTerceiros.AbsolutePage)
    End If
Else
    ProcExibePaginaTerceiros (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt11_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas11.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diarioPedidoAtraso.AbsolutePage <> 2 Then
    If TBLista_aviso_diarioPedidoAtraso.AbsolutePage = -3 Then
        ProcExibePaginaPedidoAtraso (TBLista_aviso_diarioPedidoAtraso.PageCount - 1)
    Else
        TBLista_aviso_diarioPedidoAtraso.AbsolutePage = TBLista_aviso_diarioPedidoAtraso.AbsolutePage - 2
        ProcExibePaginaPedidoAtraso (TBLista_aviso_diarioPedidoAtraso.AbsolutePage)
    End If
Else
    ProcExibePaginaPedidoAtraso (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt9_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas9.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diarioAnalise.AbsolutePage <> 2 Then
    If TBLista_aviso_diarioAnalise.AbsolutePage = -3 Then
        Select Case Tab_Analise.Tab
            Case 0: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.PageCount - 1, Lista_Analise_Engenharia
            Case 1: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.PageCount - 1, Lista_Analise_Processo
            Case 2: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.PageCount - 1, Lista_Analise_PCP
            Case 3: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.PageCount - 1, Lista_Analise_Qualidade
            Case 4: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.PageCount - 1, Lista_Analise_Compras
        End Select
    Else
        TBLista_aviso_diarioAnalise.AbsolutePage = TBLista_aviso_diarioAnalise.AbsolutePage - 2
        Select Case Tab_Analise.Tab
            Case 0: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Engenharia
            Case 1: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Processo
            Case 2: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_PCP
            Case 3: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Qualidade
            Case 4: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Compras
        End Select
    End If
Else
    Select Case Tab_Analise.Tab
        Case 0: ProcExibePaginaAnalise 1, Lista_Analise_Engenharia
        Case 1: ProcExibePaginaAnalise 1, Lista_Analise_Processo
        Case 2: ProcExibePaginaAnalise 1, Lista_Analise_PCP
        Case 3: ProcExibePaginaAnalise 1, Lista_Analise_Qualidade
        Case 4: ProcExibePaginaAnalise 1, Lista_Analise_Compras
    End Select
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
    TBLista_aviso_diario.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLista_aviso_diario.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr_Instrumentos_Click()
On Error GoTo tratar_erro

If txtPagIr_Instrumentos = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas_Instrumentos.Caption, 4))
If Quant <= 9 Or cmdPagIr_Instrumentos > Quant Then Exit Sub
If txtPagIr_Instrumentos >= 1 And txtPagIr_Instrumentos.Text <= Quant Then
    TBLista_aviso_diario_Instrumentos.AbsolutePage = txtPagIr_Instrumentos.Text
    ProcExibePaginaInstrumento TBLista_aviso_diario_Instrumentos.AbsolutePage
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr_NaoConforme_Click()
On Error GoTo tratar_erro

If txtPagIr_NaoConforme = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas_NaoConforme.Caption, 4))
If Quant <= 9 Or cmdPagIr_NaoConforme > Quant Then Exit Sub
If txtPagIr_NaoConforme >= 1 And txtPagIr_NaoConforme.Text <= Quant Then
    TBLista_aviso_diario_NaoConformidade.AbsolutePage = txtPagIr_NaoConforme.Text
    ProcExibePaginaNaoConforme TBLista_aviso_diario_NaoConformidade.AbsolutePage
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr10_Click()
On Error GoTo tratar_erro

If txtPagIr10 = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas10.Caption, 4))
If Quant <= 9 Or cmdPagIr10 > Quant Then Exit Sub
If txtPagIr10 >= 1 And txtPagIr10.Text <= Quant Then
    TBLista_aviso_diarioTerceiros.AbsolutePage = txtPagIr10.Text
    ProcExibePaginaTerceiros TBLista_aviso_diarioTerceiros.AbsolutePage
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr11_Click()
On Error GoTo tratar_erro

If txtPagIr11 = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas11.Caption, 4))
If Quant <= 9 Or cmdPagIr11 > Quant Then Exit Sub
If txtPagIr11 >= 1 And txtPagIr11.Text <= Quant Then
    TBLista_aviso_diarioPedidoAtraso.AbsolutePage = txtPagIr11.Text
    ProcExibePaginaPedidoAtraso TBLista_aviso_diarioPedidoAtraso.AbsolutePage
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr9_Click()
On Error GoTo tratar_erro

If txtPagIr9 = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas9.Caption, 4))
If Quant <= 9 Or txtPagIr9 > Quant Then Exit Sub
If txtPagIr9.Text >= 1 And txtPagIr9.Text <= Quant Then
    TBLista_aviso_diarioAnalise.AbsolutePage = txtPagIr9.Text
    Select Case Tab_Analise.Tab
        Case 0: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Engenharia
        Case 1: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Processo
        Case 2: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_PCP
        Case 3: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Qualidade
        Case 4: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Compras
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario.AbsolutePage = 1
ProcExibePagina (TBLista_aviso_diario.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Instrumentos_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas11.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario_Instrumentos.AbsolutePage = 1
ProcExibePaginaInstrumento (TBLista_aviso_diario_Instrumentos.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_NaoConforme_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas_NaoConforme.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario_NaoConformidade.AbsolutePage = 1
ProcExibePaginaNaoConforme (TBLista_aviso_diario_NaoConformidade.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim10_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas10.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diarioTerceiros.AbsolutePage = 1
ProcExibePaginaTerceiros (TBLista_aviso_diarioTerceiros.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim11_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas11.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diarioPedidoAtraso.AbsolutePage = 1
ProcExibePaginaPedidoAtraso (TBLista_aviso_diarioPedidoAtraso.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim9_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas9.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diarioAnalise.AbsolutePage = 1
Select Case Tab_Analise.Tab
    Case 0: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Engenharia
    Case 1: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Processo
    Case 2: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_PCP
    Case 3: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Qualidade
    Case 4: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Compras
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario.AbsolutePage <> -3 Then
    If TBLista_aviso_diario.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLista_aviso_diario.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLista_aviso_diario.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Instrumentos_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas_Instrumentos.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario_Instrumentos.AbsolutePage <> -3 Then
    If TBLista_aviso_diario_Instrumentos.AbsolutePage = 1 Then
        ProcExibePaginaInstrumento (2)
    Else
        ProcExibePaginaInstrumento (TBLista_aviso_diario_Instrumentos.AbsolutePage)
    End If
Else
    ProcExibePaginaInstrumento (TBLista_aviso_diario_Instrumentos.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_NaoConforme_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas_NaoConforme.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario_NaoConformidade.AbsolutePage <> -3 Then
    If TBLista_aviso_diario_NaoConformidade.AbsolutePage = 1 Then
        ProcExibePaginaNaoConforme (2)
    Else
        ProcExibePaginaNaoConforme (TBLista_aviso_diario_NaoConformidade.AbsolutePage)
    End If
Else
    ProcExibePaginaNaoConforme (TBLista_aviso_diario_NaoConformidade.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx10_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas10.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diarioTerceiros.AbsolutePage <> -3 Then
    If TBLista_aviso_diarioTerceiros.AbsolutePage = 1 Then
        ProcExibePaginaTerceiros (2)
    Else
        ProcExibePaginaTerceiros (TBLista_aviso_diarioTerceiros.AbsolutePage)
    End If
Else
    ProcExibePaginaTerceiros (TBLista_aviso_diarioTerceiros.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx11_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas11.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diarioPedidoAtraso.AbsolutePage <> -3 Then
    If TBLista_aviso_diarioPedidoAtraso.AbsolutePage = 1 Then
        ProcExibePaginaPedidoAtraso (2)
    Else
        ProcExibePaginaPedidoAtraso (TBLista_aviso_diarioPedidoAtraso.AbsolutePage)
    End If
Else
    ProcExibePaginaPedidoAtraso (TBLista_aviso_diarioPedidoAtraso.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx9_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas9.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diarioAnalise.AbsolutePage <> -3 Then
    If TBLista_aviso_diarioAnalise.AbsolutePage = 1 Then
        Select Case Tab_Analise.Tab
            Case 0: ProcExibePaginaAnalise 2, Lista_Analise_Engenharia
            Case 1: ProcExibePaginaAnalise 2, Lista_Analise_Processo
            Case 2: ProcExibePaginaAnalise 2, Lista_Analise_PCP
            Case 3: ProcExibePaginaAnalise 2, Lista_Analise_Qualidade
            Case 4: ProcExibePaginaAnalise 2, Lista_Analise_Compras
        End Select
    Else
        Select Case Tab_Analise.Tab
            Case 0: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Engenharia
            Case 1: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Processo
            Case 2: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_PCP
            Case 3: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Qualidade
            Case 4: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Compras
        End Select
    End If
Else
    Select Case Tab_Analise.Tab
        Case 0: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.PageCount, Lista_Analise_Engenharia
        Case 1: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.PageCount, Lista_Analise_Processo
        Case 2: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.PageCount, Lista_Analise_PCP
        Case 3: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.PageCount, Lista_Analise_Qualidade
        Case 4: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.PageCount, Lista_Analise_Compras
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario.AbsolutePage = TBLista_aviso_diario.PageCount
ProcExibePagina (TBLista_aviso_diario.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Instrumentos_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas_Instrumentos.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario_Instrumentos.AbsolutePage = TBLista_aviso_diario_Instrumentos.PageCount
ProcExibePaginaInstrumento (TBLista_aviso_diario_Instrumentos.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_NaoConforme_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas_NaoConforme.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario_NaoConformidade.AbsolutePage = TBLista_aviso_diario_NaoConformidade.PageCount
ProcExibePaginaNaoConforme (TBLista_aviso_diario_NaoConformidade.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt10_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas10.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diarioTerceiros.AbsolutePage = TBLista_aviso_diarioTerceiros.PageCount
ProcExibePaginaTerceiros (TBLista_aviso_diarioTerceiros.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt11_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas11.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diarioPedidoAtraso.AbsolutePage = TBLista_aviso_diarioPedidoAtraso.PageCount
ProcExibePaginaPedidoAtraso (TBLista_aviso_diarioPedidoAtraso.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt9_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas9.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diarioAnalise.AbsolutePage = TBLista_aviso_diarioAnalise.PageCount
Select Case Tab_Analise.Tab
    Case 0: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Engenharia
    Case 1: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Processo
    Case 2: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_PCP
    Case 3: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Qualidade
    Case 4: ProcExibePaginaAnalise TBLista_aviso_diarioAnalise.AbsolutePage, Lista_Analise_Compras
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub CmdSalvar_Click()
On Error GoTo tratar_erro

If Lista8.ListItems.Count = 0 Then Exit Sub
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Conexao.Execute "Update tbl_Dados_Nota_Fiscal Set Obs = '" & txtObs & "' where ID = '" & Lista8.SelectedItem & "'"
USMsgBox ("Observação da ordem de faturamento salva com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Avisos diário/Faturamento/Carteira de faturamento"
Acao = "Salvar observações da ordem de faturamento"
ID_documento = Lista8.SelectedItem
Documento = "Ordem de faturamento: " & Lista8.SelectedItem.ListSubItems.Item(3)
Documento1 = ""
ProcGravaEvento
'==================================
ProcCarregaListaFaturar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdSalvar_Terceiros_Click()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If Lista_Terceiros.ListItems.Count = 0 Then Exit Sub
Conexao.Execute "Update compras_pedido_lista Set Obs_AvisoDiario = '" & txtObs_Terceiros & "' where IDlista = " & Lista_Terceiros.SelectedItem
USMsgBox ("Observação salva com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = "Avisos diário/Terceiros"
Acao = "Salvar observações de terceiros"
ID_documento = Lista_Terceiros.SelectedItem
Documento = "Nº pedido: " & Lista_Terceiros.SelectedItem.ListSubItems.Item(2) & " - Cód. interno: " & Lista_Terceiros.SelectedItem.ListSubItems.Item(3)
Documento1 = ""
ProcGravaEvento
'==================================
ProcCarregaListaTerceiros

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub FlexGridCC_Click()
On Error GoTo tratar_erro
Dim lngImage As Boolean, L As Long, GroupoID As Long

'With FlexGridCC
'    .Redraw = False
'    If .RowGroupHeader = True Then
'        GroupoID = .RowData()
'        lngImage = (.RowImage() = 1)
'        For L = 0 To .rows - 1
'            If Not .RowGroupHeader(L) Then
'                If .RowData(L) = GroupoID Then
'                    .RowVisible(L) = lngImage
'                End If
'            End If
'        Next L
'        If lngImage Then
'            .RowImage() = 2
'        Else
'            .RowImage() = 1
'        End If
'    End If
'    .Redraw = True
'    .ColumnsForceFit
'End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub FlexGridOS_DblClick()
On Error GoTo tratar_erro

Formulario = "PCP/Gerenciamento de ordem"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

'With FlexGridOS
'    If Not .RowGroupHeader Then
'        G = .Row
'        Ordem = .CellText(G, 6)
'        OS = .CellText(G, 1)
'        frmprod.Show
'        frmprod.ProcLimpar True
'        frmprod.ProcCarregaOrdem
'        frmprod.ProcLimpaCamposAP
'        frmprod.Proclimpaevento
'        frmprod.ProcCarregaAPOS
'        frmprod.SSTab1.Tab = 4
'        frmprod.cmbAPOS = OS
'    End If
'End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Analise_Compras_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_Analise_Compras, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Analise_Compras_DblClick()
On Error GoTo tratar_erro

If Lista_Analise_Compras.ListItems.Count = 0 Then Exit Sub
Formulario = "Outros/Análise crítica/Compras"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With frmVendas_analise
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from vendas_analise where id = " & Lista_Analise_Compras.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        .ProcLimpaCampos
        .ProcPuxaDados
        .ProcCarregaLista_Engenharia_Prod
        .ProcCarregaLista_Engenharia_Checklist
        .ProcCarregaLista_Engenharia_Normas
        .ProcCarregaLista_processos_item
        .ProcCarregaLista_PCP_Checklist
        .ProcCarregaLista_Instrumentos
        .ProcCarregaLista_Qualidade_Checklist
        .ProcCarregalista_Compras
        .ProcCarregaLista_Compras_Checklist
        .ProcCarregaLista_Doc
        .SSTab1.Tab = 5
    End If
    Unload Me
End With
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Analise_Engenharia_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_Analise_Engenharia, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Analise_Engenharia_DblClick()
On Error GoTo tratar_erro

If Lista_Analise_Engenharia.ListItems.Count = 0 Then Exit Sub
Formulario = "Outros/Análise crítica/Engenharia"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With frmVendas_analise
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from vendas_analise where id = " & Lista_Analise_Engenharia.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        .ProcLimpaCampos
        .ProcPuxaDados
        .ProcCarregaLista_Engenharia_Prod
        .ProcCarregaLista_Engenharia_Checklist
        .ProcCarregaLista_Engenharia_Normas
        .ProcCarregaLista_processos_item
        .ProcCarregaLista_processos
        .ProcCarregaLista_PCP_Checklist
        .ProcCarregaLista_Instrumentos
        .ProcCarregaLista_Qualidade_Checklist
        .ProcCarregalista_Compras
        .ProcCarregaLista_Compras_Checklist
        .SSTab1.Tab = 1
    End If
    Unload Me
End With
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Analise_PCP_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_Analise_PCP, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Analise_PCP_DblClick()
On Error GoTo tratar_erro

If Lista_Analise_PCP.ListItems.Count = 0 Then Exit Sub
Formulario = "Outros/Análise crítica/Pcp"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With frmVendas_analise
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from vendas_analise where id = " & Lista_Analise_PCP.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        .ProcLimpaCampos
        .ProcPuxaDados
        .ProcCarregaLista_Engenharia_Prod
        .ProcCarregaLista_Engenharia_Checklist
        .ProcCarregaLista_Engenharia_Normas
        .ProcCarregaLista_processos_item
        .ProcCarregaLista_PCP_Checklist
        .ProcCarregaLista_Instrumentos
        .ProcCarregaLista_Qualidade_Checklist
        .ProcCarregalista_Compras
        .ProcCarregaLista_Compras_Checklist
        .ProcCarregaLista_Doc
        .SSTab1.Tab = 3
    End If
    Unload Me
End With
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Analise_Processo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_Analise_Processo, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Analise_Processo_DblClick()
On Error GoTo tratar_erro

If Lista_Analise_Processo.ListItems.Count = 0 Then Exit Sub
Formulario = "Outros/Análise crítica/Processos"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With frmVendas_analise
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from vendas_analise where id = " & Lista_Analise_Processo.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        .ProcLimpaCampos
        .ProcPuxaDados
        .ProcCarregaLista_Engenharia_Prod
        .ProcCarregaLista_Engenharia_Checklist
        .ProcCarregaLista_Engenharia_Normas
        .ProcCarregaLista_processos_item
        .ProcCarregaLista_PCP_Checklist
        .ProcCarregaLista_Instrumentos
        .ProcCarregaLista_Qualidade_Checklist
        .ProcCarregalista_Compras
        .ProcCarregaLista_Compras_Checklist
        .ProcCarregaLista_Doc
        .SSTab1.Tab = 2
    End If
    Unload Me
End With
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Analise_Qualidade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_Analise_Qualidade, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Analise_Qualidade_DblClick()
On Error GoTo tratar_erro

If Lista_Analise_Qualidade.ListItems.Count = 0 Then Exit Sub
Formulario = "Outros/Análise crítica/Qualidade"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With frmVendas_analise
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from vendas_analise where id = " & Lista_Analise_Qualidade.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        .ProcLimpaCampos
        .ProcPuxaDados
        .ProcCarregaLista_Engenharia_Prod
        .ProcCarregaLista_Engenharia_Checklist
        .ProcCarregaLista_Engenharia_Normas
        .ProcCarregaLista_processos_item
        .ProcCarregaLista_PCP_Checklist
        .ProcCarregaLista_Instrumentos
        .ProcCarregaLista_Qualidade_Checklist
        .ProcCarregalista_Compras
        .ProcCarregaLista_Compras_Checklist
        .ProcCarregaLista_Doc
        .SSTab1.Tab = 4
    End If
    Unload Me
End With
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_instrumentos_DblClick()
On Error GoTo tratar_erro

If Lista_instrumentos.ListItems.Count = 0 Then Exit Sub
Formulario = "Qualidade/Instrumentos"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With frmInstrumentos
    .StrSql_Instrumentos_Localizar = "Select I.CODIGO, I.Numero, EC.ref, EC.Numero_serie, I.Descricao, I.Data_Aquisicao, I.Fabricante, I.Familia from Instrumentos I INNER JOIN Estoque_controle EC ON EC.IDestoque = I.IDestoque where I.Codigo = " & Lista_instrumentos.SelectedItem
    .ProcCarregaLista (1)

    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select I.*, EC.Numero_serie, EC.ref from Instrumentos I LEFT JOIN Estoque_controle EC ON EC.IDestoque = I.IDestoque where I.Codigo = " & Lista_instrumentos.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        .ProcLimpar
        .ProcPuxaDados
    End If
    TBProduto.Close
    Unload Me
End With
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_NaoConforme_DblClick()
On Error GoTo tratar_erro

If Lista_NaoConforme.ListItems.Count = 0 Then Exit Sub
PCP_Ordem = False
Formulario = "Qualidade/Não conformidade"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With frmcqnc
    .StrSql_CQ_NC = "Select * from cq_nc_fabrica where Codigo = " & Lista_NaoConforme.SelectedItem
    .ProcCarregaLista (1)

    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select * from cq_nc_fabrica where Codigo = " & Lista_NaoConforme.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        .ProcLimpaCampos
        .ProcPuxaDados
        .Frame1.Enabled = True
        .Frame3.Enabled = True
        .Novo_CQNC = False
    End If
    TBLISTA.Close
    Unload Me
End With
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_pedidosAtraso_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_pedidosAtraso, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_instrumentos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_instrumentos, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_NaoConforme_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_NaoConforme, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_pedidosAtraso_DblClick()
On Error GoTo tratar_erro

If Lista_pedidosAtraso.ListItems.Count = 0 Then Exit Sub
ProcAbrirPedidoCompra Lista_pedidosAtraso.SelectedItem.SubItems(1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirPedidoCompra(IDpedido As Long)
On Error GoTo tratar_erro

Formulario = "Compras/Pedido"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With frmCompras_Pedido
    Aviso_diario_utiliza_formularel = True
    .Sql_Pedido_Localizar = "Select CP.IDpedido, CP.Data, CP.Pedido, CC.Cotacaotexto, CP.Fornecedor, CP.Status_pedido, CP.DtValidacao, CP.Data_aprovado from Compras_pedido CP LEFT JOIN Compras_cotacao CC ON CC.ID_cotacao = CP.IDcotacao where CP.IDpedido = " & IDpedido
    .FormulaRel_Pedido = "{compras_pedido.IDpedido} = " & IDpedido
    .listapedido.ListItems.Clear
    .ProcAtualizalistapedido (1)
    
    Set TBCompras_Pedido = CreateObject("adodb.recordset")
    TBCompras_Pedido.Open "Select * from compras_pedido where idpedido = " & IDpedido, Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Pedido.EOF = False Then
        .ProcLimpar
        .ProcLimpaCamposItem True
        .ProcLimpaCamposServ True
        .ProcPuxaDados
    End If
    TBCompras_Pedido.Close
    .SSTab1.Tab = 0
    Unload Me
End With
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Terceiros_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista_Terceiros, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Terceiros_DblClick()
On Error GoTo tratar_erro

If Lista_Terceiros.ListItems.Count = 0 Then Exit Sub
ProcAbrirPedidoCompra Lista_Terceiros.SelectedItem.SubItems(1)
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_Terceiros_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista_Terceiros.ListItems.Count = 0 Then Exit Sub
txtObs_Terceiros = Lista_Terceiros.SelectedItem.ListSubItems.Item(10)
frameOBS_Terceiros.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista8_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista8.ListItems.Count = 0 Then Exit Sub
txtObs = Lista8.SelectedItem.ListSubItems.Item(8)
procCarregaListaItensNota
FrameObs.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Tab_Analise_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

ProcCarregaListaAnalise

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

Private Sub txtNreg_Instrumentos_Change()
On Error GoTo tratar_erro

If txtNreg_Instrumentos <> "" Then
    VerifNumero = txtNreg_Instrumentos
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg_Instrumentos = ""
        txtNreg_Instrumentos.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg_NaoConforme_Change()
On Error GoTo tratar_erro

If txtNreg_NaoConforme <> "" Then
    VerifNumero = txtNreg_NaoConforme
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg_NaoConforme = ""
        txtNreg_NaoConforme.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg10_Change()
On Error GoTo tratar_erro

If txtNreg10 <> "" Then
    VerifNumero = txtNreg10
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg10 = ""
        txtNreg10.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg11_Change()
On Error GoTo tratar_erro

If txtNreg11 <> "" Then
    VerifNumero = txtNreg11
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg11 = ""
        txtNreg11.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg9_Change()
On Error GoTo tratar_erro

If txtNreg9 <> "" Then
    VerifNumero = txtNreg9
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg9 = ""
        txtNreg9.SetFocus
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

Private Sub cmdPagAnt1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario1.AbsolutePage <> 2 Then
    If TBLista_aviso_diario1.AbsolutePage = -3 Then
        ProcExibePagina1 (TBLista_aviso_diario1.PageCount - 1)
    Else
        TBLista_aviso_diario1.AbsolutePage = TBLista_aviso_diario1.AbsolutePage - 2
        ProcExibePagina1 (TBLista_aviso_diario1.AbsolutePage)
    End If
Else
    ProcExibePagina1 (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr1_Click()
On Error GoTo tratar_erro

If txtPagIr1 = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas1.Caption, 4))
If Quant <= 1 Or txtPagIr1 > Quant Then Exit Sub
If txtPagIr1.Text >= 1 And txtPagIr1.Text <= Quant Then
    TBLista_aviso_diario1.AbsolutePage = txtPagIr1.Text
    ProcExibePagina1 (TBLista_aviso_diario1.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario1.AbsolutePage = 1
ProcExibePagina1 (TBLista_aviso_diario1.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario1.AbsolutePage <> -3 Then
    If TBLista_aviso_diario1.AbsolutePage = 1 Then
        ProcExibePagina1 (2)
    Else
        ProcExibePagina1 (TBLista_aviso_diario1.AbsolutePage)
    End If
Else
    ProcExibePagina1 (TBLista_aviso_diario1.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt1_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas1.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario1.AbsolutePage = TBLista_aviso_diario1.PageCount
ProcExibePagina1 (TBLista_aviso_diario1.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg1_Change()
On Error GoTo tratar_erro

If txtNreg1 <> "" Then
    VerifNumero = txtNreg1
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg1 = ""
        txtNreg1.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr_Instrumentos_Change()
On Error GoTo tratar_erro

If txtPagIr_Instrumentos <> "" Then
    VerifNumero = txtPagIr_Instrumentos
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr_Instrumentos = ""
        txtPagIr_Instrumentos.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr_NaoConforme_Change()
On Error GoTo tratar_erro

If txtPagIr_NaoConforme <> "" Then
    VerifNumero = txtPagIr_NaoConforme
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr_NaoConforme = ""
        txtPagIr_NaoConforme.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr1_Change()
On Error GoTo tratar_erro

If txtPagIr1 <> "" Then
    VerifNumero = txtPagIr1
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr1 = ""
        txtPagIr1.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt2_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas2.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario2.AbsolutePage <> 2 Then
    If TBLista_aviso_diario2.AbsolutePage = -3 Then
        ProcExibePagina2 (TBLista_aviso_diario2.PageCount - 1)
    Else
        TBLista_aviso_diario2.AbsolutePage = TBLista_aviso_diario2.AbsolutePage - 2
        ProcExibePagina2 (TBLista_aviso_diario2.AbsolutePage)
    End If
Else
    ProcExibePagina2 (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr2_Click()
On Error GoTo tratar_erro

If txtPagIr2 = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas2.Caption, 4))
If Quant <= 1 Or txtPagIr2 > Quant Then Exit Sub
If txtPagIr2.Text >= 1 And txtPagIr2.Text <= Quant Then
    TBLista_aviso_diario2.AbsolutePage = txtPagIr2.Text
    ProcExibePagina2 (TBLista_aviso_diario2.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim2_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas2.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario2.AbsolutePage = 1
ProcExibePagina2 (TBLista_aviso_diario2.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx2_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas2.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario2.AbsolutePage <> -3 Then
    If TBLista_aviso_diario2.AbsolutePage = 1 Then
        ProcExibePagina2 (2)
    Else
        ProcExibePagina2 (TBLista_aviso_diario2.AbsolutePage)
    End If
Else
    ProcExibePagina2 (TBLista_aviso_diario2.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt2_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas2.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario2.AbsolutePage = TBLista_aviso_diario2.PageCount
ProcExibePagina (TBLista_aviso_diario2.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg2_Change()
On Error GoTo tratar_erro

If txtNreg2 <> "" Then
    VerifNumero = txtNreg2
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg2 = ""
        txtNreg2.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr10_Change()
On Error GoTo tratar_erro

If txtPagIr10 <> "" Then
    VerifNumero = txtPagIr10
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr10 = ""
        txtPagIr10.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr11_Change()
On Error GoTo tratar_erro

If txtPagIr11 <> "" Then
    VerifNumero = txtPagIr11
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr11 = ""
        txtPagIr11.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr2_Change()
On Error GoTo tratar_erro

If txtPagIr2 <> "" Then
    VerifNumero = txtPagIr2
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr2 = ""
        txtPagIr2.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt3_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas3.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario3.AbsolutePage <> 2 Then
    If TBLista_aviso_diario3.AbsolutePage = -3 Then
        ProcExibePagina3 (TBLista_aviso_diario3.PageCount - 1)
    Else
        TBLista_aviso_diario3.AbsolutePage = TBLista_aviso_diario3.AbsolutePage - 2
        ProcExibePagina3 (TBLista_aviso_diario3.AbsolutePage)
    End If
Else
    ProcExibePagina3 (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr3_Click()
On Error GoTo tratar_erro

If txtPagIr3 = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas3.Caption, 4))
If Quant <= 1 Or txtPagIr3 > Quant Then Exit Sub
If txtPagIr3.Text >= 1 And txtPagIr3.Text <= Quant Then
    TBLista_aviso_diario3.AbsolutePage = txtPagIr3.Text
    ProcExibePagina3 (TBLista_aviso_diario3.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim3_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas3.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario3.AbsolutePage = 1
ProcExibePagina3 (TBLista_aviso_diario3.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx3_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas3.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario3.AbsolutePage <> -3 Then
    If TBLista_aviso_diario3.AbsolutePage = 1 Then
        ProcExibePagina3 (2)
    Else
        ProcExibePagina3 (TBLista_aviso_diario3.AbsolutePage)
    End If
Else
    ProcExibePagina3 (TBLista_aviso_diario3.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt3_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas3.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario3.AbsolutePage = TBLista_aviso_diario3.PageCount
ProcExibePagina3 (TBLista_aviso_diario3.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg3_Change()
On Error GoTo tratar_erro

If txtNreg3 <> "" Then
    VerifNumero = txtNreg3
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg3 = ""
        txtNreg3.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr3_Change()
On Error GoTo tratar_erro

If txtPagIr3 <> "" Then
    VerifNumero = txtPagIr3
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr3 = ""
        txtPagIr3.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt4_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas4.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario4.AbsolutePage <> 2 Then
    If TBLista_aviso_diario4.AbsolutePage = -3 Then
        ProcExibePagina4 (TBLista_aviso_diario4.PageCount - 1)
    Else
        TBLista_aviso_diario4.AbsolutePage = TBLista_aviso_diario4.AbsolutePage - 2
        ProcExibePagina4 (TBLista_aviso_diario4.AbsolutePage)
    End If
Else
    ProcExibePagina4 (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr4_Click()
On Error GoTo tratar_erro

If txtPagIr4 = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas4.Caption, 4))
If Quant <= 1 Or txtPagIr4 > Quant Then Exit Sub
If txtPagIr4.Text >= 1 And txtPagIr4.Text <= Quant Then
    TBLista_aviso_diario4.AbsolutePage = txtPagIr4.Text
    ProcExibePagina4 (TBLista_aviso_diario4.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim4_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas4.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario4.AbsolutePage = 1
ProcExibePagina4 (TBLista_aviso_diario4.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx4_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas4.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario4.AbsolutePage <> -3 Then
    If TBLista_aviso_diario4.AbsolutePage = 1 Then
        ProcExibePagina4 (2)
    Else
        ProcExibePagina4 (TBLista_aviso_diario4.AbsolutePage)
    End If
Else
    ProcExibePagina4 (TBLista_aviso_diario4.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt4_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas4.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario4.AbsolutePage = TBLista_aviso_diario4.PageCount
ProcExibePagina4 (TBLista_aviso_diario4.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg4_Change()
On Error GoTo tratar_erro

If txtNreg4 <> "" Then
    VerifNumero = txtNreg4
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg4 = ""
        txtNreg4.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr4_Change()
On Error GoTo tratar_erro

If txtPagIr4 <> "" Then
    VerifNumero = txtPagIr4
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr4 = ""
        txtPagIr4.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt5_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas5.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario5.AbsolutePage <> 2 Then
    If TBLista_aviso_diario5.AbsolutePage = -3 Then
        ProcExibePagina5 (TBLista_aviso_diario5.PageCount - 1)
    Else
        TBLista_aviso_diario5.AbsolutePage = TBLista_aviso_diario5.AbsolutePage - 2
        ProcExibePagina5 (TBLista_aviso_diario5.AbsolutePage)
    End If
Else
    ProcExibePagina5 (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr5_Click()
On Error GoTo tratar_erro

If txtPagIr5 = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas5.Caption, 4))
If Quant <= 1 Or txtPagIr5 > Quant Then Exit Sub
If txtPagIr5.Text >= 1 And txtPagIr5.Text <= Quant Then
    TBLista_aviso_diario5.AbsolutePage = txtPagIr5.Text
    ProcExibePagina5 (TBLista_aviso_diario5.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim5_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas5.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario5.AbsolutePage = 1
ProcExibePagina5 (TBLista_aviso_diario5.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx5_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas5.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario5.AbsolutePage <> -3 Then
    If TBLista_aviso_diario5.AbsolutePage = 1 Then
        ProcExibePagina5 (2)
    Else
        ProcExibePagina5 (TBLista_aviso_diario5.AbsolutePage)
    End If
Else
    ProcExibePagina5 (TBLista_aviso_diario5.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt5_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas5.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario5.AbsolutePage = TBLista_aviso_diario5.PageCount
ProcExibePagina5 (TBLista_aviso_diario5.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg5_Change()
On Error GoTo tratar_erro

If txtNreg5 <> "" Then
    VerifNumero = txtNreg5
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg5 = ""
        txtNreg5.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr5_Change()
On Error GoTo tratar_erro

If txtPagIr5 <> "" Then
    VerifNumero = txtPagIr5
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr5 = ""
        txtPagIr5.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt6_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas6.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario6.AbsolutePage <> 2 Then
    If TBLista_aviso_diario6.AbsolutePage = -3 Then
        ProcExibePagina6 (TBLista_aviso_diario6.PageCount - 1)
    Else
        TBLista_aviso_diario6.AbsolutePage = TBLista_aviso_diario6.AbsolutePage - 2
        ProcExibePagina6 (TBLista_aviso_diario6.AbsolutePage)
    End If
Else
    ProcExibePagina6 (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr6_Click()
On Error GoTo tratar_erro

If txtPagIr6 = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas6.Caption, 4))
If Quant <= 1 Or txtPagIr6 > Quant Then Exit Sub
If txtPagIr6.Text >= 1 And txtPagIr6.Text <= Quant Then
    TBLista_aviso_diario6.AbsolutePage = txtPagIr6.Text
    ProcExibePagina6 (TBLista_aviso_diario6.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim6_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas6.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario6.AbsolutePage = 1
ProcExibePagina6 (TBLista_aviso_diario6.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx6_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas6.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario6.AbsolutePage <> -3 Then
    If TBLista_aviso_diario6.AbsolutePage = 1 Then
        ProcExibePagina6 (2)
    Else
        ProcExibePagina6 (TBLista_aviso_diario6.AbsolutePage)
    End If
Else
    ProcExibePagina6 (TBLista_aviso_diario6.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt6_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas6.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario6.AbsolutePage = TBLista_aviso_diario6.PageCount
ProcExibePagina6 (TBLista_aviso_diario6.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg6_Change()
On Error GoTo tratar_erro

If txtNreg6 <> "" Then
    VerifNumero = txtNreg6
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg6 = ""
        txtNreg6.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr6_Change()
On Error GoTo tratar_erro

If txtPagIr6 <> "" Then
    VerifNumero = txtPagIr6
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr6 = ""
        txtPagIr6.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt7_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas7.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario7.AbsolutePage <> 2 Then
    If TBLista_aviso_diario7.AbsolutePage = -3 Then
        ProcExibePagina7 (TBLista_aviso_diario7.PageCount - 1)
    Else
        TBLista_aviso_diario7.AbsolutePage = TBLista_aviso_diario7.AbsolutePage - 2
        ProcExibePagina7 (TBLista_aviso_diario7.AbsolutePage)
    End If
Else
    ProcExibePagina7 (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr7_Click()
On Error GoTo tratar_erro

If txtPagIr7 = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas7.Caption, 4))
If Quant <= 1 Or txtPagIr7 > Quant Then Exit Sub
If txtPagIr7.Text >= 1 And txtPagIr7.Text <= Quant Then
    TBLista_aviso_diario7.AbsolutePage = txtPagIr7.Text
    ProcExibePagina7 (TBLista_aviso_diario7.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim7_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas7.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario7.AbsolutePage = 1
ProcExibePagina7 (TBLista_aviso_diario7.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx7_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas7.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario7.AbsolutePage <> -3 Then
    If TBLista_aviso_diario7.AbsolutePage = 1 Then
        ProcExibePagina7 (2)
    Else
        ProcExibePagina7 (TBLista_aviso_diario7.AbsolutePage)
    End If
Else
    ProcExibePagina7 (TBLista_aviso_diario7.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt7_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas7.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario7.AbsolutePage = TBLista_aviso_diario7.PageCount
ProcExibePagina7 (TBLista_aviso_diario7.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg7_Change()
On Error GoTo tratar_erro

If txtNreg7 <> "" Then
    VerifNumero = txtNreg7
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg7 = ""
        txtNreg7.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr7_Change()
On Error GoTo tratar_erro

If txtPagIr7 <> "" Then
    VerifNumero = txtPagIr7
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr7 = ""
        txtPagIr7.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt8_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas8.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario8.AbsolutePage <> 2 Then
    If TBLista_aviso_diario8.AbsolutePage = -3 Then
        ProcExibePagina8 (TBLista_aviso_diario8.PageCount - 1)
    Else
        TBLista_aviso_diario8.AbsolutePage = TBLista_aviso_diario8.AbsolutePage - 2
        ProcExibePagina8 (TBLista_aviso_diario8.AbsolutePage)
    End If
Else
    ProcExibePagina8 (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr8_Click()
On Error GoTo tratar_erro

If txtPagIr8 = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas8.Caption, 4))
If Quant <= 1 Or txtPagIr8 > Quant Then Exit Sub
If txtPagIr8.Text >= 1 And txtPagIr8.Text <= Quant Then
    TBLista_aviso_diario8.AbsolutePage = txtPagIr8.Text
    ProcExibePagina8 (TBLista_aviso_diario8.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim8_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas8.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario8.AbsolutePage = 1
ProcExibePagina8 (TBLista_aviso_diario8.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx8_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas8.Caption, 4)) <= 1 Then Exit Sub
If TBLista_aviso_diario8.AbsolutePage <> -3 Then
    If TBLista_aviso_diario8.AbsolutePage = 1 Then
        ProcExibePagina8 (2)
    Else
        ProcExibePagina8 (TBLista_aviso_diario8.AbsolutePage)
    End If
Else
    ProcExibePagina8 (TBLista_aviso_diario8.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt8_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas8.Caption, 4)) <= 1 Then Exit Sub
TBLista_aviso_diario8.AbsolutePage = TBLista_aviso_diario8.PageCount
ProcExibePagina8 (TBLista_aviso_diario8.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg8_Change()
On Error GoTo tratar_erro

If txtNreg8 <> "" Then
    VerifNumero = txtNreg8
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg8 = ""
        txtNreg8.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr8_Change()
On Error GoTo tratar_erro

If txtPagIr8 <> "" Then
    VerifNumero = txtPagIr8
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr8 = ""
        txtPagIr8.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: If SSTab1.TabVisible(8) = True And FrameObs.Enabled = True Then CmdSalvar_Click
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 5, True
Direitos
Contador = 18
ProcVerifMostrarEsconderTab SSTab1, 0, "Avisos diário/Contas a pagar"
ProcVerifMostrarEsconderTab SSTab1, 1, "Avisos diário/Contas a receber"
ProcVerifMostrarEsconderTab SSTab1, 2, "Avisos diário/Solicitação"
ProcVerifMostrarEsconderTab SSTab1, 3, "Avisos diário/Manutenção"
ProcVerifMostrarEsconderTab SSTab1, 4, "Avisos diário/Requisição de materiais"
ProcVerifMostrarEsconderTab SSTab1, 5, "Avisos diário/Compras/Necessidade"
ProcVerifMostrarEsconderTab SSTab1, 6, "Avisos diário/PCP/Necessidade"
ProcVerifMostrarEsconderTab SSTab1, 7, "Avisos diário/Estoque/Necessidade"
ProcVerifMostrarEsconderTab SSTab1, 8, "Avisos diário/Faturamento/Carteira de faturamento"
ProcVerifMostrarEsconderTab SSTab1, 9, "Avisos diário/PCP/OSs em atraso"
ProcVerifMostrarEsconderTab SSTab1, 10, "Avisos diário/Custos/Centro de custo"
ProcVerifMostrarEsconderTab SSTab1, 11, ""
ProcVerifMostrarEsconderTab SSTab1, 12, "Avisos diário/Terceiros"
ProcVerifMostrarEsconderTab SSTab1, 13, "Avisos diário/Compras/Pedidos em atraso"
ProcVerifMostrarEsconderTab SSTab1, 14, "Avisos diário/Qualidade/Calibração a vencer"
ProcVerifMostrarEsconderTab SSTab1, 15, "Avisos diário/Qualidade/Não conformidades"
ProcVerifMostrarEsconderTab SSTab1, 16, "Avisos diário/Estoque/Produtos á vencer"
ProcVerifMostrarEsconderTab SSTab1, 17, "Avisos diário/Processos/Sugestões"

With SSTab1
    If .TabVisible(17) = True Then .Tab = 17
    If .TabVisible(16) = True Then .Tab = 16
    If .TabVisible(15) = True Then .Tab = 15
    If .TabVisible(14) = True Then .Tab = 14
    If .TabVisible(13) = True Then .Tab = 13
    If .TabVisible(12) = True Then .Tab = 12
    If .TabVisible(11) = True Then .Tab = 11
    If .TabVisible(10) = True Then .Tab = 10
    If .TabVisible(9) = True Then .Tab = 9
    If .TabVisible(8) = True Then .Tab = 8
    If .TabVisible(7) = True Then .Tab = 7
    If .TabVisible(6) = True Then .Tab = 6
    If .TabVisible(5) = True Then .Tab = 5
    If .TabVisible(4) = True Then .Tab = 4
    If .TabVisible(3) = True Then .Tab = 3
    If .TabVisible(2) = True Then .Tab = 2
    If .TabVisible(1) = True Then .Tab = 1
    If .TabVisible(0) = True Then .Tab = 0
End With

Contador = 18
ProcVerifMostrarEsconderTab Tab_Analise, 0, "Avisos diário/Análise crítica/Engenharia"
ProcVerifMostrarEsconderTab Tab_Analise, 1, "Avisos diário/Análise crítica/Processos"
ProcVerifMostrarEsconderTab Tab_Analise, 2, "Avisos diário/Análise crítica/Pcp"
ProcVerifMostrarEsconderTab Tab_Analise, 3, "Avisos diário/Análise crítica/Qualidade"
ProcVerifMostrarEsconderTab Tab_Analise, 4, "Avisos diário/Análise crítica/Compras"

ProcLimpaVariaveisPrincipais
ProcFiltrar

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaProdVencer()
On Error GoTo tratar_erro

ListaProdVencer.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Set TBLista_aviso_diario = CreateObject("adodb.recordset")
TBLista_aviso_diario.Open "Select * from Estoque_Produtos_Vencer", Conexao, adOpenKeyset, adLockOptimistic
If TBLista_aviso_diario.EOF = False Then
    SSTab1.TabVisible(16) = True
    ProcExibePagina16 (1)
Else
    SSTab1.TabVisible(16) = False
    ContadorTAB = ContadorTAB - 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaSugestoes()
On Error GoTo tratar_erro

ListaSugestoes.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Set TBLista_aviso_diario = CreateObject("adodb.recordset")
StrSql = "Select FS.ID,FS.Status, PR.Nprocesso as Processo, PP.Desenho as Codigo, PP.descricao, FA.Fase, FS.Data, FS.Sugestao , FS.Responsavel from processos PR Inner join projproduto PP ON PR.CodProduto = PP.codproduto INNER JOIN Fases FA ON PR.IDProcesso = FA.IDProcesso INNER JOIN Fases_Sugestao FS ON FA.IDFase = FS.IDFase WHERE FS.Status = 1"
TBLista_aviso_diario.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLista_aviso_diario.EOF = False Then
    SSTab1.TabVisible(17) = True
    ProcExibePagina17 (1)
Else
    SSTab1.TabVisible(17) = False
    ContadorTAB = ContadorTAB - 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaContasPagar()
On Error GoTo tratar_erro

Lista.ListItems.Clear
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Set TBLista_aviso_diario = CreateObject("adodb.recordset")
TBLista_aviso_diario.Open "Select * from tbl_ContasPagar where dt_Pagamento <= '" & Format(Date, "Short Date") & "' and Logsit = 'N' and Bloqueado = 'False' and Status <> 'TÍTULO LIQUIDADO ANTECIPADO' order by dt_Pagamento", Conexao, adOpenKeyset, adLockOptimistic
If TBLista_aviso_diario.EOF = False Then
    SSTab1.TabVisible(0) = True
    ProcExibePagina (1)
Else
    SSTab1.TabVisible(0) = False
    ContadorTAB = ContadorTAB - 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLista_aviso_diario.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLista_aviso_diario.AbsolutePage = Pagina
TamanhoPagina = TBLista_aviso_diario.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLista_aviso_diario.RecordCount - IIf(Pagina > 1, (TBLista_aviso_diario.PageSize * (Pagina - 1)), 0), TBLista_aviso_diario.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLista_aviso_diario.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBLista_aviso_diario!IDintconta
        .Item(.Count).SubItems(1) = Format(TBLista_aviso_diario!dbl_valorpagto, "###,##0.00")
        .Item(.Count).SubItems(2) = Format(TBLista_aviso_diario!dt_Pagamento, "dd/mm/yy")
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLista_aviso_diario!Txt_fornecedor), "", TBLista_aviso_diario!Txt_fornecedor)
    End With
    TBLista_aviso_diario.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLista_aviso_diario.RecordCount
If TBLista_aviso_diario.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLista_aviso_diario.PageCount
ElseIf TBLista_aviso_diario.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLista_aviso_diario.PageCount & " de: " & TBLista_aviso_diario.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLista_aviso_diario.AbsolutePage - 1 & " de: " & TBLista_aviso_diario.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina16(Pagina)
On Error GoTo tratar_erro

ListaProdVencer.ListItems.Clear
TBLista_aviso_diario.PageSize = IIf(txtNreg = "", 30, txtNreg)
'TBLista_aviso_diario.AbsolutePage = Pagina
TamanhoPagina = TBLista_aviso_diario.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLista_aviso_diario.RecordCount - IIf(Pagina > 1, (TBLista_aviso_diario.PageSize * (Pagina - 1)), 0), TBLista_aviso_diario.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLista_aviso_diario.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListaProdVencer.ListItems
        .Add , , TBLista_aviso_diario!IDEstoque
        .Item(.Count).SubItems(1) = TBLista_aviso_diario!LOTE
        .Item(.Count).SubItems(2) = TBLista_aviso_diario!Desenho
        .Item(.Count).SubItems(3) = TBLista_aviso_diario!Descricao
        .Item(.Count).SubItems(4) = TBLista_aviso_diario!Vencimento
        .Item(.Count).SubItems(5) = Format(TBLista_aviso_diario!Saldo, "###,##0.00")
    End With
    TBLista_aviso_diario.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLista_aviso_diario.RecordCount
If TBLista_aviso_diario.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLista_aviso_diario.PageCount
ElseIf TBLista_aviso_diario.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLista_aviso_diario.PageCount & " de: " & TBLista_aviso_diario.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLista_aviso_diario.AbsolutePage - 1 & " de: " & TBLista_aviso_diario.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina17(Pagina)
On Error GoTo tratar_erro

ListaSugestoes.ListItems.Clear
TBLista_aviso_diario.PageSize = IIf(txtNreg = "", 30, txtNreg)
'TBLista_aviso_diario.AbsolutePage = Pagina
TamanhoPagina = TBLista_aviso_diario.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLista_aviso_diario.RecordCount - IIf(Pagina > 1, (TBLista_aviso_diario.PageSize * (Pagina - 1)), 0), TBLista_aviso_diario.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLista_aviso_diario.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListaSugestoes.ListItems
        .Add , , TBLista_aviso_diario!ID
        .Item(.Count).SubItems(1) = TBLista_aviso_diario!Processo
        .Item(.Count).SubItems(2) = TBLista_aviso_diario!CODIGO
        .Item(.Count).SubItems(3) = TBLista_aviso_diario!Descricao
        .Item(.Count).SubItems(4) = TBLista_aviso_diario!Fase
        .Item(.Count).SubItems(5) = Format(TBLista_aviso_diario!Data, "dd/mm/yyyy")
        .Item(.Count).SubItems(6) = TBLista_aviso_diario!Sugestao
        .Item(.Count).SubItems(7) = TBLista_aviso_diario!Responsavel
    End With
    TBLista_aviso_diario.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLista_aviso_diario.RecordCount
If TBLista_aviso_diario.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLista_aviso_diario.PageCount
ElseIf TBLista_aviso_diario.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLista_aviso_diario.PageCount & " de: " & TBLista_aviso_diario.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLista_aviso_diario.AbsolutePage - 1 & " de: " & TBLista_aviso_diario.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaContasReceber()
On Error GoTo tratar_erro

Lista1.ListItems.Clear
lblRegistros1.Caption = "Nº de registros: 0"
lblPaginas1.Caption = "Página: 0 de: 0"
Set TBLista_aviso_diario1 = CreateObject("adodb.recordset")
TBLista_aviso_diario1.Open "Select * from tbl_contas_receber where Vencimento <= '" & Format(Date, "Short Date") & "' and Logsit = 'N' and Bloqueado = 'False' and Status <> 'TÍTULO LIQUIDADO ANTECIPADO' order by vencimento", Conexao, adOpenKeyset, adLockOptimistic
If TBLista_aviso_diario1.EOF = False Then
    SSTab1.TabVisible(1) = True
    ProcExibePagina1 (1)
Else
    SSTab1.TabVisible(1) = False
    ContadorTAB = ContadorTAB - 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina1(Pagina)
On Error GoTo tratar_erro

Lista1.ListItems.Clear
TBLista_aviso_diario1.PageSize = IIf(txtNreg1 = "", 30, txtNreg1)
TBLista_aviso_diario1.AbsolutePage = Pagina
TamanhoPagina = TBLista_aviso_diario1.PageSize
ContadorReg = 1

PBLista1.Min = 0
PBLista1.Max = FunVerifMaxPBListaPaginacao(TBLista_aviso_diario1.RecordCount - IIf(Pagina > 1, (TBLista_aviso_diario1.PageSize * (Pagina - 1)), 0), TBLista_aviso_diario1.PageSize)
PBLista1.Value = 1
Contador = 0
Do While TBLista_aviso_diario1.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista1.ListItems
        .Add , , TBLista_aviso_diario1!IDintconta
        .Item(.Count).SubItems(1) = Format(TBLista_aviso_diario1!valor, "###,##0.00")
        .Item(.Count).SubItems(2) = Format(TBLista_aviso_diario1!Vencimento, "dd/mm/yy")
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLista_aviso_diario1!Nome_Razao), "", TBLista_aviso_diario1!Nome_Razao)
    End With
    TBLista_aviso_diario1.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista1.Value = Contador
Loop
lblRegistros1.Caption = "Nº de registros: " & TBLista_aviso_diario1.RecordCount
If TBLista_aviso_diario1.AbsolutePage = adPosBOF Then
   lblPaginas1.Caption = "Página: 1 de: " & TBLista_aviso_diario1.PageCount
ElseIf TBLista_aviso_diario1.AbsolutePage = adPosEOF Then
        lblPaginas1.Caption = "Página: " & TBLista_aviso_diario1.PageCount & " de: " & TBLista_aviso_diario1.PageCount
    Else
        lblPaginas1.Caption = "Página: " & TBLista_aviso_diario1.AbsolutePage - 1 & " de: " & TBLista_aviso_diario1.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaSolicitacao()
On Error GoTo tratar_erro

Lista2.ListItems.Clear
lblRegistros2.Caption = "Nº de registros: 0"
lblPaginas2.Caption = "Página: 0 de: 0"
Set TBLista_aviso_diario2 = CreateObject("adodb.recordset")
TBLista_aviso_diario2.Open "Select * from Compras_requisicao where Data_Solicitacao <= '" & Format(Date, "Short Date") & "' and Status = 'ABERTA' and DtValidacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBLista_aviso_diario2.EOF = False Then
    SSTab1.TabVisible(2) = True
    ProcExibePagina2 (1)
Else
    SSTab1.TabVisible(2) = False
    ContadorTAB = ContadorTAB - 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina2(Pagina)
On Error GoTo tratar_erro

Lista2.ListItems.Clear
TBLista_aviso_diario2.PageSize = IIf(txtNreg2 = "", 30, txtNreg2)
TBLista_aviso_diario2.AbsolutePage = Pagina
TamanhoPagina = TBLista_aviso_diario2.PageSize
ContadorReg = 1

PBLista2.Min = 0
PBLista2.Max = FunVerifMaxPBListaPaginacao(TBLista_aviso_diario2.RecordCount - IIf(Pagina > 1, (TBLista_aviso_diario2.PageSize * (Pagina - 1)), 0), TBLista_aviso_diario2.PageSize)
PBLista2.Value = 1
Contador = 0
Do While TBLista_aviso_diario2.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista2.ListItems
        .Add , , TBLista_aviso_diario2!ID_Requisicao
        .Item(.Count).SubItems(1) = TBLista_aviso_diario2!Requisicaotexto
        .Item(.Count).SubItems(2) = Format(TBLista_aviso_diario2!Data_Solicitacao, "dd/mm/yy")
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLista_aviso_diario2!solicitado), "", TBLista_aviso_diario2!solicitado)
    End With
    TBLista_aviso_diario2.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista2.Value = Contador
Loop
lblRegistros2.Caption = "Nº de registros: " & TBLista_aviso_diario2.RecordCount
If TBLista_aviso_diario2.AbsolutePage = adPosBOF Then
   lblPaginas2.Caption = "Página: 1 de: " & TBLista_aviso_diario2.PageCount
ElseIf TBLista_aviso_diario2.AbsolutePage = adPosEOF Then
        lblPaginas2.Caption = "Página: " & TBLista_aviso_diario2.PageCount & " de: " & TBLista_aviso_diario2.PageCount
    Else
        lblPaginas2.Caption = "Página: " & TBLista_aviso_diario2.AbsolutePage - 1 & " de: " & TBLista_aviso_diario2.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaManutencao()
On Error GoTo tratar_erro

Lista3.ListItems.Clear
lblRegistros3.Caption = "Nº de registros: 0"
lblPaginas3.Caption = "Página: 0 de: 0"
Set TBLista_aviso_diario3 = CreateObject("adodb.recordset")
StrSql = "Select M.codigo, M.Tipo, M.IDMaquina, M.Descricao, MD.data, MD.Dias_proxima from Manutencao_data MD INNER JOIN Manutencao M on MD.idManutencao = M.codigo where (MD.Data <= '" & Format(Date, "Short Date") & "' or (MD.Data + MD.Dias_proxima) <= '" & Format(Date, "Short Date") & "') and MD.Status = 'ABERTA' and M.Tipo <> 'C'"

'Debug.print StrSql


TBLista_aviso_diario3.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLista_aviso_diario3.EOF = False Then
    SSTab1.TabVisible(3) = True
    ProcExibePagina3 (1)
Else
    SSTab1.TabVisible(3) = False
    ContadorTAB = ContadorTAB - 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina3(Pagina)
On Error GoTo tratar_erro

Lista3.ListItems.Clear
TBLista_aviso_diario3.PageSize = IIf(txtNreg3 = "", 30, txtNreg3)
TBLista_aviso_diario3.AbsolutePage = Pagina
TamanhoPagina = TBLista_aviso_diario3.PageSize
ContadorReg = 1

PBLista3.Min = 0
PBLista3.Max = FunVerifMaxPBListaPaginacao(TBLista_aviso_diario3.RecordCount - IIf(Pagina > 1, (TBLista_aviso_diario3.PageSize * (Pagina - 1)), 0), TBLista_aviso_diario3.PageSize)
PBLista3.Value = 1
Contador = 0
Do While TBLista_aviso_diario3.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista3.ListItems
        .Add , , TBLista_aviso_diario3!CODIGO
        Tipomanutencao = ""
        Select Case TBLista_aviso_diario3!Tipo
            Case "P": Tipomanutencao = "Preventiva"
            Case "S": Tipomanutencao = "Solicitação"
        End Select
        .Item(.Count).SubItems(1) = Tipomanutencao
        .Item(.Count).SubItems(2) = TBLista_aviso_diario3!IDMaquina
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLista_aviso_diario3!Descricao), "", TBLista_aviso_diario3!Descricao)
        .Item(.Count).SubItems(4) = Format(TBLista_aviso_diario3!Data, "dd/mm/yy")
        .Item(.Count).SubItems(5) = Format(TBLista_aviso_diario3!Data + TBLista_aviso_diario3!Dias_proxima, "dd/mm/yy")
    End With
    TBLista_aviso_diario3.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista3.Value = Contador
Loop
lblRegistros3.Caption = "Nº de registros: " & TBLista_aviso_diario3.RecordCount
If TBLista_aviso_diario3.AbsolutePage = adPosBOF Then
   lblPaginas3.Caption = "Página: 1 de: " & TBLista_aviso_diario3.PageCount
ElseIf TBLista_aviso_diario3.AbsolutePage = adPosEOF Then
        lblPaginas3.Caption = "Página: " & TBLista_aviso_diario3.PageCount & " de: " & TBLista_aviso_diario3.PageCount
    Else
        lblPaginas3.Caption = "Página: " & TBLista_aviso_diario3.AbsolutePage - 1 & " de: " & TBLista_aviso_diario3.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaRM()
On Error GoTo tratar_erro

Lista4.ListItems.Clear
lblRegistros4.Caption = "Nº de registros: 0"
lblPaginas4.Caption = "Página: 0 de: 0"
Set TBLista_aviso_diario4 = CreateObject("adodb.recordset")
'TBLista_aviso_diario4.Open "Select RM.requisicao, RML.* from (Requisicao_materiais RM INNER JOIN Requisicao_materiais_lista RML ON RM.ID = RML.IDrequisicao) LEFT JOIN Usuarios_Setor_Responsavel ON Usuarios_Setor_Responsavel.ID_CC = RML.ID_CC where Year(RM.data) >= '2012' and RM.Status = 'REQUISIT.' and Usuarios_Setor_Responsavel.Responsavel_CC = '" & pubUsuario & "' and RML.Data_autorizacao is null order by RM.requisicao, RML.idlista desc", Conexao, adOpenKeyset, adLockOptimistic
'TBLista_aviso_diario4.Open "Select RM.requisicao, RML.* from (Requisicao_materiais RM INNER JOIN Requisicao_materiais_lista RML ON RM.ID = RML.IDrequisicao) LEFT JOIN Usuarios_Setor_Responsavel ON Usuarios_Setor_Responsavel.ID_CC = RML.ID_CC where (RM.Status = 'REQUISIT.' and Usuarios_Setor_Responsavel.Responsavel_CC = '" & pubUsuario & "' and RML.Data_autorizacao is null) or  (RML.Data_autorizacao IS NULL AND RML.Status = 'REQUISIT.' and RML.ID_CC IS NULL) order by RM.requisicao, RML.idlista desc", Conexao, adOpenKeyset, adLockOptimistic
TBLista_aviso_diario4.Open "Select RM.requisicao, RML.* from (Requisicao_materiais RM INNER JOIN Requisicao_materiais_lista RML ON RM.ID = RML.IDrequisicao) LEFT JOIN Usuarios_Setor_Responsavel ON Usuarios_Setor_Responsavel.ID_CC = RML.ID_CC where RML.Status = 'REQUISIT.' and RML.Data_autorizacao is null and (Usuarios_Setor_Responsavel.Responsavel_CC = '" & pubUsuario & "' or RML.ID_CC IS NULL) order by RM.requisicao, RML.idlista desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLista_aviso_diario4.EOF = False Then
    SSTab1.TabVisible(4) = True
    ProcExibePagina4 (1)
Else
    SSTab1.TabVisible(4) = False
    ContadorTAB = ContadorTAB - 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina4(Pagina)
On Error GoTo tratar_erro

Lista4.ListItems.Clear
TBLista_aviso_diario4.PageSize = IIf(txtNreg4 = "", 30, txtNreg4)
TBLista_aviso_diario4.AbsolutePage = Pagina
TamanhoPagina = TBLista_aviso_diario4.PageSize
ContadorReg = 1

PBLista4.Min = 0
PBLista4.Max = FunVerifMaxPBListaPaginacao(TBLista_aviso_diario4.RecordCount - IIf(Pagina > 1, (TBLista_aviso_diario4.PageSize * (Pagina - 1)), 0), TBLista_aviso_diario4.PageSize)
PBLista4.Value = 1
Contador = 0
Do While TBLista_aviso_diario4.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista4.ListItems
        .Add , , TBLista_aviso_diario4!IDRequisicao
        .Item(.Count).SubItems(1) = TBLista_aviso_diario4!requisicao
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLista_aviso_diario4!Desenho), "", TBLista_aviso_diario4!Desenho)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLista_aviso_diario4!Descricao), "", TBLista_aviso_diario4!Descricao)
    End With
    TBLista_aviso_diario4.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista4.Value = Contador
Loop
lblRegistros4.Caption = "Nº de registros: " & TBLista_aviso_diario4.RecordCount
If TBLista_aviso_diario4.AbsolutePage = adPosBOF Then
   lblPaginas4.Caption = "Página: 1 de: " & TBLista_aviso_diario4.PageCount
ElseIf TBLista_aviso_diario4.AbsolutePage = adPosEOF Then
        lblPaginas4.Caption = "Página: " & TBLista_aviso_diario4.PageCount & " de: " & TBLista_aviso_diario4.PageCount
    Else
        lblPaginas4.Caption = "Página: " & TBLista_aviso_diario4.AbsolutePage - 1 & " de: " & TBLista_aviso_diario4.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaNecessidadeCompras()
On Error GoTo tratar_erro

Lista5.ListItems.Clear
lblRegistros5.Caption = "Nº de registros: 0"
lblPaginas5.Caption = "Página: 0 de: 0"
Set TBLista_aviso_diario5 = CreateObject("adodb.recordset")
TBLista_aviso_diario5.Open "Select * from Estoque_necessidade_resumido where Compras = 'True' and Necessidade > 0 order by Desenho", Conexao, adOpenKeyset, adLockReadOnly
If TBLista_aviso_diario5.EOF = False Then
    SSTab1.TabVisible(5) = True
    ProcExibePagina5 (1)
Else
    SSTab1.TabVisible(5) = False
    ContadorTAB = ContadorTAB - 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina5(Pagina)
On Error GoTo tratar_erro

Lista5.ListItems.Clear
TBLista_aviso_diario5.PageSize = IIf(txtNreg5 = "", 30, txtNreg5)
TBLista_aviso_diario5.AbsolutePage = Pagina
TamanhoPagina = TBLista_aviso_diario5.PageSize
ContadorReg = 1

PBLista5.Min = 0
PBLista5.Max = FunVerifMaxPBListaPaginacao(TBLista_aviso_diario5.RecordCount - IIf(Pagina > 1, (TBLista_aviso_diario5.PageSize * (Pagina - 1)), 0), TBLista_aviso_diario5.PageSize)
PBLista5.Value = 1
Contador = 0
Do While TBLista_aviso_diario5.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista5.ListItems
        .Add , , TBLista_aviso_diario5!Desenho
        .Item(.Count).SubItems(1) = TBLista_aviso_diario5!Desenho
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLista_aviso_diario5!Descricao), "", TBLista_aviso_diario5!Descricao)
        .Item(.Count).SubItems(3) = Format(TBLista_aviso_diario5!Necessidade, "###,##0.0000")
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLista_aviso_diario5!Unidade), "", TBLista_aviso_diario5!Unidade)
    End With
    TBLista_aviso_diario5.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista5.Value = Contador
Loop
lblRegistros5.Caption = "Nº de registros: " & TBLista_aviso_diario5.RecordCount
If TBLista_aviso_diario5.AbsolutePage = adPosBOF Then
   lblPaginas5.Caption = "Página: 1 de: " & TBLista_aviso_diario5.PageCount
ElseIf TBLista_aviso_diario5.AbsolutePage = adPosEOF Then
        lblPaginas5.Caption = "Página: " & TBLista_aviso_diario5.PageCount & " de: " & TBLista_aviso_diario5.PageCount
    Else
        lblPaginas5.Caption = "Página: " & TBLista_aviso_diario5.AbsolutePage - 1 & " de: " & TBLista_aviso_diario5.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaNecessidadePCP()
On Error GoTo tratar_erro

Lista6.ListItems.Clear
lblRegistros6.Caption = "Nº de registros: 0"
lblPaginas6.Caption = "Página: 0 de: 0"
Set TBLista_aviso_diario6 = CreateObject("adodb.recordset")
TBLista_aviso_diario6.Open "Select * from Estoque_necessidade_resumido where Producao = 'True' and Necessidade > 0 order by Desenho", Conexao, adOpenKeyset, adLockReadOnly
If TBLista_aviso_diario6.EOF = False Then
    SSTab1.TabVisible(6) = True
    ProcExibePagina6 (1)
Else
    SSTab1.TabVisible(6) = False
    ContadorTAB = ContadorTAB - 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina6(Pagina)
On Error GoTo tratar_erro

Lista6.ListItems.Clear
TBLista_aviso_diario6.PageSize = IIf(txtNreg6 = "", 30, txtNreg6)
TBLista_aviso_diario6.AbsolutePage = Pagina
TamanhoPagina = TBLista_aviso_diario6.PageSize
ContadorReg = 1

PBLista6.Min = 0
PBLista6.Max = FunVerifMaxPBListaPaginacao(TBLista_aviso_diario6.RecordCount - IIf(Pagina > 1, (TBLista_aviso_diario6.PageSize * (Pagina - 1)), 0), TBLista_aviso_diario6.PageSize)
PBLista6.Value = 1
Contador = 0
Do While TBLista_aviso_diario6.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista6.ListItems
        .Add , , TBLista_aviso_diario6!Desenho
        .Item(.Count).SubItems(1) = TBLista_aviso_diario6!Desenho
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLista_aviso_diario6!Descricao), "", TBLista_aviso_diario6!Descricao)
        .Item(.Count).SubItems(3) = Format(TBLista_aviso_diario6!Necessidade, "###,##0.0000")
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLista_aviso_diario6!Unidade), "", TBLista_aviso_diario6!Unidade)
    End With
    TBLista_aviso_diario6.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista6.Value = Contador
Loop
lblRegistros6.Caption = "Nº de registros: " & TBLista_aviso_diario6.RecordCount
If TBLista_aviso_diario6.AbsolutePage = adPosBOF Then
   lblPaginas6.Caption = "Página: 1 de: " & TBLista_aviso_diario6.PageCount
ElseIf TBLista_aviso_diario6.AbsolutePage = adPosEOF Then
        lblPaginas6.Caption = "Página: " & TBLista_aviso_diario6.PageCount & " de: " & TBLista_aviso_diario6.PageCount
    Else
        lblPaginas6.Caption = "Página: " & TBLista_aviso_diario6.AbsolutePage - 1 & " de: " & TBLista_aviso_diario6.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaNecessidadeEstoque()
On Error GoTo tratar_erro

Lista7.ListItems.Clear
lblRegistros7.Caption = "Nº de registros: 0"
lblPaginas7.Caption = "Página: 0 de: 0"
Set TBLista_aviso_diario7 = CreateObject("adodb.recordset")
TBLista_aviso_diario7.Open "Select * from Estoque_necessidade_resumido where Necessidade > 0 order by Desenho", Conexao, adOpenKeyset, adLockReadOnly
If TBLista_aviso_diario7.EOF = False Then
    SSTab1.TabVisible(7) = True
    ProcExibePagina7 (1)
Else
    SSTab1.TabVisible(7) = False
    ContadorTAB = ContadorTAB - 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina7(Pagina)
On Error GoTo tratar_erro

Lista7.ListItems.Clear
TBLista_aviso_diario7.PageSize = IIf(txtNreg7 = "", 30, txtNreg7)
TBLista_aviso_diario7.AbsolutePage = Pagina
TamanhoPagina = TBLista_aviso_diario7.PageSize
ContadorReg = 1

PBLista7.Min = 0
PBLista7.Max = FunVerifMaxPBListaPaginacao(TBLista_aviso_diario7.RecordCount - IIf(Pagina > 1, (TBLista_aviso_diario7.PageSize * (Pagina - 1)), 0), TBLista_aviso_diario7.PageSize)
PBLista7.Value = 1
Contador = 0
Do While TBLista_aviso_diario7.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista7.ListItems
        .Add , , TBLista_aviso_diario7!Desenho
        .Item(.Count).SubItems(1) = TBLista_aviso_diario7!Desenho
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLista_aviso_diario7!Descricao), "", TBLista_aviso_diario7!Descricao)
        .Item(.Count).SubItems(3) = Format(TBLista_aviso_diario7!Necessidade, "###,##0.0000")
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLista_aviso_diario7!Unidade), "", TBLista_aviso_diario7!Unidade)
    End With
    TBLista_aviso_diario7.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista7.Value = Contador
Loop
lblRegistros7.Caption = "Nº de registros: " & TBLista_aviso_diario7.RecordCount
If TBLista_aviso_diario7.AbsolutePage = adPosBOF Then
   lblPaginas7.Caption = "Página: 1 de: " & TBLista_aviso_diario7.PageCount
ElseIf TBLista_aviso_diario7.AbsolutePage = adPosEOF Then
        lblPaginas7.Caption = "Página: " & TBLista_aviso_diario7.PageCount & " de: " & TBLista_aviso_diario7.PageCount
    Else
        lblPaginas7.Caption = "Página: " & TBLista_aviso_diario7.AbsolutePage - 1 & " de: " & TBLista_aviso_diario7.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaFaturar()
On Error GoTo tratar_erro

txtObs = ""
FrameObs.Enabled = False
Lista8.ListItems.Clear
lblRegistros8.Caption = "Nº de registros: 0"
lblPaginas8.Caption = "Página: 0 de: 0"
Set TBLista_aviso_diario8 = CreateObject("adodb.recordset")
TBLista_aviso_diario8.Open "Select NF.ID, E.Empresa, NF.dt_DataEmissao, NF.int_NotaFiscal, NF.TipoNF, NF.Serie, NF.Id_Int_Cliente, NF.txt_Razao_Nome, NF.Obs from tbl_Dados_Nota_Fiscal NF INNER JOIN Empresa E ON NF.ID_empresa = E.Codigo where NF.Aplicacao = 'P' and NF.DtValidacaoOF IS NOT NULL and NF.int_NotaFiscal IS NULL order by NF.ID", Conexao, adOpenKeyset, adLockReadOnly
If TBLista_aviso_diario8.EOF = False Then
    SSTab1.TabVisible(8) = True
    ProcExibePagina8 (1)
Else
    SSTab1.TabVisible(8) = False
    ContadorTAB = ContadorTAB - 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaTerceiros()
On Error GoTo tratar_erro

txtObs_Terceiros = ""
frameOBS_Terceiros.Enabled = False
Lista_Terceiros.ListItems.Clear
lblRegistros10.Caption = "Nº de registros: 0"
lblPaginas10.Caption = "Página: 0 de: 0"
Set TBLista_aviso_diarioTerceiros = CreateObject("adodb.recordset")
TBLista_aviso_diarioTerceiros.Open "Select A.Dias_Terceiros from Usuarios U INNER JOIN Acessos A on U.IDusuario = A.IDusuario where U.IDusuario = " & pubIDUsuario & " and A.Dias_Terceiros IS NOT NULL and A.Acesso = 'Avisos diário/Terceiros'", Conexao, adOpenKeyset, adLockReadOnly
If TBLista_aviso_diarioTerceiros.EOF = False Then
    Dataini = Date + TBLista_aviso_diarioTerceiros!Dias_Terceiros
    Set TBLista_aviso_diarioTerceiros = CreateObject("adodb.recordset")
    TBLista_aviso_diarioTerceiros.Open "SELECT CP.IDPedido, CP.Pedido, CP.Fornecedor, CPL.IDlista, CPL.Desenho, CPL.Descricao, CPL.Prazo, SUM(ECR.Recebido) AS Total, CPL.Quant_Comp, CPL.Ordem, CPL.OS, CPL.Obs_AvisoDiario FROM (Compras_pedido CP INNER JOIN Compras_pedido_lista CPL ON CP.IDPedido = CPL.IDPedido) LEFT JOIN Estoque_controle_recebimento ECR ON ECR.IdLista = CPL.IdLista WHERE CPL.OS IS NOT NULL AND (CPL.Status_Item = 'N_RECEBIDO' or CPL.Status_Item = 'PARCIAL') GROUP BY CP.IDPedido, CP.Pedido, CP.Fornecedor, CPL.IDlista, CPL.Desenho, CPL.Descricao, CPL.Prazo, CPL.Quant_Comp, CPL.Ordem, CPL.OS, CPL.Obs_AvisoDiario HAVING CPL.Prazo <= '" & Format(Dataini, "Short Date") & "' ORDER BY CP.IDPedido", Conexao, adOpenKeyset, adLockReadOnly
    If TBLista_aviso_diarioTerceiros.EOF = False Then
        SSTab1.TabVisible(12) = True
        ProcExibePaginaTerceiros (1)
    Else
Proximo:
        SSTab1.TabVisible(12) = False
        ContadorTAB = ContadorTAB - 1
    End If
Else
    GoTo Proximo
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaPedidoAtraso()
On Error GoTo tratar_erro

Lista_pedidosAtraso.ListItems.Clear
lblRegistros11.Caption = "Nº de registros: 0"
lblPaginas11.Caption = "Página: 0 de: 0"
Set TBLista_aviso_diarioPedidoAtraso = CreateObject("adodb.recordset")
TBLista_aviso_diarioPedidoAtraso.Open "SELECT CP.IDpedido, CP.Pedido, CP.Fornecedor, CPL.IdLista, CPL.Desenho, CPL.Descricao, CPL.Prazo, SUM(ECR.Recebido) AS Total, CPL.Quant_Comp FROM (Compras_pedido CP INNER JOIN Compras_pedido_lista CPL ON CP.IDPedido = CPL.IDPedido) LEFT JOIN Estoque_controle_recebimento ECR ON ECR.IdLista = CPL.IdLista WHERE CPL.Remessa = 'False' and CPL.OS IS NULL and (CPL.Status_Item = 'N_RECEBIDO' or CPL.Status_Item = 'PARCIAL') GROUP BY CP.IDPedido, CP.Pedido, CP.Fornecedor, CPL.IdLista, CPL.Desenho, CPL.Descricao, CPL.Prazo, CPL.Quant_Comp HAVING CPL.Prazo < '" & Format(Date, "Short Date") & "' ORDER BY CP.IDPedido", Conexao, adOpenKeyset, adLockReadOnly
If TBLista_aviso_diarioPedidoAtraso.EOF = False Then
    SSTab1.TabVisible(13) = True
    ProcExibePaginaPedidoAtraso (1)
Else
    SSTab1.TabVisible(13) = False
    ContadorTAB = ContadorTAB - 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaInstrumentos()
On Error GoTo tratar_erro
Dim Data_instrumento As Date

Lista_instrumentos.ListItems.Clear
lblRegistros_instrumentos.Caption = "Nº de registros: 0"
lblPaginas_Instrumentos.Caption = "Página: 0 de: 0"

Data_instrumento = Date + 5
CamposFiltro = "I.CODIGO, I.Numero, EC.Numero_serie, I.Descricao, I.Data_Aquisicao, I.Fabricante, A.Aferido, A.Orgao, A.Proxima_afericao, A.Certificado"

Set TBLista_aviso_diario_Instrumentos = CreateObject("adodb.recordset")
TBLista_aviso_diario_Instrumentos.Open "Select " & CamposFiltro & " from (((Instrumentos I LEFT JOIN Estoque_controle EC ON EC.IDestoque = I.IDestoque) LEFT JOIN Projproduto P ON P.Desenho = I.Numero) LEFT JOIN item_aplicacoes IA ON IA.Codproduto = P.Codproduto) INNER JOIN Afericao A ON I.Codigo = A.ID_inst and I.ID_ultima_afericao = A.Codigo where A.proxima_afericao <= '" & Format(Data_instrumento, "Short Date") & "' order by I.numero", Conexao, adOpenKeyset, adLockOptimistic
If TBLista_aviso_diario_Instrumentos.EOF = False Then
    SSTab1.TabVisible(14) = True
    ProcExibePaginaInstrumento (1)
Else
    SSTab1.TabVisible(14) = False
    ContadorTAB = ContadorTAB - 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaNaoConforme()
On Error GoTo tratar_erro

Lista_NaoConforme.ListItems.Clear
lblRegistros_NaoConforme.Caption = "Nº de registros: 0"
lblPaginas_NaoConforme.Caption = "Página: 0 de: 0"
Set TBLista_aviso_diario_NaoConformidade = CreateObject("adodb.recordset")
TBLista_aviso_diario_NaoConformidade.Open "SELECT NC.CODIGO, NC.Ordem, OS.Fase, NC.OS, NC.LOTE, NC.TTNC, NC.Data, NC.Operador, NC.PARECERCQ FROM cq_nc_fabrica NC LEFT JOIN ordemservico OS ON NC.OS = OS.idproducao WHERE NC.analizada = 'False' ORDER BY NC.OS", Conexao, adOpenKeyset, adLockReadOnly
If TBLista_aviso_diario_NaoConformidade.EOF = False Then
    SSTab1.TabVisible(15) = True
    ProcExibePaginaNaoConforme (1)
Else
    SSTab1.TabVisible(15) = False
    ContadorTAB = ContadorTAB - 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaAnalise()
On Error GoTo tratar_erro

Set TBLista_aviso_diarioAnalise = CreateObject("adodb.recordset")
TBLista_aviso_diarioAnalise.Open "Select ID from Vendas_analise where Status = 'ABERTA EM ANALISE' and (DtValidacao_Engenharia IS NULL or DtValidacao_Processo IS NULL or DtValidacao_PCP IS NULL or DtValidacao_Qualidade IS NULL or DtValidacao_Compras IS NULL) order by ID", Conexao, adOpenKeyset, adLockReadOnly
If TBLista_aviso_diarioAnalise.EOF = False Then
    SSTab1.TabVisible(11) = True
    ProcCarregaListaAnalise
Else
    SSTab1.TabVisible(11) = False
    ContadorTAB = ContadorTAB - 1
End If
TBLista_aviso_diarioAnalise.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaAnalise()
On Error GoTo tratar_erro

Lista_Analise_Compras.ListItems.Clear
Lista_Analise_Engenharia.ListItems.Clear
Lista_Analise_PCP.ListItems.Clear
Lista_Analise_Processo.ListItems.Clear
Lista_Analise_Qualidade.ListItems.Clear

Select Case Tab_Analise.Tab
    Case 0: TextoFiltro = "DtValidacao_Engenharia IS NULL"
    Case 1: TextoFiltro = "DtValidacao_Processo IS NULL"
    Case 2: TextoFiltro = "DtValidacao_PCP IS NULL"
    Case 3: TextoFiltro = "DtValidacao_Qualidade IS NULL"
    Case 4: TextoFiltro = "DtValidacao_Compras IS NULL"
End Select

lblRegistros9.Caption = "Nº de registros: 0"
lblPaginas9.Caption = "Página: 0 de: 0"
Set TBLista_aviso_diarioAnalise = CreateObject("adodb.recordset")
TBLista_aviso_diarioAnalise.Open "Select ID, Nanalise, Revisao, Data, Codinterno, Descricao, Cliente from Vendas_analise where " & TextoFiltro & " and Status = 'ABERTA EM ANALISE' order by ID", Conexao, adOpenKeyset, adLockReadOnly
If TBLista_aviso_diarioAnalise.EOF = False Then
    Select Case Tab_Analise.Tab
        Case 0: ProcExibePaginaAnalise 1, Lista_Analise_Engenharia
        Case 1: ProcExibePaginaAnalise 1, Lista_Analise_Processo
        Case 2: ProcExibePaginaAnalise 1, Lista_Analise_PCP
        Case 3: ProcExibePaginaAnalise 1, Lista_Analise_Qualidade
        Case 4: ProcExibePaginaAnalise 1, Lista_Analise_Compras
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePaginaAnalise(Pagina, ListaAnalise As ListView)
On Error GoTo tratar_erro

ListaAnalise.ListItems.Clear
TBLista_aviso_diarioAnalise.PageSize = IIf(txtNreg9 = "", 30, txtNreg9)
TBLista_aviso_diarioAnalise.AbsolutePage = Pagina
TamanhoPagina = TBLista_aviso_diarioAnalise.PageSize
ContadorReg = 1

If TBLista_aviso_diarioAnalise.AbsolutePage = adPosBOF Then
    PBLista11.Min = 0
    PBLista11.Max = TBLista_aviso_diarioAnalise.PageSize
    PBLista11.Value = 1
    Contador = 0
End If
Do While TBLista_aviso_diarioAnalise.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListaAnalise.ListItems
        .Add , , TBLista_aviso_diarioAnalise!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLista_aviso_diarioAnalise!Nanalise), "", TBLista_aviso_diarioAnalise!Nanalise)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLista_aviso_diarioAnalise!Revisao), 0, TBLista_aviso_diarioAnalise!Revisao)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLista_aviso_diarioAnalise!Data), "", Format(TBLista_aviso_diarioAnalise!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLista_aviso_diarioAnalise!Codinterno), "", TBLista_aviso_diarioAnalise!Codinterno)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLista_aviso_diarioAnalise!Descricao), "", TBLista_aviso_diarioAnalise!Descricao)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLista_aviso_diarioAnalise!Cliente), "", TBLista_aviso_diarioAnalise!Cliente)
    End With
    TBLista_aviso_diarioAnalise.MoveNext
    ContadorReg = ContadorReg + 1
    If TBLista_aviso_diarioAnalise.AbsolutePage = adPosBOF Then
        Contador = Contador + 1
        PBLista11.Value = Contador
    End If
Loop
lblRegistros9.Caption = "Nº de registros: " & TBLista_aviso_diarioAnalise.RecordCount
If TBLista_aviso_diarioAnalise.AbsolutePage = adPosBOF Then
   lblPaginas9.Caption = "Página: 1 de: " & TBLista_aviso_diarioAnalise.PageCount
ElseIf TBLista_aviso_diarioAnalise.AbsolutePage = adPosEOF Then
        lblPaginas9.Caption = "Página: " & TBLista_aviso_diarioAnalise.PageCount & " de: " & TBLista_aviso_diarioAnalise.PageCount
    Else
        lblPaginas9.Caption = "Página: " & TBLista_aviso_diarioAnalise.AbsolutePage - 1 & " de: " & TBLista_aviso_diarioAnalise.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePaginaTerceiros(Pagina)
On Error GoTo tratar_erro

Lista_Terceiros.ListItems.Clear
TBLista_aviso_diarioTerceiros.PageSize = IIf(txtNreg10 = "", 30, txtNreg10)
TBLista_aviso_diarioTerceiros.AbsolutePage = Pagina
TamanhoPagina = TBLista_aviso_diarioTerceiros.PageSize
ContadorReg = 1

If TBLista_aviso_diarioTerceiros.AbsolutePage = adPosBOF Then
    PBLista12.Min = 0
    PBLista12.Max = TBLista_aviso_diarioTerceiros.PageSize
    PBLista12.Value = 1
    Contador = 0
End If
Do While TBLista_aviso_diarioTerceiros.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista_Terceiros.ListItems
        .Add , , TBLista_aviso_diarioTerceiros!IDlista
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLista_aviso_diarioTerceiros!IDpedido), "", TBLista_aviso_diarioTerceiros!IDpedido)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLista_aviso_diarioTerceiros!Pedido), "", TBLista_aviso_diarioTerceiros!Pedido)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLista_aviso_diarioTerceiros!Fornecedor), "", TBLista_aviso_diarioTerceiros!Fornecedor)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLista_aviso_diarioTerceiros!Desenho), "", TBLista_aviso_diarioTerceiros!Desenho)
        
        Qtde = IIf(IsNull(TBLista_aviso_diarioTerceiros!Quant_Comp), 0, TBLista_aviso_diarioTerceiros!Quant_Comp) - IIf(IsNull(TBLista_aviso_diarioTerceiros!Total), 0, TBLista_aviso_diarioTerceiros!Total)
        .Item(.Count).SubItems(5) = Format(Qtde, "###,##0.0000")
        
        TextoData = ""
        Set TBLista_aviso_diarioTerceirosData = CreateObject("adodb.recordset")
        TBLista_aviso_diarioTerceirosData.Open "Select E.Data, E.Saida from tbl_Detalhes_Nota_pedidos DNP INNER JOIN Estoque_movimentacao E ON DNP.ID_prod_NF = E.ID_prod_NF where DNP.Codinterno = '" & TBLista_aviso_diarioTerceiros!Desenho & "' and DNP.ID_carteira = " & TBLista_aviso_diarioTerceiros!IDlista & " group by E.Data, E.Saida order by E.Data", Conexao, adOpenKeyset, adLockReadOnly
        Do While TBLista_aviso_diarioTerceirosData.EOF = False
            If TextoData = "" Then
                TextoData = IIf(IsNull(TBLista_aviso_diarioTerceirosData!Data), "", Format(TBLista_aviso_diarioTerceirosData!Data, "dd/mm/yy")) & " - " & IIf(IsNull(TBLista_aviso_diarioTerceirosData!Saida), "", Format(TBLista_aviso_diarioTerceirosData!Saida, "###,##0.0000"))
            Else
                TextoData = TextoData & " | " & IIf(IsNull(TBLista_aviso_diarioTerceirosData!Data), "", Format(TBLista_aviso_diarioTerceirosData!Data, "dd/mm/yy")) & " - " & IIf(IsNull(TBLista_aviso_diarioTerceirosData!Saida), "", Format(TBLista_aviso_diarioTerceirosData!Saida, "###,##0.0000"))
            End If
            TBLista_aviso_diarioTerceirosData.MoveNext
        Loop
        TBLista_aviso_diarioTerceirosData.Close
        .Item(.Count).SubItems(6) = TextoData
        
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLista_aviso_diarioTerceiros!Prazo), "", Format(TBLista_aviso_diarioTerceiros!Prazo, "dd/mm/yy"))
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLista_aviso_diarioTerceiros!Ordem), "", TBLista_aviso_diarioTerceiros!Ordem)
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLista_aviso_diarioTerceiros!OS), "", TBLista_aviso_diarioTerceiros!OS)
        .Item(.Count).SubItems(10) = IIf(IsNull(TBLista_aviso_diarioTerceiros!Obs_AvisoDiario), "", TBLista_aviso_diarioTerceiros!Obs_AvisoDiario)
    End With
    TBLista_aviso_diarioTerceiros.MoveNext
    ContadorReg = ContadorReg + 1
    If TBLista_aviso_diarioTerceiros.AbsolutePage = adPosBOF Then
        Contador = Contador + 1
        PBLista12.Value = Contador
    End If
Loop
lblRegistros10.Caption = "Nº de registros: " & TBLista_aviso_diarioTerceiros.RecordCount
If TBLista_aviso_diarioTerceiros.AbsolutePage = adPosBOF Then
   lblPaginas10.Caption = "Página: 1 de: " & TBLista_aviso_diarioTerceiros.PageCount
ElseIf TBLista_aviso_diarioTerceiros.AbsolutePage = adPosEOF Then
        lblPaginas10.Caption = "Página: " & TBLista_aviso_diarioTerceiros.PageCount & " de: " & TBLista_aviso_diarioTerceiros.PageCount
    Else
        lblPaginas10.Caption = "Página: " & TBLista_aviso_diarioTerceiros.AbsolutePage - 1 & " de: " & TBLista_aviso_diarioTerceiros.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePaginaPedidoAtraso(Pagina)
On Error GoTo tratar_erro

Lista_pedidosAtraso.ListItems.Clear
TBLista_aviso_diarioPedidoAtraso.PageSize = IIf(txtNreg11 = "", 30, txtNreg11)
TBLista_aviso_diarioPedidoAtraso.AbsolutePage = Pagina
TamanhoPagina = TBLista_aviso_diarioPedidoAtraso.PageSize
ContadorReg = 1

If TBLista_aviso_diarioPedidoAtraso.AbsolutePage = adPosBOF Then
    PBLista13.Min = 0
    PBLista13.Max = TBLista_aviso_diarioPedidoAtraso.PageSize
    PBLista13.Value = 1
    Contador = 0
End If
Do While TBLista_aviso_diarioPedidoAtraso.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista_pedidosAtraso.ListItems
        .Add , , TBLista_aviso_diarioPedidoAtraso!IDlista
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLista_aviso_diarioPedidoAtraso!IDpedido), "", TBLista_aviso_diarioPedidoAtraso!IDpedido)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLista_aviso_diarioPedidoAtraso!Pedido), "", TBLista_aviso_diarioPedidoAtraso!Pedido)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLista_aviso_diarioPedidoAtraso!Fornecedor), "", TBLista_aviso_diarioPedidoAtraso!Fornecedor)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLista_aviso_diarioPedidoAtraso!Desenho), "", TBLista_aviso_diarioPedidoAtraso!Desenho)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLista_aviso_diarioPedidoAtraso!Quant_Comp), 0, Format(TBLista_aviso_diarioPedidoAtraso!Quant_Comp, "###,##0.0000"))
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLista_aviso_diarioPedidoAtraso!Total), 0, Format(TBLista_aviso_diarioPedidoAtraso!Total, "###,##0.0000"))
        Qtde = IIf(IsNull(TBLista_aviso_diarioPedidoAtraso!Quant_Comp), 0, TBLista_aviso_diarioPedidoAtraso!Quant_Comp) - IIf(IsNull(TBLista_aviso_diarioPedidoAtraso!Total), 0, TBLista_aviso_diarioPedidoAtraso!Total)
        .Item(.Count).SubItems(7) = Format(Qtde, "###,##0.0000")
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLista_aviso_diarioPedidoAtraso!Prazo), "", Format(TBLista_aviso_diarioPedidoAtraso!Prazo, "dd/mm/yy"))
    End With
    TBLista_aviso_diarioPedidoAtraso.MoveNext
    ContadorReg = ContadorReg + 1
    If TBLista_aviso_diarioPedidoAtraso.AbsolutePage = adPosBOF Then
        Contador = Contador + 1
        PBLista13.Value = Contador
    End If
Loop
lblRegistros11.Caption = "Nº de registros: " & TBLista_aviso_diarioPedidoAtraso.RecordCount
If TBLista_aviso_diarioPedidoAtraso.AbsolutePage = adPosBOF Then
   lblPaginas11.Caption = "Página: 1 de: " & TBLista_aviso_diarioPedidoAtraso.PageCount
ElseIf TBLista_aviso_diarioPedidoAtraso.AbsolutePage = adPosEOF Then
        lblPaginas11.Caption = "Página: " & TBLista_aviso_diarioPedidoAtraso.PageCount & " de: " & TBLista_aviso_diarioPedidoAtraso.PageCount
    Else
        lblPaginas11.Caption = "Página: " & TBLista_aviso_diarioPedidoAtraso.AbsolutePage - 1 & " de: " & TBLista_aviso_diarioPedidoAtraso.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina8(Pagina)
On Error GoTo tratar_erro

Lista8.ListItems.Clear
TBLista_aviso_diario8.PageSize = IIf(txtNreg8 = "", 30, txtNreg8)
TBLista_aviso_diario8.AbsolutePage = Pagina
TamanhoPagina = TBLista_aviso_diario8.PageSize
ContadorReg = 1

PBLista8.Min = 0
PBLista8.Max = FunVerifMaxPBListaPaginacao(TBLista_aviso_diario8.RecordCount - IIf(Pagina > 1, (TBLista_aviso_diario8.PageSize * (Pagina - 1)), 0), TBLista_aviso_diario8.PageSize)
PBLista8.Value = 1
Contador = 0
Do While TBLista_aviso_diario8.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista8.ListItems
        .Add , , TBLista_aviso_diario8!ID
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLista_aviso_diario8!Empresa), "", TBLista_aviso_diario8!Empresa)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLista_aviso_diario8!dt_DataEmissao), "", (Format(TBLista_aviso_diario8!dt_DataEmissao, "dd/mm/yy")))
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLista_aviso_diario8!ID), "", TBLista_aviso_diario8!ID)
        If IsNull(TBLista_aviso_diario8!TipoNF) = False Then
            If TBLista_aviso_diario8!TipoNF = "M1" Then TipoNF2 = "Produto(s)"
            If TBLista_aviso_diario8!TipoNF = "SA" Then TipoNF2 = "Serviço(s)"
            If TBLista_aviso_diario8!TipoNF = "M1SA" Then TipoNF2 = "Prod./Serv."
        End If
        .Item(.Count).SubItems(4) = TipoNF2
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLista_aviso_diario8!Serie), "", TBLista_aviso_diario8!Serie)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLista_aviso_diario8!Id_Int_Cliente), "", TBLista_aviso_diario8!Id_Int_Cliente)
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLista_aviso_diario8!txt_Razao_Nome), "", TBLista_aviso_diario8!txt_Razao_Nome)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLista_aviso_diario8!Obs), "", TBLista_aviso_diario8!Obs)
    End With
    TBLista_aviso_diario8.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista8.Value = Contador
Loop
lblRegistros8.Caption = "Nº de registros: " & TBLista_aviso_diario8.RecordCount
If TBLista_aviso_diario8.AbsolutePage = adPosBOF Then
   lblPaginas8.Caption = "Página: 1 de: " & TBLista_aviso_diario8.PageCount
ElseIf TBLista_aviso_diario8.AbsolutePage = adPosEOF Then
        lblPaginas8.Caption = "Página: " & TBLista_aviso_diario8.PageCount & " de: " & TBLista_aviso_diario8.PageCount
    Else
        lblPaginas8.Caption = "Página: " & TBLista_aviso_diario8.AbsolutePage - 1 & " de: " & TBLista_aviso_diario8.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePaginaInstrumento(Pagina)
On Error GoTo tratar_erro

Lista_instrumentos.ListItems.Clear
TBLista_aviso_diario_Instrumentos.PageSize = IIf(txtNreg_Instrumentos = "", 30, txtNreg_Instrumentos)
TBLista_aviso_diario_Instrumentos.AbsolutePage = Pagina
TamanhoPagina = TBLista_aviso_diario_Instrumentos.PageSize
ContadorReg = 1

PBLista_instrumentos.Min = 0
PBLista_instrumentos.Max = FunVerifMaxPBListaPaginacao(TBLista_aviso_diario_Instrumentos.RecordCount - IIf(Pagina > 1, (TBLista_aviso_diario_Instrumentos.PageSize * (Pagina - 1)), 0), TBLista_aviso_diario_Instrumentos.PageSize)
PBLista_instrumentos.Value = 1
Contador = 0
Do While TBLista_aviso_diario_Instrumentos.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista_instrumentos.ListItems
        .Add , , TBLista_aviso_diario_Instrumentos!CODIGO
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLista_aviso_diario_Instrumentos!Numero), "", TBLista_aviso_diario_Instrumentos!Numero)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLista_aviso_diario_Instrumentos!Numero_serie), "", TBLista_aviso_diario_Instrumentos!Numero_serie)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLista_aviso_diario_Instrumentos!Descricao), "", TBLista_aviso_diario_Instrumentos!Descricao)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLista_aviso_diario_Instrumentos!Data_Aquisicao), "", Format(TBLista_aviso_diario_Instrumentos!Data_Aquisicao, "dd/mm/yy"))
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLista_aviso_diario_Instrumentos!Fabricante), "", TBLista_aviso_diario_Instrumentos!Fabricante)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLista_aviso_diario_Instrumentos!Aferido), "", Format(TBLista_aviso_diario_Instrumentos!Aferido, "dd/mm/yy"))
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLista_aviso_diario_Instrumentos!Orgao), "", TBLista_aviso_diario_Instrumentos!Orgao)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLista_aviso_diario_Instrumentos!Proxima_afericao), "", Format(TBLista_aviso_diario_Instrumentos!Proxima_afericao, "dd/mm/yy"))
        .Item(.Count).SubItems(9) = IIf(IsNull(TBLista_aviso_diario_Instrumentos!Certificado), "", TBLista_aviso_diario_Instrumentos!Certificado)
    End With
    TBLista_aviso_diario_Instrumentos.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista_instrumentos.Value = Contador
Loop
lblRegistros_instrumentos.Caption = "Nº de registros: " & TBLista_aviso_diario_Instrumentos.RecordCount
If TBLista_aviso_diario_Instrumentos.AbsolutePage = adPosBOF Then
   lblPaginas_Instrumentos.Caption = "Página: 1 de: " & TBLista_aviso_diario_Instrumentos.PageCount
ElseIf TBLista_aviso_diario_Instrumentos.AbsolutePage = adPosEOF Then
        lblPaginas_Instrumentos.Caption = "Página: " & TBLista_aviso_diario_Instrumentos.PageCount & " de: " & TBLista_aviso_diario_Instrumentos.PageCount
    Else
        lblPaginas_Instrumentos.Caption = "Página: " & TBLista_aviso_diario_Instrumentos.AbsolutePage - 1 & " de: " & TBLista_aviso_diario_Instrumentos.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePaginaNaoConforme(Pagina)
On Error GoTo tratar_erro

Lista_NaoConforme.ListItems.Clear
TBLista_aviso_diario_NaoConformidade.PageSize = IIf(txtNreg_NaoConforme = "", 30, txtNreg_NaoConforme)
TBLista_aviso_diario_NaoConformidade.AbsolutePage = Pagina
TamanhoPagina = TBLista_aviso_diario_NaoConformidade.PageSize
ContadorReg = 1

PBLista_NaoConforme.Min = 0
PBLista_NaoConforme.Max = FunVerifMaxPBListaPaginacao(TBLista_aviso_diario_NaoConformidade.RecordCount - IIf(Pagina > 1, (TBLista_aviso_diario_NaoConformidade.PageSize * (Pagina - 1)), 0), TBLista_aviso_diario_NaoConformidade.PageSize)
PBLista_NaoConforme.Value = 1
Contador = 0
Do While TBLista_aviso_diario_NaoConformidade.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista_NaoConforme.ListItems
        .Add , , TBLista_aviso_diario_NaoConformidade!CODIGO
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLista_aviso_diario_NaoConformidade!Ordem), "", TBLista_aviso_diario_NaoConformidade!Ordem)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLista_aviso_diario_NaoConformidade!OS), "", TBLista_aviso_diario_NaoConformidade!OS)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLista_aviso_diario_NaoConformidade!Fase), "", TBLista_aviso_diario_NaoConformidade!Fase)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLista_aviso_diario_NaoConformidade!LOTE), "", TBLista_aviso_diario_NaoConformidade!LOTE)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLista_aviso_diario_NaoConformidade!TTNC), "", TBLista_aviso_diario_NaoConformidade!TTNC)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBLista_aviso_diario_NaoConformidade!Data), "", Format(TBLista_aviso_diario_NaoConformidade!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(7) = IIf(IsNull(TBLista_aviso_diario_NaoConformidade!Operador), "", TBLista_aviso_diario_NaoConformidade!Operador)
        .Item(.Count).SubItems(8) = IIf(IsNull(TBLista_aviso_diario_NaoConformidade!ParecerCQ), "", TBLista_aviso_diario_NaoConformidade!ParecerCQ)
    End With
    TBLista_aviso_diario_NaoConformidade.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista_NaoConforme.Value = Contador
Loop
lblRegistros_NaoConforme.Caption = "Nº de registros: " & TBLista_aviso_diario_NaoConformidade.RecordCount
If TBLista_aviso_diario_NaoConformidade.AbsolutePage = adPosBOF Then
   lblPaginas_NaoConforme.Caption = "Página: 1 de: " & TBLista_aviso_diario_NaoConformidade.PageCount
ElseIf TBLista_aviso_diario_NaoConformidade.AbsolutePage = adPosEOF Then
        lblPaginas_NaoConforme.Caption = "Página: " & TBLista_aviso_diario_NaoConformidade.PageCount & " de: " & TBLista_aviso_diario_NaoConformidade.PageCount
    Else
        lblPaginas_NaoConforme.Caption = "Página: " & TBLista_aviso_diario_NaoConformidade.AbsolutePage - 1 & " de: " & TBLista_aviso_diario_NaoConformidade.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procCarregaListaItensNota()
On Error GoTo tratar_erro

ListaItensNota.ListItems.Clear
Set TBLista_aviso_diarioItensNota = CreateObject("adodb.recordset")
TBLista_aviso_diarioItensNota.Open "Select Int_codigo, int_Cod_Produto, Txt_descricao, txt_Unid , int_Qtd from tbl_Detalhes_Nota where ID_Nota = " & Lista8.SelectedItem & " order by int_codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLista_aviso_diarioItensNota.EOF = False Then
    PBLista8.Min = 0
    PBLista8.Max = TBLista_aviso_diarioItensNota.RecordCount
    PBLista8.Value = 1
    Contador = 0
    Do While TBLista_aviso_diarioItensNota.EOF = False
        With ListaItensNota.ListItems
            .Add , , TBLista_aviso_diarioItensNota!Int_codigo
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLista_aviso_diarioItensNota!int_Cod_Produto), "", TBLista_aviso_diarioItensNota!int_Cod_Produto)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLista_aviso_diarioItensNota!Txt_descricao), "", TBLista_aviso_diarioItensNota!Txt_descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLista_aviso_diarioItensNota!txt_Unid), "", TBLista_aviso_diarioItensNota!txt_Unid)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLista_aviso_diarioItensNota!int_Qtd), "", Format(TBLista_aviso_diarioItensNota!int_Qtd, "###,##0.0000"))
            
            Set TBLista_aviso_diarioItensNota_pedido = CreateObject("adodb.recordset")
            TBLista_aviso_diarioItensNota_pedido.Open "Select VP.Ncotacao, VP.Revisao, VC.Prazofinal from (tbl_Detalhes_Nota_pedidos DNP INNER JOIN Vendas_carteira VC on DNP.ID_carteira = VC.Codigo) INNER JOIN vendas_proposta VP on VP.Cotacao = VC.cotacao where DNP.ID_prod_NF = " & TBLista_aviso_diarioItensNota!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
            If TBLista_aviso_diarioItensNota_pedido.EOF = False Then
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLista_aviso_diarioItensNota_pedido!Ncotacao), "", TBLista_aviso_diarioItensNota_pedido!Ncotacao)
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLista_aviso_diarioItensNota_pedido!Revisao), "", TBLista_aviso_diarioItensNota_pedido!Revisao)
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLista_aviso_diarioItensNota_pedido!PrazoFinal), "", Format(TBLista_aviso_diarioItensNota_pedido!PrazoFinal, "dd/mm/yy"))
            End If
            TBLista_aviso_diarioItensNota_pedido.Close
        End With
        TBLista_aviso_diarioItensNota.MoveNext
        Contador = Contador + 1
        PBLista8.Value = Contador
    Loop
End If
TBLista_aviso_diarioItensNota.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo tratar_erro

frmMDI.ProcVerificaAvisoDiario

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

Private Sub Lista2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista2, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista3, ColumnHeader

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

ProcOrdenaListView Lista5, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista6_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista6, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista7_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista7, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista8_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista8, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Formulario = "Financeiro/Contas a pagar"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With frmContas_Pagar
    Aviso_diario_utiliza_formularel = True
    .ProcConstruirFiltroPadrao "CP.IdIntConta = " & Lista.SelectedItem, "{tbl_ContasPagar.IdIntConta} = " & Lista.SelectedItem, True, True
    .lst_contas.ListItems.Clear
    .ProcCarregaLista (1)
    
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from tbl_ContasPagar where IdIntConta = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        .ProcLimpaCampos
        .ProcCarregaDados
    End If
    TBContas.Close
    Unload Me
End With
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista1_DblClick()
On Error GoTo tratar_erro

If Lista1.ListItems.Count = 0 Then Exit Sub
Formulario = "Financeiro/Contas a receber"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With frmContas_Receber
    Aviso_diario_utiliza_formularel = True
    .ProcConstruirFiltroPadrao "CR.IdIntConta = " & Lista1.SelectedItem, "{tbl_contas_receber.IdIntConta} = " & Lista1.SelectedItem, True, True
    .Lista.ListItems.Clear
    .ProcLimpaVariaveisCarregaLista
    .ProcCarregaLista (1)
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from tbl_contas_receber where idintconta = " & Lista1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        .ProcLimpaCampos
        .ProcCarregaDados
    End If
    TBProduto.Close
    Unload Me
End With
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista2_DblClick()
On Error GoTo tratar_erro

If Lista2.ListItems.Count = 0 Then Exit Sub
Formulario = "Outros/Solicitação"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With frmCompras_Requisicao
    .StrSql_solicitacao = "Select * from Compras_requisicao where ID_Requisicao = " & Lista2.SelectedItem
    .ProcCarregaLista_Req (1)
    
    Set TBCompras = CreateObject("adodb.recordset")
    TBCompras.Open "Select * from compras_requisicao where id_requisicao = " & Lista2.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras.EOF = False Then
        .ProcLimpaCampos
        .ProcAbrir
    End If
    TBCompras.Close
    Unload Me
End With
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista3_DblClick()
On Error GoTo tratar_erro

If Lista3.ListItems.Count = 0 Then Exit Sub
Formulario = "Manutenção/Equipamentos"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With frmManutencao
    Aviso_diario_utiliza_formularel = True
    .Sql_Manutencao_Localizar = "Select * from manutencao where Codigo = " & Lista3.SelectedItem
    .FormulaRel_Manutencao = "{manutencao.codigo} = " & Lista3.SelectedItem
    .ProcCarregaLista
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "select * from manutencao where Codigo = " & Lista3.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        .ProcDesabilitar
        .ProcLimpaCampos
        .ProcPuxaDados
    End If
    TBAbrir.Close
    Unload Me
End With
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista4_DblClick()
On Error GoTo tratar_erro

If Lista4.ListItems.Count = 0 Then Exit Sub
Formulario = "Estoque/Requisição de materiais"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With frmRequisicao_materiais
    .StrSql_Localizar_Requisicao = "Select RM.ID, RM.requisicao, RM.Data, RM.Responsavel, RM.Status, RM.DtValidacao, E.Empresa from Requisicao_materiais RM INNER JOIN Empresa E ON E.Codigo = RM.ID_empresa where ID = " & Lista4.SelectedItem & " group by RM.ID, RM.requisicao, RM.Data, RM.Responsavel, RM.Status, RM.DtValidacao, E.Empresa"
    .ProcCarregaLista (1)

    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Requisicao_materiais where id = " & Lista4.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        .ProcLimpaCampos
        .ProcPuxaDados
    End If
    TBAbrir.Close
    Unload Me
End With
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista5_DblClick()
On Error GoTo tratar_erro

If Lista5.ListItems.Count = 0 Then Exit Sub
Formulario = "Compras/Necessidade"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With Frm_necessidade
    Aviso_diario_utiliza_formularel = True
    Compras = True
    PCP_Ordem = False
    .StrSql_EstoqueNecessidade = "Select * from Estoque_necessidade_resumido where Desenho = '" & Lista5.SelectedItem & "' and Compras = 'True' and Necessidade > 0"
    .FormulaRel_EstoqueNecessidade = "{Estoque_necessidade_resumido.Desenho} = '" & Lista5.SelectedItem & "' and {Estoque_necessidade_resumido.Compras} = True and {Estoque_necessidade_resumido.Necessidade} > 0"
    .ProcCarregaLista
    Unload Me
End With
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista6_DblClick()
On Error GoTo tratar_erro

If Lista6.ListItems.Count = 0 Then Exit Sub
Formulario = "PCP/Necessidade"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With Frm_necessidade
    Aviso_diario_utiliza_formularel = True
    Compras = False
    PCP_Ordem = True
    .StrSql_EstoqueNecessidade = "Select * from Estoque_necessidade_resumido where Desenho = '" & Lista6.SelectedItem & "' and Producao = 'True' and Necessidade > 0"
    .FormulaRel_EstoqueNecessidade = "{Estoque_necessidade_resumido.Desenho} = '" & Lista6.SelectedItem & "' and {Estoque_necessidade_resumido.Producao} = True and {Estoque_necessidade_resumido.Necessidade} > 0"
    .ProcCarregaLista
    Unload Me
End With
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista7_DblClick()
On Error GoTo tratar_erro

If Lista7.ListItems.Count = 0 Then Exit Sub
Formulario = "Estoque/Necessidade"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

With Frm_necessidade
    Aviso_diario_utiliza_formularel = True
    Compras = False
    PCP_Ordem = False
    .StrSql_EstoqueNecessidade = "Select * from Estoque_necessidade_resumido where Desenho = '" & Lista7.SelectedItem & "' and Necessidade > 0"
    .FormulaRel_EstoqueNecessidade = "{Estoque_necessidade_resumido.Desenho} = '" & Lista7.SelectedItem & "' and {Estoque_necessidade_resumido.Necessidade} > 0"
    .ProcCarregaLista
    Unload Me
End With
Aviso_diario_utiliza_formularel = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub FlexGridOS_Click()
On Error GoTo tratar_erro
Dim lngImage As Boolean, L As Long, GroupoID As Long

With FlexGridOS
'    .Redraw = False
'    If .RowGroupHeader = True Then
'        GroupoID = .RowData()
'        lngImage = (.RowImage() = 1)
'        For L = 0 To .rows - 1
'            If Not .RowGroupHeader(L) Then
'                If .RowData(L) = GroupoID Then
'                    .RowVisible(L) = lngImage
'                End If
'            End If
'        Next L
'        If lngImage Then
'            .RowImage() = 2
'        Else
'            .RowImage() = 1
'        End If
'    End If
'    .Redraw = True
'    .ColumnsForceFit
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr9_Change()
On Error GoTo tratar_erro

If txtPagIr9 <> "" Then
    VerifNumero = txtPagIr9
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr9 = ""
        txtPagIr9.SetFocus
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
    Case 1: ProcFiltrar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaOS()
On Error GoTo tratar_erro
Dim L As Long, G As Long

Randomize
With FlexGridOS
    .Clear
    .ImageList = ImageList1
    Contador = 1
    'Cria o grupo
    Set TBLista_aviso_diario9 = CreateObject("adodb.recordset")
    TBLista_aviso_diario9.Open "Select MA.Grupo from ordemservico OS INNER JOIN Cadmaquinas MA on OS.Maquina = MA.Maquina where OS.Prazofinal < '" & Format(Date, "Short Date") & "' and OS.pronto = 'NÃO' group by MA.Grupo", Conexao, adOpenKeyset, adLockReadOnly
    If TBLista_aviso_diario9.EOF = False Then
        SSTab1.TabVisible(9) = True
        
        PBLista9.Min = 0
        PBLista9.Max = TBLista_aviso_diario9.RecordCount
        PBLista9.Value = 1
        Contador2 = 0
        Do While TBLista_aviso_diario9.EOF = False
    
            L = .AddItem(vbTab & TBLista_aviso_diario9!Grupo)
            .RowData(L) = Contador
            .RowGroupHeader(L) = True
            .RowImage(L) = 1
            .RowHeight(L) = 25
            .CellFontBold(L, 1) = True
        
            Set TBLista_aviso_diario10 = CreateObject("adodb.recordset")
            TBLista_aviso_diario10.Open "Select OS.Maquina, OS.IDproducao, OS.Prazofinal, OS.ordem, OS.status, OS.Tempototallote from ordemservico OS INNER JOIN Cadmaquinas MA on OS.Maquina = MA.Maquina where MA.grupo = '" & TBLista_aviso_diario9!Grupo & "' and  OS.Prazofinal < '" & Format(Date, "Short Date") & "' and OS.pronto = 'NÃO' order by OS.Idproducao", Conexao, adOpenKeyset, adLockOptimistic
            If TBLista_aviso_diario10.EOF = False Then
                Do While TBLista_aviso_diario10.EOF = False
                    G = .AddItem(, , , , , False)
                    .RowData(G) = Contador
                    .CellText(G, 1) = TBLista_aviso_diario10!IDProducao
                    .CellText(G, 2) = IIf(IsNull(TBLista_aviso_diario10!PrazoFinal), "", Format(TBLista_aviso_diario10!PrazoFinal, "dd/mm/yy"))
                    .CellText(G, 3) = IIf(IsNull(TBLista_aviso_diario10!TempoTotalLote), "", TBLista_aviso_diario10!TempoTotalLote)
                    .CellText(G, 4) = IIf(IsNull(TBLista_aviso_diario10!maquina), "", TBLista_aviso_diario10!maquina)
                    Set TBLista_aviso_diario11 = CreateObject("adodb.recordset")
                    TBLista_aviso_diario11.Open "Select Descricao from Cadmaquinas where maquina = '" & TBLista_aviso_diario10!maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBLista_aviso_diario11.EOF = False Then .CellText(G, 5) = IIf(IsNull(TBLista_aviso_diario11!Descricao), "", TBLista_aviso_diario11!Descricao)
                    
                    .CellText(G, 6) = IIf(IsNull(TBLista_aviso_diario10!Ordem), "", TBLista_aviso_diario10!Ordem)
                    
                    Set TBLista_aviso_diario12 = CreateObject("adodb.recordset")
                    TBLista_aviso_diario12.Open "Select Prazoentrega from Producao where ordem = " & TBLista_aviso_diario10!Ordem, Conexao, adOpenKeyset, adLockOptimistic
                    If TBLista_aviso_diario12.EOF = False Then .CellText(G, 7) = IIf(IsNull(TBLista_aviso_diario12!PrazoEntrega), "", Format(TBLista_aviso_diario12!PrazoEntrega, "dd/mm/yy"))
                    .CellText(G, 8) = IIf(IsNull(TBLista_aviso_diario10!status), "", TBLista_aviso_diario10!status)
                
                    TBLista_aviso_diario10.MoveNext
                Loop
            End If
            TBLista_aviso_diario10.Close
            Contador = Contador + 1
            TBLista_aviso_diario9.MoveNext
            Contador2 = Contador2 + 1
            PBLista9.Value = Contador2
        Loop
    Else
        SSTab1.TabVisible(9) = False
        ContadorTAB = ContadorTAB - 1
    End If
    TBLista_aviso_diario9.Close
    .Redraw = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaCC()
On Error GoTo tratar_erro
Dim L As Long, G As Long

Randomize
With FlexGridCC
    .Clear
    .ImageList = ImageList2
    Contador = 1
    'Cria o grupo
    Set TBLista_aviso_diario13 = CreateObject("adodb.recordset")
    Set TBLista_aviso_diarioAcesso = CreateObject("adodb.recordset")
    TBLista_aviso_diarioAcesso.Open "Select IDAcesso from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = 'Custos/Centro de custo/Visualizar todos'", Conexao, adOpenKeyset, adLockOptimistic
    If TBLista_aviso_diarioAcesso.EOF = False Then
        TBLista_aviso_diario13.Open "Select US.Setor from Usuarios_setor US INNER JOIN CC_realizado CC on US.ID = CC.ID_CC where CC.Data = '" & Format(Date, "Short Date") & "' group by US.Setor", Conexao, adOpenKeyset, adLockOptimistic
    Else
        TBLista_aviso_diario13.Open "Select US.Setor from (Usuarios_setor US INNER JOIN Usuarios_Setor_Responsavel USR ON US.ID = USR.ID_CC) INNER JOIN CC_realizado CC on US.ID = CC.ID_CC where USR.Responsavel_CC = '" & pubUsuario & "' and CC.Data = '" & Format(Date, "Short Date") & "' group by US.Setor", Conexao, adOpenKeyset, adLockOptimistic
    End If
    TBLista_aviso_diarioAcesso.Close
    If TBLista_aviso_diario13.EOF = False Then
        SSTab1.TabVisible(10) = True
        
        PBLista10.Min = 0
        PBLista10.Max = TBLista_aviso_diario13.RecordCount
        PBLista10.Value = 1
        Contador2 = 0
        Do While TBLista_aviso_diario13.EOF = False
    
            L = .AddItem(vbTab & TBLista_aviso_diario13!Setor)
            .RowData(L) = Contador
            .RowGroupHeader(L) = True
            .RowImage(L) = 1
            .RowHeight(L) = 25
            .CellFontBold(L, 1) = True
        
            Set TBLista_aviso_diario14 = CreateObject("adodb.recordset")
            Set TBLista_aviso_diarioAcesso = CreateObject("adodb.recordset")
            TBLista_aviso_diarioAcesso.Open "Select IDAcesso from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = 'Custos/Centro de custo/Visualizar todos'", Conexao, adOpenKeyset, adLockOptimistic
            If TBLista_aviso_diarioAcesso.EOF = False Then
                TBLista_aviso_diario14.Open "Select CC.ID_estoque, CC.ID_financeiro, CC.Operacao, CC.Valor, CC.Data, CC.Responsavel from Usuarios_setor US INNER JOIN CC_realizado CC on US.ID = CC.ID_CC where US.Setor = '" & TBLista_aviso_diario13!Setor & "' and CC.Data = '" & Format(Date, "Short Date") & "' order by US.ID", Conexao, adOpenKeyset, adLockOptimistic
            Else
                TBLista_aviso_diario14.Open "Select CC.ID_estoque, CC.ID_financeiro, CC.Operacao, CC.Valor, CC.Data, CC.Responsavel from (Usuarios_setor US INNER JOIN Usuarios_Setor_Responsavel USR ON US.ID = USR.ID_CC) INNER JOIN CC_realizado CC on US.ID = CC.ID_CC where US.Setor = '" & TBLista_aviso_diario13!Setor & "' and USR.Responsavel_CC = '" & pubUsuario & "' and CC.Data = '" & Format(Date, "Short Date") & "' order by US.ID", Conexao, adOpenKeyset, adLockOptimistic
            End If
            TBLista_aviso_diarioAcesso.Close
            If TBLista_aviso_diario14.EOF = False Then
                Do While TBLista_aviso_diario14.EOF = False
                    G = .AddItem(, , , , , False)
                    .RowData(G) = Contador
                    
                    If IsNull(TBLista_aviso_diario14!ID_estoque) = False And TBLista_aviso_diario14!ID_estoque <> 0 Then
                        Modulo_texto = "Estoque"
                        
                        Set TBLista_aviso_diarioCC = CreateObject("adodb.recordset")
                        TBLista_aviso_diarioCC.Open "Select Entrada, Documento, Lote from Estoque_movimentacao where Idoperacao = " & TBLista_aviso_diario14!ID_estoque, Conexao, adOpenKeyset, adLockOptimistic
                        If TBLista_aviso_diarioCC.EOF = False Then
                            If TBLista_aviso_diarioCC!Entrada > 0 Then DocumentoRef = "Ped. " & TBLista_aviso_diarioCC!LOTE Else DocumentoRef = TBLista_aviso_diarioCC!Documento
                        End If
                        TBLista_aviso_diarioCC.Close
                    Else
                        Modulo_texto = "Financeiro"
                        
                        'Verifica número do documento e se a conta já foi paga
                        Set TBLista_aviso_diarioCC = CreateObject("adodb.recordset")
                        TBLista_aviso_diarioCC.Open "Select txt_NDocumento, Logsit from tbl_ContasPagar where Idintconta = " & IIf(IsNull(TBLista_aviso_diario14!ID_financeiro), 0, TBLista_aviso_diario14!ID_financeiro), Conexao, adOpenKeyset, adLockOptimistic
                        If TBLista_aviso_diarioCC.EOF = False Then
                            If TBLista_aviso_diarioCC!Logsit = "N" Then DocumentoRef = "Doc. " & TBLista_aviso_diarioCC!txt_ndocumento & " | N" Else DocumentoRef = "Doc. " & TBLista_aviso_diarioCC!txt_ndocumento & " | S"
                        End If
                        TBLista_aviso_diarioCC.Close
                    End If
                    .CellText(G, 1) = IIf(IsNull(TBLista_aviso_diario14!Data), "", TBLista_aviso_diario14!Data)
                    .CellText(G, 2) = IIf(IsNull(TBLista_aviso_diario14!Responsavel), "", TBLista_aviso_diario14!Responsavel)
                    .CellText(G, 3) = Modulo_texto
                    .CellText(G, 4) = DocumentoRef
            
                    Select Case TBLista_aviso_diario14!Operacao
                        Case "Crédito": ValorTexto = "-" & IIf(IsNull(TBLista_aviso_diario14!valor), "", Format(TBLista_aviso_diario14!valor, "###,##0.00"))
                        Case "Débito": ValorTexto = IIf(IsNull(TBLista_aviso_diario14!valor), "", Format(TBLista_aviso_diario14!valor, "###,##0.00"))
                    End Select
                    .CellText(G, 5) = ValorTexto

                    TBLista_aviso_diario14.MoveNext
                Loop
            End If
            TBLista_aviso_diario14.Close
            Contador = Contador + 1
            TBLista_aviso_diario13.MoveNext
            Contador2 = Contador2 + 1
            PBLista10.Value = Contador2
        Loop
    Else
        SSTab1.TabVisible(10) = False
        ContadorTAB = ContadorTAB - 1
    End If
    TBLista_aviso_diario13.Close
    .Redraw = True
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifMostrarEsconderTab(Tabs As SSTab, NTab As Integer, Formulario As String)
On Error GoTo tratar_erro

Set TBAcessos = CreateObject("adodb.recordset")
If Formulario = "" Then
    TBAcessos.Open "Select Acesso from Acessos where IDUsuario = " & pubIDUsuario & " and (Acesso = 'Avisos diário/Análise crítica/Engenharia' or Acesso = 'Avisos diário/Análise crítica/Processos' or Acesso = 'Avisos diário/Análise crítica/Pcp' or Acesso = 'Avisos diário/Análise crítica/Qualidade' or Acesso = 'Avisos diário/Análise crítica/Compras')", Conexao, adOpenKeyset, adLockOptimistic
Else
    TBAcessos.Open "Select Acesso from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
End If
If TBAcessos.EOF = True Then
    Tabs.TabVisible(NTab) = False
    Contador = Contador - 1
Else
    Tabs.TabVisible(NTab) = True
End If
TBAcessos.Close
Tabs.TabsPerRow = IIf(Contador < 1, 1, Contador)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
