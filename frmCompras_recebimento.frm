VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompras_recebimento 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Qualidade - Inspeção de recebimento"
   ClientHeight    =   10035
   ClientLeft      =   300
   ClientTop       =   450
   ClientWidth     =   15360
   ControlBox      =   0   'False
   Icon            =   "frmCompras_recebimento.frx":0000
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   0
      TabIndex        =   115
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   17701
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
      TabCaption(0)   =   "Lista dos lotes"
      TabPicture(0)   =   "frmCompras_recebimento.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(1)=   "Frame15"
      Tab(0).Control(2)=   "Lista"
      Tab(0).Control(3)=   "PBLista2"
      Tab(0).Control(4)=   "USToolBar1"
      Tab(0).Control(5)=   "Frame10"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Lista de produtos"
      TabPicture(1)   =   "frmCompras_recebimento.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "USToolBar2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame7"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "SSTab2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   -74925
         TabIndex        =   127
         Top             =   1320
         Width           =   15190
         Begin VB.OptionButton Opt_inspecionados 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Inspecionados"
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
            Left            =   13470
            TabIndex        =   131
            Top             =   210
            Width           =   1515
         End
         Begin VB.OptionButton Opt_inspecionar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Inspecionar"
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
            Left            =   11910
            TabIndex        =   130
            Top             =   210
            Value           =   -1  'True
            Width           =   1275
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
            ItemData        =   "frmCompras_recebimento.frx":0044
            Left            =   1170
            List            =   "frmCompras_recebimento.frx":0046
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   128
            ToolTipText     =   "Empresa."
            Top             =   150
            Width           =   10575
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa :"
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
            TabIndex        =   129
            Top             =   150
            Width           =   825
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74925
         TabIndex        =   110
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
            TabIndex        =   5
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
            TabIndex        =   4
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   9
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_recebimento.frx":0048
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
            TabIndex        =   8
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_recebimento.frx":37EC
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
            TabIndex        =   6
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
            TabIndex        =   7
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_recebimento.frx":72F5
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
            TabIndex        =   10
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmCompras_recebimento.frx":B3E4
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
         Begin VB.Label Label18 
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
            Index           =   1
            Left            =   4410
            TabIndex        =   135
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
            TabIndex        =   113
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
            TabIndex        =   112
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label36 
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
            Left            =   3090
            TabIndex        =   111
            Top             =   240
            Width           =   645
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2355
         Left            =   75
         TabIndex        =   120
         Top             =   7620
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   4154
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Dados do produto"
         TabPicture(0)   =   "frmCompras_recebimento.frx":EC70
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Amostragem/check list de verificação"
         TabPicture(1)   =   "frmCompras_recebimento.frx":EC8C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame6"
         Tab(1).Control(1)=   "Frame5"
         Tab(1).ControlCount=   2
         Begin VB.Frame Frame2 
            BackColor       =   &H00E0E0E0&
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
            Height          =   2025
            Left            =   30
            TabIndex        =   94
            Top             =   330
            Width           =   11565
            Begin VB.TextBox Txt_certificado 
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
               MaxLength       =   255
               TabIndex        =   35
               TabStop         =   0   'False
               ToolTipText     =   "Certificado."
               Top             =   1560
               Width           =   1695
            End
            Begin VB.TextBox Txt_corrida 
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
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   30
               TabStop         =   0   'False
               ToolTipText     =   "Corrida."
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox Txt_nota_fiscal 
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
               MaxLength       =   255
               TabIndex        =   28
               TabStop         =   0   'False
               ToolTipText     =   "Número da nota fiscal."
               Top             =   960
               Width           =   1125
            End
            Begin VB.TextBox Txt_data_emissao_NF 
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
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   29
               TabStop         =   0   'False
               ToolTipText     =   "Data de emissão."
               Top             =   960
               Width           =   1125
            End
            Begin VB.CommandButton Cmd_visualizar_arquivo1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   11160
               Picture         =   "frmCompras_recebimento.frx":ECA8
               Style           =   1  'Graphical
               TabIndex        =   39
               ToolTipText     =   "Visualizar arquivo."
               Top             =   1560
               Width           =   315
            End
            Begin VB.CommandButton Cmd_limpar_caminho1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   10830
               Picture         =   "frmCompras_recebimento.frx":F26A
               Style           =   1  'Graphical
               TabIndex        =   38
               ToolTipText     =   "Limpar caminho."
               Top             =   1560
               Width           =   315
            End
            Begin VB.CommandButton Cmd_visualizar_arquivo 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   11160
               Picture         =   "frmCompras_recebimento.frx":F3A8
               Style           =   1  'Graphical
               TabIndex        =   34
               ToolTipText     =   "Visualizar arquivo."
               Top             =   960
               Width           =   315
            End
            Begin VB.CommandButton Cmd_limpar_caminho 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   10830
               Picture         =   "frmCompras_recebimento.frx":F96A
               Style           =   1  'Graphical
               TabIndex        =   33
               ToolTipText     =   "Limpar caminho."
               Top             =   960
               Width           =   315
            End
            Begin VB.TextBox Txt_qtde_recebida 
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
               Left            =   7530
               Locked          =   -1  'True
               TabIndex        =   25
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade recebida."
               Top             =   375
               Width           =   1305
            End
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   9810
               Top             =   750
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.CommandButton cmdImportar2 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   10500
               Picture         =   "frmCompras_recebimento.frx":FAA8
               Style           =   1  'Graphical
               TabIndex        =   37
               ToolTipText     =   "Localizar certificado."
               Top             =   1560
               Width           =   315
            End
            Begin VB.TextBox Txt_caminho2 
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
               Left            =   1890
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   36
               TabStop         =   0   'False
               ToolTipText     =   "Caminho do certificado."
               Top             =   1560
               Width           =   8625
            End
            Begin VB.CommandButton cmdImportar 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   10500
               Picture         =   "frmCompras_recebimento.frx":FBAA
               Style           =   1  'Graphical
               TabIndex        =   32
               ToolTipText     =   "Localizar corrida."
               Top             =   960
               Width           =   315
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
               Left            =   4170
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   31
               TabStop         =   0   'False
               ToolTipText     =   "Caminho da corrida."
               Top             =   960
               Width           =   6315
            End
            Begin VB.Frame Frame9 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
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
               Height          =   405
               Left            =   5700
               TabIndex        =   95
               Top             =   975
               Width           =   1095
            End
            Begin VB.TextBox txtNomenclatura 
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
               MaxLength       =   50
               TabIndex        =   22
               TabStop         =   0   'False
               ToolTipText     =   "Código interno."
               Top             =   375
               Width           =   1530
            End
            Begin VB.TextBox txtQtdeAinspecionar 
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
               Left            =   10170
               Locked          =   -1  'True
               TabIndex        =   27
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade a inspecionar."
               Top             =   375
               Width           =   1305
            End
            Begin VB.TextBox Txt_unidade 
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
               Left            =   6915
               Locked          =   -1  'True
               TabIndex        =   24
               TabStop         =   0   'False
               ToolTipText     =   "Unidade."
               Top             =   375
               Width           =   600
            End
            Begin VB.TextBox txtEspecificacoes 
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
               Left            =   1720
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   23
               TabStop         =   0   'False
               ToolTipText     =   "Descrição."
               Top             =   375
               Width           =   5175
            End
            Begin VB.TextBox txtInspecionada 
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
               Left            =   8850
               Locked          =   -1  'True
               TabIndex        =   26
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade inspecionada."
               Top             =   375
               Width           =   1305
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
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
               Left            =   570
               TabIndex        =   126
               Top             =   1365
               Width           =   915
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
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
               Left            =   3000
               TabIndex        =   125
               Top             =   765
               Width           =   615
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Nota fiscal"
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
               Index           =   0
               Left            =   307
               TabIndex        =   124
               Top             =   765
               Width           =   870
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Dt. emissão"
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
               Left            =   1387
               TabIndex        =   123
               Top             =   765
               Width           =   990
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. receb."
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
               Left            =   7672
               TabIndex        =   121
               Top             =   180
               Width           =   1020
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Caminho do certificado"
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
               Left            =   5385
               TabIndex        =   102
               Top             =   1365
               Width           =   1635
            End
            Begin VB.Label Label30 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Caminho da corrida"
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
               Left            =   6637
               TabIndex        =   101
               Top             =   765
               Width           =   1380
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
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
               Left            =   3962
               TabIndex        =   100
               Top             =   180
               Width           =   690
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Cód. interno"
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
               Left            =   435
               TabIndex        =   99
               Top             =   180
               Width           =   1020
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. a insp."
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
               Left            =   10305
               TabIndex        =   98
               Top             =   180
               Width           =   1035
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. insp."
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
               Left            =   9060
               TabIndex        =   97
               Top             =   180
               Width           =   885
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   "Un."
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
               Left            =   7088
               TabIndex        =   96
               Top             =   180
               Width           =   255
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Amostragem (NBR5426 - Normal) / Critérios de aceitação (CA)"
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
            Height          =   1965
            Left            =   -74970
            TabIndex        =   86
            Top             =   330
            Width           =   6285
            Begin VB.TextBox Txt_ID_RNC 
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
               Left            =   4305
               Locked          =   -1  'True
               TabIndex        =   122
               TabStop         =   0   'False
               Text            =   "0"
               ToolTipText     =   "ID RNC."
               Top             =   1275
               Visible         =   0   'False
               Width           =   390
            End
            Begin VB.CommandButton cmdRNC 
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   5835
               Picture         =   "frmCompras_recebimento.frx":FCAC
               Style           =   1  'Graphical
               TabIndex        =   47
               ToolTipText     =   "Criar RNC."
               Top             =   1275
               Width           =   315
            End
            Begin VB.TextBox Txtsac 
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
               Left            =   4305
               TabIndex        =   46
               TabStop         =   0   'False
               ToolTipText     =   "N° da RNC."
               Top             =   1275
               Width           =   1530
            End
            Begin VB.TextBox Txtenc 
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
               TabIndex        =   42
               TabStop         =   0   'False
               ToolTipText     =   "Quantidade encontrada."
               Top             =   645
               Width           =   1530
            End
            Begin VB.TextBox Txtrejeitado 
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
               Left            =   2940
               TabIndex        =   45
               ToolTipText     =   "Quantidade rejeitada."
               Top             =   1275
               Width           =   1350
            End
            Begin VB.TextBox txtcondicional 
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
               Left            =   1575
               TabIndex        =   44
               ToolTipText     =   "Quantidade aceita."
               Top             =   1275
               Width           =   1350
            End
            Begin VB.ComboBox CmbNQA 
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
               ItemData        =   "frmCompras_recebimento.frx":FD8E
               Left            =   180
               List            =   "frmCompras_recebimento.frx":FDA7
               TabIndex        =   40
               ToolTipText     =   "NQA."
               Top             =   645
               Width           =   2250
            End
            Begin VB.ComboBox cmbNivel 
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
               ItemData        =   "frmCompras_recebimento.frx":FDE3
               Left            =   2430
               List            =   "frmCompras_recebimento.frx":FDFF
               TabIndex        =   41
               ToolTipText     =   "Nível."
               Top             =   645
               Width           =   2190
            End
            Begin VB.TextBox Txtamostra 
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
               TabIndex        =   43
               ToolTipText     =   "Quantidade amostra."
               Top             =   1275
               Width           =   1380
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "NQA*"
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
               Left            =   1095
               TabIndex        =   93
               Top             =   450
               Width           =   420
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nível* / CA"
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
               Left            =   3150
               TabIndex        =   92
               Top             =   450
               Width           =   795
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Amostra*"
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
               Left            =   525
               TabIndex        =   91
               Top             =   1080
               Width           =   690
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. aceita*"
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
               Left            =   1755
               TabIndex        =   90
               Top             =   1080
               Width           =   990
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. rejeitada*"
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
               Left            =   3015
               TabIndex        =   89
               Top             =   1080
               Width           =   1200
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. encontr.*"
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
               Left            =   4800
               TabIndex        =   88
               Top             =   450
               Width           =   1170
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nº RNC"
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
               Left            =   4800
               TabIndex        =   87
               Top             =   1080
               Width           =   540
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Check list de verificação"
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
            Height          =   1965
            Left            =   -68670
            TabIndex        =   79
            Top             =   330
            Width           =   5235
            Begin DrawSuite2022.USButton btnFoto 
               Height          =   225
               Left            =   1530
               TabIndex        =   137
               ToolTipText     =   "Tirar foto do item"
               Top             =   1110
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   397
               Caption         =   "Foto"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderColor     =   8421504
               BorderColorDisabled=   13160660
               BorderColorDown =   7907521
               BorderColorOver =   7907521
               GradientColor2  =   14737632
               GradientColor3  =   12632256
               GradientColor4  =   12632256
               GradientColorDisabled1=   14215660
               GradientColorDisabled2=   14215660
               GradientColorDisabled3=   14215660
               GradientColorDisabled4=   14215660
               GradientColorOver1=   14417407
               GradientColorOver2=   12317439
               GradientColorOver3=   4838399
               GradientColorOver4=   9627391
               GradientColorDown1=   10802943
               GradientColorDown2=   7979263
               GradientColorDown3=   4370174
               GradientColorDown4=   7395582
               GradientColors  =   1
               ShowFocusRect   =   0   'False
               Theme           =   1
            End
            Begin VB.CheckBox chk_outros_na 
               BackColor       =   &H00E0E0E0&
               Caption         =   "N/A"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4200
               TabIndex        =   65
               Top             =   1650
               Width           =   615
            End
            Begin VB.CheckBox chk_outros_nc 
               BackColor       =   &H00E0E0E0&
               Caption         =   "N/C"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   3360
               TabIndex        =   64
               Top             =   1650
               Width           =   615
            End
            Begin VB.CheckBox chk_outros_ok 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ok"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   2520
               TabIndex        =   63
               Top             =   1650
               Width           =   615
            End
            Begin VB.CheckBox chk_dim_na 
               BackColor       =   &H00E0E0E0&
               Caption         =   "N/A"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4200
               TabIndex        =   62
               Top             =   1368
               Width           =   615
            End
            Begin VB.CheckBox chk_dim_nc 
               BackColor       =   &H00E0E0E0&
               Caption         =   "N/C"
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
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   3360
               TabIndex        =   61
               Top             =   1368
               Width           =   615
            End
            Begin VB.CheckBox chk_dim_ok 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ok"
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
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   2520
               TabIndex        =   60
               Top             =   1368
               Width           =   615
            End
            Begin VB.CheckBox chk_visual_na 
               BackColor       =   &H00E0E0E0&
               Caption         =   "N/A"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4200
               TabIndex        =   59
               Top             =   1086
               Width           =   615
            End
            Begin VB.CheckBox chk_visual_nc 
               BackColor       =   &H00E0E0E0&
               Caption         =   "N/C"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   3360
               TabIndex        =   58
               Top             =   1086
               Width           =   615
            End
            Begin VB.CheckBox chk_visual_ok 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ok"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   2520
               TabIndex        =   57
               Top             =   1086
               Width           =   615
            End
            Begin VB.CheckBox chk_qtd_na 
               BackColor       =   &H00E0E0E0&
               Caption         =   "N/A"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4200
               TabIndex        =   56
               Top             =   804
               Width           =   615
            End
            Begin VB.CheckBox chk_qtd_nc 
               BackColor       =   &H00E0E0E0&
               Caption         =   "N/C"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   3360
               TabIndex        =   55
               Top             =   804
               Width           =   615
            End
            Begin VB.CheckBox chk_qtd_ok 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ok"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   2520
               TabIndex        =   54
               Top             =   804
               Width           =   615
            End
            Begin VB.CheckBox chk_laudo_na 
               BackColor       =   &H00E0E0E0&
               Caption         =   "N/A"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4200
               TabIndex        =   53
               Top             =   522
               Width           =   615
            End
            Begin VB.CheckBox chk_laudo_nc 
               BackColor       =   &H00E0E0E0&
               Caption         =   "N/C"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   3360
               MouseIcon       =   "frmCompras_recebimento.frx":FE24
               TabIndex        =   52
               Top             =   522
               Width           =   615
            End
            Begin VB.CheckBox chk_laudo_ok 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ok"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   2520
               MouseIcon       =   "frmCompras_recebimento.frx":FF76
               TabIndex        =   51
               Top             =   522
               Width           =   615
            End
            Begin VB.CheckBox chk_emb_na 
               BackColor       =   &H00E0E0E0&
               Caption         =   "N/A"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4200
               TabIndex        =   50
               Top             =   240
               Width           =   615
            End
            Begin VB.CheckBox chk_emb_nc 
               BackColor       =   &H00E0E0E0&
               Caption         =   "N/C"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   3360
               TabIndex        =   49
               Top             =   240
               Width           =   615
            End
            Begin VB.CheckBox chk_emb_ok 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ok"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   2520
               TabIndex        =   48
               Top             =   240
               Width           =   615
            End
            Begin DrawSuite2022.USButton btnPlano 
               Height          =   225
               Left            =   1530
               TabIndex        =   138
               ToolTipText     =   "Abrir plano de inspeção do item"
               Top             =   1380
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   397
               Caption         =   "Plano"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderColor     =   8421504
               BorderColorDisabled=   13160660
               BorderColorDown =   7907521
               BorderColorOver =   7907521
               GradientColor2  =   14737632
               GradientColor3  =   12632256
               GradientColor4  =   12632256
               GradientColorDisabled1=   14215660
               GradientColorDisabled2=   14215660
               GradientColorDisabled3=   14215660
               GradientColorDisabled4=   14215660
               GradientColorOver1=   14417407
               GradientColorOver2=   12317439
               GradientColorOver3=   4838399
               GradientColorOver4=   9627391
               GradientColorDown1=   10802943
               GradientColorDown2=   7979263
               GradientColorDown3=   4370174
               GradientColorDown4=   7395582
               GradientColors  =   1
               PicSize         =   5
               PicSizeH        =   8
               PicSizeW        =   8
               ShowFocusRect   =   0   'False
               Theme           =   1
            End
            Begin DrawSuite2022.USButton btnMedicoes 
               Height          =   225
               Left            =   1530
               TabIndex        =   139
               ToolTipText     =   "Informar medições encontradas"
               Top             =   1650
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   397
               Caption         =   "Medições"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderColor     =   8421504
               BorderColorDisabled=   13160660
               BorderColorDown =   7907521
               BorderColorOver =   7907521
               GradientColor2  =   14737632
               GradientColor3  =   12632256
               GradientColor4  =   12632256
               GradientColorDisabled1=   14215660
               GradientColorDisabled2=   14215660
               GradientColorDisabled3=   14215660
               GradientColorDisabled4=   14215660
               GradientColorOver1=   14417407
               GradientColorOver2=   12317439
               GradientColorOver3=   4838399
               GradientColorOver4=   9627391
               GradientColorDown1=   10802943
               GradientColorDown2=   7979263
               GradientColorDown3=   4370174
               GradientColorDown4=   7395582
               GradientColors  =   1
               PicSize         =   5
               PicSizeH        =   8
               PicSizeW        =   8
               ShowFocusRect   =   0   'False
               Theme           =   1
            End
            Begin VB.Image ImgAnexarplanomedicao 
               Height          =   270
               Left            =   1560
               Picture         =   "frmCompras_recebimento.frx":100C8
               Top             =   1335
               Width           =   120
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "1 - Embalagem*"
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
               Left            =   240
               TabIndex        =   85
               Top             =   270
               Width           =   1140
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "2 - Laudos / certificados*"
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
               Left            =   240
               TabIndex        =   84
               Top             =   555
               Width           =   1815
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "3 - Quantidade*"
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
               Left            =   240
               TabIndex        =   83
               Top             =   840
               Width           =   1170
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "4 - Visual*"
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
               Left            =   240
               TabIndex        =   82
               Top             =   1110
               Width           =   735
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "5 - Dimensional"
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
               Height          =   165
               Left            =   240
               TabIndex        =   81
               Top             =   1398
               Width           =   1080
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "6 - Outros*"
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
               Left            =   240
               TabIndex        =   80
               Top             =   1680
               Width           =   825
            End
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Laudo final de verificação"
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
         Height          =   2355
         Left            =   11745
         TabIndex        =   103
         Top             =   7620
         Width           =   3525
         Begin VB.TextBox txtobservacoes 
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
            Height          =   1215
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   68
            ToolTipText     =   "Observações."
            Top             =   990
            Width           =   3165
         End
         Begin VB.ComboBox cmblaudo 
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
            ItemData        =   "frmCompras_recebimento.frx":10424
            Left            =   150
            List            =   "frmCompras_recebimento.frx":10440
            Style           =   2  'Dropdown List
            TabIndex        =   66
            ToolTipText     =   "Laudo."
            Top             =   435
            Width           =   1905
         End
         Begin VB.ComboBox Cmbliberado 
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
            ItemData        =   "frmCompras_recebimento.frx":104AB
            Left            =   2070
            List            =   "frmCompras_recebimento.frx":104B8
            Style           =   2  'Dropdown List
            TabIndex        =   67
            ToolTipText     =   "Liberado."
            Top             =   435
            Width           =   1245
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
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
            Left            =   1260
            TabIndex        =   106
            Top             =   780
            Width           =   945
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Laudo*"
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
            Left            =   840
            TabIndex        =   105
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Liberado*"
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
            Left            =   2370
            TabIndex        =   104
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   825
         Left            =   75
         TabIndex        =   74
         Top             =   1305
         Width           =   15195
         Begin VB.TextBox txtDtValidacao 
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
            Left            =   4185
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação."
            Top             =   370
            Width           =   2025
         End
         Begin VB.TextBox txtRespValidacao 
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
            Left            =   6225
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   370
            Width           =   3105
         End
         Begin VB.TextBox mskData 
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
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Data da inspeção."
            Top             =   370
            Width           =   870
         End
         Begin VB.TextBox Txt_lote 
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
            Left            =   9345
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Número do lote."
            Top             =   370
            Width           =   1140
         End
         Begin VB.TextBox Txt_cliente_forn 
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
            Left            =   10495
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Cliente/fornecedor."
            Top             =   370
            Width           =   4500
         End
         Begin VB.TextBox txtResponsavel 
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
            Left            =   1065
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela inspeção."
            Top             =   370
            Width           =   3105
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Data/hora da validação"
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
            Left            =   4357
            TabIndex        =   133
            Top             =   180
            Width           =   1680
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Left            =   6787
            TabIndex        =   132
            Top             =   180
            Width           =   1980
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Data insp."
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
            Left            =   248
            TabIndex        =   78
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Lote"
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
            Left            =   9720
            TabIndex        =   77
            Top             =   180
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            BackStyle       =   0  'Transparent
            Caption         =   "Responsável pela inspeção"
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
            Left            =   1650
            TabIndex        =   76
            Top             =   180
            Width           =   1935
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente/fornecedor"
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
            Left            =   12063
            TabIndex        =   75
            Top             =   180
            Width           =   1365
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lista de produtos a inspecionar"
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
         Height          =   5475
         Left            =   75
         TabIndex        =   71
         Top             =   2130
         Width           =   7597
         Begin MSComctlLib.ListView Listprod 
            Height          =   4845
            Left            =   180
            TabIndex        =   21
            Top             =   240
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   8546
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
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "T"
               Text            =   "Cód. interno"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   6041
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Un."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Qtde."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "RE"
               Object.Width           =   1764
            EndProperty
         End
         Begin DrawSuite2022.USProgressBar PBLista 
            Height          =   255
            Left            =   180
            TabIndex        =   118
            Top             =   5100
            Width           =   7215
            _ExtentX        =   12726
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
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lista de itens inspecionados"
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
         Height          =   5475
         Left            =   7673
         TabIndex        =   72
         Top             =   2130
         Width           =   7597
         Begin VB.ComboBox cmb_Opcao_Lista 
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
            ItemData        =   "frmCompras_recebimento.frx":104D1
            Left            =   5460
            List            =   "frmCompras_recebimento.frx":104DB
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   5040
            Width           =   1965
         End
         Begin VB.TextBox txtID 
            Height          =   315
            Left            =   570
            TabIndex        =   73
            Top             =   870
            Visible         =   0   'False
            Width           =   525
         End
         Begin MSComctlLib.ListView ListProdReceb 
            Height          =   4785
            Left            =   180
            TabIndex        =   69
            Top             =   240
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   8440
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Cód. interno"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   4366
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Un."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Qtde."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "RE"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Valid."
               Object.Width           =   1147
            EndProperty
         End
         Begin DrawSuite2022.USProgressBar PBLista1 
            Height          =   255
            Left            =   180
            TabIndex        =   119
            Top             =   5100
            Width           =   3825
            _ExtentX        =   6747
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
            Left            =   4110
            TabIndex        =   134
            Top             =   5100
            Width           =   1260
         End
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   6330
         Left            =   -74925
         TabIndex        =   3
         Top             =   2745
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   11165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
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
            Text            =   "Lote"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Cliente/fornecedor"
            Object.Width           =   23292
         EndProperty
      End
      Begin DrawSuite2022.USProgressBar PBLista2 
         Height          =   255
         Left            =   -74925
         TabIndex        =   116
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
         BackColor       =   16773091
         BarColor1       =   16773091
         BarColor2       =   16769998
         BorderColor     =   14854529
         ForeColor2      =   0
         SearchText      =   ""
         Theme           =   0
         Value           =   0
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74925
         TabIndex        =   114
         Top             =   330
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
            Name            =   "MS Sans Serif"
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
         ButtonUseMaskColor3=   0   'False
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
         ButtonUseMaskColor4=   0   'False
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
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
         ButtonState5    =   5
         ButtonLeft5     =   110
         ButtonTop5      =   2
         ButtonWidth5    =   24
         ButtonHeight5   =   24
         ButtonUseMaskColor5=   0   'False
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   7710
            Top             =   90
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCompras_recebimento.frx":104F3
            Count           =   1
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   75
         TabIndex        =   117
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   8
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   83
         ButtonTop3      =   2
         ButtonWidth3    =   51
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Validação"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Validar/cancelar validação (F12)"
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
         ButtonLeft4     =   136
         ButtonTop4      =   2
         ButtonWidth4    =   53
         ButtonHeight4   =   21
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonAlignment5=   2
         ButtonType5     =   1
         ButtonStyle5    =   -1
         BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState5    =   -1
         ButtonLeft5     =   191
         ButtonTop5      =   4
         ButtonWidth5    =   2
         ButtonHeight5   =   54
         ButtonCaption6  =   "Ajuda"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Ajuda (F1)"
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
         ButtonLeft6     =   195
         ButtonTop6      =   2
         ButtonWidth6    =   36
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Sair"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Sair (Esc)"
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
         ButtonLeft7     =   233
         ButtonTop7      =   2
         ButtonWidth7    =   26
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
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
         ButtonState8    =   5
         ButtonLeft8     =   261
         ButtonTop8      =   2
         ButtonWidth8    =   24
         ButtonHeight8   =   24
         ButtonUseMaskColor8=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   7710
            Top             =   90
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCompras_recebimento.frx":126DB
            Count           =   1
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0E0E0&
         Height          =   885
         Left            =   -74925
         TabIndex        =   107
         Top             =   1845
         Width           =   15195
         Begin VB.Frame Frame11 
            BackColor       =   &H00E0E0E0&
            Height          =   510
            Left            =   4110
            TabIndex        =   136
            Top             =   210
            Width           =   4785
            Begin VB.OptionButton Optfim 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Fim frase"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2760
               TabIndex        =   13
               Top             =   180
               Width           =   1155
            End
            Begin VB.OptionButton Optinicio 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Início frase"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   180
               TabIndex        =   11
               Top             =   180
               Value           =   -1  'True
               Width           =   1275
            End
            Begin VB.OptionButton Optmeio 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Meio frase"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1470
               TabIndex        =   12
               Top             =   180
               Width           =   1275
            End
            Begin VB.OptionButton optIgual 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Igual"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3930
               TabIndex        =   14
               Top             =   180
               Width           =   705
            End
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
            ItemData        =   "frmCompras_recebimento.frx":169F5
            Left            =   180
            List            =   "frmCompras_recebimento.frx":16A0E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   3855
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
            Left            =   8970
            TabIndex        =   1
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   6045
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
            Left            =   8970
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Texto para pesquisa."
            Top             =   390
            Width           =   6045
         End
         Begin VB.Label Label35 
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
            Left            =   1687
            TabIndex        =   109
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label33 
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
            Left            =   11257
            TabIndex        =   108
            Top             =   180
            Width           =   1470
         End
      End
   End
End
Attribute VB_Name = "frmCompras_recebimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ListaInspecionar     As Boolean 'OK
Public ListaInspecionados   As Boolean 'OK
Public Consignacao          As Boolean 'OK
Dim StrSql_Localizar_Inspecao As String 'OK
Dim FormulaRel_Inspecao As String 'OK
Dim TBLISTA_Inspecao_Recebimento   As ADODB.Recordset 'OK

Private Sub ProcSair()
On Error GoTo tratar_erro

Inspecaorecebimento_AnexarPlano = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposPrincipal()
On Error GoTo tratar_erro

mskData.Text = Format(Date, "dd/mm/yy")
txtResponsavel.Text = pubUsuario
txtDtValidacao = ""
txtRespValidacao = ""
Txt_lote = ""
Txt_cliente_forn = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtLote = "0,0000"
Txt_qtde_recebida = "0,0000"
txtQtdeAinspecionar = "0,0000"
txtInspecionada = "0,0000"
Txt_nota_fiscal = ""
Txt_data_emissao_NF = ""
txt_Corrida = ""
txt_Certificado = ""
txt_Caminho = ""
Txt_caminho2 = ""
txtObservacoes.Text = ""
chk_dim_na.Value = 0
chk_dim_nc.Value = 0
chk_dim_ok.Value = 0
chk_emb_na.Value = 0
chk_emb_nc.Value = 0
chk_emb_ok.Value = 0
chk_laudo_na.Value = 0
chk_laudo_nc.Value = 0
chk_laudo_ok.Value = 0
chk_visual_na.Value = 0
chk_visual_nc.Value = 0
chk_visual_ok.Value = 0
chk_outros_na.Value = 0
chk_outros_nc.Value = 0
chk_outros_ok.Value = 0
chk_qtd_na.Value = 0
chk_qtd_nc.Value = 0
chk_qtd_ok.Value = 0
cmblaudo.ListIndex = -1
Cmbliberado.ListIndex = -1
CmbNQA.Text = "0,1"
cmbNivel.Text = "I"
Txtamostra.Text = ""
txtcondicional.Text = ""
Txtrejeitado.Text = ""
Txtenc.Text = ""
Txt_ID_RNC = 0
Txtsac.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

If ListaInspecionar = True Then
    TBCompras!IDEstoque = Listprod.SelectedItem.ListSubItems(4)
    TBCompras!Data = mskData
    TBCompras!Responsavel = txtResponsavel
End If

TBCompras!Laudo = IIf(cmblaudo = "", Null, cmblaudo)
TBCompras!Liberado = IIf(Cmbliberado = "", Null, Cmbliberado)
TBCompras!NQA = IIf(IsNumeric(CmbNQA.Text) = False, 0, CmbNQA.Text)
TBCompras!Nivel = IIf(cmbNivel.Text = "", Null, cmbNivel.Text)
TBCompras!AMOSTRA = IIf(IsNumeric(Txtamostra) = False, 0, Txtamostra)
TBCompras!AC = IIf(IsNumeric(txtcondicional) = False, 0, Format(txtcondicional, "###,##0.0000"))
TBCompras!RJ = IIf(IsNumeric(Txtrejeitado) = False, 0, Format(Txtrejeitado, "###,##0.0000"))
TBCompras!Enc = IIf(IsNumeric(Txtenc) = False, 0, Format(Txtenc, "###,##0.0000"))
TBCompras!ID_RNC = IIf(Txt_ID_RNC = 0, Null, Txt_ID_RNC)
                
ProcsaidaList
TBCompras!Obs = txtObservacoes.Text
TBCompras!Embalagem = Embalagem
TBCompras!Laudos = Laudos
TBCompras!quantidade = quantidade
TBCompras!dimensional = Dimensoes
TBCompras!Visual = Visual
TBCompras!Outros = Outros
TBCompras!Verificado = True
TBCompras!caminho = txt_Caminho
TBCompras!caminho2 = Txt_caminho2

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcsaidaList()
On Error GoTo tratar_erro

'==================
'Checa embalagem
'==================
If chk_emb_ok.Value = 1 And chk_emb_nc.Value = 0 And chk_emb_na = 0 Then Embalagem = 1
If chk_emb_ok.Value = 0 And chk_emb_nc.Value = 1 And chk_emb_na = 0 Then Embalagem = 2
If chk_emb_ok.Value = 0 And chk_emb_nc.Value = 0 And chk_emb_na = 1 Then Embalagem = 3
'==================
'Checa Laudos/certificados
'==================
If chk_laudo_ok.Value = 1 And chk_laudo_nc.Value = 0 And chk_laudo_na = 0 Then Laudos = 1
If chk_laudo_ok.Value = 0 And chk_laudo_nc.Value = 1 And chk_laudo_na = 0 Then Laudos = 2
If chk_laudo_ok.Value = 0 And chk_laudo_nc.Value = 0 And chk_laudo_na = 1 Then Laudos = 3
'==================
'Checa quantidades
'==================
If chk_qtd_ok.Value = 1 And chk_qtd_nc.Value = 0 And chk_qtd_na = 0 Then quantidade = 1
If chk_qtd_ok.Value = 0 And chk_qtd_nc.Value = 1 And chk_qtd_na = 0 Then quantidade = 2
If chk_qtd_ok.Value = 0 And chk_qtd_nc.Value = 0 And chk_qtd_na = 1 Then quantidade = 3
'==================
'Checa visual
'==================
If chk_visual_ok.Value = 1 And chk_visual_nc.Value = 0 And chk_visual_na = 0 Then Visual = 1
If chk_visual_ok.Value = 0 And chk_visual_nc.Value = 1 And chk_visual_na = 0 Then Visual = 2
If chk_visual_ok.Value = 0 And chk_visual_nc.Value = 0 And chk_visual_na = 1 Then Visual = 3
'==================
'Checa dimensoes
'==================
Dimensoes = 0
If chk_dim_ok.Value = 1 And chk_dim_nc.Value = 0 And chk_dim_na = 0 Then Dimensoes = 1
If chk_dim_ok.Value = 0 And chk_dim_nc.Value = 1 And chk_dim_na = 0 Then Dimensoes = 2
If chk_dim_ok.Value = 0 And chk_dim_nc.Value = 0 And chk_dim_na = 1 Then Dimensoes = 3
'==================
'Checa outros
'==================
If chk_outros_ok.Value = 1 And chk_outros_nc.Value = 0 And chk_outros_na = 0 Then Outros = 1
If chk_outros_ok.Value = 0 And chk_outros_nc.Value = 1 And chk_outros_na = 0 Then Outros = 2
If chk_outros_ok.Value = 0 And chk_outros_nc.Value = 0 And chk_outros_na = 1 Then Outros = 3

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnFoto_Click()
On Error GoTo tratar_erro

frmCompras_Recebimento_Foto.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnMedicoes_Click()
On Error GoTo tratar_erro

frmCompras_Recebimento_Medicoes.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnPlano_Click()
On Error GoTo tratar_erro

Acao = "anexar o controle de medição"
If txtNomenclatura = "" Then
    NomeCampo = "o produto"
    ProcVerificaAcao
    Exit Sub
End If
If cmbNivel = "" Then
    NomeCampo = "o nível"
    ProcVerificaAcao
    cmbNivel.SetFocus
    Exit Sub
End If
If Txtenc = "" Then
    NomeCampo = "a quantidade encontrada"
    ProcVerificaAcao
    Txtenc.SetFocus
    Exit Sub
End If
If Txtamostra = "" Then
    NomeCampo = "a quantidade de amostra"
    ProcVerificaAcao
    Txtamostra.SetFocus
    Exit Sub
End If
Set TBplano = CreateObject("adodb.recordset")
TBplano.Open "Select idplano from plano where desenho = '" & txtNomenclatura.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBplano.EOF = False Then
    Set TBplanomedicao = CreateObject("adodb.recordset")
    TBplanomedicao.Open "Select * from planodimensao where idplano = " & TBplano!IDPlano, Conexao, adOpenKeyset, adLockOptimistic
    If TBplanomedicao.EOF = True Then
        USMsgBox "Não foi encontrado nenhuma medição cadastrada para o plano de inspeção, favor cadastrar antes de abrir o controle de medição.", vbExclamation, "CAPRIND v5.0"
        TBplanomedicao.Close
        Exit Sub
    End If
    TBplanomedicao.Close
Else
    USMsgBox "Não foi encontrado nenhum plano de inspeção, favor cadastrar antes de abrir o controle de medição.", vbExclamation, "CAPRIND v5.0"
    TBplano.Close
    Exit Sub
End If
TBplano.Close
quantidade = 0
QTLOTE = 0
If ListaInspecionar = True Then
    USMsgBox ("Salve a inspeção de recebimento antes de anexar o controle de medição."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
Else
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Sum(quant_liberada) as quantidade from medicao where idlista = " & ListProdReceb.SelectedItem.ListSubItems(5) & " and desenho = '" & txtNomenclatura & "' and peca = '" & Txt_lote & "'", Conexao, adOpenKeyset, adLockOptimistic
End If
If TBAbrir.EOF = False Then
    quantidade = IIf(IsNull(TBAbrir!quantidade), 0, TBAbrir!quantidade)
End If
TBAbrir.Close
QTLOTE = txtLote - quantidade
Inspecaorecebimento_AnexarPlano = True
frmPlanomedicao.Show
frmPlanomedicao_ListaRNC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_dim_na_Click()
On Error GoTo tratar_erro

If chk_dim_na.Value = 1 Then
    chk_dim_ok.Value = 0
    chk_dim_nc.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_dim_nc_Click()
On Error GoTo tratar_erro

If chk_dim_nc.Value = 1 Then
    chk_dim_na.Value = 0
    chk_dim_ok.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_dim_ok_Click()
On Error GoTo tratar_erro

If chk_dim_ok.Value = 1 Then
    chk_dim_na.Value = 0
    chk_dim_nc.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_emb_na_Click()
On Error GoTo tratar_erro

If chk_emb_na.Value = 1 Then
    chk_emb_ok.Value = 0
    chk_emb_nc.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_emb_nc_Click()
On Error GoTo tratar_erro

If chk_emb_nc.Value = 1 Then
    chk_emb_na.Value = 0
    chk_emb_ok.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_emb_ok_Click()
On Error GoTo tratar_erro

If chk_emb_ok.Value = 1 Then
    chk_emb_na.Value = 0
    chk_emb_nc.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_laudo_na_Click()
On Error GoTo tratar_erro

If chk_laudo_na.Value = 1 Then
    chk_laudo_ok.Value = 0
    chk_laudo_nc.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_laudo_nc_Click()
On Error GoTo tratar_erro

If chk_laudo_nc.Value = 1 Then
    chk_laudo_na.Value = 0
    chk_laudo_ok.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_laudo_ok_Click()
On Error GoTo tratar_erro

If chk_laudo_ok.Value = 1 Then
    chk_laudo_na.Value = 0
    chk_laudo_nc.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_outros_na_Click()
On Error GoTo tratar_erro

If chk_outros_na.Value = 1 Then
    chk_outros_ok.Value = 0
    chk_outros_nc.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_outros_nc_Click()
On Error GoTo tratar_erro

If chk_outros_nc.Value = 1 Then
    chk_outros_na.Value = 0
    chk_outros_ok.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_outros_ok_Click()
On Error GoTo tratar_erro

If chk_outros_ok.Value = 1 Then
    chk_outros_na.Value = 0
    chk_outros_nc.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_qtd_na_Click()
On Error GoTo tratar_erro

If chk_qtd_na.Value = 1 Then
    chk_qtd_ok.Value = 0
    chk_qtd_nc.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_qtd_nc_Click()
On Error GoTo tratar_erro

If chk_qtd_nc.Value = 1 Then
    chk_qtd_na.Value = 0
    chk_qtd_ok.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_qtd_ok_Click()
On Error GoTo tratar_erro

If chk_qtd_ok.Value = 1 Then
    chk_qtd_na.Value = 0
    chk_qtd_nc.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_visual_na_Click()
On Error GoTo tratar_erro

If chk_visual_na.Value = 1 Then
    chk_visual_ok.Value = 0
    chk_visual_nc.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_visual_nc_Click()
On Error GoTo tratar_erro

If chk_visual_nc.Value = 1 Then
    chk_visual_na.Value = 0
    chk_visual_ok.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chk_visual_ok_Click()
On Error GoTo tratar_erro

If chk_visual_ok.Value = 1 Then
    chk_visual_na.Value = 0
    chk_visual_nc.Value = 0
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbCorrida_LostFocus()
On Error GoTo tratar_erro

If cmbCorrida = "" Then Exit Sub

If Consignacao = False Then
    If Programacao = True Then TextoFiltro = "Programacao = 'True'" Else TextoFiltro = "Programacao = 'False'"
    cmbCertificado.Clear
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select certificado from Estoque_controle_recebimento where IdLista = " & TXTIDLista.Text & " and IDPedido = " & txtIDPedido & " and Desenho = '" & txtNomenclatura & "' and " & TextoFiltro & " and nota_fiscal = '" & cmbNotaFiscal & "' and corrida = '" & cmbCorrida & "' group by certificado", Conexao, adOpenKeyset, adLockOptimistic
    Do While TBAbrir.EOF = False
        cmbCertificado.AddItem IIf(IsNull(TBAbrir!Certificado), "", TBAbrir!Certificado)
        TBAbrir.MoveNext
    Loop
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Change()
On Error GoTo tratar_erro

ProcLimpaTudoFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With ListProdReceb
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar2
    If Cmb_opcao_lista = "Validação" Then
        .ButtonState(2) = 5
        .ButtonState(4) = 0
    Else
        .ButtonState(2) = 0
        .ButtonState(4) = 5
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

ProcLimpaTudoFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ProcLimpaTudoFiltro
If cmbfiltrarpor = "Família" Then
    cmbfamilia.Visible = True
    txtTexto.Visible = False
Else
    cmbfamilia.Visible = False
    txtTexto.Visible = True
    
    If cmbfiltrarpor = "RE" And txtTexto <> "" Then
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
    Exit Sub
End Sub

Private Sub cmbNivel_Click()
On Error GoTo tratar_erro

Txtamostra = FunCalculaAmostragem(cmbNivel, IIf(Txtenc = "", 0, Txtenc))
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub CmbNQA_Click()
On Error GoTo tratar_erro

If CmbNQA.Text <> "" Then
      If CmbNQA.Text = "Liberação contra laudo" Then
        cmbNivel.Text = "EMP"
        Txtenc.Text = "N/A"
        Txtamostra.Text = "N/A"
        txtcondicional.Text = "N/A"
        Txtrejeitado.Text = "N/A"
        Txtsac.Text = "N/A"
        Txtenc.Locked = True
        Txtamostra.Locked = True
        txtcondicional.Locked = True
        Txtrejeitado.Locked = True
        Txtsac.Locked = True
      Else
        Txtenc.Text = txtQtdeAinspecionar.Text
        cmbNivel.Text = "I"
        Txtamostra.Text = "0"
        txtcondicional.Text = "0"
        Txtrejeitado.Text = "0"
        Txtsac.Text = "0"
        Txtenc.Locked = 0
        Txtamostra.Locked = 0
        txtcondicional.Locked = 0
        Txtrejeitado.Locked = 0
        Txtsac.Locked = 0
      End If
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

Private Sub Cmd_limpar_caminho1_Click()
On Error GoTo tratar_erro

Txt_caminho2 = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo1_Click()
On Error GoTo tratar_erro

If Txt_caminho2 <> "" Then ProcAbrirArquivo Txt_caminho2

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

Private Sub cmdImportar_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
txt_Caminho = caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImportar2_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
Txt_caminho2 = caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

frmCompras_recebimento_Menuimpressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

If Cmb_empresa = "" Then
    NomeCampo = "a empresa"
    Acao = "filtrar"
    ProcVerificaAcao
    Cmb_empresa.SetFocus
    Exit Sub
End If

CamposFiltro = "Lote, Cliente_forn"
INNERJOINTEXTO = "Select " & CamposFiltro & " from Qualidade_inspecao_recebimento"
OrdenarTexto = " group by " & CamposFiltro & " order by Lote"
If Opt_inspecionar.Value = True Then TextoFiltroInsp = "Laudo IS NULL" Else TextoFiltroInsp = "Laudo IS NOT NULL"

Nao_inspecionar = ""
If Opt_inspecionar = True Then
    Set TBTempo = CreateObject("adodb.recordset")
    TBTempo.Open "Select Codigo from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Nao_inspecionar = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBTempo.EOF = False Then Nao_inspecionar = " and Fornecedor IS NULL"
    TBTempo.Close
End If

If txtTexto <> "" Or cmbfamilia <> "" Then
    If cmbfiltrarpor = "Família" Then
        StrSql_Localizar_Inspecao = INNERJOINTEXTO & " where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and Classe = '" & cmbfamilia & "' and " & TextoFiltroInsp & Nao_inspecionar & OrdenarTexto
    ElseIf cmbfiltrarpor = "RE" Then
            StrSql_Localizar_Inspecao = INNERJOINTEXTO & " where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and IDestoque = " & txtTexto & " and " & TextoFiltroInsp & Nao_inspecionar & OrdenarTexto
        Else
            Select Case cmbfiltrarpor
                Case "Lote": TextoFiltro = "Lote"
                Case "Cliente/Fornecedor": TextoFiltro = "Cliente_forn"
                Case "Código interno": TextoFiltro = "Desenho"
                Case "Código de referência": TextoFiltro = "Ref"
                Case "Descrição": TextoFiltro = "descricao"
            End Select
            StrSql_Localizar_Inspecao = INNERJOINTEXTO & " where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroInsp & Nao_inspecionar & OrdenarTexto
    End If
Else
    StrSql_Localizar_Inspecao = INNERJOINTEXTO & " where ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TextoFiltroInsp & Nao_inspecionar & OrdenarTexto
End If
'Debug.print StrSql_Localizar_Inspecao

ProcCarregaListaRE

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaRE()
On Error GoTo tratar_erro

If StrSql_Localizar_Inspecao = "" Then Exit Sub
lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
Set TBLISTA_Inspecao_Recebimento = CreateObject("adodb.recordset")
TBLISTA_Inspecao_Recebimento.Open StrSql_Localizar_Inspecao, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Inspecao_Recebimento.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Inspecao_Recebimento.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Inspecao_Recebimento.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Inspecao_Recebimento.PageSize
ContadorReg = 1

PBLista2.Min = 0
PBLista2.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Inspecao_Recebimento.RecordCount - IIf(Pagina > 1, (TBLISTA_Inspecao_Recebimento.PageSize * (Pagina - 1)), 0), TBLISTA_Inspecao_Recebimento.PageSize)
PBLista2.Value = 1
Contador = 0
Do While TBLISTA_Inspecao_Recebimento.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , IIf(IsNull(TBLISTA_Inspecao_Recebimento!LOTE), "", TBLISTA_Inspecao_Recebimento!LOTE)
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Inspecao_Recebimento!Cliente_forn), "", TBLISTA_Inspecao_Recebimento!Cliente_forn)
        '.Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Inspecao_Recebimento!IDestoque), "", TBLISTA_Inspecao_Recebimento!IDestoque)
    End With
    TBLISTA_Inspecao_Recebimento.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista2.Value = Contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBLISTA_Inspecao_Recebimento.RecordCount
If TBLISTA_Inspecao_Recebimento.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBLISTA_Inspecao_Recebimento.PageCount
ElseIf TBLISTA_Inspecao_Recebimento.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBLISTA_Inspecao_Recebimento.PageCount & " de: " & TBLISTA_Inspecao_Recebimento.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBLISTA_Inspecao_Recebimento.AbsolutePage - 1 & " de: " & TBLISTA_Inspecao_Recebimento.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Inspecao_Recebimento.AbsolutePage <> 2 Then
    If TBLISTA_Inspecao_Recebimento.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Inspecao_Recebimento.PageCount - 1)
    Else
        TBLISTA_Inspecao_Recebimento.AbsolutePage = TBLISTA_Inspecao_Recebimento.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Inspecao_Recebimento.AbsolutePage)
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
    TBLISTA_Inspecao_Recebimento.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Inspecao_Recebimento.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Inspecao_Recebimento.AbsolutePage = 1
ProcExibePagina (TBLISTA_Inspecao_Recebimento.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Inspecao_Recebimento.AbsolutePage <> -3 Then
    If TBLISTA_Inspecao_Recebimento.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Inspecao_Recebimento.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Inspecao_Recebimento.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Inspecao_Recebimento.AbsolutePage = TBLISTA_Inspecao_Recebimento.PageCount
ProcExibePagina (TBLISTA_Inspecao_Recebimento.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdRNC_Click()
On Error GoTo tratar_erro

If ListaInspecionar = True Then
    USMsgBox ("Salve a inspeção de recebimento antes de criar a RNC."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Qtde = IIf(IsNumeric(Txtrejeitado) = False, 0, Txtrejeitado)
If Qtde <= 0 Then
    USMsgBox ("A quantidade rejeitada dever ser maior que zero."), vbExclamation, "CAPRIND v5.0"
    Txtrejeitado.SetFocus
    Exit Sub
End If
RNC_Inspecao_Recebimento = True
RNC_Controle_Medicao = False
RNC_Nao_Conformidade = False
RNC_Solicitacao_Desvio = False
frmQualidade_RNC.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

If SSTab1.Tab = 0 Then
    Select Case KeyCode
        Case vbKeyF2: ProcFiltrar
        Case vbKeyEscape: ProcSair
    End Select
Else
    Select Case KeyCode
        Case vbKeyF3: ProcSalvar
        Case vbKeyF4: If Cmb_opcao_lista = "Excluir" Then ProcExcluir
        Case vbKeyF5: ProcImprimir
        Case vbKeyF12: If Cmb_opcao_lista = "Validação" Then ProcValidarRegistros ListProdReceb, "Qualidade/Inspeção de recebimento"
        Case vbKeyEscape: ProcSair
    End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 5, True
ProcCarregaToolBar2 Me, 15195, 8, True

Formulario = "Qualidade/Inspeção de recebimento"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
cmbfiltrarpor = "Lote"
ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null'", True
ProcCarregaComboEmpresa Cmb_empresa, False
Cmb_opcao_lista = "Validação"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Qualidade/Inspeção de recebimento"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ImgAnexarplanomedicao_DblClick()
On Error GoTo tratar_erro

Acao = "anexar o controle de medição"
If txtNomenclatura = "" Then
    NomeCampo = "o produto"
    ProcVerificaAcao
    Exit Sub
End If
If cmbNivel = "" Then
    NomeCampo = "o nível"
    ProcVerificaAcao
    cmbNivel.SetFocus
    Exit Sub
End If
If Txtenc = "" Then
    NomeCampo = "a quantidade encontrada"
    ProcVerificaAcao
    Txtenc.SetFocus
    Exit Sub
End If
If Txtamostra = "" Then
    NomeCampo = "a quantidade de amostra"
    ProcVerificaAcao
    Txtamostra.SetFocus
    Exit Sub
End If
Set TBplano = CreateObject("adodb.recordset")
TBplano.Open "Select idplano from plano where desenho = '" & txtNomenclatura.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBplano.EOF = False Then
    Set TBplanomedicao = CreateObject("adodb.recordset")
    TBplanomedicao.Open "Select * from planodimensao where idplano = " & TBplano!IDPlano, Conexao, adOpenKeyset, adLockOptimistic
    If TBplanomedicao.EOF = True Then
        USMsgBox "Não foi encontrado nenhuma medição cadastrada para o plano de inspeção, favor cadastrar antes de abrir o controle de medição.", vbExclamation, "CAPRIND v5.0"
        TBplanomedicao.Close
        Exit Sub
    End If
    TBplanomedicao.Close
Else
    USMsgBox "Não foi encontrado nenhum plano de inspeção, favor cadastrar antes de abrir o controle de medição.", vbExclamation, "CAPRIND v5.0"
    TBplano.Close
    Exit Sub
End If
TBplano.Close
quantidade = 0
QTLOTE = 0
If ListaInspecionar = True Then
    USMsgBox ("Salve a inspeção de recebimento antes de anexar o controle de medição."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
Else
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Sum(quant_liberada) as quantidade from medicao where idlista = " & ListProdReceb.SelectedItem.ListSubItems(5) & " and desenho = '" & txtNomenclatura & "' and peca = '" & Txt_lote & "'", Conexao, adOpenKeyset, adLockOptimistic
End If
If TBAbrir.EOF = False Then
    quantidade = IIf(IsNull(TBAbrir!quantidade), 0, TBAbrir!quantidade)
End If
TBAbrir.Close
QTLOTE = txtLote - quantidade
Inspecaorecebimento_AnexarPlano = True
frmPlanomedicao.Show
frmPlanomedicao_ListaRNC.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListProdReceb
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) inspeção(ões) de recebimento?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from Compras_recebimento where ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Qualidade/Inspeção de recebimento"
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                Documento = "Lote: " & Txt_lote & " - Cliente/fornecedor: " & Txt_cliente_forn & " - Cód. interno: " & .ListItems(InitFor).ListSubItems(1)
                Documento1 = ""
                ProcGravaEvento
                '==================================
                
                Set TBplano = CreateObject("adodb.recordset")
                TBplano.Open "Select IDplano from medicao where idlista = " & IIf(IsNull(TBFI!IDEstoque), 0, TBFI!IDEstoque) & " and desenho = '" & .ListItems(InitFor).ListSubItems(1) & "' and id_inspecionado = " & TBFI!ID, Conexao, adOpenKeyset, adLockOptimistic
                If TBplano.EOF = False Then
                    Conexao.Execute "DELETE from Medicaodimensao where idplano = " & TBplano!IDPlano
                    TBplano.Delete
                End If
                TBplano.Close
                Conexao.Execute "DELETE from CQ_RNC where ID = " & IIf(IsNull(TBFI!ID_RNC), 0, TBFI!ID_RNC)
                Conexao.Execute "DELETE from compras_recebimento where ID =  " & TBFI!ID
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) inspeção(ões) de recebimento antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Inspeção(ões) de recebimento excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaListaInspecionar
    ProcCarregaListaInspecionados
    ProcBloquearCampos
    ProcCarregaListaRE
    
    ListaInspecionar = True
    ListaInspecionados = False
End If
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBloquearCampos()
On Error GoTo tratar_erro

Frame2.Enabled = False
Frame5.Enabled = False
Frame6.Enabled = False
Frame7.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLiberarCampos()
On Error GoTo tratar_erro

Frame2.Enabled = True
Frame5.Enabled = True
Frame6.Enabled = True
Frame7.Enabled = True

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
If Txt_lote = "" Then Exit Sub
Acao = "salvar"
If txtEspecificacoes.Text = "" Then
    NomeCampo = "o produto"
    ProcVerificaAcao
    Exit Sub
End If
If CmbNQA = "" Then
    NomeCampo = "o NQA"
    ProcVerificaAcao
    SSTab2.Tab = 1
    CmbNQA.SetFocus
    Exit Sub
End If
If cmbNivel = "" Then
    NomeCampo = "o nível"
    ProcVerificaAcao
    SSTab2.Tab = 1
    cmbNivel.SetFocus
    Exit Sub
End If
If Txtenc = "" Then
    NomeCampo = "a quantidade encontrada"
    ProcVerificaAcao
    SSTab2.Tab = 1
    Txtenc.SetFocus
    Exit Sub
End If
If Txtamostra = "" Then
    NomeCampo = "a quantidade da amostra"
    ProcVerificaAcao
    SSTab2.Tab = 1
    Txtamostra.SetFocus
    Exit Sub
End If
If ListaInspecionados = True Then
    If txtcondicional = "" Then
        NomeCampo = "a quantidade aceita"
        ProcVerificaAcao
        SSTab2.Tab = 1
        txtcondicional.SetFocus
        Exit Sub
    End If
    If Txtrejeitado = "" Then
        NomeCampo = "a quantidade rejeitada"
        ProcVerificaAcao
        SSTab2.Tab = 1
        Txtrejeitado.SetFocus
        Exit Sub
    End If
    If chk_emb_ok.Value = 0 And chk_emb_na.Value = 0 And chk_emb_nc = 0 Then
        NomeCampo = "o status da embalagem"
        ProcVerificaAcao
        SSTab2.Tab = 1
        Exit Sub
    End If
    If chk_laudo_ok.Value = 0 And chk_laudo_na.Value = 0 And chk_laudo_nc = 0 Then
        NomeCampo = "o status do laudo"
        ProcVerificaAcao
        SSTab2.Tab = 1
        Exit Sub
    End If
    If chk_qtd_ok.Value = 0 And chk_qtd_na.Value = 0 And chk_qtd_nc = 0 Then
        NomeCampo = "o status da quantidade"
        ProcVerificaAcao
        SSTab2.Tab = 1
        Exit Sub
    End If
    If chk_visual_ok.Value = 0 And chk_visual_na.Value = 0 And chk_visual_nc = 0 Then
        NomeCampo = "o status do visual"
        ProcVerificaAcao
        SSTab2.Tab = 1
        Exit Sub
    End If
    If chk_outros_ok.Value = 0 And chk_outros_na.Value = 0 And chk_outros_nc = 0 Then
        NomeCampo = "o status de outros"
        ProcVerificaAcao
        SSTab2.Tab = 1
        Exit Sub
    End If
    If cmblaudo = "" Then
        NomeCampo = "o laudo"
        ProcVerificaAcao
        cmblaudo.SetFocus
        Exit Sub
    End If
    If Cmbliberado = "" Then
        NomeCampo = "a liberação"
        ProcVerificaAcao
        Cmbliberado.SetFocus
        Exit Sub
    End If
End If

quantidade = 0
QuantEmpenho = 0
Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from Compras_recebimento where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = True Then
    TBCompras.AddNew
    Evento = "Nova"
    USMsgBox ("Produto inspecionado com sucesso."), vbInformation, "CAPRIND v5.0"
Else
    If FunVerificaRegistroValidado("Compras_recebimento", "ID = " & txtId, "mesma", "esta inspeção de recebimento", "alterar", False, True) = False Then Exit Sub
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
End If
ProcEnviaDados
TBCompras.Update
txtId = TBCompras!ID
'==================================
Modulo = "Qualidade/Inspeção de recebimento"
ID_documento = txtId
Documento = "Lote: " & Txt_lote & " - Cliente/fornecedor: " & Txt_cliente_forn & " - Cód. interno: " & txtNomenclatura
Documento1 = ""
ProcGravaEvento
'==================================

ProcCarregaListaInspecionar
ProcCarregaListaInspecionados
ListaInspecionar = False
ListaInspecionados = True
ProcCarregaListaRE

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function ProcVerifMovSaidaEst(IDEstoque As Long, Acao As String, MostrarMsg As Boolean) As Boolean
On Error GoTo tratar_erro

ProcVerifMovSaidaEst = False
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select IDoperacao from Estoque_movimentacao where IDestoque = " & IDEstoque & " and Saida > 0", Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido " & Acao & " desta inspeção, pois este RE já possui movimentação de saída."), vbExclamation, "CAPRIND v5.0"
    ProcVerifMovSaidaEst = True
End If
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select ID from Estoque_Controle_Empenho_Vendas where ID_estoque = " & IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido " & Acao & " desta inspeção, pois este RE está empenhado."), vbExclamation, "CAPRIND v5.0"
    ProcVerifMovSaidaEst = True
End If
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select ID from Producao_NF_Consignada where IDestoque = " & IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    If MostrarMsg = True Then USMsgBox ("Não é permitido " & Acao & " desta inspeção, pois este RE está empenhado."), vbExclamation, "CAPRIND v5.0"
    ProcVerifMovSaidaEst = True
End If
TBEstoque.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Sub ProcEntradaList()
On Error GoTo tratar_erro

'==========================
' Entrada de embalagem
'==========================
Select Case TBRecebidos!Embalagem
    Case 1: chk_emb_ok.Value = 1
    Case 2: chk_emb_nc.Value = 1
    Case 3: chk_emb_na.Value = 1
End Select
'==========================
' Entrada de laudo
'==========================
Select Case TBRecebidos!Laudos
    Case 1: chk_laudo_ok.Value = 1
    Case 2: chk_laudo_nc.Value = 1
    Case 3: chk_laudo_na.Value = 1
End Select
'==========================
' Entrada de quantidade
'==========================
Select Case TBRecebidos!quantidade
    Case 1: chk_qtd_ok.Value = 1
    Case 2: chk_qtd_nc.Value = 1
    Case 3: chk_qtd_na.Value = 1
End Select
'==========================
' Entrada de visual
'==========================
Select Case TBRecebidos!Visual
    Case 1: chk_visual_ok.Value = 1
    Case 2: chk_visual_nc.Value = 1
    Case 3: chk_visual_na.Value = 1
End Select
'==========================
' Entrada de dimensional
'==========================
Select Case TBRecebidos!dimensional
    Case 1: chk_dim_ok.Value = 1
    Case 2: chk_dim_nc.Value = 1
    Case 3: chk_dim_na.Value = 1
End Select
'==========================
' Entrada de outros
'==========================
Select Case TBRecebidos!Outros
    Case 1: chk_outros_ok.Value = 1
    Case 2: chk_outros_nc.Value = 1
    Case 3: chk_outros_na.Value = 1
End Select

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

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
ProcLimpaCamposPrincipal
ProcLimpaCampos
Listprod.ListItems.Clear
ListProdReceb.ListItems.Clear
mskData.Text = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
Txt_lote = Lista.SelectedItem
Txt_cliente_forn = Lista.SelectedItem.ListSubItems(1)
ProcCarregaListaInspecionar
ProcCarregaListaInspecionados
Frame3.Enabled = True
Frame8.Enabled = True
SSTab2.Tab = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listprod_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Listprod, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listprod_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
RE = 0
mskData = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
txtDtValidacao = ""
txtRespValidacao = ""
RE = Listprod.SelectedItem.ListSubItems(4)

ProcLimpaCampos
ProcLiberarCampos
Set TBCompras_Lista = CreateObject("adodb.recordset")
TBCompras_Lista.Open "Select * from Qualidade_inspecao_recebimento where IDestoque = " & Listprod.SelectedItem.ListSubItems(4), Conexao, adOpenKeyset, adLockOptimistic
If TBCompras_Lista.EOF = False Then
    txtNomenclatura = IIf(IsNull(TBCompras_Lista!Desenho), "", TBCompras_Lista!Desenho)
    txtEspecificacoes = IIf(IsNull(TBCompras_Lista!Descricao), "", TBCompras_Lista!Descricao)
    Txt_unidade = IIf(IsNull(TBCompras_Lista!Un), "", TBCompras_Lista!Un)
    Txt_qtde_recebida = Format(TBCompras_Lista!Qtde_recebida, "###,##0.0000")
    txtInspecionada = Format(TBCompras_Lista!Qtde_inspecionada, "###,##0.0000")
    txtQtdeAinspecionar = Format(TBCompras_Lista!Qtde_inspecionar, "###,##0.0000")
    Txt_nota_fiscal = IIf(IsNull(TBCompras_Lista!Nota_fiscal), "", TBCompras_Lista!Nota_fiscal)
    Txt_data_emissao_NF = IIf(IsNull(TBCompras_Lista!Data_emissao), "", Format(TBCompras_Lista!Data_emissao, "dd/mm/yy"))
    txt_Corrida = IIf(IsNull(TBCompras_Lista!Corrida), "", TBCompras_Lista!Corrida)
    txt_Certificado = IIf(IsNull(TBCompras_Lista!Certificado), "", TBCompras_Lista!Certificado)
    Txtenc = Format(TBCompras_Lista!Qtde_recebida, "###,##0.0000")
    Txtamostra = FunCalculaAmostragem(cmbNivel, IIf(Txtenc = "", 0, Txtenc))
End If
TBCompras_Lista.Close
Set TBplano = CreateObject("adodb.recordset")
TBplano.Open "Select Nivel from plano where desenho = '" & txtNomenclatura & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBplano.EOF = False Then
    If IsNull(TBplano!Nivel) = False And TBplano!Nivel <> "" Then cmbNivel = TBplano!Nivel
End If
TBplano.Close

ListaInspecionar = True
ListaInspecionados = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListProdReceb_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListProdReceb
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If Cmb_opcao_lista = "Excluir" Then
                    If FunVerificaRegistroValidadoSemMsg("Compras_recebimento", "ID = " & .ListItems(InitFor), True) = False Then
                        .ListItems.Item(InitFor).Checked = False
                        GoTo Proximo
                    End If
                ElseIf .ListItems(InitFor).ListSubItems(6) = "Sim" And .ListItems(InitFor).ListSubItems(5) <> "" Then
                        If ProcVerifMovSaidaEst(.ListItems(InitFor).ListSubItems(5), "cancelar validação", False) = True Then
                            .ListItems.Item(InitFor).Checked = False
                            GoTo Proximo
                        End If
                    ElseIf .ListItems(InitFor).ListSubItems(6) = "Não" Then
                            Set TBRecebidos = CreateObject("adodb.recordset")
                            TBRecebidos.Open "Select Liberado from Compras_recebimento where ID = " & .ListItems(InitFor) & " and (Liberado IS NULL or Liberado = N'')", Conexao, adOpenKeyset, adLockOptimistic
                            If TBRecebidos.EOF = False Then
                                .ListItems.Item(InitFor).Checked = False
                                TBRecebidos.Close
                                GoTo Proximo
                            End If
                            TBRecebidos.Close
                        
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListProdReceb, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListProdReceb_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListProdReceb
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If Cmb_opcao_lista = "Excluir" Then
                If FunVerificaRegistroValidado("Compras_recebimento", "ID = " & .ListItems(InitFor), "mesma", "inspeção de recebimento", "excluir esta", False, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
            ElseIf .ListItems(InitFor).ListSubItems(6) = "Sim" And .ListItems(InitFor).ListSubItems(5) <> "" Then
                    If ProcVerifMovSaidaEst(.ListItems(InitFor).ListSubItems(5), "cancelar a validação", True) = True Then
                        .ListItems.Item(InitFor).Checked = False
                        Exit Sub
                    End If
                ElseIf .ListItems(InitFor).ListSubItems(6) = "Não" Then
                        Set TBRecebidos = CreateObject("adodb.recordset")
                        TBRecebidos.Open "Select Liberado from Compras_recebimento where ID = " & .ListItems(InitFor) & " and (Liberado IS NULL or Liberado = N'')", Conexao, adOpenKeyset, adLockOptimistic
                        If TBRecebidos.EOF = False Then
                            USMsgBox ("É necessário informar se a inspeção foi liberada antes de validar."), vbExclamation, "CAPRIND v5.0"
                            .ListItems.Item(InitFor).Checked = False
                            TBRecebidos.Close
                             Exit Sub
                        End If
                        TBRecebidos.Close
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListProdReceb_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

ProcCarregaDadosInsp ListProdReceb.SelectedItem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaDadosInsp(IDinsp As Long)
On Error GoTo tratar_erro

ProcLimpaCampos
ProcLiberarCampos
Set TBRecebidos = CreateObject("adodb.recordset")
TBRecebidos.Open "Select QIR.Desenho, QIR.Descricao as Descricao_prod, QIR.Un, QIR.Qtde_recebida, QIR.Qtde_inspecionada, QIR.Qtde_inspecionar, QIR.Nota_fiscal As NF, QIR.Data_emissao As DE, QIR.Corrida As Corr, QIR.Certificado AS Cert, CR.* from Qualidade_inspecao_recebimento QIR INNER JOIN Compras_recebimento CR ON CR.IDestoque = QIR.IDestoque where CR.ID = " & ListProdReceb.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBRecebidos.EOF = False Then
    txtId = TBRecebidos!ID
    mskData = IIf(IsNull(TBRecebidos!Data), "", Format(TBRecebidos!Data, "dd/mm/yy"))
    txtResponsavel = IIf(IsNull(TBRecebidos!Responsavel), "", TBRecebidos!Responsavel)
    txtDtValidacao = IIf(IsNull(TBRecebidos!DtValidacao), "", TBRecebidos!DtValidacao)
    txtRespValidacao = IIf(IsNull(TBRecebidos!RespValidacao), "", TBRecebidos!RespValidacao)
    
    txtNomenclatura = IIf(IsNull(TBRecebidos!Desenho), "", TBRecebidos!Desenho)
    txtEspecificacoes = IIf(IsNull(TBRecebidos!Descricao_prod), "", TBRecebidos!Descricao_prod)
    Txt_unidade = IIf(IsNull(TBRecebidos!Un), "", TBRecebidos!Un)
    Txt_qtde_recebida = Format(TBRecebidos!Qtde_recebida, "###,##0.0000")
    txtInspecionada = Format(TBRecebidos!Qtde_inspecionada, "###,##0.0000")
    txtQtdeAinspecionar = Format(TBRecebidos!Qtde_inspecionar, "###,##0.0000")
    Txt_nota_fiscal = IIf(IsNull(TBRecebidos!NF), "", TBRecebidos!NF)
    Txt_data_emissao_NF = IIf(IsNull(TBRecebidos!De), "", Format(TBRecebidos!De, "dd/mm/yy"))
    txt_Corrida = IIf(IsNull(TBRecebidos!Corr), "", TBRecebidos!Corr)
    txt_Certificado = IIf(IsNull(TBRecebidos!Cert), "", TBRecebidos!Cert)
    txtObservacoes = IIf(IsNull(TBRecebidos!Obs), "", TBRecebidos!Obs)
    CmbNQA = IIf(TBRecebidos!NQA = 0, "Liberação contra laudo", TBRecebidos!NQA)
    cmbNivel = IIf(IsNull(TBRecebidos!Nivel), "", TBRecebidos!Nivel)
    txtcondicional = IIf(IsNull(TBRecebidos!AC), "", Format(TBRecebidos!AC, "###,##0.0000"))
    Txtrejeitado = IIf(IsNull(TBRecebidos!RJ), "", Format(TBRecebidos!RJ, "###,##0.0000"))
    Txtenc = IIf(IsNull(TBRecebidos!Enc), "", Format(TBRecebidos!Enc, "###,##0.0000"))
    Txtamostra = IIf(IsNull(TBRecebidos!AMOSTRA), "", Format(TBRecebidos!AMOSTRA, "###,##0.0000"))
    
    Txt_ID_RNC = IIf(IsNull(TBRecebidos!ID_RNC), 0, TBRecebidos!ID_RNC)
    Set TBCompras_Pedido = CreateObject("adodb.recordset")
    TBCompras_Pedido.Open "Select ID_texto, Seq FROM CQ_RNC where ID = " & Txt_ID_RNC, Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Pedido.EOF = False Then
        Txtsac = IIf(IsNull(TBCompras_Pedido!Seq), TBCompras_Pedido!id_texto, TBCompras_Pedido!id_texto & "/" & IIf(TBCompras_Pedido!Seq < 10, "0" & TBCompras_Pedido!Seq, TBCompras_Pedido!Seq))
    End If
    TBCompras_Pedido.Close
    
    If IsNull(TBRecebidos!Laudo) = False And TBRecebidos!Laudo <> "" Then cmblaudo.Text = TBRecebidos!Laudo
    If IsNull(TBRecebidos!Liberado) = False And TBRecebidos!Liberado <> "" Then Cmbliberado.Text = TBRecebidos!Liberado
    txt_Caminho = IIf(IsNull(TBRecebidos!caminho), "", TBRecebidos!caminho)
    Txt_caminho2 = IIf(IsNull(TBRecebidos!caminho2), "", TBRecebidos!caminho2)
    
    ProcEntradaList
End If
TBRecebidos.Close
ListaInspecionar = False
ListaInspecionados = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Sub ProcCarregaListaInspecionar()
On Error GoTo tratar_erro

Listprod.ListItems.Clear
CamposFiltro = "Desenho, Descricao, Un, Qtde_inspecionar, IDestoque"
Set TBLISTA = CreateObject("adodb.recordset")
'StrSql = "Select " & CamposFiltro & " from Qualidade_inspecao_recebimento where Lote = '" & Txt_lote & "' and Cliente_forn = '" & Txt_cliente_forn & "' and Qtde_inspecionar > 0 group by " & CamposFiltro & " order by Desenho, IDestoque"
StrSql = "Select " & CamposFiltro & " from Qualidade_inspecao_recebimento where Lote = '" & Txt_lote & "' and Cliente_forn = '" & Txt_cliente_forn & "'  group by " & CamposFiltro & " order by Desenho, IDestoque"

'Debug.print StrSql

TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With Listprod.ListItems
            .Add , , TBLISTA!Desenho
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Un), "", TBLISTA!Un)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Qtde_inspecionar), 0, Format(TBLISTA!Qtde_inspecionar, "###,##0.0000"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!IDEstoque), "", TBLISTA!IDEstoque)
        End With
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaInspecionados()
On Error GoTo tratar_erro

ListProdReceb.ListItems.Clear
CamposFiltro = "CR.ID, QIR.Desenho, QIR.Descricao, QIR.Un, CR.Enc, CR.IDestoque, CR.Dtvalidacao"
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select " & CamposFiltro & " from Qualidade_inspecao_recebimento QIR INNER JOIN Compras_recebimento CR ON CR.IDestoque = QIR.IDestoque where QIR.Lote = '" & Txt_lote & "' and QIR.Cliente_forn = '" & Txt_cliente_forn & "' and CR.Data IS NOT NULL group by " & CamposFiltro & "  order by QIR.Desenho, CR.IDestoque", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With ListProdReceb.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Desenho), "", TBLISTA!Desenho)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Un), "", TBLISTA!Un)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Enc), 0, Format(TBLISTA!Enc, "###,##0.0000"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!IDEstoque), "", TBLISTA!IDEstoque)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!DtValidacao), "Não", "Sim")
        End With
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_inspecionados_Click()
On Error GoTo tratar_erro

ProcLimpaTudoFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_inspecionar_Click()
On Error GoTo tratar_erro

ProcLimpaTudoFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ProcLimpaTudoFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ProcLimpaTudoFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ProcLimpaTudoFiltro

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0: If Lista.Visible = True Then Lista.SetFocus
    Case 1: If Listprod.Enabled = True Then Listprod.SetFocus
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Select Case SSTab2.Tab
    Case 0: If Frame2.Enabled = True Then txtNomenclatura.SetFocus
    Case 1: If Frame5.Enabled = True Then CmbNQA.SetFocus
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txtamostra_Change()
On Error GoTo tratar_erro

If Txtamostra.Text <> "" Then
    VerifNumero = Txtamostra.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        'Txtamostra.Text = ""
        'Txtamostra.SetFocus
       ' Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txtamostra_LostFocus()
On Error GoTo tratar_erro

Txtamostra.Text = Format(Txtamostra.Text, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtcondicional_Change()
On Error GoTo tratar_erro

If txtcondicional.Text <> "" Then
    VerifNumero = txtcondicional.Text
    ProcVerificaNumero
    If VerifNumero = False Then
       ' txtcondicional.Text = ""
       ' txtcondicional.SetFocus
       ' Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtcondicional_LostFocus()
On Error GoTo tratar_erro

txtcondicional.Text = Format(txtcondicional.Text, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txtenc_LostFocus()
On Error GoTo tratar_erro

Txtenc.Text = Format(Txtenc.Text, "###,##0.0000")

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

Private Sub Txtrejeitado_Change()
On Error GoTo tratar_erro

If Txtrejeitado.Text <> "" Then
    VerifNumero = Txtrejeitado.Text
    ProcVerificaNumero
    If VerifNumero = False Then
      '  Txtrejeitado.Text = ""
      '  Txtrejeitado.SetFocus
      '  Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txtrejeitado_LostFocus()
On Error GoTo tratar_erro

Txtrejeitado.Text = Format(Txtrejeitado.Text, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

ProcLimpaTudoFiltro
If cmbfiltrarpor = "RE" And txtTexto <> "" Then
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
    Exit Sub
End Sub

Private Sub ProcLimpaTudoFiltro()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Listprod.ListItems.Clear
ListProdReceb.ListItems.Clear
ProcLimpaCamposPrincipal
ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_LostFocus()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Nota fiscal" And txtTexto <> "" Then txtTexto = FunTamanhoTextoZeroEsq(txtTexto, 9)

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

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    Case 2: ProcExcluir
    Case 3: ProcImprimir
    Case 4: ProcValidarRegistros ListProdReceb, "Qualidade/Inspeção de recebimento"
    'Case 6: ProcAjuda
    Case 7: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
