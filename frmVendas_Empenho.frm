VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_Empenho 
   Caption         =   "Administrativo - Vendas - Empenho"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15360
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
   Icon            =   "frmVendas_Empenho.frx":0000
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
      TabIndex        =   75
      Top             =   30
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   17701
      _Version        =   393216
      Tabs            =   5
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
      TabCaption(0)   =   "Lista de produtos"
      TabPicture(0)   =   "frmVendas_Empenho.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "PBLista"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "USToolBar1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "USImageList1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Estrutura"
      TabPicture(1)   =   "frmVendas_Empenho.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "USImageList2"
      Tab(1).Control(2)=   "USToolBar2"
      Tab(1).Control(3)=   "Grid1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Vendidos"
      TabPicture(2)   =   "frmVendas_Empenho.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "ListaVendidos"
      Tab(2).Control(2)=   "USImageList3"
      Tab(2).Control(3)=   "USToolBar3"
      Tab(2).Control(4)=   "PBLista1"
      Tab(2).Control(5)=   "PBLista2"
      Tab(2).Control(6)=   "SSTab2"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Estoque"
      TabPicture(3)   =   "frmVendas_Empenho.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).Control(1)=   "USImageList4"
      Tab(3).Control(2)=   "USToolBar4"
      Tab(3).Control(3)=   "PBLista3"
      Tab(3).Control(4)=   "SSTab3"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Produzindo"
      TabPicture(4)   =   "frmVendas_Empenho.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "SSTab4"
      Tab(4).Control(1)=   "PBLista4"
      Tab(4).Control(2)=   "USToolBar6"
      Tab(4).Control(3)=   "USImageList6"
      Tab(4).Control(4)=   "Frame6"
      Tab(4).ControlCount=   5
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
         Height          =   675
         Left            =   -74945
         TabIndex        =   135
         Top             =   1320
         Width           =   15195
         Begin VB.ComboBox cmbVersao_pesquisar_estrutura 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            ItemData        =   "frmVendas_Empenho.frx":0098
            Left            =   2070
            List            =   "frmVendas_Empenho.frx":00EA
            Style           =   2  'Dropdown List
            TabIndex        =   17
            ToolTipText     =   "Versão."
            Top             =   210
            Width           =   795
         End
         Begin VB.Image imgFile 
            Height          =   240
            Left            =   10830
            Picture         =   "frmVendas_Empenho.frx":013C
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgFolder 
            Height          =   240
            Left            =   10560
            Picture         =   "frmVendas_Empenho.frx":06C6
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesquisa por versão :"
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
            Index           =   19
            Left            =   180
            TabIndex        =   136
            Top             =   210
            Width           =   1800
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
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
         Left            =   -74940
         TabIndex        =   108
         Top             =   9120
         Width           =   15195
         Begin VB.TextBox Txt_qtde_total_disp_produzindo 
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
            Left            =   13440
            Locked          =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade disponível."
            Top             =   420
            Width           =   1575
         End
         Begin VB.TextBox Txt_qtde_total_emp_produzindo 
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
            Left            =   11550
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            ToolTipText     =   "Quatidade total empenhada."
            Top             =   420
            Width           =   1575
         End
         Begin VB.TextBox Txt_qtde_total_produzindo 
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
            Left            =   9720
            Locked          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade total produzindo."
            Top             =   420
            Width           =   1575
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. disponível"
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
            Left            =   13552
            TabIndex        =   113
            Top             =   210
            Width           =   1350
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. empenhada"
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
            Left            =   11580
            TabIndex        =   112
            Top             =   210
            Width           =   1500
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. produzindo"
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
            Left            =   9787
            TabIndex        =   111
            Top             =   210
            Width           =   1440
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            Left            =   0
            TabIndex        =   110
            Top             =   0
            Width           =   75
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-                                       ="
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
            Left            =   11370
            TabIndex        =   109
            Top             =   480
            Width           =   1965
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
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
         Left            =   -74940
         TabIndex        =   89
         Top             =   9120
         Width           =   15195
         Begin VB.TextBox Txt_qtde_total_estoque 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   9720
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade total em estoque."
            Top             =   420
            Width           =   1575
         End
         Begin VB.TextBox Txt_qtde_total_emp_estoque 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   11550
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            ToolTipText     =   "Quatidade total empenhada."
            Top             =   420
            Width           =   1575
         End
         Begin VB.TextBox Txt_qtde_total_disp_estoque 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   13440
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade disponível."
            Top             =   420
            Width           =   1575
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-                                       ="
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
            Left            =   11370
            TabIndex        =   99
            Top             =   480
            Width           =   1965
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
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
            Left            =   0
            TabIndex        =   93
            Top             =   0
            Width           =   75
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total em estoque"
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
            Left            =   9885
            TabIndex        =   92
            Top             =   210
            Width           =   1245
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total empenhado"
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
            Left            =   11700
            TabIndex        =   91
            Top             =   210
            Width           =   1245
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total disponível"
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
            Left            =   13635
            TabIndex        =   90
            Top             =   210
            Width           =   1110
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
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
         Left            =   -74940
         TabIndex        =   87
         Top             =   5460
         Width           =   15195
         Begin VB.TextBox Txt_necessidade_vendidos 
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
            Left            =   13350
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "Necessidade."
            Top             =   390
            Width           =   1665
         End
         Begin VB.TextBox Txt_qtde_total_emp_prod 
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
            Left            =   11370
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade total empenhada produzindo."
            Top             =   390
            Width           =   1665
         End
         Begin VB.TextBox Txt_qtde_total_emp_est 
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
            Left            =   9450
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade total empenhada em estoque."
            Top             =   390
            Width           =   1665
         End
         Begin VB.TextBox Txt_qtde_total_vendidos 
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
            Left            =   7500
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade total vendida."
            Top             =   390
            Width           =   1665
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " -                                         -                                         ="
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
            Left            =   9240
            TabIndex        =   97
            Top             =   450
            Width           =   4020
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Necessidade"
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
            Left            =   13657
            TabIndex        =   96
            Top             =   180
            Width           =   1050
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. emp. prod."
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
            Left            =   11505
            TabIndex        =   95
            Top             =   180
            Width           =   1395
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. emp. est."
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
            Left            =   9645
            TabIndex        =   94
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. vendida"
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
            Left            =   7747
            TabIndex        =   88
            Top             =   180
            Width           =   1170
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   55
         TabIndex        =   62
         Top             =   1330
         Width           =   15195
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            Height          =   510
            Left            =   10260
            TabIndex        =   134
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
               TabIndex        =   8
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
               TabIndex        =   6
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
               TabIndex        =   7
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
               TabIndex        =   9
               Top             =   180
               Width           =   705
            End
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
            ItemData        =   "frmVendas_Empenho.frx":0C50
            Left            =   180
            List            =   "frmVendas_Empenho.frx":0C52
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Empresa."
            Top             =   390
            Width           =   6555
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
            ItemData        =   "frmVendas_Empenho.frx":0C54
            Left            =   6750
            List            =   "frmVendas_Empenho.frx":0C7F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   1
            ToolTipText     =   "Opções para filtro."
            Top             =   390
            Width           =   3435
         End
         Begin VB.TextBox txtTexto 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   180
            TabIndex        =   2
            ToolTipText     =   "Texto para pesquisa."
            Top             =   1020
            Width           =   12585
         End
         Begin VB.ComboBox Cmb_ordenar 
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
            ItemData        =   "frmVendas_Empenho.frx":0D36
            Left            =   12780
            List            =   "frmVendas_Empenho.frx":0D40
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "Ordenar por."
            Top             =   1020
            Width           =   2265
         End
         Begin VB.ComboBox cmbfamilia 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   180
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "Texto para pesquisa."
            Top             =   1020
            Visible         =   0   'False
            Width           =   12585
         End
         Begin VB.Label Label44 
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
            Left            =   3090
            TabIndex        =   86
            Top             =   180
            Width           =   735
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
            Index           =   0
            Left            =   5737
            TabIndex        =   65
            Top             =   810
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
            Left            =   8047
            TabIndex        =   64
            Top             =   180
            Width           =   840
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ordenar por"
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
            Left            =   13402
            TabIndex        =   63
            Top             =   810
            Width           =   1020
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   55
         TabIndex        =   58
         Top             =   9090
         Width           =   15195
         Begin VB.TextBox txtPagIr 
            Height          =   315
            Left            =   9540
            TabIndex        =   11
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txtNreg 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   3780
            TabIndex        =   10
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   15
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_Empenho.frx":0D5F
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
            TabIndex        =   14
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_Empenho.frx":4503
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
            TabIndex        =   12
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
            TabIndex        =   13
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_Empenho.frx":800C
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
            TabIndex        =   16
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmVendas_Empenho.frx":C0FB
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
         Begin VB.Label Label3 
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
            TabIndex        =   137
            Top             =   240
            Width           =   1440
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
            TabIndex        =   61
            Top             =   240
            Width           =   1275
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
            TabIndex        =   60
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label24 
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
            TabIndex        =   59
            Top             =   240
            Width           =   645
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6225
         Left            =   60
         TabIndex        =   5
         Top             =   2835
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   10980
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "Cód."
            Object.Width           =   0
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
            Object.Width           =   6880
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Un. est."
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Un. com."
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   3621
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Qtde. vend."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Qtde. est."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Emp. est."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Disp. est."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Object.Tag             =   "N"
            Text            =   "Qtde. prod."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Object.Tag             =   "N"
            Text            =   "Emp. prod."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Object.Tag             =   "N"
            Text            =   "Disp. prod."
            Object.Width           =   1587
         EndProperty
      End
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   5310
         Top             =   465
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_Empenho.frx":F987
         Count           =   1
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   55
         TabIndex        =   66
         Top             =   350
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
      End
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   -70950
         Top             =   555
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_Empenho.frx":11B6F
         Count           =   1
      End
      Begin MSComctlLib.ListView ListaVendidos 
         Height          =   3810
         Left            =   -74940
         TabIndex        =   19
         Top             =   1350
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   6720
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Ped. int."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Rev."
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Cliente"
            Object.Width           =   5300
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Vendedor interno"
            Object.Width           =   5027
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Vendedor externo"
            Object.Width           =   5027
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "D"
            Text            =   "Pr. final"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Qtde. vend."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Emp. est."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Emp. prod."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Object.Tag             =   "N"
            Text            =   "Necessidade"
            Object.Width           =   1764
         EndProperty
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74945
         TabIndex        =   67
         Top             =   345
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   3
         GradientColor2  =   14737632
         GradientColorOverRight1=   16315633
         GradientColorOverRight2=   15195350
         GripperColor    =   15195350
         IsStrech        =   -1  'True
         RightColor1     =   0
         RightColor2     =   0
         ShowEndPanel    =   0   'False
         Theme           =   1
         ButtonCaption1  =   "Ajuda"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Ajuda (F1)"
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
         ButtonCaption2  =   "Sair"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Sair (Esc)"
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
         ButtonWidth2    =   26
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonKey3      =   "3"
         BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState3    =   5
         ButtonLeft3     =   68
         ButtonTop3      =   2
         ButtonWidth3    =   24
         ButtonHeight3   =   24
         ButtonUseMaskColor3=   0   'False
      End
      Begin DrawSuite2022.USImageList USImageList3 
         Left            =   -70980
         Top             =   645
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_Empenho.frx":13194
         Count           =   1
      End
      Begin DrawSuite2022.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74940
         TabIndex        =   68
         Top             =   345
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
         ButtonCaption1  =   "Excluir"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Excluir (F4)"
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
         ButtonWidth1    =   39
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonAlignment2=   2
         ButtonType2     =   1
         ButtonStyle2    =   -1
         BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState2    =   -1
         ButtonLeft2     =   43
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
         ButtonLeft3     =   47
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
         ButtonLeft4     =   85
         ButtonTop4      =   2
         ButtonWidth4    =   26
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonKey5      =   "3"
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
         ButtonLeft5     =   113
         ButtonTop5      =   2
         ButtonWidth5    =   24
         ButtonHeight5   =   24
         ButtonUseMaskColor5=   0   'False
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   55
         TabIndex        =   69
         Top             =   9705
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
      Begin DrawSuite2022.USImageList USImageList4 
         Left            =   -71040
         Top             =   480
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_Empenho.frx":1551A
         Count           =   1
      End
      Begin DrawSuite2022.USToolBar USToolBar4 
         Height          =   975
         Left            =   -74940
         TabIndex        =   84
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   3
         GradientColor2  =   14737632
         GradientColorOverRight1=   16315633
         GradientColorOverRight2=   15195350
         GripperColor    =   15195350
         IsStrech        =   -1  'True
         RightColor1     =   0
         RightColor2     =   0
         ShowEndPanel    =   0   'False
         Theme           =   1
         ButtonCaption1  =   "Ajuda"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Ajuda (F1)"
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
         ButtonCaption2  =   "Sair"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Sair (Esc)"
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
         ButtonWidth2    =   26
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonKey3      =   "3"
         BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState3    =   5
         ButtonLeft3     =   68
         ButtonTop3      =   2
         ButtonWidth3    =   24
         ButtonHeight3   =   24
         ButtonUseMaskColor3=   0   'False
      End
      Begin DrawSuite2022.USImageList USImageList6 
         Left            =   -71010
         Top             =   630
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_Empenho.frx":16B3F
         Count           =   1
      End
      Begin DrawSuite2022.USToolBar USToolBar6 
         Height          =   975
         Left            =   -74940
         TabIndex        =   85
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   3
         GradientColor2  =   14737632
         GradientColorOverRight1=   16315633
         GradientColorOverRight2=   15195350
         GripperColor    =   15195350
         IsStrech        =   -1  'True
         RightColor1     =   0
         RightColor2     =   0
         ShowEndPanel    =   0   'False
         Theme           =   1
         ButtonCaption1  =   "Ajuda"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Ajuda (F1)"
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
         ButtonCaption2  =   "Sair"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Sair (Esc)"
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
         ButtonWidth2    =   26
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonKey3      =   "3"
         BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState3    =   5
         ButtonLeft3     =   68
         ButtonTop3      =   2
         ButtonWidth3    =   24
         ButtonHeight3   =   24
         ButtonUseMaskColor3=   0   'False
      End
      Begin DrawSuite2022.USProgressBar PBLista1 
         Height          =   255
         Left            =   -74940
         TabIndex        =   98
         Top             =   5160
         Width           =   15225
         _ExtentX        =   26855
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
      Begin DrawSuite2022.USProgressBar PBLista2 
         Height          =   255
         Left            =   -74940
         TabIndex        =   100
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
      Begin TabDlg.SSTab SSTab2 
         Height          =   3735
         Left            =   -75000
         TabIndex        =   70
         Top             =   6330
         Width           =   15390
         _ExtentX        =   27146
         _ExtentY        =   6588
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
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
         TabCaption(0)   =   "RE's empenhados"
         TabPicture(0)   =   "frmVendas_Empenho.frx":18164
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "ListaEmpenhoVendidosRE"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Ordens empenhadas"
         TabPicture(1)   =   "frmVendas_Empenho.frx":18180
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "ListaEmpenhoVendidosOP"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin MSComctlLib.ListView ListaEmpenhoVendidosRE 
            Height          =   3045
            Left            =   -74940
            TabIndex        =   24
            Top             =   330
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   5371
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
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "N"
               Text            =   "RE"
               Object.Width           =   2469
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
               Text            =   "Lote"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Local de armazenamento"
               Object.Width           =   5830
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Corrida"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Certificado"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Object.Tag             =   "N"
               Text            =   "Qtde. emp."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Object.Tag             =   "N"
               Text            =   "Qtde. saída"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Object.Tag             =   "N"
               Text            =   "Saldo"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSComctlLib.ListView ListaEmpenhoVendidosOP 
            Height          =   3045
            Left            =   60
            TabIndex        =   25
            Top             =   330
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   5371
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
               Object.Tag             =   "N"
               Text            =   "Ordem"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Cliente"
               Object.Width           =   15178
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Object.Tag             =   "D"
               Text            =   "Pr. final"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Qtde. emp."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "Qtde. entr."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Saldo"
               Object.Width           =   2117
            EndProperty
         End
      End
      Begin DrawSuite2022.USProgressBar PBLista3 
         Height          =   255
         Left            =   -74940
         TabIndex        =   101
         Top             =   8850
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
      Begin TabDlg.SSTab SSTab3 
         Height          =   8745
         Left            =   -75000
         TabIndex        =   71
         Top             =   1320
         Width           =   15390
         _ExtentX        =   27146
         _ExtentY        =   15425
         _Version        =   393216
         Tab             =   1
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
         TabCaption(0)   =   "Em estoque"
         TabPicture(0)   =   "frmVendas_Empenho.frx":1819C
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "ListaEstoque"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Empenhos (Pedido interno)"
         TabPicture(1)   =   "frmVendas_Empenho.frx":181B8
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "ListaEmpenhoEstoquePed"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "USToolBar5"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Frame3"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Txt_ID_emp_est"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Txt_ID_carteira_est"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "Empenhos (Ordem)"
         TabPicture(2)   =   "frmVendas_Empenho.frx":181D4
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "ListaEmpenhoEstoqueOrdem"
         Tab(2).ControlCount=   1
         Begin VB.TextBox Txt_ID_carteira_est 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5310
            TabIndex        =   107
            Text            =   "0"
            Top             =   3630
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox Txt_ID_emp_est 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   4650
            TabIndex        =   102
            Text            =   "0"
            Top             =   3630
            Visible         =   0   'False
            Width           =   645
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
            Height          =   1425
            Left            =   60
            TabIndex        =   73
            Top             =   1320
            Width           =   15195
            Begin VB.TextBox Txt_responsavel_est 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1380
               Locked          =   -1  'True
               TabIndex        =   130
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   955
               Width           =   10845
            End
            Begin VB.TextBox Txt_data_est 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   180
               Locked          =   -1  'True
               MaxLength       =   25
               TabIndex        =   129
               TabStop         =   0   'False
               ToolTipText     =   "Data do cadastro."
               Top             =   955
               Width           =   1185
            End
            Begin VB.TextBox Txt_estoque_disp_est 
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
               Left            =   12240
               Locked          =   -1  'True
               TabIndex        =   32
               TabStop         =   0   'False
               ToolTipText     =   "Estoque disponível."
               Top             =   955
               Width           =   1305
            End
            Begin VB.TextBox Txt_pedido_est 
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
               TabIndex        =   27
               TabStop         =   0   'False
               ToolTipText     =   "Pedido interno."
               Top             =   390
               Width           =   1335
            End
            Begin VB.TextBox Txt_rev_est 
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
               Left            =   1530
               Locked          =   -1  'True
               TabIndex        =   28
               TabStop         =   0   'False
               ToolTipText     =   "Revisão."
               Top             =   390
               Width           =   555
            End
            Begin VB.TextBox Txt_necessidade_vendas_est 
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
               Left            =   13710
               Locked          =   -1  'True
               TabIndex        =   31
               TabStop         =   0   'False
               ToolTipText     =   "Necessidade do pedido."
               Top             =   390
               Width           =   1305
            End
            Begin VB.TextBox Txt_cliente_est 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   2100
               Locked          =   -1  'True
               TabIndex        =   29
               TabStop         =   0   'False
               ToolTipText     =   "Cliente."
               Top             =   390
               Width           =   10425
            End
            Begin VB.TextBox Txt_prazo_est 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   12540
               Locked          =   -1  'True
               TabIndex        =   30
               TabStop         =   0   'False
               ToolTipText     =   "Prazo final."
               Top             =   390
               Width           =   1155
            End
            Begin VB.TextBox Txt_qtde_emp_est 
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
               Left            =   13560
               TabIndex        =   33
               ToolTipText     =   "Quantidade à empenhar."
               Top             =   955
               Width           =   1455
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Index           =   2
               Left            =   6345
               TabIndex        =   128
               Top             =   765
               Width           =   915
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   600
               TabIndex        =   127
               Top             =   765
               Width           =   345
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Estoque disp."
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
               Left            =   12330
               TabIndex        =   104
               Top             =   765
               Width           =   1110
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Necessidade"
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
               Left            =   11925
               TabIndex        =   103
               Top             =   -240
               Width           =   900
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ped. interno"
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
               Left            =   337
               TabIndex        =   80
               Top             =   180
               Width           =   1020
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
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
               Left            =   1620
               TabIndex        =   79
               Top             =   180
               Width           =   375
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Necessidade"
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
               Left            =   13830
               TabIndex        =   78
               Top             =   180
               Width           =   1050
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
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
               Left            =   7065
               TabIndex        =   77
               Top             =   180
               Width           =   495
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Prazo final"
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
               Left            =   12735
               TabIndex        =   76
               Top             =   180
               Width           =   750
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. empenhar"
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
               Left            =   13620
               TabIndex        =   74
               Top             =   765
               Width           =   1365
            End
         End
         Begin DrawSuite2022.USProgressBar USProgressBar2 
            Height          =   255
            Left            =   -74970
            TabIndex        =   72
            Top             =   3390
            Width           =   15225
            _ExtentX        =   26855
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
         Begin DrawSuite2022.USToolBar USToolBar5 
            Height          =   975
            Left            =   60
            TabIndex        =   81
            Top             =   330
            Width           =   15195
            _ExtentX        =   26802
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
            ButtonToolTipText2=   "Salvar"
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
            ButtonToolTipText3=   "Excluir"
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
            BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
               Left            =   3480
               Top             =   240
               _ExtentX        =   900
               _ExtentY        =   767
               Img1            =   "frmVendas_Empenho.frx":181F0
               Count           =   1
            End
         End
         Begin MSComctlLib.ListView ListaEstoque 
            Height          =   7185
            Left            =   -74940
            TabIndex        =   26
            Top             =   330
            Width           =   15285
            _ExtentX        =   26961
            _ExtentY        =   12674
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
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Text            =   "RE"
               Object.Width           =   1587
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
               Text            =   "Lote"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Local de armazenamento"
               Object.Width           =   6412
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Corrida"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Certificado"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Qtde. est."
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Object.Tag             =   "N"
               Text            =   "Qtde. emp."
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Object.Tag             =   "N"
               Text            =   "Est. disp."
               Object.Width           =   2293
            EndProperty
         End
         Begin MSComctlLib.ListView ListaEmpenhoEstoquePed 
            Height          =   4755
            Left            =   60
            TabIndex        =   34
            Top             =   2760
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   8387
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
            NumItems        =   14
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "N"
               Text            =   "ID_carteira"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Object.Tag             =   "D"
               Text            =   "Data"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Responsável"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Ped. int."
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Rev."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Cliente"
               Object.Width           =   3888
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Tag             =   "T"
               Text            =   "Vendedor interno"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Tag             =   "T"
               Text            =   "Vendedor externo"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Object.Tag             =   "D"
               Text            =   "Pr. final"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Object.Tag             =   "N"
               Text            =   "Qtde. vend."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Object.Tag             =   "N"
               Text            =   "Qtde. emp."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   12
               Object.Tag             =   "N"
               Text            =   "Qtde. saída"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   13
               Object.Tag             =   "N"
               Text            =   "Saldo"
               Object.Width           =   1587
            EndProperty
         End
         Begin MSComctlLib.ListView ListaEmpenhoEstoqueOrdem 
            Height          =   7185
            Left            =   -74940
            TabIndex        =   35
            Top             =   330
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   12674
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
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Object.Tag             =   "D"
               Text            =   "Data"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Responsável"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Ordem"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Cliente"
               Object.Width           =   11298
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Object.Tag             =   "D"
               Text            =   "Pr. final"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Qtde."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Object.Tag             =   "N"
               Text            =   "Qtde. emp."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Object.Tag             =   "N"
               Text            =   "Qtde. saída"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Object.Tag             =   "N"
               Text            =   "Saldo"
               Object.Width           =   1587
            EndProperty
         End
      End
      Begin DrawSuite2022.USProgressBar PBLista4 
         Height          =   255
         Left            =   -74940
         TabIndex        =   105
         Top             =   8850
         Width           =   15225
         _ExtentX        =   26855
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
      Begin TabDlg.SSTab SSTab4 
         Height          =   8745
         Left            =   -75000
         TabIndex        =   82
         Top             =   1320
         Width           =   15390
         _ExtentX        =   27146
         _ExtentY        =   15425
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
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
         TabCaption(0)   =   "Produzindo"
         TabPicture(0)   =   "frmVendas_Empenho.frx":19FAB
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "MSFlexGrid1"
         Tab(0).Control(1)=   "ListaProduzindo"
         Tab(0).Control(2)=   "Frame10"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Empenhos"
         TabPicture(1)   =   "frmVendas_Empenho.frx":19FC7
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "ListaEmpenhoProduzindo"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "USToolBar7"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Txt_ID_carteira_prod"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Txt_ID_emp_prod"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Frame8"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         Begin VB.Frame Frame10 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Legenda"
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
            Height          =   1425
            Left            =   -61290
            TabIndex        =   133
            Top             =   6090
            Width           =   1545
            Begin VB.CheckBox chkCancelada 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Cancelada"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   150
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   1050
               Width           =   1215
            End
            Begin VB.CheckBox chkProduzir 
               BackColor       =   &H000000FF&
               Caption         =   "A produzir"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   150
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   210
               Width           =   1215
            End
            Begin VB.CheckBox chkConcluida 
               BackColor       =   &H0000FF00&
               Caption         =   "Concluída"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   150
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   765
               Width           =   1215
            End
            Begin VB.CheckBox chkProduzindo 
               BackColor       =   &H0000FFFF&
               Caption         =   "Produzindo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   150
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   495
               Width           =   1215
            End
         End
         Begin VB.Frame Frame8 
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
            Height          =   1425
            Left            =   60
            TabIndex        =   116
            Top             =   1320
            Width           =   15195
            Begin VB.TextBox Txt_data_prod 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   180
               Locked          =   -1  'True
               MaxLength       =   25
               TabIndex        =   50
               TabStop         =   0   'False
               ToolTipText     =   "Data do cadastro."
               Top             =   955
               Width           =   1185
            End
            Begin VB.TextBox Txt_responsavel_prod 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1380
               Locked          =   -1  'True
               TabIndex        =   51
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   955
               Width           =   10845
            End
            Begin VB.TextBox Txt_qtde_emp_prod 
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
               Left            =   13560
               TabIndex        =   53
               ToolTipText     =   "Quantidade à empenhar."
               Top             =   955
               Width           =   1455
            End
            Begin VB.TextBox Txt_prazo_prod 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   12540
               Locked          =   -1  'True
               TabIndex        =   48
               TabStop         =   0   'False
               ToolTipText     =   "Prazo final."
               Top             =   390
               Width           =   1155
            End
            Begin VB.TextBox Txt_cliente_prod 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   2100
               Locked          =   -1  'True
               TabIndex        =   47
               TabStop         =   0   'False
               ToolTipText     =   "Cliente."
               Top             =   390
               Width           =   10425
            End
            Begin VB.TextBox Txt_necessidade_vendas_prod 
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
               Left            =   13710
               Locked          =   -1  'True
               TabIndex        =   49
               TabStop         =   0   'False
               ToolTipText     =   "Necessidade do pedido."
               Top             =   390
               Width           =   1305
            End
            Begin VB.TextBox Txt_rev_prod 
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
               Left            =   1530
               Locked          =   -1  'True
               TabIndex        =   46
               TabStop         =   0   'False
               ToolTipText     =   "Revisão."
               Top             =   390
               Width           =   555
            End
            Begin VB.TextBox Txt_pedido_prod 
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
               TabIndex        =   45
               TabStop         =   0   'False
               ToolTipText     =   "Pedido interno."
               Top             =   390
               Width           =   1335
            End
            Begin VB.TextBox Txt_produzindo_disp_prod 
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
               Left            =   12240
               Locked          =   -1  'True
               TabIndex        =   52
               TabStop         =   0   'False
               ToolTipText     =   "Produzindo disponível."
               Top             =   955
               Width           =   1305
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
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
               Index           =   4
               Left            =   600
               TabIndex        =   132
               Top             =   765
               Width           =   345
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Index           =   3
               Left            =   6345
               TabIndex        =   131
               Top             =   765
               Width           =   915
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde. empenhar"
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
               Left            =   13620
               TabIndex        =   124
               Top             =   765
               Width           =   1365
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Prazo final"
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
               Left            =   12742
               TabIndex        =   123
               Top             =   180
               Width           =   750
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
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
               Left            =   7065
               TabIndex        =   122
               Top             =   180
               Width           =   495
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Necessidade"
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
               Left            =   13830
               TabIndex        =   121
               Top             =   180
               Width           =   1050
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
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
               Left            =   1620
               TabIndex        =   120
               Top             =   180
               Width           =   375
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ped. interno"
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
               Left            =   337
               TabIndex        =   119
               Top             =   180
               Width           =   1020
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Necessidade"
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
               Left            =   11925
               TabIndex        =   118
               Top             =   -240
               Width           =   900
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Prod. disp."
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
               Left            =   12457
               TabIndex        =   117
               Top             =   765
               Width           =   870
            End
         End
         Begin VB.TextBox Txt_ID_emp_prod 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   4590
            TabIndex        =   115
            Text            =   "0"
            Top             =   3180
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox Txt_ID_carteira_prod 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5250
            TabIndex        =   114
            Text            =   "0"
            Top             =   3180
            Visible         =   0   'False
            Width           =   645
         End
         Begin DrawSuite2022.USProgressBar USProgressBar4 
            Height          =   255
            Left            =   -74970
            TabIndex        =   83
            Top             =   3390
            Width           =   15225
            _ExtentX        =   26855
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
         Begin DrawSuite2022.USToolBar USToolBar7 
            Height          =   975
            Left            =   60
            TabIndex        =   106
            Top             =   330
            Width           =   15195
            _ExtentX        =   26802
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
            ButtonToolTipText2=   "Salvar"
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
            ButtonToolTipText3=   "Excluir"
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
            BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            Begin DrawSuite2022.USImageList USImageList7 
               Left            =   3480
               Top             =   240
               _ExtentX        =   900
               _ExtentY        =   767
               Img1            =   "frmVendas_Empenho.frx":19FE3
               Count           =   1
            End
         End
         Begin MSComctlLib.ListView ListaProduzindo 
            Height          =   5745
            Left            =   -74940
            TabIndex        =   39
            Top             =   330
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   10134
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
               Text            =   "Ordem"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "T"
               Text            =   "Cliente"
               Object.Width           =   15760
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Object.Tag             =   "D"
               Text            =   "Pr. final"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Qtde. prod."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Qtde. emp."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "Saldo"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSComctlLib.ListView ListaEmpenhoProduzindo 
            Height          =   4755
            Left            =   60
            TabIndex        =   54
            Top             =   2760
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   8387
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
            NumItems        =   14
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Tag             =   "N"
               Text            =   "ID_carteira"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Object.Tag             =   "D"
               Text            =   "Data"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Responsável"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Ped. int."
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Rev."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Cliente"
               Object.Width           =   3889
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Tag             =   "T"
               Text            =   "Vendedor interno"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Tag             =   "T"
               Text            =   "Vendedor externo"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Object.Tag             =   "T"
               Text            =   "Pr. final"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Object.Tag             =   "N"
               Text            =   "Qtde. vend."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Object.Tag             =   "N"
               Text            =   "Qtde. emp."
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   12
               Object.Tag             =   "N"
               Text            =   "Qtde. entr."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   13
               Object.Tag             =   "N"
               Text            =   "Saldo"
               Object.Width           =   1587
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   1425
            Left            =   -74940
            TabIndex        =   40
            Top             =   6090
            Width           =   13635
            _ExtentX        =   24051
            _ExtentY        =   2514
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   16777215
            ForeColorFixed  =   0
            BackColorSel    =   12632064
            ForeColorSel    =   16777215
            BackColorBkg    =   14737632
            FocusRect       =   0
            Appearance      =   0
         End
      End
      Begin FlexCell.Grid Grid1 
         Height          =   7965
         Left            =   -74945
         TabIndex        =   18
         Top             =   2010
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   14049
         Cols            =   2
         DefaultFontSize =   8.25
         GridColor       =   12632256
         ReadOnly        =   -1  'True
         Rows            =   2
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Index           =   9
      Left            =   3375
      TabIndex        =   126
      Top             =   3750
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Index           =   5
      Left            =   750
      TabIndex        =   125
      Top             =   3750
      Width           =   345
   End
End
Attribute VB_Name = "frmVendas_Empenho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_Vendas_Empenho_Localizar As String 'OK
Dim Novo_Vendas_Empenho As Boolean 'OK
Dim Novo_Vendas_Empenho1 As Boolean 'OK
Dim TBLISTA_Vendas_Empenho As ADODB.Recordset 'OK

'GridEstrutura
Public m_Tree As New Node
Public m_Row As Long
Public m_Col As Long
Dim tempNode As Node
Dim intIndex, i As Integer
Dim CodRef As String, DataValidacao As String, RespValidacao As String

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_ordenar_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
With cmbfamilia
    If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Cliente" Or cmbfiltrarpor = "Grupo do cliente" Then
        txtTexto.Visible = False
        .Visible = True
        .Clear
        .AddItem ""
        If cmbfiltrarpor = "Família" Then
            ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and Vendas = 'True'", True
        ElseIf cmbfiltrarpor = "Cliente" Then
                Set TBClientes = CreateObject("adodb.recordset")
                TBClientes.Open "Select IDCliente, Cliente from vendas_proposta where Cliente <> 'Null' and (Tipo = 'PRPE' or Tipo = 'PE') group by IDCliente, Cliente order by Cliente", Conexao, adOpenKeyset, adLockOptimistic
                If TBClientes.EOF = False Then
                    Do While TBClientes.EOF = False
                        .AddItem TBClientes!Cliente
                        .ItemData(.NewIndex) = TBClientes!IDCliente
                        TBClientes.MoveNext
                    Loop
                End If
                TBClientes.Close
            Else
                Set TBFamilia = CreateObject("adodb.recordset")
                TBFamilia.Open "Select * from Clientes_grupos where Texto <> 'Null' order by Texto", Conexao, adOpenKeyset, adLockOptimistic
                If TBFamilia.EOF = False Then
                    Do While TBFamilia.EOF = False
                        .AddItem TBFamilia!Texto
                        .ItemData(.NewIndex) = TBFamilia!ID
                        TBFamilia.MoveNext
                    Loop
                End If
                TBFamilia.Close
        End If
    Else
        txtTexto.Visible = True
        .Visible = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

CamposFiltro = "P.codProduto, P.Desenho, P.Descricao, P.Unidade, P.Unidade_com, P.classe, P.Producao, P.Compras, P.Vendas"
INNERJOINTEXTO = "Select " & CamposFiltro & " from ((Projproduto P LEFT JOIN item_aplicacoes IA ON IA.codproduto = P.codproduto) LEFT JOIN Projproduto_clientes PC ON PC.codproduto = P.codproduto) LEFT JOIN Projproduto_fabricante PFAB ON PFAB.Codproduto = P.codproduto"
If Cmb_ordenar = "Código interno" Then Ordenar = "P.desenho" Else Ordenar = "P.Descricao"
TextoFiltroPadrao = "(P.tipo = 'P' or P.Tipo = 'S') and P.Vendas = 'True' and P.bloqueado = 'False' group by " & CamposFiltro & " order by " & Ordenar

If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
    If cmbfiltrarpor = "Cliente" Then
        StrSql_Vendas_Empenho_Localizar = INNERJOINTEXTO & " where PC.IDCliente = " & cmbfamilia.ItemData(cmbfamilia.ListIndex) & " and " & TextoFiltroPadrao
    ElseIf cmbfiltrarpor = "Grupo do cliente" Then
            StrSql_Vendas_Empenho_Localizar = INNERJOINTEXTO & " where IA.IDGrupo = " & cmbfamilia.ItemData(cmbfamilia.ListIndex) & " and " & TextoFiltroPadrao
        ElseIf cmbfiltrarpor = "Família" Then
                StrSql_Vendas_Empenho_Localizar = INNERJOINTEXTO & " where P.classe = '" & cmbfamilia & "' and " & TextoFiltroPadrao
            ElseIf cmbfiltrarpor = "Comprimento" Or cmbfiltrarpor = "Largura" Or cmbfiltrarpor = "Espessura" Then
                    Select Case cmbfiltrarpor
                        Case "Comprimento": TextoFiltro = "P.Comprimento"
                        Case "Largura": TextoFiltro = "P.Largura"
                        Case "Espessura": TextoFiltro = "P.Espessura"
                    End Select
                    valor = txtTexto
                    NovoValor = Replace(valor, ",", ".")
                    StrSql_Vendas_Empenho_Localizar = INNERJOINTEXTO & " where " & TextoFiltro & " = " & NovoValor & " and " & TextoFiltroPadrao
                Else
                    Select Case cmbfiltrarpor
                        Case "Código interno": TextoFiltro = "P.desenho"
                        Case "Código de referência": TextoFiltro = "IA.N_referencia"
                        Case "Número do desenho": TextoFiltro = "IA.desenho"
                        Case "Descrição": TextoFiltro = "P.descricao"
                        Case "Descrição comercial": TextoFiltro = "P.Descricaotecnica"
                        Case "Dureza": TextoFiltro = "P.Dureza"
                        Case "Part number": TextoFiltro = "PFAB.Part_number"
                    End Select
                    StrSql_Vendas_Empenho_Localizar = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
    End If
Else
    StrSql_Vendas_Empenho_Localizar = INNERJOINTEXTO & " where " & TextoFiltroPadrao
End If

ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbVersao_pesquisar_estrutura_Click()
On Error GoTo tratar_erro

ProcCarregaEstrutura

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Vendas_Empenho.AbsolutePage <> 2 Then
    If TBLISTA_Vendas_Empenho.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Vendas_Empenho.PageCount - 1)
    Else
        TBLISTA_Vendas_Empenho.AbsolutePage = TBLISTA_Vendas_Empenho.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Vendas_Empenho.AbsolutePage)
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
    TBLISTA_Vendas_Empenho.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Vendas_Empenho.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Vendas_Empenho.AbsolutePage = 1
ProcExibePagina (TBLISTA_Vendas_Empenho.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Vendas_Empenho.AbsolutePage <> -3 Then
    If TBLISTA_Vendas_Empenho.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBLISTA_Vendas_Empenho.AbsolutePage)
    End If
Else
    ProcExibePagina (TBLISTA_Vendas_Empenho.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Vendas_Empenho.AbsolutePage = TBLISTA_Vendas_Empenho.PageCount
ProcExibePagina (TBLISTA_Vendas_Empenho.AbsolutePage)

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
            Case vbKeyF2: ProcFiltrar
            Case vbKeyReturn: ListView1_DblClick
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyF4: ProcExcluirEmpPed
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 3:
        Select Case KeyCode
            Case vbKeyInsert: If SSTab3.Tab = 1 Then ProcNovoEmpEst
            Case vbKeyF3: If SSTab3.Tab = 1 Then ProcSalvarEmpEst
            Case vbKeyF4: If SSTab3.Tab = 1 Then ProcExcluirEmpEst
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 4:
        Select Case KeyCode
            Case vbKeyInsert: If SSTab4.Tab = 1 Then ProcNovoEmpProd
            Case vbKeyF3: If SSTab4.Tab = 1 Then ProcSalvarEmpProd
            Case vbKeyF4: If SSTab4.Tab = 1 Then ProcExcluirEmpProd
            'Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 5, True
ProcCarregaToolBar2 Me, 15195, 3, True
ProcCarregaToolBar3 Me, 15195, 4, True
ProcCarregaToolBar4 Me, 15195, 3, True
ProcCarregaToolBar5 Me, 15195, 4, True
ProcCarregaToolBar6 Me, 15195, 3, True
ProcCarregaToolBar7 Me, 15195, 4, True
ProcCarregaComboEmpresa Cmb_empresa, False
Formulario = "Vendas/Empenho"
Direitos
SSTab1.Tab = 0
SSTab2.Tab = 0
SSTab3.Tab = 0
SSTab4.Tab = 0

ProcFiltroPadrao cmbfiltrarpor, Optmeio, Optfim, optIgual, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), "Produtos/Serviços", "V", True
If Permitido = False Then cmbfiltrarpor = "Código interno"

Cmb_ordenar = "Código interno"
cmbVersao_pesquisar_estrutura = "A"

ProcRemoveObjetosResize Me
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Vendas_Empenho = True Then
    If USMsgBox("O empenho do estoque ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvarEmpEst
        If Novo_Vendas_Empenho = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Vendas_Empenho1 = True Then
    If USMsgBox("O empenho da produção ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvarEmpProd
        If Novo_Vendas_Empenho1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Vendas_Empenho = False
Novo_Vendas_Empenho1 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Vendas/Empenho"
Direitos
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoEstoqueOrdem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListaEmpenhoEstoqueOrdem, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoEstoquePed_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaEmpenhoEstoquePed
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
'                If .ListItems.Item(InitFor).ListSubItems(12) <> 0 Then
'                    .ListItems.Item(InitFor).Checked = False
'                    GoTo Proximo
'                End If

                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaEmpenhoEstoquePed, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoEstoquePed_DblClick()
On Error GoTo tratar_erro

With ListaEmpenhoEstoquePed
    If .ListItems.Count = 0 Then Exit Sub
    ProcVerifQtdeFaturadaProdServ .SelectedItem.ListSubItems(1), ListView1.SelectedItem.ListSubItems(1), False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoEstoquePed_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

'With ListaEmpenhoEstoquePed
'    For InitFor = 1 To .ListItems.Count
'        If .ListItems.Item(InitFor).Checked = True Then
'            If .ListItems.Item(InitFor).ListSubItems(12) <> 0 Then
'                usMsgbox ("Não é permitido excluir este empenho, pois o mesmo já foi baixado no estoque."), vbExclamation, "CAPRIND v5.0"
'                .ListItems.Item(InitFor).Checked = False
'                Exit Sub
'            End If
'        End If
'    Next InitFor
'End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoEstoquePed_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaEmpenhoEstoquePed
    If .ListItems.Count = 0 Then Exit Sub
    ProcLimpaCamposEmpEst
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from Estoque_Controle_Empenho_Vendas where ID = " & .SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Txt_ID_emp_est = TBFI!ID
        Txt_ID_carteira_est = .SelectedItem.ListSubItems(1)
        Txt_data_est = .SelectedItem.ListSubItems(2)
        Txt_responsavel_est = .SelectedItem.ListSubItems(3)
        Txt_pedido_est = .SelectedItem.ListSubItems(4)
        Txt_rev_est = .SelectedItem.ListSubItems(5)
        Txt_cliente_est = .SelectedItem.ListSubItems(6)
        Txt_prazo_est = .SelectedItem.ListSubItems(9)
        ProcVerifNecessVendas TBFI!ID_carteira, .SelectedItem.ListSubItems(10)
        Txt_necessidade_vendas_est = Qtd - (qtdeliberada + qtdeliberar)
        Txt_qtde_emp_est = IIf(IsNull(TBFI!Qtde_empenhada), 0, TBFI!Qtde_empenhada)
    End If
    TBFI.Close
    CodigoLista = .SelectedItem.index
    Frame3.Enabled = True
End With
Novo_Vendas_Empenho = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifNecessVendas(ID_carteira As Long, Qtde_vendida As Double)
On Error GoTo tratar_erro

Qtd = Qtde_vendida
qtdeliberada = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(Qtde_empenhada - Qtde_saida) as qtdeliberada from Estoque_Controle_Empenho_Vendas where ID_carteira = " & ID_carteira & " and Qtde_empenhada - Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    qtdeliberada = IIf(IsNull(TBAbrir!qtdeliberada), 0, TBAbrir!qtdeliberada)
End If
qtdeliberar = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(PP.Qtde_empenho - PP.Qtde_entrada) as qtdeliberar from (Producao_pedidos PP INNER JOIN Vendas_carteira VC ON VC.Codigo = PP.IDCarteira) INNER JOIN Producao P ON P.Ordem = PP.Ordem where PP.IDcarteira = " & ID_carteira & " and P.desenho = '" & ListView1.SelectedItem.ListSubItems(1) & "' and PP.Qtde_empenho - PP.Qtde_entrada > 0", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    qtdeliberar = IIf(IsNull(TBAbrir!qtdeliberar), 0, TBAbrir!qtdeliberar)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoProduzindo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaEmpenhoProduzindo
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
'                If .ListItems.Item(InitFor).ListSubItems(12) <> 0 Then
'                    .ListItems.Item(InitFor).Checked = False
'                    GoTo Proximo
'                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaEmpenhoProduzindo, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoProduzindo_DblClick()
On Error GoTo tratar_erro

With ListaEmpenhoProduzindo
    If .ListItems.Count = 0 Then Exit Sub
    ProcVerifQtdeFaturadaProdServ .SelectedItem.ListSubItems(1), ListView1.SelectedItem.ListSubItems(1), False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoProduzindo_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

'With ListaEmpenhoProduzindo
'    For InitFor = 1 To .ListItems.Count
'        If .ListItems.Item(InitFor).Checked = True Then
'            If .ListItems.Item(InitFor).ListSubItems(12) <> 0 Then
'                usMsgbox ("Não é permitido excluir este empenho, pois já foi dado entrada no estoque."), vbExclamation, "CAPRIND v5.0"
'                .ListItems.Item(InitFor).Checked = False
'                Exit Sub
'            End If
'        End If
'    Next InitFor
'End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoProduzindo_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaEmpenhoProduzindo
    If .ListItems.Count = 0 Then Exit Sub
    ProcLimpaCamposEmpProd
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from Producao_pedidos where ID = " & .SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Txt_ID_emp_prod = TBFI!ID
        Txt_ID_carteira_prod = .SelectedItem.ListSubItems(1)
        Txt_data_prod = .SelectedItem.ListSubItems(2)
        Txt_responsavel_prod = .SelectedItem.ListSubItems(3)
        Txt_pedido_prod = .SelectedItem.ListSubItems(4)
        Txt_rev_prod = .SelectedItem.ListSubItems(5)
        Txt_cliente_prod = .SelectedItem.ListSubItems(6)
        Txt_prazo_prod = .SelectedItem.ListSubItems(9)
        ProcVerifNecessVendas TBFI!IDcarteira, .SelectedItem.ListSubItems(10)
        Txt_necessidade_vendas_prod = Qtd - (qtdeliberada + qtdeliberar)
        Txt_qtde_emp_prod = IIf(IsNull(TBFI!Qtde_empenho), 0, TBFI!Qtde_empenho)
    End If
    TBFI.Close
    CodigoLista1 = .SelectedItem.index
    Frame8.Enabled = True
End With
Novo_Vendas_Empenho1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoVendidosOP_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaEmpenhoVendidosOP
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If .ListItems.Item(InitFor).ListSubItems(5) <> 0 Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaEmpenhoVendidosOP, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoVendidosOP_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaEmpenhoVendidosOP
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If .ListItems.Item(InitFor).ListSubItems(5) <> 0 Then
                USMsgBox ("Não é permitido excluir este empenho, pois já foi dado entrada no estoque."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoVendidosRE_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With ListaEmpenhoVendidosRE
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If .ListItems.Item(InitFor).ListSubItems(8) <> 0 Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If

                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView ListaEmpenhoVendidosRE, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEmpenhoVendidosRE_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaEmpenhoVendidosRE
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If .ListItems.Item(InitFor).ListSubItems(8) <> 0 Then
                USMsgBox ("Não é permitido excluir este empenho, pois o mesmo já foi baixado no estoque."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEstoque_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListaEstoque, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaEstoque_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

ProcLimpaCamposEmpEst
Frame3.Enabled = False
Novo_Vendas_Empenho = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaProduzindo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListaProduzindo, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaProduzindo_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

ProcLimpaCamposEmpProd
Frame8.Enabled = False
Novo_Vendas_Empenho1 = False
FamiliaAntiga = ""
If ListaProduzindo.ListItems.Count <> 0 Then ProcCarregaGridSitProd MSFlexGrid1, "Select P.*, OS.Idproducao, OS.Maquina, CM.Setor FROM ((Producao P INNER JOIN Projproduto PP ON P.Desenho = PP.Desenho) INNER JOIN Ordemservico OS ON OS.Ordem = P.Ordem) INNER JOIN CadMaquinas CM ON CM.Maquina = OS.Maquina where P.ordem = " & ListaProduzindo.SelectedItem & " order by P.ordem, OS.IDproducao", PBLista4, False, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaVendidos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListaVendidos, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaVendidos_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

ProcCarregaListaREEmpPed
ProcCarregaListaOPEmpPed

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListView1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_DblClick()
On Error GoTo tratar_erro

If ListView1.ListItems.Count = 0 Then Exit Sub
'Verifica se o produto/serviço pertence ao cliente
If PI_Produtos = True Or PI_Servicos = True Then
    With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
        IDempresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
        IDCliente = .txtIDcliente
        Cliente = .txtCliente
    End With
ElseIf Vendas_Programacao = True Then
        IDempresa = frmVendas_programacao.Cmb_empresa.ItemData(frmVendas_programacao.Cmb_empresa.ListIndex)
        IDCliente = frmVendas_programacao.txtID_cli
        Cliente = frmVendas_programacao.txtCliente
End If
If IDCliente <> 0 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Projproduto_clientes where Codproduto = " & ListView1.SelectedItem & " and IDCliente <> 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from Projproduto_clientes where Codproduto = " & ListView1.SelectedItem & " and IDCliente = " & IDCliente, Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then
            If PI_Servicos = True Or Proposta_Servicos = True Then NomeCampo = "serviço" Else NomeCampo = "Produto"
            Set TBCiclo = CreateObject("adodb.recordset")
            TBCiclo.Open "Select * from empresa where codigo = " & IDempresa & " and Bloquear_produtos = 'False'", Conexao, adOpenKeyset, adLockOptimistic
            If TBCiclo.EOF = False Then
                If USMsgBox("Este " & NomeCampo & " não pertence ao cliente " & Cliente & ", deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                    TBCiclo.Close
                    Exit Sub
                End If
            Else
                USMsgBox ("Este " & NomeCampo & " não pertence ao cliente " & Cliente & "."), vbExclamation, "CAPRIND v5.0"
                TBCiclo.Close
                Exit Sub
            End If
            TBCiclo.Close
            
        End If
        TBProduto.Close
    End If
    TBAbrir.Close
End If

If PI_Produtos = True Then
    With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
        .ProcLimparProdutos False
        .txtNomenclatura = ListView1.SelectedItem.ListSubItems.Item(1).Text
        .ProcPuxaDadosProduto
    End With
ElseIf PI_Servicos = True Then
        With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
            .ProcLimparServicos False
            .txtcodservico = ListView1.SelectedItem.ListSubItems.Item(1).Text
            .ProcPuxadadosServico
        End With
    ElseIf Vendas_Programacao = True Then
        With frmVendas_programacao
            .ProcLimpaCampos_Item
            .txtCodigo = ListView1.SelectedItem.ListSubItems.Item(1)
            .txtdescricao = ListView1.SelectedItem.ListSubItems.Item(2)
        End With
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de reg.: 0"
lblPaginas.Caption = "Página: 0 de: 0"
ListView1.ListItems.Clear
If StrSql_Vendas_Empenho_Localizar = "" And StrSqlLocServVendas = "" Then Exit Sub
Set TBLISTA_Vendas_Empenho = CreateObject("adodb.recordset")
TBLISTA_Vendas_Empenho.Open StrSql_Vendas_Empenho_Localizar, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Vendas_Empenho.EOF = False Then ProcExibePagina (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

IDConta = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

ListView1.ListItems.Clear
TBLISTA_Vendas_Empenho.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Vendas_Empenho.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Vendas_Empenho.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Vendas_Empenho.RecordCount - IIf(Pagina > 1, (TBLISTA_Vendas_Empenho.PageSize * (Pagina - 1)), 0), TBLISTA_Vendas_Empenho.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Vendas_Empenho.EOF = False And (ContadorReg <= TamanhoPagina)
    With ListView1.ListItems
        .Add , , TBLISTA_Vendas_Empenho!Codproduto
        .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA_Vendas_Empenho!Desenho), "", TBLISTA_Vendas_Empenho!Desenho)
        .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA_Vendas_Empenho!Descricao), "", TBLISTA_Vendas_Empenho!Descricao)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA_Vendas_Empenho!Unidade), "", TBLISTA_Vendas_Empenho!Unidade)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Vendas_Empenho!Unidade_com), "", TBLISTA_Vendas_Empenho!Unidade_com)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Vendas_Empenho!Classe), "", TBLISTA_Vendas_Empenho!Classe)
        
        valor = FunVerificaNecessidadeVenda(TBLISTA_Vendas_Empenho!Desenho, IDConta)
        .Item(.Count).SubItems(6) = Format(valor, "###,##0.0000")
        
        Valor1 = FunVerificaQtdeEstoque(TBLISTA_Vendas_Empenho!Desenho, IDConta, "")
        .Item(.Count).SubItems(7) = Format(Valor1, "###,##0.0000")
                
        Valor2 = FunVerificaQtdeEmpenhoEstVenda(TBLISTA_Vendas_Empenho!Desenho, IDConta)
        TTE = FunVerificaQtdeEmpenhoEst(TBLISTA_Vendas_Empenho!Desenho, IDConta)
        .Item(.Count).SubItems(8) = Format(Valor2 + TTE, "###,##0.0000")
        Valor_Cofins_Prod = Valor1 - (Valor2 + TTE)
        .Item(.Count).SubItems(9) = Format(Valor_Cofins_Prod, "###,##0.0000")
        
        Valor3 = FunVerificaQtdeProduzindo(TBLISTA_Vendas_Empenho!Desenho, IDConta)
        .Item(.Count).SubItems(10) = Format(Valor3, "###,##0.0000")
        Valor_Cofins_Serv = FunVerificaQtdeEmpenhoProduzindo(TBLISTA_Vendas_Empenho!Desenho, IDConta)
        .Item(.Count).SubItems(11) = Format(Valor_Cofins_Serv, "###,##0.0000")
        Valor_CSLL_Prod = Valor3 - Valor_Cofins_Serv
        .Item(.Count).SubItems(12) = Format(Valor_CSLL_Prod, "###,##0.0000")
    End With
    TBLISTA_Vendas_Empenho.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de reg.: " & TBLISTA_Vendas_Empenho.RecordCount
If TBLISTA_Vendas_Empenho.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Pág.: 1 de: " & TBLISTA_Vendas_Empenho.PageCount
ElseIf TBLISTA_Vendas_Empenho.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Pág.: " & TBLISTA_Vendas_Empenho.PageCount & " de: " & TBLISTA_Vendas_Empenho.PageCount
    Else
        lblPaginas.Caption = "Pág.: " & TBLISTA_Vendas_Empenho.AbsolutePage - 1 & " de: " & TBLISTA_Vendas_Empenho.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

SSTab3.Tab = 0
SSTab4.Tab = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optfim_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optIgual_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optinicio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Optmeio_Click()
On Error GoTo tratar_erro

ListView1.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If ListView1.ListItems.Count = 0 Then
    SSTab1.Tab = 0
    Exit Sub
End If
If SSTab1.Tab = 0 Then Caption = "Administrativo - Vendas - Empenho" Else Caption = "Administrativo - Vendas - Empenho (Cód. interno : " & ListView1.SelectedItem.ListSubItems(1) & " - Descrição : " & ListView1.SelectedItem.ListSubItems(2) & ")"
Select Case SSTab1.Tab
    Case 0:
        If ListView1.Visible = True Then ListView1.SetFocus
        If ListView1.ListItems.Count <> 0 Then ProcCarregaLista
    Case 1:
        ProcCarregaEstrutura
    Case 2:
        ListaVendidos.SetFocus
        ProcCarregaListaVendidos
    Case 3:
        ListaEstoque.SetFocus
        SSTab3.Tab = 0
        ProcCarregaListaEstoque
        ProcLimpaCamposEmpEst
        Frame3.Enabled = False
        Novo_Vendas_Empenho = False
    Case 4:
        ListaProduzindo.SetFocus
        SSTab4.Tab = 0
        ProcCarregaListaProduzindo
        With MSFlexGrid1
            .rows = 0
            .Cols = 0
            .Refresh
        End With
        ProcLimpaCamposEmpProd
        Frame8.Enabled = False
        Novo_Vendas_Empenho1 = False
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaVendidos()
On Error GoTo tratar_erro

Quant = 0
valor = 0
Valor1 = 0
ListaVendidos.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select VP.Ncotacao, VP.Revisao, VP.Cliente, VP.vend_int, VP.vend_ext, VC.Codigo, VC.Quantidade - VC.Qtdefaturada as Qtd, VC.Prazofinal, VC.Unidade, VC.Unidade_com from vendas_carteira VC INNER JOIN vendas_proposta VP on VC.cotacao = VP.cotacao where VC.Desenho = '" & ListView1.SelectedItem.ListSubItems(1) & "' and VP.DtValidacaoPI IS NOT NULL and (VC.Liberacao = 'VENDIDA' or VC.Liberacao = 'FATURAR' or VC.Liberacao = 'FATURAR PARCIAL' or VC.Liberacao = 'FATURADO PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista1.Min = 0
    PBLista1.Max = TBLISTA.RecordCount
    PBLista1.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With ListaVendidos.ListItems
            .Add , , TBLISTA!CODIGO
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Ncotacao), "", TBLISTA!Ncotacao)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Revisao), "", TBLISTA!Revisao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!vend_int), "", TBLISTA!vend_int)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Vend_ext), "", TBLISTA!Vend_ext)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!PrazoFinal), "", Format(TBLISTA!PrazoFinal, "dd/mm/yy"))
            Qtd = IIf(IsNull(TBLISTA!Qtd), 0, TBLISTA!Qtd) / FunVerificaTabelaConversaoUnidade(TBLISTA!Unidade, TBLISTA!Unidade_com)
            .Item(.Count).SubItems(7) = Format(IIf(Qtd < 0, 0, Qtd), "###,##0.0000")
            ProcVerifNecessVendas TBLISTA!CODIGO, Qtd
            .Item(.Count).SubItems(8) = Format(qtdeliberada, "###,##0.0000") 'Empenho estoque
            .Item(.Count).SubItems(9) = Format(qtdeliberar, "###,##0.0000") 'Empenho produzindo
            .Item(.Count).SubItems(10) = Format(Qtd - (qtdeliberada + qtdeliberar), "###,##0.0000")
        End With
        Quant = Quant + Qtd
        valor = valor + qtdeliberada
        Valor1 = Valor1 + qtdeliberar
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista1.Value = Contador
    Loop
End If
TBLISTA.Close
Txt_qtde_total_vendidos = Format(Quant, "###,##0.0000")
Txt_qtde_total_emp_est = Format(valor, "###,##0.0000")
Txt_qtde_total_emp_prod = Format(Valor1, "###,##0.0000")
Txt_necessidade_vendidos = Format(Quant - (valor + Valor1), "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaREEmpPed()
On Error GoTo tratar_erro

ListaEmpenhoVendidosRE.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select ECEV.*, EC.* from Estoque_Controle_Empenho_Vendas ECEV INNER JOIN Estoque_Controle EC ON ECEV.ID_Estoque = EC.IDEstoque where ECEV.ID_Carteira = '" & ListaVendidos.SelectedItem & "' and Qtde_empenhada - Qtde_saida > 0", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista2.Min = 0
    PBLista2.Max = TBLISTA.RecordCount
    PBLista2.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With ListaEmpenhoVendidosRE.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!IDEstoque), "", TBLISTA!IDEstoque)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!LOTE), "", TBLISTA!LOTE)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!local_armaz), "", TBLISTA!local_armaz)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Corrida), "", TBLISTA!Corrida)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Certificado), "", TBLISTA!Certificado)
            valor = IIf(IsNull(TBLISTA!Qtde_empenhada), 0, TBLISTA!Qtde_empenhada)
            .Item(.Count).SubItems(7) = Format(valor, "###,##0.0000")
            Valor1 = IIf(IsNull(TBLISTA!Qtde_saida), 0, TBLISTA!Qtde_saida)
            .Item(.Count).SubItems(8) = Format(Valor1, "###,##0.0000")
            .Item(.Count).SubItems(9) = Format(valor - Valor1, "###,##0.0000")
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista2.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaOPEmpPed()
On Error GoTo tratar_erro

ListaEmpenhoVendidosOP.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select PP.*, P.* from (Producao_pedidos PP INNER JOIN Vendas_carteira VC ON VC.Codigo = PP.IDCarteira) INNER JOIN Producao P ON P.Ordem = PP.Ordem and P.Desenho = VC.Desenho where PP.IDCarteira = " & ListaVendidos.SelectedItem & " and VC.Desenho = '" & ListView1.SelectedItem.ListSubItems(1) & "' and PP.Qtde_empenho - PP.Qtde_entrada > 0", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista2.Min = 0
    PBLista2.Max = TBLISTA.RecordCount
    PBLista2.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With ListaEmpenhoVendidosOP.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Ordem), "", TBLISTA!Ordem)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!PrazoEntrega), "", Format(TBLISTA!PrazoEntrega, "dd/mm/yy"))
            valor = IIf(IsNull(TBLISTA!Qtde_empenho), 0, TBLISTA!Qtde_empenho)
            .Item(.Count).SubItems(4) = valor
            Valor1 = IIf(IsNull(TBLISTA!Qtde_entrada), 0, TBLISTA!Qtde_entrada)
            .Item(.Count).SubItems(5) = Valor1
            .Item(.Count).SubItems(6) = valor - Valor1
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista2.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaEmpprod()
On Error GoTo tratar_erro

Valor3 = 0
ListaEmpenhoProduzindo.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select PP.ID, PP.IDcarteira, PP.Data as Dataemp, PP.Responsavel as Respemp, PP.Qtde_Empenho, PP.Qtde_entrada, VP.*, VC.Codigo, VC.Qtde_produzir - VC.Qtdefaturada as Qtd, VC.Prazofinal from (Producao_Pedidos PP INNER JOIN vendas_carteira VC ON PP.IDCarteira =  VC.Codigo) INNER JOIN vendas_proposta VP ON VP.cotacao = VC.Cotacao where PP.Ordem = " & ListaProduzindo.SelectedItem & " and VC.Desenho = '" & ListView1.SelectedItem.ListSubItems(1) & "' and PP.Qtde_empenho - PP.Qtde_entrada > 0", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista4.Min = 0
    PBLista4.Max = TBLISTA.RecordCount
    PBLista4.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With ListaEmpenhoProduzindo.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = TBLISTA!IDcarteira
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Dataemp), "", Format(TBLISTA!Dataemp, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Respemp), "", TBLISTA!Respemp)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Ncotacao), "", TBLISTA!Ncotacao)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Revisao), "", TBLISTA!Revisao)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!vend_int), "", TBLISTA!vend_int)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!Vend_ext), "", TBLISTA!Vend_ext)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!PrazoFinal), "", Format(TBLISTA!PrazoFinal, "dd/mm/yy"))
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!Qtd), "", TBLISTA!Qtd)
            valor = IIf(IsNull(TBLISTA!Qtde_empenho), 0, TBLISTA!Qtde_empenho)
            .Item(.Count).SubItems(11) = valor
            Valor1 = IIf(IsNull(TBLISTA!Qtde_entrada), 0, TBLISTA!Qtde_entrada)
            .Item(.Count).SubItems(12) = Valor1
            .Item(.Count).SubItems(13) = valor - Valor1
        End With
         Valor3 = Valor3 + (valor - Valor1)
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista4.Value = Contador
    Loop
End If
TBLISTA.Close
valor = ListaProduzindo.SelectedItem.ListSubItems(3)
Txt_qtde_total_produzindo = valor
Txt_qtde_total_emp_produzindo = Valor3
Txt_qtde_total_disp_produzindo = valor - Valor3

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaEmpEstPed()
On Error GoTo tratar_erro

Valor3 = 0
ListaEmpenhoEstoquePed.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
'TBLISTA.Open "Select ECEV.ID, ECEV.ID_carteira, ECEV.Data1 as Dataemp, ECEV.Responsavel as Respemp, ECEV.Qtde_empenhada, ECEV.Qtde_saida, VP.*, VC.Codigo, VC.Qtde_produzir - VC.Qtdefaturada as Qtd, VC.Prazofinal from (Estoque_Controle_Empenho_Vendas ECEV INNER JOIN vendas_carteira VC ON ECEV.ID_Carteira =  VC.Codigo) INNER JOIN vendas_proposta VP ON VP.cotacao = VC.Cotacao where ECEV.ID_Estoque = " & ListaEstoque.SelectedItem & " and ECEV.Qtde_empenhada - ECEV.Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
TBLISTA.Open "Select ECEV.ID, ECEV.ID_carteira, ECEV.Data as Dataemp, ECEV.Responsavel as Respemp, ECEV.Qtde_empenhada, ECEV.Qtde_saida, VP.*, VC.Codigo, VC.Qtde_produzir as Qtd, VC.Prazofinal from (Estoque_Controle_Empenho_Vendas ECEV INNER JOIN vendas_carteira VC ON ECEV.ID_Carteira =  VC.Codigo) INNER JOIN vendas_proposta VP ON VP.cotacao = VC.Cotacao where ECEV.ID_Estoque = " & ListaEstoque.SelectedItem & " and ECEV.Qtde_empenhada - ECEV.Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista3.Min = 0
    PBLista3.Max = TBLISTA.RecordCount
    PBLista3.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With ListaEmpenhoEstoquePed.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = TBLISTA!ID_carteira
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Dataemp), "", Format(TBLISTA!Dataemp, "dd/mm/yy"))
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Respemp), "", TBLISTA!Respemp)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Ncotacao), "", TBLISTA!Ncotacao)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Revisao), "", TBLISTA!Revisao)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!vend_int), "", TBLISTA!vend_int)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!Vend_ext), "", TBLISTA!Vend_ext)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!PrazoFinal), "", Format(TBLISTA!PrazoFinal, "dd/mm/yy"))
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!Qtd), "", TBLISTA!Qtd)
            valor = IIf(IsNull(TBLISTA!Qtde_empenhada), 0, TBLISTA!Qtde_empenhada)
            .Item(.Count).SubItems(11) = valor
            Valor1 = IIf(IsNull(TBLISTA!Qtde_saida), 0, TBLISTA!Qtde_saida)
            .Item(.Count).SubItems(12) = Valor1
            .Item(.Count).SubItems(13) = valor - Valor1
        End With
        Valor3 = Valor3 + (valor - Valor1)
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista3.Value = Contador
    Loop
End If
TBLISTA.Close
valor = ListaEstoque.SelectedItem.ListSubItems(6)
Txt_qtde_total_estoque = valor
QuantEmpenho = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select SUM(Quantidade - Qtde_saida) as QuantEmpenho from Producao_NF_Consignada where IDEstoque = " & ListaEstoque.SelectedItem & " and Quantidade - Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    QuantEmpenho = IIf(IsNull(TBAbrir!QuantEmpenho), 0, TBAbrir!QuantEmpenho)
End If
Txt_qtde_total_emp_estoque = Valor3 + QuantEmpenho
Txt_qtde_total_disp_estoque = valor - (Valor3 + QuantEmpenho)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaEmpEstOrdem()
On Error GoTo tratar_erro

Valor3 = 0
ListaEmpenhoEstoqueOrdem.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select PNFC.ID, PNFC.Data as Dataemp, PNFC.Responsavel as Respemp, PNFC.Ordem, PNFC.Quantidade, PNFC.Qtde_saida, P.* from Producao_NF_Consignada PNFC INNER JOIN Producao P ON PNFC.Ordem = P.Ordem where PNFC.IDEstoque = " & ListaEstoque.SelectedItem & " and PNFC.Quantidade - PNFC.Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista3.Min = 0
    PBLista3.Max = TBLISTA.RecordCount
    PBLista3.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With ListaEmpenhoEstoqueOrdem.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Dataemp), "", Format(TBLISTA!Dataemp, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Respemp), "", TBLISTA!Respemp)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Ordem), "", TBLISTA!Ordem)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!PrazoEntrega), "", TBLISTA!PrazoEntrega)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Quant), "", TBLISTA!Quant)
            valor = IIf(IsNull(TBLISTA!quantidade), 0, TBLISTA!quantidade)
            .Item(.Count).SubItems(7) = valor
            Valor1 = IIf(IsNull(TBLISTA!Qtde_saida), 0, TBLISTA!Qtde_saida)
            .Item(.Count).SubItems(8) = Valor1
            .Item(.Count).SubItems(9) = valor - Valor1
        End With
        Valor3 = Valor3 + (valor - Valor1)
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista3.Value = Contador
    Loop
End If
TBLISTA.Close
valor = ListaEstoque.SelectedItem.ListSubItems(6)
Txt_qtde_total_estoque = valor
QuantEmpenho = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select SUM(Qtde_empenhada - Qtde_saida) as QuantEmpenho from Estoque_Controle_Empenho_Vendas where ID_Estoque = " & ListaEstoque.SelectedItem & " and Qtde_empenhada - Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    QuantEmpenho = IIf(IsNull(TBAbrir!QuantEmpenho), 0, TBAbrir!QuantEmpenho)
End If
Txt_qtde_total_emp_estoque = Valor3 + QuantEmpenho
Txt_qtde_total_disp_estoque = valor - (Valor3 + QuantEmpenho)
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaEstoque()
On Error GoTo tratar_erro

valor = 0
Valor1 = 0
ListaEstoque.ListItems.Clear
CamposFiltro = "EC.IDestoque, EC.Data, EC.Lote, EC.Local_armaz, EC.Corrida, EC.Certificado, EC.Estoque_real, EC.Estoque_disponivel"
'StrSql = "Select DISTINCT COUNT(EC.IDestoque) OVER () AS TotalRecords, " & CamposFiltro & " FROM ((Estoque_Controle EC LEFT JOIN Estoque_Controle_Empenho_Vendas ECEV ON ECEV.ID_estoque = EC.IDestoque) LEFT JOIN vendas_carteira VC ON VC.Codigo = ECEV.ID_carteira) LEFT JOIN Estoque_Localarmazenamento_criar LA ON EC.local_armaz = LA.Descricao WHERE EC.desenho = '" & ListView1.SelectedItem.ListSubItems(1) & "' and (EC.estoque_real > 0 or ECEV.ID IS NOT NULL and VC.Quantidade > VC.QtdeFaturada) AND (LA.Estoque <> 'True') group by " & CamposFiltro & ""
StrSql = "Select DISTINCT COUNT(EC.IDestoque) OVER () AS TotalRecords, " & CamposFiltro & " FROM ((Estoque_produtos EC LEFT JOIN Estoque_Controle_Empenho_Vendas ECEV ON ECEV.ID_estoque = EC.IDestoque) LEFT JOIN vendas_carteira VC ON VC.Codigo = ECEV.ID_carteira) LEFT JOIN Estoque_Localarmazenamento_criar LA ON EC.local_armaz = LA.Descricao WHERE EC.desenho = '" & ListView1.SelectedItem.ListSubItems(1) & "' and (EC.estoque_disponivel > 0 or ECEV.ID IS NOT NULL and VC.Quantidade > VC.QtdeFaturada) AND (LA.Estoque <> 'True') group by " & CamposFiltro & ""

'Debug.print StrSql

Set TBLISTA = CreateObject("adodb.recordset")
'TBLISTA.Open "Select DISTINCT COUNT(EC.IDestoque) OVER () AS TotalRecords, " & CamposFiltro & " FROM ((Estoque_Controle EC LEFT JOIN Estoque_Controle_Empenho_Vendas ECEV ON ECEV.ID_estoque = EC.IDestoque) LEFT JOIN vendas_carteira VC ON VC.Codigo = ECEV.ID_carteira) LEFT JOIN Estoque_Localarmazenamento_criar LA ON EC.local_armaz = LA.Descricao WHERE EC.desenho = '" & ListView1.SelectedItem.ListSubItems(1) & "' and (EC.estoque_real > 0 or ECEV.ID IS NOT NULL and VC.Quantidade > VC.QtdeFaturada) AND (LA.Estoque <> 'True') group by " & CamposFiltro & "", Conexao, adOpenKeyset, adLockOptimistic
TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

If TBLISTA.EOF = False Then
    PBLista3.Min = 0
    PBLista3.Max = TBLISTA!TotalRecords
    PBLista3.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With ListaEstoque.ListItems
            .Add , , TBLISTA!IDEstoque
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!LOTE), "", TBLISTA!LOTE)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!local_armaz), "", TBLISTA!local_armaz)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Corrida), "", TBLISTA!Corrida)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Certificado), "", TBLISTA!Certificado)
            
            Valor2 = IIf(IsNull(TBLISTA!Estoque_Disponivel), 0, TBLISTA!Estoque_Disponivel)
            .Item(.Count).SubItems(6) = Format(Valor2, "###,##0.0000") 'Valor2
            
            ProcVerifQtdeEmpEst TBLISTA!IDEstoque
            .Item(.Count).SubItems(7) = Format(Valor3, "###,##0.0000") 'Valor3
            
            .Item(.Count).SubItems(8) = Format(Valor2 - Valor3, "###,##0.0000") 'Valor2 - Valor3
        End With
        valor = valor + Valor2
        Valor1 = Valor1 + Valor3
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista3.Value = Contador
    Loop
End If
TBLISTA.Close
Txt_qtde_total_estoque = Format(valor, "###,##0.0000")
Txt_qtde_total_emp_estoque = Format(Valor1, "###,##0.0000")
Txt_qtde_total_disp_estoque = Format(valor - Valor1, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifQtdeEmpEst(ID_estoque As Long)
On Error GoTo tratar_erro

Valor3 = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "SELECT SUM(Qtde_Empenhada - Qtde_saida) as Valor3 from Estoque_Controle_Empenho_Vendas where ID_Estoque = " & ID_estoque & " and Qtde_Empenhada - Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Valor3 = IIf(IsNull(TBAbrir!Valor3), 0, TBAbrir!Valor3)
End If
QuantEmpenho = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "SELECT SUM(Quantidade - Qtde_saida) as QuantEmpenho from Producao_NF_Consignada where IDEstoque = " & ID_estoque & " and Quantidade - Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    QuantEmpenho = IIf(IsNull(TBAbrir!QuantEmpenho), 0, TBAbrir!QuantEmpenho)
End If
Valor3 = Valor3 + QuantEmpenho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaProduzindo()
On Error GoTo tratar_erro

valor = 0
Valor1 = 0
ListaProduzindo.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select PP.Qtde_produzindo, P.Ordem, P.Cliente, P.PrazoEntrega from Qtde_produzindo_produto PP INNER JOIN Producao P ON P.Ordem = PP.Ordem where P.Desenho = '" & ListView1.SelectedItem.ListSubItems(1) & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista4.Min = 0
    PBLista4.Max = TBLISTA.RecordCount
    PBLista4.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With ListaProduzindo.ListItems
            .Add , , IIf(IsNull(TBLISTA!Ordem), "", TBLISTA!Ordem)
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!PrazoEntrega), "", Format(TBLISTA!PrazoEntrega, "dd/mm/yy"))
            
            Valor2 = IIf(IsNull(TBLISTA!Qtde_produzindo), 0, TBLISTA!Qtde_produzindo)
            .Item(.Count).SubItems(3) = Valor2
            
            ProcVerifQtdeEmpProd TBLISTA!Ordem
            .Item(.Count).SubItems(4) = Valor3
            
            .Item(.Count).SubItems(5) = Valor2 - Valor3
        End With
        valor = valor + Valor2
        Valor1 = Valor1 + Valor3
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista4.Value = Contador
    Loop
End If
TBLISTA.Close
Txt_qtde_total_produzindo = valor
Txt_qtde_total_emp_produzindo = Valor1
Txt_qtde_total_disp_produzindo = valor - Valor1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifQtdeEmpProd(Ordem As Long)
On Error GoTo tratar_erro

Valor3 = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "SELECT SUM(PP.Qtde_empenho - ISNULL(PP.Qtde_entrada, 0)) as Valor3 from Producao_pedidos PP INNER JOIN Vendas_carteira VC ON VC.Codigo = PP.IDCarteira where PP.ordem = " & Ordem & " and VC.Desenho = '" & ListView1.SelectedItem.ListSubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Valor3 = IIf(IsNull(TBAbrir!Valor3), 0, TBAbrir!Valor3)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaEstrutura()
On Error GoTo tratar_erro
''ReDim arrNodes(2000)

If ListView1.ListItems.Count = 0 Then Exit Sub
Call m_Tree.Nodes.Clear
Grid1.rows = 1

m_Row = 1
m_Col = 1

Contador1 = -1
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Projproduto where Desenho = '" & ListView1.SelectedItem.ListSubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    CodRef = ""
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select n_referencia from item_aplicacoes where codproduto = " & TBLISTA!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        CodRef = TBFI!N_referencia
    End If
    TBFI.Close
    
    DataValidacao = ""
    RespValidacao = ""
    If TBLISTA!SubTipoItem <> 0 Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from Projconjunto_desc_versao where codproduto = " & TBLISTA!Codproduto & " and Versao = '" & cmbVersao_pesquisar_estrutura & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            DataValidacao = IIf(IsNull(TBFI!DtValidacao), "", TBFI!DtValidacao)
            RespValidacao = IIf(IsNull(TBFI!RespValidacao), "", TBFI!RespValidacao)
        End If
    End If
    Contador1 = Contador1 + 1
    arrNodes(Contador1).Level = 0
    arrNodes(Contador1).Text = TBLISTA!Desenho & vbTab & "" & vbTab & TBLISTA!Codproduto & vbTab & CodRef & vbTab & TBLISTA!Descricao & vbTab & TBLISTA!Unidade & vbTab & cmbVersao_pesquisar_estrutura & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(valor, "###,##0.00000000") & vbTab & DataValidacao & vbTab & RespValidacao
    
    Codproduto = TBLISTA!Codproduto
    
    ProcNivel2Estrutura frmVendas_ListaProduto, cmbVersao_pesquisar_estrutura, False, False, True, False

    With Grid1
        .AutoRedraw = False
        .AllowUserPaste = cellTextOnly
        .ExtendLastCol = True
        .DrawMode = cellOwnerDraw
        .Cols = 20
        .rows = m_Row
        .Cell(0, 1).Text = "Cód. interno"
        .Cell(0, 2).Text = "Pos."
        .Cell(0, 3).Text = "ID"
        .Cell(0, 4).Text = "Cód. de ref."
        .Cell(0, 5).Text = "Descrição"
        .Cell(0, 6).Text = "Un."
        .Cell(0, 7).Text = "Ver."
        .Cell(0, 8).Text = "Vlr./un"
        .Cell(0, 9).Text = "Un/vlr."
        .Cell(0, 10).Text = "Dim/mm"
        .Cell(0, 11).Text = "Vlr./pç"
        .Cell(0, 12).Text = "Qtde."
        .Cell(0, 13).Text = "Total"
        .Cell(0, 14).Text = "Vlr. custo"
        .Cell(0, 15).Text = "Dt. validação"
        .Cell(0, 16).Text = "Resp. validação"
        .Cell(0, 17).Text = "ID estr."
        .Cell(0, 18).Text = "Part number"
        .Cell(0, 19).Text = "Observações"
        .Range(0, 1, 0, 19).Alignment = cellCenterCenter
        .Column(1).Width = 200
        .Column(2).Width = 30
        .Column(3).Width = 0
        .Column(4).Width = 80
        .Column(5).Width = 300
        .Column(6).Width = 40
        .Column(6).Alignment = cellCenterCenter
        .Column(7).Width = 40
        .Column(7).Alignment = cellCenterCenter
        .Column(8).Width = 100
        .Column(8).Alignment = cellRightCenter
        .Column(9).Width = 40
        .Column(10).Width = 100
        .Column(10).Alignment = cellRightCenter
        .Column(11).Width = 100
        .Column(11).Alignment = cellRightCenter
        .Column(12).Width = 100
        .Column(12).Alignment = cellRightCenter
        .Column(13).Width = 100
        .Column(13).Alignment = cellRightCenter
        .Column(14).Width = 100
        .Column(14).Alignment = cellRightCenter
        .Column(15).Width = 120
        .Column(16).Width = 100
        .Column(17).Width = 0
        .Column(18).Width = 150
        .Column(19).Width = 400
        
        'First node
        Set tempNode = m_Tree.Nodes.Add("")
        .AddItem arrNodes(0).Text
        
        'Other nodes
        For intIndex = 1 To Contador1 'UBound(arrNodes)
            If arrNodes(intIndex).Level = arrNodes(intIndex - 1).Level Then
                Set tempNode = tempNode.Parent.Nodes.Add("")
            ElseIf arrNodes(intIndex).Level > arrNodes(intIndex - 1).Level Then
                Set tempNode = tempNode.Nodes.Add("")
            ElseIf arrNodes(intIndex).Level < arrNodes(intIndex - 1).Level Then
                For i = arrNodes(intIndex).Level To arrNodes(intIndex - 1).Level
                    Set tempNode = tempNode.Parent
                Next
                Set tempNode = tempNode.Nodes.Add("")
            End If
            .AddItem arrNodes(intIndex).Text
        Next
        
        .AutoRedraw = True
        .Refresh
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab3_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If ListaEstoque.ListItems.Count = 0 Then
    SSTab3.Tab = 0
    Exit Sub
End If
Select Case SSTab3.Tab
    Case 0:
        ListaEstoque.SetFocus
        ProcCarregaListaEstoque
    Case 1:
        ListaEmpenhoEstoquePed.SetFocus
        Txt_estoque_disp_est = ListaEstoque.SelectedItem.ListSubItems(8)
        ProcCarregaListaEmpEstPed
    Case 2:
        ListaEmpenhoEstoqueOrdem.SetFocus
        Txt_estoque_disp_est = ListaEstoque.SelectedItem.ListSubItems(8)
        ProcCarregaListaEmpEstOrdem
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposEmpEst()
On Error GoTo tratar_erro

Txt_ID_emp_est = 0
Txt_ID_carteira_est = 0
Txt_data_est = Format(Date, "dd/mm/yy")
Txt_responsavel_est = pubUsuario
Txt_pedido_est = ""
Txt_rev_est = ""
Txt_cliente_est = ""
Txt_prazo_est = ""
Txt_necessidade_vendas_est = ""
Txt_qtde_emp_est = ""
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab4_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If ListaProduzindo.ListItems.Count = 0 Then
    SSTab4.Tab = 0
    Exit Sub
End If
Select Case SSTab4.Tab
    Case 0:
        ListaProduzindo.SetFocus
        ProcCarregaListaProduzindo
        With MSFlexGrid1
            .rows = 0
            .Cols = 0
            .Refresh
        End With
    Case 1:
        ListaEmpenhoProduzindo.SetFocus
        Txt_produzindo_disp_prod = ListaProduzindo.SelectedItem.ListSubItems(5)
        ProcCarregaListaEmpprod
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposEmpProd()
On Error GoTo tratar_erro

Txt_ID_emp_prod = 0
Txt_ID_carteira_prod = 0
Txt_data_prod = Format(Date, "dd/mm/yy")
Txt_responsavel_prod = pubUsuario
Txt_pedido_prod = ""
Txt_rev_prod = ""
Txt_cliente_prod = ""
Txt_prazo_prod = ""
Txt_necessidade_vendas_prod = ""
Txt_qtde_emp_prod = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_emp_est_Change()
On Error GoTo tratar_erro

If Txt_qtde_emp_est <> "" Then
    VerifNumero = Txt_qtde_emp_est
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_qtde_emp_est = ""
        Txt_qtde_emp_est.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_emp_est_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_qtde_emp_est

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_emp_prod_Change()
On Error GoTo tratar_erro

If Txt_qtde_emp_prod <> "" Then
    VerifNumero = Txt_qtde_emp_prod
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_qtde_emp_prod = ""
        Txt_qtde_emp_prod.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_qtde_emp_prod_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus Txt_qtde_emp_prod

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

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

ListView1.ListItems.Clear
If txtTexto <> "" Then cmbfamilia.ListIndex = -1

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
    'Case 1: ProcAjuda
    Case 2: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcExcluirEmpPed
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar4_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    'Case 1: ProcAjuda
    Case 2: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirEmpProd()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaEmpenhoProduzindo
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) empenho(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from producao_pedidos where ID = " & .ListItems(InitFor)

            '==================================
            Modulo = "Vendas/Empenho"
            Evento = "Excluir empenho da produção"
            ID_documento = .ListItems(InitFor)
            Documento = "Ordem: " & ListaProduzindo.SelectedItem
            Documento1 = "Nº pedido: " & .ListItems(InitFor).SubItems(4) & " - Rev.: " & .ListItems(InitFor).SubItems(5)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) empenho(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Empenho(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposEmpProd
    ProcCarregaListaEmpprod
    With ListaProduzindo
        Valor2 = .SelectedItem.ListSubItems(3)
        ProcVerifQtdeEmpProd .SelectedItem
        Txt_produzindo_disp_prod = Valor2 - Valor3
    End With
    ProcLimpaCamposEmpEst
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirEmpEst()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With ListaEmpenhoEstoquePed
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) empenho(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from Estoque_controle_empenho_vendas where ID = " & .ListItems(InitFor)
            
            '==================================
            Modulo = "Vendas/Empenho"
            Evento = "Excluir empenho do estoque"
            ID_documento = .ListItems(InitFor)
            Documento = "RE: " & ListaEstoque.SelectedItem
            Documento1 = "Nº pedido: " & .ListItems(InitFor).SubItems(4) & " - Rev.: " & .ListItems(InitFor).SubItems(5)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) empenho(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Empenho(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCamposEmpEst
    ProcCarregaListaEmpEstPed
    With ListaEstoque
        Valor2 = .SelectedItem.ListSubItems(6)
        ProcVerifQtdeEmpEst .SelectedItem
        Txt_estoque_disp_est = Valor2 - Valor3
    End With
    ProcLimpaCamposEmpProd
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirEmpPed()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
If SSTab2.Tab = 0 Then ProcExcluirEmpPedEst Else ProcExcluirEmpPedProd

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirEmpPedEst()
On Error GoTo tratar_erro

With ListaEmpenhoVendidosRE
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) empenho(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from Estoque_controle_empenho_vendas where ID = " & .ListItems(InitFor)
            
            '==================================
            Modulo = "Vendas/Empenho"
            Evento = "Excluir empenho do estoque"
            ID_documento = .ListItems(InitFor)
            Documento = "Nº RE: " & .ListItems(InitFor).SubItems(1)
            Documento1 = "Nº pedido: " & ListaVendidos.SelectedItem.ListSubItems(1) & " - Rev.: " & ListaVendidos.SelectedItem.ListSubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) empenho(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Empenho(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaListaREEmpPed
    ProcCarregaListaVendidos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirEmpPedProd()
On Error GoTo tratar_erro

With ListaEmpenhoVendidosOP
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) empenho(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from producao_pedidos where ID = " & .ListItems(InitFor)

            '==================================
            Modulo = "Vendas/Empenho"
            Evento = "Excluir empenho da produção"
            ID_documento = .ListItems(InitFor)
            Documento = "Ordem: " & .ListItems(InitFor).SubItems(1)
            Documento1 = "Nº pedido: " & ListaVendidos.SelectedItem.ListSubItems(1) & " - Rev.: " & ListaVendidos.SelectedItem.ListSubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) empenho(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Empenho(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaListaOPEmpPed
    ProcCarregaListaVendidos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarEmpProd()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame8.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If Txt_necessidade_vendas_prod = 0 Then
    USMsgBox ("Não é permitido empenhar, pois o pedido interno já está empenhado integralmente."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "empenhar"
valor = IIf(Txt_qtde_emp_prod = "", 0, Txt_qtde_emp_prod)
If valor <= 0 Then
    NomeCampo = "a quantidade empenhada"
    ProcVerificaAcao
    Txt_qtde_emp_prod.SetFocus
    Exit Sub
End If
Valor1 = Txt_necessidade_vendas_prod
Valor2 = Txt_produzindo_disp_prod
If Novo_Vendas_Empenho1 = True Then
    If valor > Valor1 Then
        USMsgBox ("A quantidade empenhada não pode ser maior que a necessidade."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    If Valor2 < valor Then
        USMsgBox ("Não é permitido empenhar, pois a quantidade disponível produzindo é menor que a quantidade empenhada."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Producao_pedidos where ID = " & Txt_ID_emp_prod, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    If valor > (Valor1 + TBGravar!Qtde_empenho) Then
        USMsgBox ("A quantidade empenhada não pode ser maior que a necessidade."), vbExclamation, "CAPRIND v5.0"
        TBGravar.Close
        Exit Sub
    End If
    If (Valor2 + TBGravar!Qtde_empenho) < valor Then
        USMsgBox ("Não é permitido empenhar, pois a quantidade disponível produzindo é menor que a quantidade empenhada."), vbExclamation, "CAPRIND v5.0"
        TBGravar.Close
        Exit Sub
    End If
    If TBGravar!Qtde_entrada > 0 And TBGravar!Qtde_entrada > valor Then
        USMsgBox ("A quantidade empenhada não pode ser menor que a quantidade de entrada."), vbExclamation, "CAPRIND v5.0"
        TBGravar.Close
        Exit Sub
    End If
    If FunVerificaQtdePedFaturado(TBGravar!IDcarteira, TBGravar!Qtde_empenho) = False Then Exit Sub
Else
    TBGravar.AddNew
End If
TBGravar!Data = IIf(Txt_data_prod = "", Date, Txt_data_prod)
TBGravar!Responsavel = IIf(Txt_responsavel_prod = "", pubUsuario, Txt_responsavel_prod)
TBGravar!Ordem = ListaProduzindo.SelectedItem
TBGravar!IDcarteira = Txt_ID_carteira_prod
TBGravar!Qtde_empenho = valor
TBGravar.Update

ProcAtualizaQtdeEntEmpProd ListaProduzindo.SelectedItem, ListView1.SelectedItem.ListSubItems(1)
ProcCarregaListaEmpprod
If Novo_Vendas_Empenho1 = True Then
    USMsgBox ("Novo empenho da produção cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo empenho produção"
Else
    Evento = "Alterar empenho estoque"
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    If CodigoLista1 <> 0 And ListaEmpenhoProduzindo.ListItems.Count <> 0 Then
        ListaEmpenhoProduzindo.SelectedItem = ListaEmpenhoProduzindo.ListItems(CodigoLista1)
1:
        ListaEmpenhoProduzindo.SetFocus
    End If
End If
'==================================
Modulo = "Vendas/Empenho"
ID_documento = Txt_ID_emp_prod
Documento = "Nº pedido: " & Txt_pedido_prod & " - Rev.: " & Txt_rev_prod
Documento1 = ""
ProcGravaEvento
'==================================
If Novo_Vendas_Empenho1 = True Then
    With ListaVendidos
        ProcVerifNecessVendas .SelectedItem, .SelectedItem.ListSubItems(7)
    End With
Else
    ProcVerifNecessVendas Txt_ID_carteira_prod, ListaEmpenhoEstoquePed.SelectedItem.ListSubItems(10)
End If
Txt_necessidade_vendas_prod = Qtd - (qtdeliberada + qtdeliberar)
Txt_necessidade_vendas_est = Txt_necessidade_vendas_prod

With ListaProduzindo
    Valor2 = .SelectedItem.ListSubItems(3)
    ProcVerifQtdeEmpProd .SelectedItem
    Txt_produzindo_disp_prod = Valor2 - Valor3
End With
Novo_Vendas_Empenho1 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoEmpEst()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With ListaVendidos
    If .ListItems.Count = 0 Then
        USMsgBox ("Informe o pedido interno na lista vendidos antes de criar novo empenho."), vbExclamation, "CAPRIND v5.0"
        Frame3.Enabled = False
        Exit Sub
    Else
        ProcLimpaCamposEmpEst
        Txt_ID_carteira_est = .SelectedItem
        Txt_pedido_est = .SelectedItem.SubItems(1)
        Txt_rev_est = .SelectedItem.SubItems(2)
        Txt_cliente_est = .SelectedItem.SubItems(3)
        Txt_prazo_est = .SelectedItem.SubItems(6)
        
        With ListaVendidos
            ProcVerifNecessVendas .SelectedItem, .SelectedItem.ListSubItems(7)
        End With
        Txt_necessidade_vendas_est = Qtd - (qtdeliberada + qtdeliberar)
        
        Frame3.Enabled = True
        Txt_qtde_emp_est.SetFocus
    End If
End With
Novo_Vendas_Empenho = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoEmpProd()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With ListaVendidos
    If .ListItems.Count = 0 Then
        USMsgBox ("Informe o pedido interno na lista vendidos antes de criar novo empenho."), vbExclamation, "CAPRIND v5.0"
        Frame8.Enabled = False
        Exit Sub
    Else
        ProcLimpaCamposEmpProd
        Txt_ID_carteira_prod = .SelectedItem
        Txt_pedido_prod = .SelectedItem.SubItems(1)
        Txt_rev_prod = .SelectedItem.SubItems(2)
        Txt_cliente_prod = .SelectedItem.SubItems(3)
        Txt_prazo_prod = .SelectedItem.SubItems(6)
        
        With ListaVendidos
            ProcVerifNecessVendas .SelectedItem, .SelectedItem.ListSubItems(7)
        End With
        Txt_necessidade_vendas_prod = Qtd - (qtdeliberada + qtdeliberar)
        
        Frame8.Enabled = True
        Txt_qtde_emp_prod.SetFocus
    End If
End With
Novo_Vendas_Empenho1 = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarEmpEst()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame3.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If Txt_necessidade_vendas_est = 0 Then
    USMsgBox ("Não é permitido empenhar, pois o pedido interno já está empenhado integralmente."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "empenhar"
valor = IIf(Txt_qtde_emp_est = "", 0, Txt_qtde_emp_est)
If valor <= 0 Then
    NomeCampo = "a quantidade empenhada"
    ProcVerificaAcao
    Txt_qtde_emp_est.SetFocus
    Exit Sub
End If
Valor1 = Txt_necessidade_vendas_est
Valor2 = Txt_estoque_disp_est
If Novo_Vendas_Empenho = True Then
    If valor > Valor1 Then
        USMsgBox ("A quantidade empenhada não pode ser maior que a necessidade."), vbExclamation, "CAPRIND v5.0"
        Txt_qtde_emp_est.SetFocus
        Exit Sub
    End If
    If Valor2 < valor Then
        USMsgBox ("Não é permitido empenhar, pois a quantidade disponível em estoque é menor que a quantidade empenhada."), vbExclamation, "CAPRIND v5.0"
        Txt_qtde_emp_est.SetFocus
        Exit Sub
    End If
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from estoque_controle_empenho_vendas where ID = " & Txt_ID_emp_est, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    If valor > (Valor1 + TBGravar!Qtde_empenhada) Then
        USMsgBox ("A quantidade empenhada não pode ser maior que a necessidade."), vbExclamation, "CAPRIND v5.0"
        TBGravar.Close
        Exit Sub
    End If
    If (Valor2 + TBGravar!Qtde_empenhada) < valor Then
        USMsgBox ("Não é permitido empenhar, pois a quantidade disponível em estoque é menor que a quantidade empenhada."), vbExclamation, "CAPRIND v5.0"
        TBGravar.Close
        Exit Sub
    End If
    If TBGravar!Qtde_saida > 0 And TBGravar!Qtde_saida > valor Then
        USMsgBox ("A quantidade empenhada não pode ser menor que a quantidade de saída."), vbExclamation, "CAPRIND v5.0"
        TBGravar.Close
        Exit Sub
    End If
    If FunVerificaQtdePedFaturado(TBGravar!ID_carteira, TBGravar!Qtde_empenhada) = False Then Exit Sub
Else
    TBGravar.AddNew
End If
TBGravar!Data = IIf(Txt_data_est = "", Date, Txt_data_est)
TBGravar!Responsavel = IIf(Txt_responsavel_est = "", pubUsuario, Txt_responsavel_est)
TBGravar!ID_estoque = ListaEstoque.SelectedItem
TBGravar!ID_carteira = Txt_ID_carteira_est
TBGravar!Qtde_empenhada = valor
TBGravar.Update
ProcCarregaListaEmpEstPed
If Novo_Vendas_Empenho = True Then
    USMsgBox ("Novo empenho do estoque cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo empenho estoque"
Else
    Evento = "Alterar empenho estoque"
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    If CodigoLista <> 0 And ListaEmpenhoEstoquePed.ListItems.Count <> 0 Then
        ListaEmpenhoEstoquePed.SelectedItem = ListaEmpenhoEstoquePed.ListItems(CodigoLista)
1:
        ListaEmpenhoEstoquePed.SetFocus
    End If
End If
'==================================
Modulo = "Vendas/Empenho"
ID_documento = Txt_ID_emp_est
Documento = "Nº pedido: " & Txt_pedido_est & " - Rev.: " & Txt_rev_est
Documento1 = ""
ProcGravaEvento
'==================================
If Novo_Vendas_Empenho = True Then
    With ListaVendidos
        ProcVerifNecessVendas .SelectedItem, .SelectedItem.ListSubItems(7)
    End With
Else
    ProcVerifNecessVendas Txt_ID_carteira_est, ListaEmpenhoEstoquePed.SelectedItem.ListSubItems(10)
End If
Txt_necessidade_vendas_est = Qtd - (qtdeliberada + qtdeliberar)
Txt_necessidade_vendas_prod = Txt_necessidade_vendas_est

With ListaEstoque
    Valor2 = .SelectedItem.ListSubItems(6)
    ProcVerifQtdeEmpEst .SelectedItem
    Txt_estoque_disp_est = Valor2 - Valor3
End With
Novo_Vendas_Empenho = False

Exit Sub
tratar_erro:
    If Err.Number = "35600" Then GoTo 1
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Function FunVerificaQtdePedFaturado(ID_carteira As String, Qtde_empenho_antigo) As Boolean
On Error GoTo tratar_erro

FunVerificaQtdePedFaturado = True
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from vendas_carteira where Codigo = " & ID_carteira & " and (Liberacao = 'FATURADO' or Liberacao = 'FATURADO PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    qtdeliberar = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "SELECT SUM(Qtde_Empenhada) as qtdeliberar from Estoque_Controle_Empenho_Vendas where ID_carteira = " & TBFI!CODIGO, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        qtdeliberar = IIf(IsNull(TBAbrir!qtdeliberar), 0, TBAbrir!qtdeliberar)
    End If
    qtdeliberada = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "SELECT SUM(PP.Qtde_Empenho) as qtdeliberada from Producao_pedidos PP INNER JOIN Vendas_carteira VC ON VC.Codigo = PP.IDCarteira where PP.IDcarteira = " & ID_carteira & " and VC.desenho = '" & ListView1.SelectedItem.ListSubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        qtdeliberada = IIf(IsNull(TBAbrir!qtdeliberada), 0, TBAbrir!qtdeliberada)
    End If
    QTLOTE = (qtdeliberar + qtdeliberada + valor) - Qtde_empenho_antigo
    If QTLOTE < TBFI!QtdeFaturada Then
        USMsgBox ("A quantidade empenhada não pode ser menor que a quantidade faturada."), vbExclamation, "CAPRIND v5.0"
        FunVerificaQtdePedFaturado = False
        Exit Function
    End If
End If
TBFI.Close

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub USToolBar5_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoEmpEst
    Case 2: ProcSalvarEmpEst
    Case 3: ProcExcluirEmpEst
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar6_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    'Case 1: ProcAjuda
    Case 2: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar7_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoEmpProd
    Case 2: ProcSalvarEmpProd
    Case 3: ProcExcluirEmpProd
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Grid1_Click()
On Error GoTo tratar_erro
Dim point As POINTAPI
Dim objCell As FlexCell.Cell
Dim intWidth As Integer

If FunCheckEditStatus() Then Exit Sub
intWidth = 20

Call GetCursorPos(point)
Call ScreenToClient(Grid1.hWnd, point)
Set objCell = Grid1.HitTest(point.x, point.Y)

If Not objCell Is Nothing Then
    If objCell.Row >= m_Row And objCell.Col = m_Col Then
        Dim objNode As Node
        Set objNode = m_Tree.FindNode(objCell.Row - m_Row + 2)
        If Not objNode Is Nothing Then
            Dim i As Long, x As Long, Y As Long
            x = objCell.Left + 2 + (objNode.Level - 1) * intWidth
            Y = objCell.Top + (objCell.Height - 9) / 2
            If point.x >= x And point.x <= x + 9 And point.Y >= Y And point.Y <= Y + 9 Then
                If objNode.Expanded Then
                    objNode.Collapse
                    Grid1.AutoRedraw = False
                    For i = 1 To objNode.ChildrenCount
                        Grid1.RowHeight(objCell.Row + i) = 0
                    Next
                    Grid1.AutoRedraw = True
                    Grid1.Refresh
                Else
                    objNode.Expand
                    Grid1.AutoRedraw = False
                    For i = 1 To objNode.ChildrenCount
                        If objNode.FindNode(i + 1).Visible Then
                            Grid1.RowHeight(objCell.Row + i) = -1 'DefaultRowHeight
                        End If
                    Next
                    Grid1.AutoRedraw = True
                    Grid1.Refresh
                End If
            End If
        End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Grid1_OwnerDrawCell(ByVal Row As Long, ByVal Col As Long, ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Handled As Boolean)
On Error GoTo tratar_erro
Dim i As Long, j As Long
Dim x As Long, Y As Long
Dim hPen As Long, hOldPen As Long
Dim hBrush As Long, hOldBrush As Long
Dim lngLevel As Long
Dim blnDrawLine As Boolean
Dim objNode As Node, tmpNode As Node
Dim intWidth As Integer
Dim intAdd As Integer

If Row < m_Row Or Col <> m_Col Then Exit Sub

intWidth = 20
intAdd = 26
    
Set objNode = m_Tree.FindNode(Row - m_Row + 2)
If Not objNode Is Nothing Then
    lngLevel = objNode.Level - 1

    'Tree lines
    hPen = CreatePen(0, 1, RGB(128, 128, 128))
    hOldPen = SelectObject(hdc, hPen)
    For i = 0 To lngLevel
        If i < lngLevel - 1 Then
            blnDrawLine = True
            Set tmpNode = objNode
            For j = i To lngLevel - 2
                Set tmpNode = tmpNode.Parent
            Next
            If tmpNode.NextNode Is Nothing Then
                blnDrawLine = False
            End If
            If blnDrawLine Then
                'All
                Call DrawLine(hdc, Left + intWidth * i + intAdd, Top - 1, Left + intWidth * i + intAdd, Bottom + 1)
            End If
        ElseIf i = lngLevel - 1 Then
            'Top
            Call DrawLine(hdc, Left + intWidth * i + intAdd, Top - 1, Left + intWidth * i + intAdd, Top + (Bottom - Top) / 2)
            If Not objNode.NextNode Is Nothing Then
                'Bottom
                Call DrawLine(hdc, Left + intWidth * i + intAdd, Top + (Bottom - Top) / 2, Left + intWidth * i + intAdd, Bottom + 1)
            End If
        ElseIf i = lngLevel Then
            'Top
            If objNode.VisibleNodesCount > 1 Then
                Call DrawLine(hdc, Left + intWidth * i + intAdd, Top + (Bottom - Top) / 2, Left + intWidth * i + intAdd, Bottom + 1)
            End If
        End If
        'Horizontal line
        If lngLevel > 0 Then
            Call DrawLine(hdc, Left + intWidth * (lngLevel - 1) + intAdd, Top + (Bottom - Top) / 2, Left + intWidth * (lngLevel - 1) + intAdd + 10, Top + (Bottom - Top) / 2)
        End If
    Next
    
    Call SelectObject(hdc, hOldPen)
    Call DeleteObject(hPen)

    '+/-
    If objNode.ChildrenCount > 0 Then
        hPen = CreatePen(0, 1, 0)
        hOldPen = SelectObject(hdc, hPen)
        hBrush = CreateSolidBrush(RGB(255, 255, 255))
        hOldPen = SelectObject(hdc, hBrush)
        
        x = Left + 2 + intWidth * lngLevel
        Y = Top + (Bottom - Top - 9) / 2
        
        Call Rectangle(hdc, x, Y, x + 9, Y + 9)
        If objNode.Expanded Then
            Call DrawLine(hdc, x + 2, Y + 4, x + 7, Y + 4)
        Else
            Call DrawLine(hdc, x + 2, Y + 4, x + 7, Y + 4)
            Call DrawLine(hdc, x + 4, Y + 2, x + 4, Y + 7)
        End If
    
        Call SelectObject(hdc, hOldPen)
        Call DeleteObject(hPen)
        Call SelectObject(hdc, hOldBrush)
        Call DeleteObject(hBrush)
    End If
    
    'Icon
    If objNode.ChildrenCount > 0 Then
        DrawIconEx hdc, Left + intWidth * lngLevel + 18, Top + (Bottom - Top - 16) / 2, imgFolder.Picture, 16, 16, 0, 0, DI_NORMAL
    Else
        DrawIconEx hdc, Left + intWidth * lngLevel + 18, Top + (Bottom - Top - 16) / 2, imgFile.Picture, 16, 16, 0, 0, DI_NORMAL
    End If
    
    'Text
    With Grid1.Cell(Row, Col)
        Dim rc As rect
        Call SetRect(rc, Left + intWidth * lngLevel + 37, Top, Right, Bottom)
        Call DrawText(hdc, .Text, -1, rc, DT_SINGLELINE Or DT_VCENTER)
    End With

    Handled = True
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Function FunCheckEditStatus() As Boolean
On Error GoTo tratar_erro
Dim hWnd As Long
Dim strClassName As String
Dim intPos As Integer

strClassName = Space(256)
hWnd = GetFocus()
Call GetClassName(hWnd, strClassName, 256)
intPos = InStr(1, strClassName, Chr(0))
strClassName = Left(strClassName, intPos - 1)
If strClassName = "ThunderRT6TextBox" Then FunCheckEditStatus = True    'Editing Else    FunCheckEditStatus = False

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

