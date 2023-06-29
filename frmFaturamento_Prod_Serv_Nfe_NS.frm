VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmFaturamento_Prod_Serv_NFe_NS 
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
   Icon            =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
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
   Begin TabDlg.SSTab SStab_nfe 
      Height          =   10035
      Left            =   0
      TabIndex        =   25
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
      TabPicture(0)   =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Txt_ID_cobranca"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtdatacancelamento"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "USToolBar1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtID_nota"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "SSTab1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Lista de produtos"
      TabPicture(1)   =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "listaProdutos"
      Tab(1).Control(1)=   "USToolBar2"
      Tab(1).Control(2)=   "FrameCST"
      Tab(1).Control(3)=   "txtID_item"
      Tab(1).Control(4)=   "Frame_comb_lub"
      Tab(1).ControlCount=   5
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Certificado digital"
         Height          =   840
         Left            =   11580
         TabIndex        =   103
         Top             =   1260
         Width           =   3705
         Begin VB.TextBox txtSerialCertificado 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000040&
            Height          =   315
            Left            =   90
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   106
            TabStop         =   0   'False
            ToolTipText     =   "Serial Certificado"
            Top             =   390
            Width           =   1545
         End
         Begin VB.TextBox txtTPCertificado 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000040&
            Height          =   315
            Left            =   1650
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   105
            TabStop         =   0   'False
            ToolTipText     =   "Status NFe."
            Top             =   390
            Width           =   315
         End
         Begin VB.TextBox txtValidade 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   1980
            MaxLength       =   60
            TabIndex        =   104
            TabStop         =   0   'False
            ToolTipText     =   "Data de validade do certificado digital"
            Top             =   390
            Width           =   1575
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Número serial"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   375
            TabIndex        =   109
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   1650
            TabIndex        =   108
            Top             =   180
            Width           =   300
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Validade"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   2467
            TabIndex        =   107
            Top             =   180
            Width           =   600
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
         Left            =   60
         TabIndex        =   33
         Top             =   1260
         Width           =   11505
         Begin DrawSuite2022.USButton cmdConsultar 
            Height          =   315
            Left            =   5700
            TabIndex        =   89
            ToolTipText     =   "Consultar nota fiscal no SEFAZ com chave de acesso."
            Top             =   390
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":107A
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
            PicAlign        =   0
            Theme           =   1
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin VB.TextBox txt_nProt 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   7170
            MaxLength       =   60
            TabIndex        =   61
            TabStop         =   0   'False
            ToolTipText     =   "Status NFe."
            Top             =   390
            Width           =   1455
         End
         Begin VB.TextBox txtnsNrec 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   6030
            MaxLength       =   60
            TabIndex        =   59
            TabStop         =   0   'False
            ToolTipText     =   "Status NFe."
            Top             =   390
            Width           =   795
         End
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":820D
            Left            =   4950
            List            =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":8265
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            ToolTipText     =   "UF."
            Top             =   1830
            Width           =   630
         End
         Begin VB.TextBox txtLocal_embarque 
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
            Left            =   1470
            MaxLength       =   60
            TabIndex        =   8
            ToolTipText     =   "Local onde ocorrerá o embarque dos produtos."
            Top             =   1830
            Width           =   3465
         End
         Begin VB.TextBox txtSerie 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   990
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Série."
            Top             =   390
            Width           =   435
         End
         Begin VB.TextBox txtchNFe 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   1440
            MaxLength       =   44
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Chave de acesso NFe."
            Top             =   390
            Width           =   4245
         End
         Begin VB.TextBox txtNota 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   90
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Número da NFe."
            Top             =   390
            Width           =   885
         End
         Begin VB.TextBox txtcStat 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   8655
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Status NFe."
            Top             =   390
            Width           =   2745
         End
         Begin DrawSuite2022.USButton BTNnsnrec 
            Height          =   315
            Left            =   6840
            TabIndex        =   90
            ToolTipText     =   "Consultar recibo nsNRec no SEFAZ com chave de acesso."
            Top             =   390
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":82D7
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
            PicAlign        =   0
            Theme           =   1
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Chave proteção"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   7320
            TabIndex        =   62
            Top             =   180
            Width           =   1155
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Recibo"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   6210
            TabIndex        =   60
            Top             =   180
            Width           =   480
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Local onde ocorrerá o embarque dos produtos"
            Height          =   195
            Index           =   1
            Left            =   1545
            TabIndex        =   53
            Top             =   1620
            Width           =   3315
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UF"
            Height          =   195
            Left            =   5085
            TabIndex        =   52
            Top             =   1620
            Width           =   195
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Série"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   1035
            TabIndex        =   38
            Top             =   180
            Width           =   360
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Chave de acesso"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   2947
            TabIndex        =   36
            Top             =   180
            Width           =   1230
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nota fiscal"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   165
            TabIndex        =   35
            Top             =   180
            Width           =   750
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Status Nota fiscal eletrônica"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   8992
            TabIndex        =   34
            Top             =   180
            Width           =   2070
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5745
         Left            =   60
         TabIndex        =   64
         Top             =   3600
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   10134
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         WordWrap        =   0   'False
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Lista de notas fiscais"
         TabPicture(0)   =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":F46A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ListaNota"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Retorno SEFAZ"
         TabPicture(1)   =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":F486
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame8"
         Tab(1).Control(1)=   "Frame7"
         Tab(1).Control(2)=   "Frame6"
         Tab(1).Control(3)=   "Frame5"
         Tab(1).Control(4)=   "Frame11"
         Tab(1).Control(5)=   "Frame12"
         Tab(1).Control(6)=   "Frame13"
         Tab(1).Control(7)=   "Frame14"
         Tab(1).Control(8)=   "Frame15"
         Tab(1).Control(9)=   "BtnValidadorXML"
         Tab(1).Control(10)=   "TxtRetorno"
         Tab(1).ControlCount=   11
         Begin VB.TextBox TxtRetorno 
            BorderStyle     =   0  'None
            Height          =   1965
            Left            =   -72480
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   99
            Top             =   450
            Width           =   11085
         End
         Begin DrawSuite2022.USButton BtnValidadorXML 
            Height          =   1965
            Left            =   -61230
            TabIndex        =   87
            Top             =   390
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   3466
            DibPicture      =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":F4A2
            Caption         =   "Validar XML"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
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
            PicSize         =   5
            PicSizeH        =   64
            PicSizeW        =   64
            ShowFocusRect   =   0   'False
            Theme           =   3
         End
         Begin VB.Frame Frame15 
            Caption         =   "Diretório de arquivos log de retorno"
            Height          =   615
            Left            =   -63630
            TabIndex        =   66
            Top             =   2370
            Width           =   3735
            Begin VB.TextBox txtD4 
               Enabled         =   0   'False
               Height          =   320
               Left            =   90
               TabIndex        =   67
               Top             =   210
               Width           =   3195
            End
            Begin DrawSuite2022.USButton cmdD4 
               Height          =   315
               Left            =   3300
               TabIndex        =   94
               ToolTipText     =   "Abrir diretório de arquivos log de retorno SEFAZ..."
               Top             =   210
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               DibPicture      =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":1288A
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
               PicAlign        =   0
               Theme           =   1
               ToolTipTitle    =   "CAPRIND v5.0"
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Diretório de retorno"
            Height          =   615
            Left            =   -67350
            TabIndex        =   68
            Top             =   2370
            Width           =   3705
            Begin VB.TextBox txtD3 
               Enabled         =   0   'False
               Height          =   320
               Left            =   90
               TabIndex        =   69
               ToolTipText     =   "Abrir diretório de retorno..."
               Top             =   210
               Width           =   3225
            End
            Begin DrawSuite2022.USButton cmdD3 
               Height          =   315
               Left            =   3330
               TabIndex        =   93
               ToolTipText     =   "Abrir diretório de retorno"
               Top             =   210
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               DibPicture      =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":3098F
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
               PicAlign        =   0
               Theme           =   1
               ToolTipTitle    =   "CAPRIND v5.0"
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "Diretório DANFE - XML"
            Height          =   615
            Left            =   -71130
            TabIndex        =   70
            Top             =   2370
            Width           =   3765
            Begin VB.TextBox txtD2 
               Enabled         =   0   'False
               Height          =   320
               Left            =   90
               TabIndex        =   71
               Top             =   210
               Width           =   3225
            End
            Begin DrawSuite2022.USButton cmdD2 
               Height          =   315
               Left            =   3330
               TabIndex        =   92
               ToolTipText     =   "Abrir diretório DANFE-XML..."
               Top             =   210
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               DibPicture      =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":4EA94
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
               PicAlign        =   0
               Theme           =   1
               ToolTipTitle    =   "CAPRIND v5.0"
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Diretório de envio"
            Height          =   615
            Left            =   -74880
            TabIndex        =   83
            Top             =   2370
            Width           =   3735
            Begin DrawSuite2022.USButton cmdD1 
               Height          =   315
               Left            =   3330
               TabIndex        =   91
               ToolTipText     =   "Abrir diretório de envio..."
               Top             =   210
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               DibPicture      =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":6CB99
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
               PicAlign        =   0
               Theme           =   1
               ToolTipTitle    =   "CAPRIND v5.0"
            End
            Begin VB.TextBox txtD1 
               Enabled         =   0   'False
               Height          =   320
               Left            =   90
               TabIndex        =   84
               Top             =   210
               Width           =   3225
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Arquivos log..."
            Height          =   2565
            Left            =   -61440
            TabIndex        =   85
            Top             =   3060
            Width           =   1545
            Begin VB.FileListBox File4 
               Height          =   2040
               Left            =   60
               TabIndex        =   86
               Top             =   330
               Width           =   1395
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Operações NFe (SEFAZ)"
            Height          =   2055
            Left            =   -74880
            TabIndex        =   78
            Top             =   330
            Width           =   2295
            Begin DrawSuite2022.USButton BtnCriarXML 
               Height          =   405
               Left            =   120
               TabIndex        =   80
               ToolTipText     =   "Criar XML da NFe pra assinatura e envio"
               Top             =   240
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   714
               Caption         =   "Criar XML"
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
               ShowFocusRect   =   0   'False
               Theme           =   6
            End
            Begin VB.CheckBox chk_Cp 
               Height          =   285
               Left            =   1890
               TabIndex        =   79
               Top             =   -30
               Width           =   165
            End
            Begin DrawSuite2022.USButton BtnAssinarXML 
               Height          =   405
               Left            =   120
               TabIndex        =   81
               ToolTipText     =   "Assinar XML criado para envio"
               Top             =   690
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   714
               Caption         =   "Assinar XML"
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
               ShowFocusRect   =   0   'False
               Theme           =   3
            End
            Begin DrawSuite2022.USButton BtnEnviarXML 
               Height          =   375
               Left            =   120
               TabIndex        =   82
               ToolTipText     =   "Enviar XML da NFe assinado para o SEFAZ."
               Top             =   1590
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   661
               Caption         =   "Enviar XML"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderColor     =   8388608
               BorderColorDisabled=   13160660
               BorderColorDown =   12582912
               BorderColorOver =   16711680
               GradientColor1  =   8388608
               GradientColor2  =   8388608
               GradientColor3  =   8388608
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
               ShowFocusRect   =   0   'False
               Theme           =   4
            End
            Begin DrawSuite2022.USButton btnPrevia 
               Height          =   405
               Left            =   120
               TabIndex        =   100
               ToolTipText     =   "Prévia da DANFE em pdf..."
               Top             =   1140
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   714
               Caption         =   "Previa da DANFE"
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
               ShowFocusRect   =   0   'False
               Theme           =   5
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Arquivos xml de envio"
            Height          =   2565
            Left            =   -74880
            TabIndex        =   76
            Top             =   3060
            Width           =   2145
            Begin VB.FileListBox File1 
               Height          =   2040
               Left            =   120
               TabIndex        =   77
               Top             =   330
               Width           =   1935
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Arquivos Danfe e XML"
            Height          =   2565
            Left            =   -72720
            TabIndex        =   74
            Top             =   3060
            Width           =   5385
            Begin VB.FileListBox File2 
               Height          =   2040
               Left            =   120
               TabIndex        =   75
               Top             =   330
               Width           =   5145
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Arquivos carta de correção"
            Height          =   2565
            Left            =   -67320
            TabIndex        =   72
            Top             =   3060
            Width           =   5865
            Begin VB.FileListBox File3 
               Height          =   2040
               Left            =   120
               TabIndex        =   73
               Top             =   330
               Width           =   5655
            End
         End
         Begin MSComctlLib.ListView ListaNota 
            Height          =   5310
            Left            =   60
            TabIndex        =   65
            Top             =   360
            Width           =   15135
            _ExtentX        =   26696
            _ExtentY        =   9366
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
            Appearance      =   0
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
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Object.Tag             =   "D"
               Text            =   "Dt. emissão"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Nota fiscal"
               Object.Width           =   1940
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
               Alignment       =   2
               SubItemIndex    =   8
               Object.Tag             =   "T"
               Text            =   "Status"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Object.Tag             =   "T"
               Text            =   "Status NFe"
               Object.Width           =   6174
            EndProperty
         End
      End
      Begin VB.Frame Frame_comb_lub 
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
         TabIndex        =   44
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
            TabIndex        =   57
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":8AC9E
            Left            =   12720
            List            =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":8ACB4
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   16
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":8AD14
            Left            =   180
            List            =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":8AD16
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   14
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":8AD18
            Left            =   11790
            List            =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":8AD1A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   15
            ToolTipText     =   "UF de consumo."
            Top             =   450
            Width           =   915
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição do produto da ANP"
            Height          =   195
            Left            =   6135
            TabIndex        =   58
            Top             =   240
            Width           =   2100
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo do produto"
            Height          =   195
            Left            =   13297
            TabIndex        =   47
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UF cons."
            Height          =   195
            Left            =   11925
            TabIndex        =   46
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código do produto da ANP"
            Height          =   195
            Left            =   435
            TabIndex        =   45
            Top             =   240
            Width           =   1905
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   60
         TabIndex        =   41
         Top             =   9360
         Width           =   15255
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
            Left            =   2730
            TabIndex        =   18
            Text            =   "24"
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
            TabIndex        =   19
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   555
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   11760
            TabIndex        =   23
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":8AD1C
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
            TabIndex        =   22
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":8E4C0
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
            TabIndex        =   20
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
            TabIndex        =   21
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":91FC9
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
            TabIndex        =   24
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":960B8
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
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "registros por página"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3360
            TabIndex        =   54
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carregar"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2040
            TabIndex        =   48
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lblRegistros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de registros: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   43
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label lblPaginas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Página: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13050
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox txtID_nota 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2130
         TabIndex        =   29
         Text            =   "0"
         Top             =   7530
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txtID_item 
         Alignment       =   2  'Center
         Height          =   335
         Left            =   -72870
         TabIndex        =   26
         Text            =   "0"
         ToolTipText     =   "id do produto."
         Top             =   4950
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000080&
         Height          =   1575
         Left            =   55
         TabIndex        =   27
         Top             =   2040
         Width           =   15225
         Begin VB.ComboBox cmbOperacao 
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":99944
            Left            =   2340
            List            =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":99951
            Style           =   2  'Dropdown List
            TabIndex        =   117
            ToolTipText     =   "Forma de emissão."
            Top             =   1080
            Width           =   1425
         End
         Begin VB.TextBox txtID_cobranca 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7950
            TabIndex        =   115
            Text            =   "0"
            Top             =   390
            Width           =   645
         End
         Begin VB.TextBox txtID_entrega 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   114
            Text            =   "0"
            Top             =   390
            Width           =   645
         End
         Begin VB.TextBox txtCobranca 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   8610
            TabIndex        =   112
            TabStop         =   0   'False
            ToolTipText     =   "Chave de acesso NFe."
            Top             =   390
            Width           =   6165
         End
         Begin VB.TextBox txtEntrega 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   780
            TabIndex        =   110
            TabStop         =   0   'False
            ToolTipText     =   "Chave de acesso NFe."
            Top             =   390
            Width           =   6765
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":99985
            Left            =   10860
            List            =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":99992
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "Indicador da forma de pagamento."
            Top             =   1080
            Width           =   1845
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":999C6
            Left            =   6510
            List            =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":999FA
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Forma de pagamento."
            Top             =   1080
            Width           =   4335
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":99B8A
            Left            =   12720
            List            =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":99BA3
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "Indicador de presença do comprador no estabelecimento comercial no momento da operação."
            Top             =   1080
            Width           =   2415
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":99CB4
            Left            =   5610
            List            =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":99CBE
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "Operação com consumidor final."
            Top             =   1080
            Width           =   885
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":99CD4
            Left            =   120
            List            =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":99CED
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Forma de emissão."
            Top             =   1080
            Width           =   2205
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":99F1E
            Left            =   3780
            List            =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":99F2E
            Style           =   2  'Dropdown List
            TabIndex        =   1
            ToolTipText     =   "Finalidade de emissão."
            Top             =   1080
            Width           =   1815
         End
         Begin XtremeSuiteControls.CheckBox Chk_DA_entrega 
            Height          =   195
            Left            =   270
            TabIndex        =   101
            Top             =   180
            Width           =   4845
            _Version        =   1245187
            _ExtentX        =   8546
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Endereço de entrega (Imprimir nos dados adicionais)"
            ForeColor       =   128
            UseVisualStyle  =   -1  'True
            DrawFocusRect   =   0   'False
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox Chk_DA_cobranca 
            Height          =   195
            Left            =   8160
            TabIndex        =   102
            Top             =   180
            Width           =   4845
            _Version        =   1245187
            _ExtentX        =   8546
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Endereço de cobrança (Imprimir nos dados adicionais)"
            ForeColor       =   128
            UseVisualStyle  =   -1  'True
            DrawFocusRect   =   0   'False
            Value           =   1
         End
         Begin DrawSuite2022.USButton btnEntrega 
            Height          =   315
            Left            =   7560
            TabIndex        =   111
            ToolTipText     =   "Consultar nota fiscal no SEFAZ com chave de acesso."
            Top             =   390
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":99F73
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
            PicAlign        =   0
            Theme           =   1
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin DrawSuite2022.USButton btnCobranca 
            Height          =   315
            Left            =   14790
            TabIndex        =   113
            ToolTipText     =   "Consultar nota fiscal no SEFAZ com chave de acesso."
            Top             =   390
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
            DibPicture      =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":A1106
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
            PicAlign        =   0
            Theme           =   1
            ToolTipTitle    =   "CAPRIND v5.0"
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operação"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   2565
            TabIndex        =   118
            Top             =   870
            Width           =   705
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ind. forma de pagto.*"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   10995
            TabIndex        =   56
            Top             =   870
            Width           =   1605
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forma de pagamento*"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   8070
            TabIndex        =   55
            Top             =   870
            Width           =   1620
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Presença do comprador*"
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   4
            Left            =   13050
            TabIndex        =   51
            Top             =   870
            Width           =   1965
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cons.final*"
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   3
            Left            =   5625
            TabIndex        =   50
            Top             =   870
            Width           =   900
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forma de emissão*"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   532
            TabIndex        =   37
            Top             =   870
            Width           =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Finalidade de emissão*"
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   0
            Left            =   3855
            TabIndex        =   28
            Top             =   870
            Width           =   1650
         End
      End
      Begin VB.Frame FrameCST 
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
         TabIndex        =   30
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":A8299
            Left            =   7590
            List            =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":A82AF
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   13
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
            ItemData        =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":A835E
            Left            =   180
            List            =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":A836E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            ToolTipText     =   "Modalidade de determinação da BC."
            Top             =   450
            Width           =   7395
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modalidade de determinação da BC ST"
            Height          =   195
            Left            =   9945
            TabIndex        =   32
            Top             =   240
            Width           =   2745
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modalidade de determinação da BC"
            Height          =   195
            Left            =   2617
            TabIndex        =   31
            Top             =   240
            Width           =   2520
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   60
         TabIndex        =   39
         Top             =   330
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   1720
         ButtonCount     =   10
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
         ButtonCaption3  =   "Relatório"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Relatório (F5)"
         ButtonKey3      =   "4"
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
         ButtonWidth3    =   51
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Consultar status"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Consultar status da nota fiscal.."
         ButtonKey4      =   "5"
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
         ButtonLeft4     =   135
         ButtonTop4      =   2
         ButtonWidth4    =   87
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Enviar DFE"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Enviar Danfe e XML por email"
         ButtonKey5      =   "6"
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
         ButtonLeft5     =   224
         ButtonTop5      =   2
         ButtonWidth5    =   60
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Baixar DFE"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Fazer download dos arquivos Danfe e XML"
         ButtonKey6      =   "7"
         ButtonAlignment6=   2
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
         ButtonLeft6     =   286
         ButtonTop6      =   2
         ButtonWidth6    =   61
         ButtonHeight6   =   21
         ButtonCaption7  =   "Cancelar"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Cancelar nota fiscal (F4)"
         ButtonKey7      =   "3"
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
         ButtonLeft7     =   349
         ButtonTop7      =   2
         ButtonWidth7    =   50
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Inutilizar NFe"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Inutilizar numero da NFe"
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
         ButtonLeft8     =   401
         ButtonTop8      =   2
         ButtonWidth8    =   71
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Ajuda"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Ajuda (F1)"
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
         ButtonLeft9     =   474
         ButtonTop9      =   2
         ButtonWidth9    =   36
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonCaption10 =   "Sair"
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonToolTipText10=   "Sair (Esc)"
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
         ButtonLeft10    =   512
         ButtonTop10     =   2
         ButtonWidth10   =   26
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         Begin VB.CheckBox chkOperacaoExterna 
            Alignment       =   1  'Right Justify
            Caption         =   "Operação Externa?"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   13410
            TabIndex        =   116
            Top             =   720
            Width           =   1695
         End
         Begin DrawSuite2022.USImageList USImageList1 
            Left            =   0
            Top             =   0
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":A83DC
            Count           =   1
         End
         Begin VB.CheckBox chkUsuario 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            Caption         =   "Me enviar uma cópia da Danfe e XML"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   12150
            TabIndex        =   98
            Top             =   0
            Width           =   2955
         End
         Begin VB.CheckBox chkTransportadora 
            Alignment       =   1  'Right Justify
            Caption         =   "Enviar Danfe e XML para transportadora"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   11850
            TabIndex        =   97
            Top             =   240
            Width           =   3255
         End
         Begin VB.CheckBox chkCodRef 
            Alignment       =   1  'Right Justify
            Caption         =   "Cód. referência na DANFE"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   12900
            TabIndex        =   96
            Top             =   480
            Width           =   2205
         End
         Begin VB.CheckBox chkTPAmb 
            Alignment       =   1  'Right Justify
            Caption         =   "Ambiente de testes"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   13380
            TabIndex        =   95
            Top             =   960
            Width           =   1725
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cert."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   12840
            TabIndex        =   88
            Top             =   270
            Width           =   375
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74945
         TabIndex        =   40
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
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   13980
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmFaturamento_Prod_Serv_Nfe_NS.frx":ADEE2
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView listaProdutos 
         Height          =   6530
         Left            =   -74940
         TabIndex        =   17
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
      Begin VB.TextBox txtdatacancelamento 
         Height          =   405
         Left            =   1080
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   5580
         Width           =   1365
      End
      Begin VB.TextBox Txt_ID_cobranca 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5550
         TabIndex        =   49
         Text            =   "0"
         Top             =   2220
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmFaturamento_Prod_Serv_NFe_NS"
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
Dim NomeArquivo As String
Dim TBLISTA_Faturamento_NFe As ADODB.Recordset
Dim SerialCertificado As String

Dim CnpjNF As String
Dim DAPartilhaICMS As String
Dim TextoCancelamento As String

Dim objDom As MSXML2.DOMDocument50
Dim objLinha As IXMLDOMElement
Dim objCabecalho As IXMLDOMElement
Dim objNFe As IXMLDOMElement
Dim objinfNFe As IXMLDOMElement
Dim objversao As IXMLDOMElement
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
Dim objdetPag As IXMLDOMElement
Dim objInfAdic As IXMLDOMElement
Dim objExporta As IXMLDOMElement
Dim objCompra As IXMLDOMElement
Dim objIPINT As IXMLDOMElement

Dim ArquivoXMLEnvio As String

Sub Proc_XML_Formatar(Parent As IXMLDOMNode, Optional Level As Integer)
  Dim Node As IXMLDOMNode
  Dim Indent As IXMLDOMText

  If Not Parent.parentNode Is Nothing And Parent.childNodes.Length > 0 Then
    For Each Node In Parent.childNodes
      Set Indent = Node.ownerDocument.createTextNode(vbNewLine & String(Level, vbTab))

      If Node.nodeType = NODE_TEXT Then
        If Trim(Node.Text) = "" Then
          Parent.removeChild Node
        End If
      ElseIf Node.previousSibling Is Nothing Then
        Parent.InsertBefore Indent, Node
      ElseIf Node.previousSibling.nodeType <> NODE_TEXT Then
        Parent.InsertBefore Indent, Node
      End If
    Next Node
  End If

  If Parent.childNodes.Length > 0 Then
    For Each Node In Parent.childNodes
      If Node.nodeType <> NODE_TEXT Then Proc_XML_Formatar Node, Level + 1
    Next Node
  End If
End Sub

Public Sub ProcCriarPastaDanfe()
On Error GoTo tratar_erro
Dim Mes As String
Dim Ano As String
Dim NomeEmpresa As String

Mes = Month(Date)
Ano = Year(Date)
NomeEmpresa = Trim(frmFaturamento_Prod_Serv.txtEmpresa.Text)

Select Case Mes
Case 1: Mes = "JANEIRO"
Case 2: Mes = "FEVEREIRO"
Case 3: Mes = "MARÇO"
Case 4: Mes = "ABRIL"
Case 5: Mes = "MAIO"
Case 6: Mes = "JUNHO"
Case 7: Mes = "JULHO"
Case 8: Mes = "AGOSTO"
Case 9: Mes = "SETEMBRO"
Case 10: Mes = "OUTUBRO"
Case 11: Mes = "NOVEMBRO"
Case 12: Mes = "DEZEMBRO"
End Select

DiretorioDanfe = DiretorioXMLDanfe & NomeEmpresa & "\" & Ano & "\" & Mes & "\DANFE\"

If DS.FileOrDirExists(DiretorioDanfe) = False Then

If DS.FileOrDirExists(DiretorioXMLDanfe & NomeEmpresa) = False Then
DiretorioDanfe = DiretorioXMLDanfe & NomeEmpresa
MkDir DiretorioDanfe
End If


If DS.FileOrDirExists(DiretorioXMLDanfe & NomeEmpresa & "\" & Ano) = False Then
DiretorioDanfe = DiretorioXMLDanfe & NomeEmpresa & "\" & Ano
MkDir DiretorioDanfe
End If

If DS.FileOrDirExists(DiretorioXMLDanfe & NomeEmpresa & "\" & Ano & "\" & Mes) = False Then
DiretorioDanfe = DiretorioXMLDanfe & NomeEmpresa & "\" & Ano & "\" & Mes
MkDir DiretorioDanfe
End If

If DS.FileOrDirExists(DiretorioXMLDanfe & NomeEmpresa & "\" & Ano & "\" & Mes & "\DANFE\") = False Then
DiretorioDanfe = DiretorioXMLDanfe & NomeEmpresa & "\" & Ano & "\" & Mes & "\DANFE\"
MkDir DiretorioDanfe
End If

End If

DiretorioDanfe = DiretorioXMLDanfe & NomeEmpresa & "\" & Ano & "\" & Mes & "\DANFE\"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Screen.MousePointer = vbDefault
End Sub

Private Sub ProcCarregaEmailCliente()
On Error GoTo tratar_erro
'=======================================
'Se o destinatário for Cliente
'=======================================
    Set TBClientes = CreateObject("adodb.recordset")

    TBClientes.Open "Select Email from Clientes where IDcliente = " & TBproducao!Id_Int_Cliente & " and NomeRazao = '" & TBproducao!txt_Razao_Nome & "' and Enviar_NF = 'True'", Conexao, adOpenKeyset, adLockReadOnly
    If TBClientes.EOF = False Then
        EmailCliente = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
'============================================
'Buscar email dos contatos do Cliente
'============================================
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select NomeContato, Email from Clientes_Contatos where IDcliente = " & TBproducao!Id_Int_Cliente & TextoFiltro & " and Enviar_NFe = 'True' and EMail is not null", Conexao, adOpenKeyset, adLockReadOnly
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    If IsNull(TBFI!Email) = False And TBFI!Email <> "" Then
                        EmailCliente = EmailCliente & ", " & TBFI!Email
                    End If
                    TBFI.MoveNext
                Loop
            End If
            TBFI.Close
    Else
'=======================================
'Se o destinatário for Fornecedor
'=======================================
        TBClientes.Close
        Set TBFornecedores = CreateObject("adodb.recordset")
        TBFornecedores.Open "Select Email, Pais, Codigo_pais from Compras_fornecedores where IDcliente = " & TBproducao!Id_Int_Cliente & " and Nome_Razao = '" & TBproducao!txt_Razao_Nome & "' and Enviar_NF = 'True'", Conexao, adOpenKeyset, adLockReadOnly
        If TBFornecedores.EOF = False Then
            EmailCliente = IIf(IsNull(TBFornecedores!Email), "", TBFornecedores!Email)
'============================================
'Buscar email dos contatos do Fornecedor
'============================================
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Email from Contatos_fornecedor where IdFornecedor = " & TBproducao!Id_Int_Cliente & " and Enviar_NFe = 'True' and Email is not null", Conexao, adOpenKeyset, adLockReadOnly
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    If IsNull(TBFI!Email) = False And TBFI!Email <> "" Then
                       EmailCliente = EmailCliente & "," & TBFI!Email
                    End If
                    TBFI.MoveNext
                Loop
            End If
            TBFI.Close
        End If
End If

'TBFornecedores.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaEmailTransportadora()
On Error GoTo tratar_erro


'Verifica se tem transportadora na NF para consultar o e-mail
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select IdIntTransp, txt_Razao from tbl_Dados_Transp where ID_nota = " & frmFaturamento_Prod_Serv_NFe_NS.txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
    If TBFIltro.EOF = False Then
        'Verifica se a transportadora é o Cliente
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select Email, IDCliente from Clientes where IDcliente = " & TBFIltro!IdIntTransp & " and NomeRazao = '" & TBFIltro!txt_Razao & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBClientes.EOF = False Then
            EmailTransportadora = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
        'Busca contatos do cliente
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select nomecontato, Email from Clientes_Contatos where IDcliente = " & TBClientes!IDCliente & "   and Enviar_NFe = 'True' and EMail is not null", Conexao, adOpenKeyset, adLockReadOnly
            If TBFI.EOF = False Then
                Do While TBFI.EOF = False
                    If IsNull(TBFI!Email) = False And TBFI!Email <> "" Then
                        EmailTransportadora = EmailTransportadora & ", " & TBFI!Email
                    End If
                    TBFI.MoveNext
                Loop
            End If
            TBFI.Close
            TBClientes.Close
        Else
        'Verifica se a transportadora é tipo Fornecedor
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select Email from Compras_fornecedores where IDcliente = " & TBFIltro!IdIntTransp & " and Nome_Razao = '" & TBFIltro!txt_Razao & "'", Conexao, adOpenKeyset, adLockReadOnly
            If TBClientes.EOF = False Then
                EmailTransportadora = IIf(IsNull(TBClientes!Email), "", TBClientes!Email)
        'Busca contatos do fornecedor
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select nomecontato, Email from Contatos_fornecedor where IdFornecedor = " & TBClientes!IDCliente & " and Enviar_NFe = 'True' and Email is not null", Conexao, adOpenKeyset, adLockReadOnly
                If TBFI.EOF = False Then
                    Do While TBFI.EOF = False
                        If IsNull(TBFI!Email) = False And TBFI!Email <> "" Then
                            EmailTransportadora = EmailTransportadora & ", " & TBFI!Email
                        End If
                        TBFI.MoveNext
                    Loop
                End If
                TBFI.Close
            End If
        End If
    End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaEmailUsuario()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Usuario,Email from Usuarios where Usuario = '" & pubUsuario & "'", Conexao, adOpenKeyset, adLockReadOnly
    If TBAbrir.EOF = False Then
       If TBAbrir!Email = "" Or IsNull(TBAbrir!Email) = True Then
          USMsgBox "Atenção " & pubUsuario & " você não tem email cadastrado para receber uma cópia!!!", vbCritical, "CAPRIND v5.0"
          Exit Sub
       End If
        EmailUsuario = IIf(IsNull(TBAbrir!Email) = False, TBAbrir!Email, "")
    End If
    TBAbrir.Close



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcCriarPastaXML()
On Error GoTo tratar_erro
Dim Mes As String
Dim Ano As String
Dim NomeEmpresa As String

Mes = Month(Date)
Ano = Year(Date)
NomeEmpresa = Trim(frmFaturamento_Prod_Serv.txtEmpresa.Text)

Select Case Mes
Case 1: Mes = "JANEIRO"
Case 2: Mes = "FEVEREIRO"
Case 3: Mes = "MARÇO"
Case 4: Mes = "ABRIL"
Case 5: Mes = "MAIO"
Case 6: Mes = "JUNHO"
Case 7: Mes = "JULHO"
Case 8: Mes = "AGOSTO"
Case 9: Mes = "SETEMBRO"
Case 10: Mes = "OUTUBRO"
Case 11: Mes = "NOVEMBRO"
Case 11: Mes = "DEZEMBRO"
End Select

DiretorioXML = DiretorioXMLDanfe & NomeEmpresa & "\" & Ano & "\" & Mes & "\XML\"

If DS.FileOrDirExists(DiretorioXML) = False Then

If DS.FileOrDirExists(DiretorioXMLDanfe & NomeEmpresa) = False Then
DiretorioXML = DiretorioXMLDanfe & NomeEmpresa
MkDir DiretorioXML
End If


If DS.FileOrDirExists(DiretorioXMLDanfe & NomeEmpresa & "\" & Ano) = False Then
DiretorioXML = DiretorioXMLDanfe & NomeEmpresa & "\" & Ano
MkDir DiretorioXML
End If

If DS.FileOrDirExists(DiretorioXMLDanfe & NomeEmpresa & "\" & Ano & "\" & Mes) = False Then
DiretorioXML = DiretorioXMLDanfe & NomeEmpresa & "\" & Ano & "\" & Mes
MkDir DiretorioXML
End If

If DS.FileOrDirExists(DiretorioXMLDanfe & NomeEmpresa & "\" & Ano & "\" & Mes & "\XML\") = False Then
DiretorioXML = DiretorioXMLDanfe & NomeEmpresa & "\" & Ano & "\" & Mes & "\XML\"
MkDir DiretorioXML
End If

End If

DiretorioXML = DiretorioXMLDanfe & NomeEmpresa & "\" & Ano & "\" & Mes & "\XML\"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Screen.MousePointer = vbDefault
End Sub

Public Sub procEnviar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If txtID_nota = 0 Then
    USMsgBox ("Informe a nota fiscal antes de enviar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select status from tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & txtID_nota & " AND status IN (100,101)", Conexao, adOpenKeyset, adLockReadOnly
If TBproducao.EOF = False Then
    USMsgBox ("Não é permitido enviar, pois a mesma já foi enviada."), vbCritical, "CAPRIND v5.0"
    TBproducao.Close
    Exit Sub
End If
TBproducao.Close

Acao = "enviar"
If funVerificaMigrate = False Then Exit Sub
If funVerificacaoEnviar = False Then Exit Sub

If USMsgBox("Deseja realmente enviar a nota fiscal n° " & txtNota & vbCrLf & " para o SEFAZ?", vbYesNo, "CAPRIND v5.0") = vbYes Then
  nfDocumento = "NF" & txtNota.Text
    Set TBproducao = CreateObject("adodb.recordset")
    TBproducao.Open "Select NF.*, T.*, E.Simples, E.Simples1, E.Cultural, E.CNPJ, E.CNAE, E.Razao, E.Empresa, E.IM, E.ie, E.Tipo_endereco, E.Endereco, E.Numero as numeroEmpresa, E.Complemento, E.Tipo_bairro, E.Bairro, E.Cidade, E.UF, E.CEP, E.Telefone, E.Email, NFE.Consumidor_final, NFE.Presenca_comprador, NFE.Forma_emissao, NFE.Finalidade_emissao, NFE.Enviar_Email, NFE.Forma_pagamento, NFE.FormaPagto, NFE.DA_entrega, NFE.DA_cobranca, NFE.ID_entrega, NFE.ID_Cobranca, NFE.xPag from (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota T ON NF.ID = T.ID_nota) INNER JOIN tbl_Dados_Nota_Fiscal_NFe NFE ON NFE.ID_Nota = NF.ID INNER JOIN Empresa E ON NF.ID_empresa = E.Codigo WHERE NF.ID = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
     If TBproducao.EOF = False Then
       NomeArquivo = "NF" & txtNota & txtSerie
       procMontaEmail 'Monta email pra envio da DANFe e do XML
       proc_XML_Criar 'Cria todo o XML
       proc_XML_Assinar 'Assina o XML
       ProcEnviarNotaSefaz 'Envia a nota fiscal para o SEFAZ
     End If
    TBproducao.Close
End If

Screen.MousePointer = vbDefault
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Proc_XML_Cabecalho()
On Error GoTo tratar_erro
'Cria o documento XML
Dim Atributo As String
Dim valor As String

Atributo = "xml"
valor = "version='1.0', encoding='UTF-8'"
Set objDom = New MSXML2.DOMDocument50

'==============================================
Dim dom, Node, PCMS
  objDom.async = False
  objDom.validateOnParse = True
  objDom.resolveExternals = False
  objDom.preserveWhiteSpace = False
'=============================================
'Cabecalho do xml
'=============================================
Set Cabecalho = objDom.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
objDom.appendChild Cabecalho
Set Cabecalho = Nothing
'Debug.print objDom.XML

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Proc_XML_InfNFE()
On Error GoTo tratar_erro

'Cria um nó filho chamado NFe dentro do documento
Set objNFe = objDom.createNode(1, "NFe", "http://www.portalfiscal.inf.br/nfe")
objDom.appendChild objNFe
'============================
'Cria o objeto infNFe
Set objinfNFe = objDom.createElement("infNFe")
objinfNFe.setAttribute "versao", "4.00"
objNFe.appendChild objinfNFe
'==============================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Proc_XML_Autorizado()
On Error GoTo tratar_erro

'Autorização para baixar o XML, hoje esta só a tranportadora
    Set TBTransporte = CreateObject("adodb.recordset")
    TBTransporte.Open "Select CF.Pessoa, DT.txt_CNPJ from tbl_Dados_Transp DT INNER JOIN Compras_fornecedores CF ON CF.IDCliente = DT.IdIntTransp and CF.Nome_Razao = DT.txt_Razao where DT.ID_Nota = " & txtID_nota & " and DT.txt_CNPJ IS NOT NULL and DT.txt_CNPJ <> N'' and DT.enviarXML = 1 AND txt_CNPJ <> '" & TBproducao!txt_CNPJ_CPF & "'", Conexao, adOpenKeyset, adLockReadOnly
    If TBTransporte.EOF = False Then
        'nó autXML dentro de InfNFe (A01)
        Set objAutXML = objDom.createElement("autXML")
        'objinfNFe.appendChild objEmit
        objinfNFe.appendChild objAutXML
        'Abre autXML=================================================================================================
            
            If Left(TBTransporte!Pessoa, 1) = "J" Then
            objAutXML.appendChild objDom.createElement("CNPJ") '0
            objAutXML.childNodes(0).Text = ReturnNumbersOnly(TBTransporte!txt_CNPJ)
            End If
            
            If Left(TBTransporte!Pessoa, 1) = "F" Then
            objAutXML.appendChild objDom.createElement("CPF") '1
            objAutXML.childNodes(0).Text = ReturnNumbersOnly(TBTransporte!txt_CNPJ)
            End If
            
'            If Left(TBTransporte!Pessoa, 1) = "J" Then objAutXML.childNodes(0).Text = ReturnNumbersOnly(TBTransporte!txt_CNPJ) Else objAutXML.childNodes(1).Text = ReturnNumbersOnly(TBTransporte!txt_CNPJ)
        'Fecha autXML================================================================================================
    End If
    TBTransporte.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Proc_XML_FormaPag()
On Error GoTo tratar_erro

If NFCe = False Then
    'Forma de pagamento Novo layout da Sefaz (4.0)
'    Set TBContas = CreateObject("adodb.recordset")
'    TBContas.Open "Select * from tbl_Detalhes_Recebimento where ID_nota = " & txtID_nota & " order by ID", Conexao, adOpenKeyset, adLockReadOnly
'    If TBContas.EOF = False Then
        'no cobr (Z01) dentro de enviar (A01)
        Set objCobr = objDom.createElement("cobr")
        objinfNFe.appendChild objCobr
        'Abre cobr==================================================================================================
            'no fat (Z02) dentro de Cobr (Z01)
            Set objFat = objDom.createElement("fat")
            objCobr.appendChild objFat
            'Abre Fat==================================================================================================
                objFat.appendChild objDom.createElement("nFat") '0
                objFat.childNodes(0).Text = txtNota
                objFat.appendChild objDom.createElement("vOrig") '1
                TotalCreditar = IIf(IsNull(TBproducao!Valor_total_receber_pagar), 0, TBproducao!Valor_total_receber_pagar) + IIf(IsNull(TBproducao!Valor_total_desconto), 0, TBproducao!Valor_total_desconto)
                objFat.childNodes(1).Text = Replace(IIf(IsNull(TBproducao!Valor_total_receber_pagar), "0.00", Format(TotalCreditar, "0.#0")), ",", ".")
                objFat.appendChild objDom.createElement("vDesc") '2
                If IsNull(TBproducao!Valor_total_desconto) = False And TBproducao!Valor_total_desconto > 0 Then
                    objFat.childNodes(2).Text = Replace(Format(TBproducao!Valor_total_desconto, "0.#0"), ",", ".")
                Else
                    objFat.childNodes(2).Text = "0.00"
                End If
                objFat.appendChild objDom.createElement("vLiq") '3
                objFat.childNodes(3).Text = Replace(IIf(IsNull(TBproducao!Valor_total_receber_pagar), "0.00", Format(TBproducao!Valor_total_receber_pagar, "0.#0")), ",", ".")
            'Fecha Fat=================================================================================================
    
'    'Forma de pagamento Novo layout da Sefaz (4.0)
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from tbl_Detalhes_Recebimento where ID_nota = " & txtID_nota & " order by ID", Conexao, adOpenKeyset, adLockReadOnly
    If TBContas.EOF = False Then
           
            'no dup (Za01) dentro de Cobr (Z01)
            'Abre Dup==================================================================================================
                Do While TBContas.EOF = False
                  Set objDup = objDom.createElement("dup")
                  objCobr.appendChild objDup
                
                    'no dupItem (Za02) dentro de Cobr (Za01)
                    'Abre DupItem==================================================================================================
                        objDup.appendChild objDom.createElement("nDup") '0
                        objDup.childNodes(0).Text = Left(TBContas!txt_Parcela, 3)
                        objDup.appendChild objDom.createElement("dVenc") '1
                        objDup.childNodes(1).Text = Format(TBContas!dt_Vencimento, "yyyy-mm-dd")
                        objDup.appendChild objDom.createElement("vDup") '2
                        objDup.childNodes(2).Text = Replace(Format(TBContas!dbl_Valor, "#0.#0"), ",", ".")
                    'Fecha DupItem=================================================================================================
                    TBContas.MoveNext
                Loop
            'Fecha Dup=================================================================================================
        'Fecha cobr=================================================================================================
    End If
    TBContas.Close
    
End If

    'no Pag (AA01) dentro de Enviar (A01)
    Set objPag = objDom.createElement("pag")
    objinfNFe.appendChild objPag
    'Abre Pag====================================================================================================
        'no detPag (AA02) dentro de Pag (AA01)
        Set objdetPag = objDom.createElement("detPag")
        objPag.appendChild objdetPag
        'Abre Pag====================================================================================================
        
        If NFCe = False Then
            If IsNull(TBproducao!Forma_pagamento) = False Then
                objdetPag.appendChild objDom.createElement("indPag") '0
                objdetPag.getElementsByTagName("indPag").Item(0).Text = TBproducao!Forma_pagamento
            End If
        End If
        
            objdetPag.appendChild objDom.createElement("tPag") '1
            objdetPag.getElementsByTagName("tPag").Item(0).Text = IIf(IsNull(TBproducao!FormaPagto), "15", TBproducao!FormaPagto)
            If TBproducao!FormaPagto <> 90 Then
            
                If NFCe = False Then
                    If TBproducao!FormaPagto = 99 Then ' Outros tem que informa a tag xPag
                        objdetPag.appendChild objDom.createElement("xPag") '2
                        objdetPag.getElementsByTagName("xPag").Item(0).Text = TBproducao!Xpag
                    End If
                End If
            
            
                objdetPag.appendChild objDom.createElement("vPag") '2
                objdetPag.getElementsByTagName("vPag").Item(0).Text = Replace(IIf(IsNull(TBproducao!dbl_Valor_Total_Nota), "00.00", Format(TBproducao!dbl_Valor_Total_Nota, "#0.#0")), ",", ".") '
                'Replace(IIf(IsNull(TBproducao!dbl_Valor_Total_Nota), 0, TBproducao!dbl_Valor_Total_Nota), ",", ".")
            Else
                objdetPag.appendChild objDom.createElement("vPag") '2
                objdetPag.getElementsByTagName("vPag").Item(0).Text = "0.00"
            End If
        'Fecha Pag===================================================================================================
    'Fecha Pag===================================================================================================
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Proc_XML_Exportacao()
On Error GoTo tratar_erro

If TBproducao!txt_UF = "EX" Then
    If TBproducao!int_TipoNota = 1 And (IsNull(TBproducao!txt_UF) = True Or TBproducao!txt_UF = "" Or TBproducao!txt_UF = "EX") Then
        Set TBTransporte = CreateObject("adodb.recordset")
        TBTransporte.Open "Select UF_embarque, Local_embarque from tbl_Dados_Transp where ID_Nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
        If TBTransporte.EOF = False Then
            'nó exporta (AD01) dentro de Enviar (A01)
            Set objExporta = objDom.createElement("exporta")
            objinfNFe.appendChild objExporta
            'Abre exporta====================================================================================================
                If IsNull(TBTransporte!UF_embarque) = False Then
                    objExporta.appendChild objDom.createElement("UFSaidaPais")
                    objExporta.getElementsByTagName("UFSaidaPais").Item(0).Text = TBTransporte!UF_embarque
                End If
                If IsNull(TBTransporte!Local_embarque) = False Then
                    objExporta.appendChild objDom.createElement("xLocExporta")
                    objExporta.getElementsByTagName("xLocExporta").Item(0).Text = TBTransporte!Local_embarque
                End If
            'Fecha exporta===================================================================================================
        End If
        TBTransporte.Close
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub proc_XML_Criar()
On Error GoTo tratar_erro
Dim Texto As String
Dim TextoAssinaturaA3 As String

ttvICMSUFDest = 0
vICMSUFDest = 0
'VarST = False

'===============================================================
' VERIFICA O REGIME TRIBUTÁRIO DO EMITENTE DA NFE
'===============================================================
ProcVerificaRegime
'===============================================================
' MONTA O XML DA NFE
'===============================================================
Proc_XML_Cabecalho
Proc_XML_InfNFE
proc_XML_Identificacao
proc_XML_Emitente

If TBproducao!Id_Int_Cliente <> 0 Then
proc_XML_Destinatario
End If

Proc_XML_Autorizado
proc_XML_Produtos
proc_XML_Totais
proc_XML_Transporte
Proc_XML_FormaPag

If NFCe = False Then
    proc_XML_Adicionais
    Proc_XML_Exportacao
End If

'================================================================
' Fim do arquivo XML
'================================================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub proc_XML_Assinar()
On Error GoTo tratar_erro
'===============================================================
' Tira todos os espacos vazios e efetua a identacao do XML
'    Proc_XML_Formatar objDom
'===============================================================
' Salva o arquivo XML
    objDom.Save (DiretorioEnvio & NomeArquivo & ".xml")
'================================================================
    ArquivoXMLEnvio = DiretorioEnvio & NomeArquivo & ".xml"
'Não apagar nada abaixo em hipotese nehuma
    Texto = "infNFe xmlns="""""
    Texto_Envio = objDom.XML
    Texto_Envio = Replace(Texto_Envio, Texto, "infNFe")

'===============================================================
'CRIAR CHAVE DE ACESSO NFE CERTIFICADO A1 e A3
'Código da UF + Data da emissão + CNPJ do Emitente + Modelo + Série + Número da NFe + Código Numérico + Dígito Verificador
'===============================================================
' Cria um código aleatorio com 8 digitos
'===============================================================
 chCodNumerico = Right(chNNfe, 8)
'===============================================================
' verifica o comprimento da serie da NFe
'===============================================================
  Select Case Len(chSerie)
   Case 1: chSerie = "00" & chSerie
   Case 2: chSerie = "0" & chSerie
  End Select
'===============================================================
' Monta a chave de acesso
'===============================================================
If txtchNFe.Text = "" Then
   chChave = chCodUF & chDTEmissao & chCNPJ & chModelo & chSerie & chNNfe & chFormaEmissao & Var3
   chdVer = CalculaDV(chChave)
   chChave = chChave & chdVer
   txtchNFe.Text = chChave
 End If
 If Len(txtchNFe.Text) = 44 Then
   chChave = txtchNFe.Text
   chdVer = Right(txtchNFe, 1)
 End If
 
'===============================================================
' Aqui começa o teste
' Gravar chave de acesso
'===============================================================
 Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set chave_acesso = '" & chChave & "' where id_nota = " & txtID_nota
 txtchNFe.Text = chChave

 Texto_Envio = Replace(Texto_Envio, "<infNFe versao=""4.00"">", "<infNFe versao=""4.00"" Id=""NFe" & chChave & """>")
'================================================================
' SE O CERTIFICADO FOR A3 MODIFICA O TEXTO DO XML PARA ASSINAR
'================================================================
If txtTPCertificado.Text = "A3" Then
'============================================================
' AQUI COMEÇA A ASSINATURA DO XML COM VB6
'==============================================================
If chk_Cp.Value = 1 Then
'==============================================================
' Aqui retira a TAG </NFe> do texto para envio
'==============================================================
   Texto_Envio = Replace(Texto_Envio, "</NFe>", "")
'==============================================================
' Aqui adiciona as tags para assinatura do XML
'==============================================================
   Texto_Envio = Replace(Texto_Envio, "<?xml version=""1.0""?>", "<?xml version=""1.0"" encoding=""UTF-8""?>")
   Texto_Envio = Replace(Texto_Envio, "<infNFe versao=""4.00"">", "<infNFe versao=""4.00"" Id=""NFe" & chChave & """>")
   Texto_Envio = Replace(Texto_Envio, "<tpAmb>" & tpAmb & "</tpAmb>", "<cDV>" & chdVer & "</cDV><tpAmb>" & tpAmb & "</tpAmb>")
   Texto_Envio = Replace(Texto_Envio, "<cNF>" & Right(txtNota, 8) & "</cNF>", "<cNF>" & chChave & "</cNF>")
   Texto_Envio = Texto_Envio & Trim("       <Signature xmlns=""http://www.w3.org/2000/09/xmldsig#"">")
   Texto_Envio = Texto_Envio & Trim("           <SignedInfo>")
   Texto_Envio = Texto_Envio & Trim("               <CanonicalizationMethod Algorithm=""http://www.w3.org/TR/2001/REC-xml-c14n-20010315""/>")
   Texto_Envio = Texto_Envio & Trim("               <SignatureMethod Algorithm=""http://www.w3.org/2000/09/xmldsig#rsa-sha1""/>")
   Texto_Envio = Texto_Envio & Trim("               <Reference URI=""#NFe" & chChave & """>")
   Texto_Envio = Texto_Envio & Trim("                   <Transforms>")
   Texto_Envio = Texto_Envio & Trim("                       <Transform Algorithm=""http://www.w3.org/2000/09/xmldsig#enveloped-signature""/>")
   Texto_Envio = Texto_Envio & Trim("                       <Transform Algorithm=""http://www.w3.org/TR/2001/REC-xml-c14n-20010315""/>")
   Texto_Envio = Texto_Envio & Trim("                   </Transforms>")
   Texto_Envio = Texto_Envio & Trim("                   <DigestMethod Algorithm=""http://www.w3.org/2000/09/xmldsig#sha1""/>")
   Texto_Envio = Texto_Envio & Trim("                   <DigestValue></DigestValue>")
   Texto_Envio = Texto_Envio & Trim("               </Reference>")
   Texto_Envio = Texto_Envio & Trim("           </SignedInfo>")
   Texto_Envio = Texto_Envio & Trim("           <SignatureValue></SignatureValue>")
   Texto_Envio = Texto_Envio & Trim("           <KeyInfo>")
   Texto_Envio = Texto_Envio & Trim("               <X509Data>")
   Texto_Envio = Texto_Envio & Trim("                   <X509Certificate></X509Certificate>")
   Texto_Envio = Texto_Envio & Trim("               </X509Data>")
   Texto_Envio = Texto_Envio & Trim("           </KeyInfo>")
   Texto_Envio = Texto_Envio & Trim("       </Signature>")
   Texto_Envio = Texto_Envio & Trim("   </NFe>")
  'Proc_XML_Formatar Texto_Envio
  Texto_Envio = Assina(Texto_Envio, txtSerialCertificado.Text)
Else
  '===================================================================
  ' AQUI COMEÇA A ASSINATURA DO XML COM CERTIFICADO A3 E A DLL .net  =
  ' NÃO APAGAR OU COMENTAR AS LINHAS ABAIXO                          =
  ' EM HIPÓTESE NENHUMA POIS ASSINA XML COM CERTIFICADO A3           =
  '===================================================================
  Texto_Envio = Replace(Texto_Envio, "<?xml version=""1.0""?>", "<?xml version=""1.0"" encoding=""UTF-8""?>")
  Texto_Envio = Replace(Texto_Envio, "<infNFe versao=""4.00"">", "<infNFe versao=""4.00"" Id=""NFe" & chChave & """>")
  Texto_Envio = Replace(Texto_Envio, "<tpAmb>" & tpAmb & "</tpAmb>", "<cDV>" & chdVer & "</cDV><tpAmb>" & tpAmb & "</tpAmb>")
  Texto_Envio = Replace(Texto_Envio, "<cNF>" & Right(txtNota, 8) & "</cNF>", "<cNF>" & chChave & "</cNF>")
  'Debug.print Texto_Envio
  
  Dim AssinaturaXML2 As New AssinaturaXML2.Principal
  Dim retorno As String
  
  retorno = AssinaturaXML2.assinarXML(Texto_Envio, "infNFe", chCNPJ)
  
If retorno = "Certificado Digital não encontrado" Then
  retorno = AssinaturaXML2.assinarXML(Texto_Envio, "infNFe", txtSerialCertificado)
End If

  Texto_Envio = retorno
End If

If Texto_Envio = "Certificado Digital não encontrado" Or Texto_Envio = "" Then
  USMsgBox Texto_Envio & " para o CNPJ " & chCNPJ & " tipo A3, por favor verifique seu certificado e tente de novo!", vbCritical, "CAPRIND v5.0"
  Permitido = False
  Exit Sub
End If

objDom.loadXML (Texto_Envio)
'===============================================================
' Tira todos os espacos vazios e efetua a identacao do XML
'    Proc_XML_Formatar objDom
'===============================================================
objDom.Save (DiretorioEnvio & NomeArquivo & ".xml")

txtRetorno.Text = objDom.XML

'Usmsgbox "XML assinado com sucesso", vbInformation, "CAPRIND v5.0"
File1.Refresh

End If

Permitido = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcEnviarNotaSefaz()
Dim XMLretorno As String
Dim Email_Enviado As Boolean

Email_Enviado = False

'===============================================
' BUSCA CNPJ DO EMITENTE DA NOTA
'===============================================
Screen.MousePointer = vbHourglass

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from empresa where codigo = '" & IDempresa & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
CnpjNF = ReturnNumbersOnly(TBAbrir!CNPJ)
End If
TBAbrir.Close

'Debug.print Texto_Envio
ProcCriarPastaDanfe
ProcCriarPastaXML

'Faz a emissão síncrona


resposta = emitirNFeSincrono(Texto_Envio, "xml", CnpjNF, "P", tpAmb, DiretorioDanfe, True)
'Debug.print resposta
  
    Mensagem = " Status de envio:" & statusEnvio & vbCrLf
    Mensagem = Mensagem & " cStat:" & cStat & vbCrLf
    Mensagem = Mensagem & " Status da consulta:" & motivo & vbCrLf
    Mensagem = Mensagem & " Motivo:" & xMotivo & vbCrLf
    
    txtRetorno.Text = Mensagem

    'Agora que você já leu os dados, é aconselhável que faça o salvamento de todos
    'eles no seu banco de dados antes de prosseguir para o teste abaixo
'================================================
' SALVAR DADOS DA NFE EMITIDA NO BANCO DE DADOS =
'===========================================================================================
' Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Status = " & cStat & ", chave_acesso = '" & chNFe & "', nsNRec ='" & nsNRec & "' , nProt ='" & nProt & "'  where id_nota = " & txtID_nota
' txtchNFe.Text = chNFe
' txtnsNrec.Text = nsNRec
' txt_nProt.Text = nProt
              
            
    'Testa se houve sucesso na emissão
    If (statusEnvio = 200) Or (statusEnvio = -6) Or (statusEnvio = 100) Then
        'Testa se houve sucesso na consulta
        If (statusConsulta = 200 Or statusConsulta = 100) Then
            'Testa se a nota foi autorizada
            If (cStat = 100) Then
            
            
                'Aqui dentro você pode realizar procedimentos como desabilitar o botão de emitir, etc
                USMsgBox "ATENÇÃO!" & vbCrLf & (motivo) & " N° " & txtNota.Text & " pela SEFAZ", vbInformation, "CAPRIND v5.0"
                '================================================
                ' SALVAR DADOS DA NFE EMITIDA NO BANCO DE DADOS =
                '===========================================================================================
                   Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Status = " & cStat & ", chave_acesso = '" & chNFe & "', nsNRec ='" & nsNRec & "' , nProt ='" & nProt & "'  where id_nota = " & txtID_nota
                 '==============================================
                 ' Aqui Movimenta estoque
                 '==============================================
                 If ID_nota = 0 Then
                 ID_nota = Int(txtID_nota.Text)
                 End If
                 
                 If ID_empresa = 0 Then
                 ID_empresa = IDempresa
                 End If
                 
                 Set TBCodigoDesc = CreateObject("adodb.recordset")
                 TBCodigoDesc.Open "Select Codigo from Empresa where Codigo = " & ID_empresa & " and Baixa_Auto_Estoque_NF = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                 If TBCodigoDesc.EOF = False Then
                 If frmFaturamento_Prod_Serv.opt_Saida = True Then
                    BaixarEstoqueNF 'Baixa estoque com NFe
                  Else
                    EntrarEstoqueNF 'Entra estoque com NFe
                 End If
                 End If
                 '==============================================
                   txtchNFe.Text = chNFe
                   txtnsNrec.Text = nsNRec
                   txt_nProt.Text = nProt
                   txtRetorno.Text = txtRetorno & vbCrLf & motivo
                   
                   If NFCe = False Then
                    Baixar = downloadNFeAndSave(txtchNFe.Text, tpAmb, "X", DiretorioXML, False)
                    Baixar = downloadNFeAndSave(txtchNFe.Text, tpAmb, "P", DiretorioDanfe, True)
                   Else
                    Baixar = NFCe_downloadESalvar(txtchNFe.Text, tpAmb, DiretorioDanfe, True)
                   End If

                   With frmFaturamento_Prod_Serv
                   .Txt_chave_acesso = chNFe
                   '=================================================================================
                   ' AQUI ENVIA EMAIL E ATUALIZA BANCO DE DADOS
                   '=================================================================================
                   
                   If tpAmb = 2 Then
                   ProcCarregaEmailUsuario
                   EmailEnvioNFe = EmailUsuario
                   End If
                   
                   ProcCarregaListaNota (1)
                   frmFaturamento_Prod_Serv.ProcCarregaListaNota (1)
                    
                                    
                   If EmailEnvioNFe <> "" Then
                        testeEmail = enviarEmail(txtchNFe.Text, "true", "true", EmailEnvioNFe)
                     If EmailEnviado = True Then
                        Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Enviar_email = '2' where id_nota = " & txtID_nota
                        Conexao.Execute "Update tbl_dados_nota_fiscal Set Imprimir = '1' where id = " & txtID_nota
                        USMsgBox "A DANFE e o XML foram enviados para " & EmailEnvioNFe & " com sucesso!", vbInformation, "CAPRIND v5.0"
                     End If
                   Else
                        USMsgBox "A Danfe e o XML não serão enviados por email automaticamente, pois não existe cadastros válidos nos contatos do cliente.", vbCritical, "CAPRIND v5.0"
                   End If
                   
                   End With
                '===========================================================================================
                 
                'Testa se o download teve problemas
                If statusDownload <> "" Then
                If (statusDownload <> 200) Then
                  'usMsgbox (motivo)
                End If
                End If
                
            Else
                'Aqui você pode mostrar alguma solução para o parceiro ou exibir opção de editar a nota
              USMsgBox (motivo), vbInformation, "CAPRIND v5.0"
              txtRetorno.Text = txtRetorno & vbCrLf & motivo
            End If
        'Caso tenha dado erro na consulta
        Else
            'Aqui você pode mostrar uma mensagem ao usuário
          USMsgBox (motivo + Chr(13) + erros), vbInformation, "CAPRIND v5.0"
          txtRetorno.Text = txtRetorno & vbCrLf & motivo
        End If
    Else
        'Aqui você pode exibir para o usuário o erro que ocorreu no envio
      USMsgBox (motivo + Chr(13) + erros + Chr(13) + xMotivo), vbCritical, "CAPRIND v5.0"
    End If
    'txtRetorno.Text = TextoRetorno


Exit Sub
tratar_erro:
USMsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, "CAPRIND v5.0", titleCTeAPI
End Sub

Sub ProcCarregaListaProdutos()
On Error GoTo tratar_erro

ListaProdutos.ListItems.Clear
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select NFP.*, CF.IDIntClasse from tbl_Detalhes_Nota NFP LEFT JOIN tbl_ClassificacaoFiscal CF ON CF.Idclass = NFP.ID_CF where NFP.id_nota = " & txtID_nota.Text & " order by NFP.int_codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Contador = 0
    Do While TBProduto.EOF = False
        With ListaProdutos.ListItems
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
        End With
    Loop
End If
TBProduto.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

ListaNota.ListItems.Clear
TBLISTA_Faturamento_NFe.PageSize = IIf(txtNreg = "", 24, txtNreg)
TBLISTA_Faturamento_NFe.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Faturamento_NFe.PageSize
ContadorReg = 1

'PBLista.Min = 0
'PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Faturamento_NFe.RecordCount - IIf(Pagina > 1, (TBLISTA_Faturamento_NFe.PageSize * (Pagina - 1)), 0), TBLISTA_Faturamento_NFe.PageSize)
'PBLista.Value = 1
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
    'PBLista.Value = Contador
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnAssinarXML_Click()
On Error GoTo tratar_erro


If USMsgBox("Deseja realmente assinar o XML da NFe n°" & txtNota.Text & "?", vbYesNo, "CAPRIND V5.0") = vbNo Then
Exit Sub
End If
'Texto_Envio = txtRetorno.Text
proc_XML_Assinar
USMsgBox "XML assinado com sucesso", vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnCobranca_Click()
On Error GoTo tratar_erro

frmFaturamento_Enderecos.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnEntrega_Click()
On Error GoTo tratar_erro

frmFaturamento_Enderecos.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnEnviarXML_Click()
Dim retorno As String
Dim Email_Enviado As Boolean

If USMsgBox("Deseja realmente enviar a o XML da nota fiscal n°" & txtNota.Text & " para o SEFAZ?", vbYesNo, "CAPRIND v5.0") = vbNo Then
Exit Sub
End If

'===============================================
' BUSCA CNPJ DO EMITENTE DA NOTA
'===============================================
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from empresa where Empresa = '" & frmFaturamento_Prod_Serv.txtEmpresa.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
CnpjNF = ReturnNumbersOnly(TBAbrir!CNPJ)
End If
TBAbrir.Close
objDom.loadXML (txtRetorno.Text)
Texto_Envio = objDom.XML

'Debug.print Texto_Envio
'Faz a emissão síncrona
retorno = emitirNFeSincrono(Texto_Envio, "xml", CnpjNF, "XP", tpAmb, DiretorioXMLDanfe, True)
    
    'Abaixo, confira um exemplo de tratamento de retorno da função emitirNFeSincrono
    
    Dim statusEnvio, statusConsulta, statusDownload, cStat, chNFe, nProt, motivo, nsNRec, erros As String
    
    'Lê o statusEnvio
    statusEnvio = LerDadosJSON(retorno, "statusEnvio", "", "")
    'Lê o statusConsulta
    statusConsulta = LerDadosJSON(retorno, "statusConsulta", "", "")
    'Lê o statusDownload
    statusDownload = LerDadosJSON(retorno, "statusDownload", "", "")
    'Lê o cStat
    cStat = LerDadosJSON(retorno, "cStat", "", "")
    'Lê a chNFe
    chNFe = LerDadosJSON(retorno, "chNFe", "", "")
    'usMsgbox cStat
    'Lê o nProt
    nProt = LerDadosJSON(retorno, "nProt", "", "")
    'Lê o motivo
    motivo = LerDadosJSON(retorno, "motivo", "", "")
    'Lê o nsNRec
    nsNRec = LerDadosJSON(retorno, "nsNRec", "", "")
    'Lê os erros
    erros = LerDadosJSON(retorno, "erros", "", "")

    'Agora que você já leu os dados, é aconselhável que faça o salvamento de todos
    'eles no seu banco de dados antes de prosseguir para o teste abaixo
'================================================
' SALVAR DADOS DA NFE EMITIDA NO BANCO DE DADOS =
'===========================================================================================
' Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Status = " & cStat & ", chave_acesso = '" & chNFe & "', nsNRec ='" & nsNRec & "' , nProt ='" & nProt & "'  where id_nota = " & txtID_nota
' txtchNFe.Text = chNFe
' txtnsNrec.Text = nsNRec
' txt_nProt.Text = nProt
              
            
    'Testa se houve sucesso na emissão
    If (statusEnvio = 200) Or (statusEnvio = -6) Then
        'Testa se houve sucesso na consulta
        If (statusConsulta = 200) Then
            'Testa se a nota foi autorizada
            If (cStat = 100) Then
                'Aqui dentro você pode realizar procedimentos como desabilitar o botão de emitir, etc
              USMsgBox "ATENÇÃO!" & vbCrLf & (motivo) & " N° " & txtNota.Text & " pela SEFAZ", vbInformation, "CAPRIND v5.0"
                '================================================
                ' SALVAR DADOS DA NFE EMITIDA NO BANCO DE DADOS =
                '===========================================================================================
                   Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Status = " & cStat & ", chave_acesso = '" & chNFe & "', nsNRec ='" & nsNRec & "' , nProt ='" & nProt & "'  where id_nota = " & txtID_nota
                   txtchNFe.Text = chNFe
                   txtnsNrec.Text = nsNRec
                   txt_nProt.Text = nProt
                   txtRetorno.Text = motivo

                   With frmFaturamento_Prod_Serv
                   '.txt_nsNRec = nsNRec
                   .Txt_chave_acesso = chNFe
                   '=================================================================================
                   ' AQUI ENVIA EMAIL E ATUALIZA BANCO DE DADOS
                   '=================================================================================
                   
                   If tpAmb = 2 Then
                   Email = EmailUsuario
                   End If
                   
                    ProcCarregaListaNota (1)
                    frmFaturamento_Prod_Serv.ProcCarregaListaNota (1)
                   
                   'Email = Email & ", vendas@caprind.com.br"
                   testeEmail = enviarEmail(txtchNFe.Text, "true", "true", Email)
                   
                   If EmailEnviado = True Then Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Enviar_email = '2' where id_nota = " & txtID_nota
                   If EmailEnviado = True Then Conexao.Execute "Update tbl_dados_nota_fiscal Set Imprimir = '1' where id = " & txtID_nota
                   
                   End With
                '===========================================================================================
                 
                'Testa se o download teve problemas
                If (statusDownload <> 200) Then
                  'usMsgbox (motivo)
                End If
            Else
                'Aqui você pode mostrar alguma solução para o parceiro ou exibir opção de editar a nota
              USMsgBox (motivo), vbInformation, "CAPRIND v5.0"
              txtRetorno.Text = motivo
            End If
        'Caso tenha dado erro na consulta
        Else
            'Aqui você pode mostrar uma mensagem ao usuário
          USMsgBox (motivo + Chr(13) + erros), vbInformation, "CAPRIND v5.0"
          txtRetorno.Text = motivo
        End If
    Else
        'Aqui você pode exibir para o usuário o erro que ocorreu no envio
      USMsgBox (motivo + Chr(13) + erros), vbCritical, "CAPRIND v5.0"
      txtRetorno.Text = motivo
    End If
    txtRetorno.Text = TextoRetorno

Exit Sub
tratar_erro:
USMsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, "CAPRIND v5.0", titleCTeAPI
End Sub

Private Sub BTNnsnrec_Click()
On Error GoTo tratar_erro
Dim RespostaNSNrec As String
Dim p As Object
NomeArquivo = frmFaturamento_Prod_Serv_NFe_NS.txtNota
nfDocumento = "NREC" & NomeArquivo

txtRetorno = ""

If USMsgBox("Deseja realmente consultar recibo NS dessa nota?", vbYesNo, "CAPRIND v5.0") = vbYes Then
If txtchNFe <> "" And Len(txtchNFe.Text) = 44 Then
RespostaNSNrec = listarNSNRecs(txtchNFe)
txtRetorno.Text = RespostaNSNrec
status = LerDadosJSON(txtRetorno.Text, "status", "", "")
'Debug.print RespostaNSNrec
   If status = "200" Then
      Set p = JSON.parse(RespostaNSNrec)
      txtnsNrec.Text = p.Item("nsNRecs").Item(1).Item("nsNRec")
   Else
      USMsgBox txtRetorno.Text, vbCritical, "CAPRIND v5.0"
   End If

Else
USMsgBox "Para consultar o recibo NS é necessário a chave de acesso da nota com 44 digitos", vbInformation, "CAPRIND v5.0"
Exit Sub
End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnPrevia_Click()
On Error GoTo tratar_erro

     Dim retorno As String
    Dim i As Integer
    If (txtD2.Text <> "") And (txtRetorno.Text <> "") Then
        'retorno = previaNFe(txtRetorno.Text, "xml")
        retorno = previaNFeESalvar(txtRetorno.Text, "xml", txtD2.Text, True)
    Else
        MsgBox ("Todos campos necessarios devem ser preenchidos ['caminho', 'tipo de conteudo', 'conteudo']")
    End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnSalvar_Click()
On Error GoTo tratar_erro

ProcSalvar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub BtnValidadorXML_Click()
On Error GoTo tratar_erro

Dim iret As Long
iret = ShellExecute(Me.hWnd, vbNullString, "http://www.sefaz.rs.gov.br/NFE/NFE-VAL.aspx", vbNullString, "c:\", SW_SHOWNORMAL)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmbFormaPag_Click()
On Error GoTo tratar_erro

'If cmbFormaPag.Text = "99 - Outros" Then
'    frmFaturamento_Prod_Serv_FormaPagamento.Show 1
'End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdD1_Click()
On Error GoTo tratar_erro

  ShellExecute 0, "open", DiretorioEnvio, "", "", vbNormalFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdD2_Click()
On Error GoTo tratar_erro

  ShellExecute 0, "open", DiretorioXMLDanfe, "", "", vbNormalFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdD3_Click()
On Error GoTo tratar_erro

  ShellExecute 0, "open", txtD3.Text, "", "", vbNormalFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdD4_Click()
On Error GoTo tratar_erro

  ShellExecute 0, "open", txtD4.Text, "", "", vbNormalFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcBuscaValidadeCertificado()
On Error GoTo tratar_erro

Dim Stor As New Store
Dim Cert As Certificate
Dim Certs As New Certificates
Dim CForNext As Integer


Stor.Open

Certs.Clear
For CForNext = 1 To Stor.Certificates.Count
Certs.Add Stor.Certificates.Item(CForNext)
Next CForNext

For Each Cert In Certs
    If txtSerialCertificado.Text = Cert.SerialNumber Then
        txtValidade.Text = Cert.ValidToDate
    End If
Next


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Function ProcBuscaIDDest(IDnota As Double)
On Error GoTo tratar_erro

If IDnota <> 0 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select E.UF, DN.txt_UF from tbl_Dados_Nota_Fiscal DN inner join Empresa E on E.codigo = DN.id_Empresa where DN.ID = " & IDnota, Conexao, adOpenKeyset, adLockReadOnly
    If TBAbrir.EOF = False Then
        UF = TBAbrir!UF
        UF_Destinatario = TBAbrir!txt_UF
        If UF = UF_Destinatario Then
            idDest = 1
        ElseIf UF_Destinatario = "EX" Then
            idDest = 3
        Else
            idDest = 2
        End If
    End If
    TBAbrir.Close
End If


Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function
Private Sub File4_DblClick()
On Error GoTo tratar_erro

  ShellExecute 0, "open", App.Path & "\Log\" & File4.filename, "", "", vbNormalFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub


Private Sub Frame14_DragDrop(Source As Control, x As Single, Y As Single)
On Error GoTo tratar_erro

  ShellExecute 0, "open", DiretorioRetorno, "", "", vbNormalFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"

End Sub

Private Sub tnEnviarNFE_Click()
On Error GoTo tratar_erro

procEnviar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub



Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    Case 2: procEnviar
    Case 3: ProcImprimir
    Case 4:
    If NFCe = False Then
    'procConsultarNFE
    ProcStatusNFe
    Else
    'status = NFCe_consultarSituacao(txtchNFe, tpAmb)
    ProcStatusNFe
    End If
    
    Case 5: procEnviaEmailDanfeXML
    Case 6: ProcBaixarDFE
    Case 7: ProcCancelar
    Case 8: ProcInutilizarNFe
    Case 10: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcInutilizarNFe()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente inutilizar o numero " & txtNota.Text & " no SEFAZ?", vbYesNo, "CAPRIND v5.0") = vbYes Then
frmFaturamento_inutilizarNFe.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcBaixarDFE()
On Error GoTo tratar_erro

With FrmFaturamento_Prod_Serv_DFE
.txtchNFe = txtchNFe
.txtD2 = DiretorioXMLDanfe
.Show 1
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnCriarXML_Click()
On Error GoTo tratar_erro
Dim Texto As String
Dim TextoAssinaturaA3 As String
txtRetorno.Text = ""

    Set TBproducao = CreateObject("adodb.recordset")
    StrSql = "Select NF.*, T.*, E.Simples, E.Simples1, E.Cultural, E.CNPJ, E.CNAE, E.Razao, E.Empresa, E.IM, E.ie, E.Tipo_endereco, E.Endereco, E.Numero as numeroEmpresa, E.Complemento, E.Tipo_bairro, E.Bairro, E.Cidade, E.UF, E.CEP, E.Telefone, E.Email, NFE.Consumidor_final, NFE.Presenca_comprador, NFE.Forma_emissao, NFE.Finalidade_emissao, NFE.Enviar_Email, NFE.Forma_pagamento, NFE.FormaPagto, NFE.DA_entrega, NFE.DA_cobranca, NFE.ID_entrega, NFE.ID_Cobranca, NFE.xPag from (tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota T ON NF.ID = T.ID_nota) INNER JOIN tbl_Dados_Nota_Fiscal_NFe NFE ON NFE.ID_Nota = NF.ID INNER JOIN Empresa E ON NF.ID_empresa = E.Codigo WHERE NF.ID = " & txtID_nota
    'Debug.print StrSql
    
    TBproducao.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
    If TBproducao.EOF = False Then
     'Da nome ao arquivo
     NomeArquivo = "NF" & txtNota & txtSerie
     'Verifica o(s) email(s) para enviar Dafe e xml
     procMontaEmail
     'Cria o XML da nota
     proc_XML_Criar
End If

'==============================================================
    Texto = "infNFe xmlns="""""
    Texto_Envio = objDom.XML
    Texto_Envio = Replace(Texto_Envio, Texto, "infNFe")
    
'===========================================================
'===============================================================
'CRIAR CHAVE DE ACESSO NFE CERTIFICADO A1 e A3
'Código da UF + Data da emissão + CNPJ do Emitente + Modelo + Série + Número da NFe + Código Numérico + Dígito Verificador
'===============================================================
' Cria um código aleatorio com 8 digitos
'===============================================================
 chCodNumerico = Right(chNNfe, 8)
'===============================================================
' verifica o comprimento da serie da NFe
'===============================================================
  Select Case Len(chSerie)
   Case 1: chSerie = "00" & chSerie
   Case 2: chSerie = "0" & chSerie
  End Select
'===============================================================
' Monta a chave de acesso
'===============================================================
If txtchNFe.Text = "" Then
   chChave = chCodUF & chDTEmissao & chCNPJ & chModelo & chSerie & chNNfe & chFormaEmissao & Var3
   chdVer = CalculaDV(chChave)
   chChave = chChave & chdVer
   txtchNFe.Text = chChave
 End If
 If Len(txtchNFe.Text) = 44 Then
   chChave = txtchNFe.Text
   chdVer = Right(txtchNFe, 1)
 End If
 
'===============================================================
' Aqui começa o teste
' Gravar chave de acesso
'===============================================================
 Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set chave_acesso = '" & chChave & "' where id_nota = " & txtID_nota
 txtchNFe.Text = chChave
 
  Texto_Envio = Replace(Texto_Envio, "<infNFe versao=""4.00"">", "<infNFe versao=""4.00"" Id=""NFe" & chChave & """>")

 '==============================================================
  Texto_Envio = Replace(Texto_Envio, "<?xml version=""1.0""?>", "<?xml version=""1.0"" encoding=""UTF-8""?>")
  Texto_Envio = Replace(Texto_Envio, "<infNFe versao=""4.00"">", "<infNFe versao=""4.00"" Id=""NFe" & chChave & """>")
  Texto_Envio = Replace(Texto_Envio, "<tpAmb>" & tpAmb & "</tpAmb>", "<cDV>" & chdVer & "</cDV><tpAmb>" & tpAmb & "</tpAmb>")
  Texto_Envio = Replace(Texto_Envio, "<cNF>" & Right(txtNota, 8) & "</cNF>", "<cNF>" & chChave & "</cNF>")
'==============================================================
    objDom.loadXML (Texto_Envio)
'===============================================================
' Tira todos os espacos vazios e efetua a identacao do XML
'    Proc_XML_Formatar objDom
'===============================================================
' Salva o arquivo XML
    objDom.Save (DiretorioEnvio & "\" & NomeArquivo & ".xml")
    DocXML = (DiretorioEnvio & NomeArquivo & ".xml")
'================================================================
txtRetorno.Text = objDom.XML
'USMsgBox "XML criado com sucesso", vbInformation, "CAPRIND v5.0"
File1.Refresh

TBproducao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub chkTPAmb_Click()
On Error GoTo tratar_erro

If chkTPAmb.Value = 1 Then
tpAmb = "2"
USMsgBox "Ambiente de emissão de notas fiscais em HOMOLOGAÇÃO - TESTES", vbInformation, "CAPRIND v5.0"
Else
tpAmb = "1"
USMsgBox "Ambiente de emissão de notas fiscais em  PRODUÇÃO", vbInformation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub


Private Sub cmbEntrega_Click()
On Error GoTo tratar_erro

If cmbEntrega <> "" Then txtID_entrega = cmbEntrega.ItemData(cmbEntrega.ListIndex) Else txtID_entrega = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Cmb_cobranca_Click()
On Error GoTo tratar_erro
  
If Cmb_cobranca <> "" Then Txt_ID_cobranca = Cmb_cobranca.ItemData(Cmb_cobranca.ListIndex) Else Txt_ID_cobranca = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcExcluirProduto()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With ListaProdutos
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir estes dados do(s) produto(s)/item(ns) da NFe?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
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
    USMsgBox ("Informe o(s) produto(s)/item(ns) da nota fiscal antes de excluir estes dados."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Dados do(s) produto(s)/item(ns) da nota fiscal excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpacamposProdutos
    ProcCarregaListaProdutos
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdConsultar_Click()
On Error GoTo tratar_erro
Dim resposta As String
Dim status As String

If USMsgBox("Deseja consultar a chave de acesso da nota fiscal N°" & txtNota.Text & "?", vbYesNo, "CAPRIND v5.0") = vbNo Then
 Exit Sub
End If

  If IsInternetOnline = False Then
  USMsgBox "Terminal sem sinal de internet, a consulta não será realizada.", vbCritical, "CAPRIND v5.0"
  Exit Sub
  End If
  
  
  If txtchNFe.Text = "" Then
  USMsgBox "Chave de acesso não informada, a consulta não será realizada.", vbCritical, "CAPRIND v5.0"
  Exit Sub
  End If
  
  If CnpjNF = "" Then
  USMsgBox "CNPJ emitente não informado, a consulta não será realizada.", vbCritical, "CAPRIND v5.0"
  Exit Sub
  End If
  
  NomeArquivo = txtNota.Text
  nfDocumento = "CSIT" & NomeArquivo
  
  resposta = consultarSituacao(ReturnNumbersOnly(CnpjNF), txtchNFe.Text, tpAmb, "4.00")
  'resposta = consultarCadastroContribuinte(ReturnNumbersOnly(CnpjNF), "SP", "05272563000152", "CNPJ")
  status = LerDadosJSON(resposta, status, "", "")
  txtRetorno.Text = resposta
  USMsgBox resposta, vbInformation, "CAPRIND v5.0"

ProcCarregaListaNota (1)
frmFaturamento_Prod_Serv.ProcCarregaListaNota (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

If txtPagIr = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBLISTA_Faturamento_NFe.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Faturamento_NFe.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Faturamento_NFe.AbsolutePage = 1
ProcExibePagina (TBLISTA_Faturamento_NFe.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Faturamento_NFe.AbsolutePage = TBLISTA_Faturamento_NFe.PageCount
ProcExibePagina (TBLISTA_Faturamento_NFe.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub File1_DblClick()
On Error GoTo tratar_erro

  ShellExecute 0, "open", DiretorioEnvio & "\" & File1.filename, "", "", vbNormalFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub File2_DblClick()
On Error GoTo tratar_erro
    
  ShellExecute 0, "open", DiretorioXMLDanfe & File2.filename, "", "", vbNormalFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub File3_DblClick()
On Error GoTo tratar_erro

  ShellExecute 0, "open", DiretorioRetorno & File3.filename, "", "", vbNormalFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaEnderecos()
On Error GoTo tratar_erro

With frmFaturamento_Prod_Serv

    If .txtIDcliente <> "" And .txt_Razao <> "" Then
        'Verifica se é cliente ou fornecedor
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from Clientes where IDCliente = " & .txtIDcliente & " and NomeRazao = '" & .txt_Razao & "'", Conexao, adOpenKeyset, adLockReadOnly
            If TBFI.EOF = False Then
                Tipo = "C"
            Else
                Tipo = "F"
            End If
            TBFI.Close
        End If
        
'===================================================================
' Endereco de entrega do cliente ou do fornecedor
'===================================================================
        Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from clientes_entrega where idcliente = " & .txtIDcliente & " and Tipo = '" & Tipo & "'", Conexao, adOpenKeyset, adLockReadOnly
            If TBFI.EOF = False Then
                    If IsNull(TBFI!Tipo_endereco) = False And TBFI!Tipo_endereco <> "" Then
                        Endereco = TBFI!Tipo_endereco & ": " & IIf(IsNull(TBFI!endereco_entrega), "", TBFI!endereco_entrega)
                    Else
                        Endereco = IIf(IsNull(TBFI!endereco_entrega), "", TBFI!endereco_entrega)
                    End If
                    If IsNull(TBFI!Tipo_bairro) = False And TBFI!Tipo_bairro <> "" Then
                        Bairro = TBFI!Tipo_bairro & ": " & IIf(IsNull(TBFI!bairro_entrega), "", TBFI!bairro_entrega)
                    Else
                        Bairro = IIf(IsNull(TBFI!bairro_entrega), "", TBFI!bairro_entrega)
                    End If
                    Endereco2 = Endereco & " - " & IIf(IsNull(TBFI!Numero), "", TBFI!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBFI!cidade_entrega), "", TBFI!cidade_entrega) & " - " & IIf(IsNull(TBFI!uf_entrega), "", TBFI!uf_entrega) & " - " & IIf(IsNull(TBFI!cep_entrega), "", TBFI!cep_entrega)
                    txtEntrega = Endereco2
                   txtID_entrega = TBFI!identrega
            End If
            TBFI.Close
            
            
'======================================================================
'Endereco cobranca do cliente ou fornecedor
'======================================================================
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from clientes_Cobranca where idcliente = " & .txtIDcliente & " and Tipo = '" & Tipo & "'", Conexao, adOpenKeyset, adLockReadOnly
            If TBFI.EOF = False Then

                    If IsNull(TBFI!Tipo_endereco) = False And TBFI!Tipo_endereco <> "" Then
                        Endereco = TBFI!Tipo_endereco & ": " & IIf(IsNull(TBFI!endereco_Cobranca), "", TBFI!endereco_Cobranca)
                    Else
                        Endereco = IIf(IsNull(TBFI!endereco_Cobranca), "", TBFI!endereco_Cobranca)
                    End If
                    If IsNull(TBFI!Tipo_bairro) = False And TBFI!Tipo_bairro <> "" Then
                        Bairro = TBFI!Tipo_bairro & ": " & IIf(IsNull(TBFI!bairro_Cobranca), "", TBFI!bairro_Cobranca)
                    Else
                        Bairro = IIf(IsNull(TBFI!bairro_Cobranca), "", TBFI!bairro_Cobranca)
                    End If
                    Endereco2 = Endereco & " - " & IIf(IsNull(TBFI!Numero), "", TBFI!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBFI!cidade_Cobranca), "", TBFI!cidade_Cobranca) & " - " & IIf(IsNull(TBFI!uf_Cobranca), "", TBFI!uf_Cobranca) & " - " & IIf(IsNull(TBFI!cep_Cobranca), "", TBFI!cep_Cobranca)
                    txtID_cobranca.Text = TBFI!idCobranca
                    txtCobranca.Text = Endereco2

            End If
            TBFI.Close

End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
Dim DiretorioLog As String
DiretorioLog = App.Path & "\Log\"

ProcCarregaToolBar1 Me, 15195, 10, True
ProcCarregaToolBar2 Me, 15195, 6, True

IDempresa = frmFaturamento_Prod_Serv.txtIDEmpresa.Text

Chk_DA_cobranca.Value = xtpChecked
Chk_DA_entrega.Value = xtpChecked


Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from empresa where codigo = '" & IDempresa & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
CnpjNF = ReturnNumbersOnly(TBAbrir!CNPJ)
tpAmb = IIf(IsNull(TBAbrir!tpAmb) = True, "2", TBAbrir!tpAmb)
chkTPAmb.Value = IIf(tpAmb = "1", "0", "1")
SerialCertificado = IIf(TBAbrir!Certificadodigital <> "", TBAbrir!Certificadodigital, "A1")
txtSerialCertificado.Text = SerialCertificado
DiretorioEnvio = TBAbrir!Caminho_Nfe
End If
TBAbrir.Close

If Formulario = "Faturamento/Nota fiscal/Própria" Then
    Caption = "Administrativo - Faturamento - Nota fiscal - Própria - Dados da NFe"
ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
    Caption = "Administrativo - Faturamento - Nota fiscal - Terceiros - Dados da NFe"
ElseIf Formulario = "Estoque/Ordem de faturamento" Then
    Caption = "Estoque - Ordem de faturamento - Dados da NFe"
Else
    Caption = "Estoque - Nota fiscal - Dados da NFe"
End If




txtTPCertificado.Text = TPCertificado
txtD1.Text = DiretorioEnvio
txtD2.Text = DiretorioXMLDanfe
txtD3.Text = DiretorioRetorno

If DS.FileOrDirExists(DiretorioLog) = True Then
txtD4.Text = DiretorioLog
End If

ProcLimpaVariaveisPrincipais
ProcRemoveObjetosResize Me

ProcCarregaListaNota (1)
ProcLimpaCampos
ProcPuxaDados

With frmFaturamento_Prod_Serv
    If .txtId <> "" And .txtId <> "0" And .txtDtValidacao <> "" Then
        txtID_nota = .txtId
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

SSTab1.Tab = 0

If NFCe = False Then
    Cmb_presenca_comprador.Text = "0 - Não se aplica"
Else
    Cmb_presenca_comprador.Text = "1 - Operação presencial"
End If

ProcBuscaValidadeCertificado
If txtcStat.Text = "" Then
ProcCarregaEnderecos
End If
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

'Timer_status_NFe.Enabled = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Acao = "salvar"
If txtID_nota = 0 Then
    USMsgBox ("Informe a nota antes de salvar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select status from tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & txtID_nota & " AND status IN (100,101)", Conexao, adOpenKeyset, adLockReadOnly
If TBproducao.EOF = False Then
    USMsgBox ("Não é permitido salvar, pois a mesma já foi enviada."), vbExclamation, "CAPRIND v5.0"
    TBproducao.Close
    Exit Sub
End If

'Verifica se é NFSe
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * FROM tbl_Dados_Nota_Fiscal WHERE ID = " & txtID_nota & " and TipoNF = 'SA'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    USMsgBox ("Não é permitido salvar, pois esta é uma nota fiscal de serviço."), vbExclamation, "CAPRIND v5.0"
    TBGravar.Close
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * From tbl_Dados_Nota_Fiscal where id = " & txtID_nota & " and DtValidacao IS NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    USMsgBox ("Não é permitido salvar, pois esta nota fiscal ainda não foi validada."), vbExclamation, "CAPRIND v5.0"
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

If NFCe = False Then

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * FROM tbl_Dados_Nota_Fiscal where ID = " & txtID_nota & " and txt_tipocliente <> 'E'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Endereco_NF = IIf(IsNull(TBAbrir!txt_Endereco), "", TBAbrir!txt_Endereco) & " - " & IIf(IsNull(TBAbrir!Numero), "", TBAbrir!Numero) & " - " & IIf(IsNull(TBAbrir!txt_Bairro), "", TBAbrir!txt_Bairro) & " - " & IIf(IsNull(TBAbrir!txt_Municipio), "", TBAbrir!txt_Municipio) & " - " & IIf(IsNull(TBAbrir!txt_UF), "", TBAbrir!txt_UF) & " - " & IIf(IsNull(TBAbrir!Txt_CEP), "", TBAbrir!Txt_CEP)
    If txtEntrega = "" Then
        NomeCampo = "o endereço de entrega"
        ProcVerificaAcao
        cmbEntrega.SetFocus
        Exit Sub
    Else
        If txtEntrega <> Endereco_NF And Chk_DA_entrega.Value = xtpChecked Then
            If USMsgBox("O endereço de entrega é diferente do endereço principal, deseja prosseguir sem imprimir o endereço de entrega nos dados adicionais?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
        ElseIf txtEntrega = Endereco_NF And Chk_DA_entrega.Value = xtpChecked Then
            If USMsgBox("O endereço de entrega é igual o endereço principal, deseja prosseguir imprimindo o endereço de entrega nos dados adicionais?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
        End If
    End If
    If txtCobranca = "" Then
        NomeCampo = "o endereço de cobrança"
        ProcVerificaAcao
        txtCobranca.SetFocus
        Exit Sub
    Else
        If txtCobranca <> Endereco_NF And Chk_DA_cobranca.Value = xtpUnchecked Then
            If USMsgBox("O endereço de cobrança é diferente do endereço principal, deseja prosseguir sem imprimir o endereço de cobrança nos dados adicionais?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
        ElseIf txtCobranca = Endereco_NF And Chk_DA_cobranca.Value = xtpChecked Then
                If USMsgBox("O endereço de cobrança é igual o endereço principal, deseja prosseguir imprimindo o endereço de cobrança nos dados adicionais?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
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

End If

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * FROM tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    'usMsgbox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar dados da nota fiscal"
Else
    TBGravar.AddNew
    USMsgBox ("Novos dados da nota fiscal cadastrados com sucesso."), vbInformation, "CAPRIND v5.0"
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
    If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtId Else NomeCampo = "N° nota: " & .txtNFiscal
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
        If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtId Else NomeCampo = "N° nota: " & .txtNFiscal
        Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
    End With
    Documento1 = "Transportadora: " & TBProduto!txt_Razao
    ProcGravaEvento
    '==================================
    TBProduto.Update
End If
TBProduto.Close
ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
If CodigoLista <> 0 And ListaNota.ListItems.Count <> 0 Then
    ListaNota.SelectedItem = ListaNota.ListItems(CodigoLista)
    ListaNota.SetFocus
End If

If cmbFormaPag.Text = "99 - Outros" And txtID_nota <> "" Then
    frmFaturamento_Prod_Serv_FormaPagamento.Show 1
End If


USMsgBox "Dados gravados com sucesso!", vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

If cmbForma_pagamento <> "" Then TBGravar!Forma_pagamento = Left(cmbForma_pagamento, 1) Else TBGravar!Forma_pagamento = Null
TBGravar!FormaPagto = Left(cmbFormaPag, 2)
TBGravar!Forma_emissao = Left(Cmb_forma_de_emissao, 1)
TBGravar!Finalidade_emissao = Left(cmbFinalidade_emissao, 1)
TBGravar!Consumidor_final = Left(Cmb_consumidor, 1)
TBGravar!Presenca_comprador = Left(Cmb_presenca_comprador, 1)
TBGravar!idDest = Left(cmbOperacao.Text, 1)

If Len(txtchNFe.Text) = 44 Or txtchNFe.Text = "" Then
TBGravar!Chave_acesso = txtchNFe.Text
End If

If Len(txt_nProt.Text) = 15 Or txt_nProt.Text = "" Then
TBGravar!nProt = txt_nProt.Text
End If

If Len(txtnsNrec.Text) = 7 Or txtnsNrec.Text = "" Then
TBGravar!nsNRec = txtnsNrec.Text
End If

TBGravar!ID_entrega = txtID_entrega.Text
If Chk_DA_entrega.Value = xtpChecked Then TBGravar!DA_entrega = True Else TBGravar!DA_entrega = False

TBGravar!ID_Cobranca = txtID_cobranca.Text
If Chk_DA_cobranca.Value = xtpChecked Then TBGravar!DA_cobranca = True Else TBGravar!DA_cobranca = False

TBGravar!Enviar_DANFE_email = IIf(Cmb_enviar_DANFE = "", "S", Left(Cmb_enviar_DANFE, 1))

If chkCodRef.Value = 1 Then TBGravar!CodRef = True Else TBGravar!CodRef = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
'ProcBuscaIDDest (txtID_nota)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Cmb_forma_de_emissao = "1 - Normal"
cmbFinalidade_emissao = "1 - Normal"
cmbFormaPag.ListIndex = -1
Cmb_consumidor.ListIndex = -1
Cmb_presenca_comprador.ListIndex = -1
txtID_entrega = 0
txtEntrega.Text = ""
Txt_ID_cobranca = 0
txtCobranca.Text = ""
Cmb_enviar_DANFE = "Sim"
txtLocal_embarque = ""
cmbUF_embarque.ListIndex = -1
txtStatus = ""
Txt_chave_acesso = ""
chkCodRef.Value = 0
txtID_nota = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ListaProdutos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListaProdutos, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub listaProdutos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With ListaProdutos
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
                USMsgBox ("Não é permitido excluir os dados do produto desta nota fiscal, pois a mesma está " & NomeCampo & "."), vbExclamation, "CAPRIND v5.0"
                .ListItems.Item(InitFor).Checked = False
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ListaProdutos_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaProdutos.ListItems.Count = 0 Then Exit Sub
ProcLimpacamposProdutos
CodigoLista1 = ListaProdutos.SelectedItem.index
txtID_item = ListaProdutos.SelectedItem
If Len(ListaProdutos.SelectedItem.SubItems(3)) = 4 Then Quant = Right(ListaProdutos.SelectedItem.SubItems(3), 3) Else Quant = Right(ListaProdutos.SelectedItem.SubItems(3), 2)
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
        If IsNull(TBFI!Tipo_Produto) = False And TBFI!Tipo_Produto <> "" Then
            With Cmb_tipo_produto
                Select Case TBFI!Tipo_Produto
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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
        ListaProdutos.SetFocus
        ProcLimpacamposProdutos
        ProcCarregaListaProdutos
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

ProcBuscaIDDest (txtID_nota)

Set TBAbrir = CreateObject("adodb.recordset")
StrSql = "Select * from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & txtID_nota

TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
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
            Case "16": .Text = "16 - Depósito Bancário"
            Case "17": .Text = "17 - Pagamento instantâneo (PIX)"
            Case "18": .Text = "18 - Transferência bancária, Carteira Digital"
            Case "19": .Text = "19 - Programa de fidelidade, Cashback, Crédito Virtual"
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
    If TBAbrir!DA_entrega = False Then Chk_DA_entrega.Value = xtpUnchecked Else Chk_DA_cobranca.Value = xtpChecked
    Txt_ID_cobranca = IIf(IsNull(TBAbrir!ID_Cobranca), 0, TBAbrir!ID_Cobranca)
    If TBAbrir!DA_cobranca = False Then Chk_DA_cobranca.Value = xtpUnchecked Else Chk_DA_cobranca.Value = xtpChecked
    txtcStat.Text = FunVerifStatusNFe(TBAbrir!ID_nota)
    txtchNFe.Text = IIf(IsNull(TBAbrir!Chave_acesso), "", TBAbrir!Chave_acesso)
    txt_nProt = IIf(IsNull(TBAbrir!nProt), "", TBAbrir!nProt)
    txtnsNrec.Text = IIf(IsNull(TBAbrir!nsNRec), "", TBAbrir!nsNRec)
    
    If IsNull(TBAbrir!CodRef) = True Then
        Set TBAliquota = CreateObject("adodb.recordset")
        TBAliquota.Open "Select Codigo_ref_DANFE from empresa where codigo = " & IDempresa, Conexao, adOpenKeyset, adLockOptimistic
        If TBAliquota.EOF = False Then
            If TBAliquota!Codigo_ref_DANFE = False Or IsNull(TBAliquota!Codigo_ref_DANFE) = True Then chkCodRef.Value = 0 Else chkCodRef.Value = 1
        End If
        TBAliquota.Close
    Else
        If TBAbrir!CodRef = False Then chkCodRef.Value = 0 Else chkCodRef.Value = 1
    End If
    
    With cmbOperacao
        If IsNull(TBAbrir!idDest) = True Then
            Select Case idDest
                Case "1": .Text = "1 - Interna"
                Case "2": .Text = "2 - Interestadual"
                Case "3": .Text = "3 - Exportação"
            End Select
        Else
            Select Case TBAbrir!idDest
                Case "1": .Text = "1 - Interna"
                Case "2": .Text = "2 - Interestadual"
                Case "3": .Text = "3 - Exportação"
            End Select
        End If
    End With
    
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaEntrega()
On Error GoTo tratar_erro

        
'Busca do cadastro do cliente na nota fiscal o local de entrega
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
                txtEntrega.Text = Endereco
            End If
            TBClientes.Close
        End If
  '  End If
    'TBFIltro.Close
'End With

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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Sub ProcCarregaCobranca()
On Error GoTo tratar_erro


        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "Select ID_Cobranca from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & txtID_nota & " and ID_Cobranca IS NOT NULL", Conexao, adOpenKeyset, adLockReadOnly
        If TBAcessos.EOF = False Then
            Set TBClientes = CreateObject("adodb.recordset")
            TBClientes.Open "Select * from clientes_cobranca where idcobranca = " & TBAcessos!ID_Cobranca, Conexao, adOpenKeyset, adLockReadOnly
            If TBClientes.EOF = False Then
                Txt_ID_cobranca = TBClientes!idCobranca
                idCobranca = TBClientes!idCobranca
                
                If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                    Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!endereco_Cobranca), "", TBClientes!endereco_Cobranca)
                Else
                    Endereco = IIf(IsNull(TBClientes!endereco_Cobranca), "", TBClientes!endereco_Cobranca)
                End If
                If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then
                    Bairro = TBClientes!Tipo_bairro & ": " & IIf(IsNull(TBClientes!bairro_Cobranca), "", TBClientes!bairro_Cobranca)
                Else
                    Bairro = IIf(IsNull(TBClientes!bairro_Cobranca), "", TBClientes!bairro_Cobranca)
                End If
                Endereco = Endereco & " - " & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & " - " & Bairro & " - " & IIf(IsNull(TBClientes!cidade_Cobranca), "", TBClientes!cidade_Cobranca) & " - " & IIf(IsNull(TBClientes!uf_Cobranca), "", TBClientes!uf_Cobranca) & " - " & IIf(IsNull(TBClientes!cep_Cobranca), "", TBClientes!cep_Cobranca)
                txtCobranca.Text = Endereco
            End If
            TBClientes.Close
    End If
    
Exit Sub
tratar_erro:
    If Err.Number = 383 Then
        With Cmb_cobranca
            .AddItem Endereco
            .ItemData(Cmb_cobranca.NewIndex) = idCobranca
            .Text = Endereco
        End With
        Exit Sub
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab

Case 0:
Case 1:
txtD1.Text = DiretorioEnvio
File1.Path = DiretorioEnvio
txtD2.Text = DiretorioXMLDanfe
File2.Path = DiretorioXMLDanfe
File2.Refresh
txtD3.Text = DiretorioRetorno
File3.Path = DiretorioRetorno
File3.Refresh
txtD4.Text = App.Path & "\Log\"
File4.Path = App.Path & "\Log\"
File4.Refresh

If txtTPCertificado.Text = "A1" Then
 btnAssinarXML.Enabled = True
Else
 btnAssinarXML.Enabled = True
End If

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

Sub ProcSalvarProduto()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Acao = "salvar"
If txtID_item = 0 Then
    NomeCampo = "Produto"
    ProcVerificaAcao
    Exit Sub
End If

If FrameCST.Enabled = True Then
    If Len(ListaProdutos.SelectedItem.SubItems(3)) = 4 Then Quant = Right(ListaProdutos.SelectedItem.SubItems(3), 3) Else Quant = Right(ListaProdutos.SelectedItem.SubItems(3), 2)
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
    TBFI!Tipo_Produto = Left(Cmb_tipo_produto, 1)
    TBFI.Update
    TBFI.Close
End If

USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
'==================================
Modulo = Formulario
Evento = "Alterar dados do produto da nota fiscal"
ID_documento = ListaProdutos.SelectedItem
With frmFaturamento_Prod_Serv
    .ProcVerificaTipoNF False
    If .txtNFiscal = "" Then NomeCampo = "N° ordem: " & .txtId Else NomeCampo = "N° nota: " & .txtNFiscal
    Documento = NomeCampo & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
End With
Documento1 = "Cód. interno: " & ListaProdutos.SelectedItem.ListSubItems(1)
ProcGravaEvento
'==================================
If CodigoLista1 <> 0 And ListaProdutos.ListItems.Count <> 0 Then
    ListaProdutos.SelectedItem = ListaProdutos.ListItems(CodigoLista1)
    ListaProdutos.SetFocus
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Function funVerifLiberacao(Mensagem As Boolean) As Boolean
On Error GoTo tratar_erro

funVerifLiberacao = True

Familiatext = ""
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal_NFe where id_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = True Then
    If Mensagem = True Then USMsgBox ("Salve os dados da NF-e antes de liberar para envio."), vbExclamation, "CAPRIND v5.0"
    funVerifLiberacao = False
    Exit Function
Else
    If IsNull(TBAbrir!Forma_emissao) = True Or TBAbrir!Forma_emissao = "" Then
        If Mensagem = True Then USMsgBox ("Salve os dados da NF-e antes de liberar para envio."), vbExclamation, "CAPRIND v5.0"
        funVerifLiberacao = False
        Exit Function
    End If
    If TBAbrir!Finalidade_emissao = 4 Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select Id from tbl_dados_transp where id_nota = " & txtID_nota & " and txt_Frete_Conta = 9", Conexao, adOpenKeyset, adLockReadOnly
        If TBFI.EOF = False Then
            If Mensagem = True Then USMsgBox ("Frete inválido para o tipo de nota."), vbExclamation, "CAPRIND v5.0"
            funVerifLiberacao = False
            Exit Function
        Else
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Id from tbl_dados_transp where id_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
            If TBFI.EOF = True Then
                If Mensagem = True Then USMsgBox ("É necessário cadastrar a transportadora antes de liberar para envio."), vbExclamation, "CAPRIND v5.0"
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
If NFCe = False Then
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
    If TBAbrir!txt_Bairro = "" Or IsNull(TBAbrir!txt_Bairro) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "Bairro do destinatário da NF" Else Familiatext = "Bairro do destinatário da NF"
        funVerifLiberacao = False
    End If
    If (TBAbrir!Txt_CEP = "" Or IsNull(TBAbrir!Txt_CEP) = True) And TBAbrir!txt_UF <> "EX" Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "CEP do destinatário da NF" Else Familiatext = "CEP do destinatário da NF"
        funVerifLiberacao = False
    End If
    If TBAbrir!txt_UF = "" Or IsNull(TBAbrir!txt_UF) = True Then
        If Familiatext <> "" Then Familiatext = Familiatext & "; " & vbCrLf & "UF do destinatário da NF" Else Familiatext = "UF do destinatário da NF"
        funVerifLiberacao = False
    End If
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

If funVerifLiberacao = False And Mensagem = True Then USMsgBox ("Informe o(s) campo(s) antes de liberar a NF para envio: " & vbCrLf & Familiatext), vbInformation, "CAPRIND v5.0"

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function


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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub procEnviaEmailDanfeXML()
On Error GoTo tratar_erro
NomeArquivo = frmFaturamento_Prod_Serv_NFe_NS.txtNota
nfDocumento = "ENV" & NomeArquivo

If txtchNFe.Text = "" Then
USMsgBox "Chave de acesso da nota não cadastrada, por favor verificar", vbCritical, "CAPRIND v5.0"
Exit Sub
Else
frmFaturamento_EnviaEmail.Show
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procLogErros()
On Error GoTo tratar_erro
Dim retorno As String

If txtID_nota = 0 Then
    USMsgBox ("Informe a nota fiscal antes de consultar log de erros."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

retorno = downloadEventoNFe(txtchNFe, tpAmb, "xml", "canc", "1")

Acao = "verificar o log"
If funVerificaMigrate = False Then Exit Sub

Sit_REG = 1
frmFaturamento_Prod_Serv_NFSe_Log.Show 1
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Function funVerificaMigrate() As Boolean
On Error GoTo tratar_erro
funVerificaMigrate = False

'If ChaveMigrate = "" Then
'    NomeCampo = "a chave de acesso da Migrate no cadastro da empresa"
'    ProcVerificaAcao
'    Exit Function
'End If

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
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Public Sub proc_XML_LerRetorno()
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
cnpjXML = ReturnNumbersOnly(CnpjNF)
NotaXML = FunTamanhoTextoZeroEsq(txtNota, 9)
SerieXML = FunTamanhoTextoZeroEsq(txtSerie, 5)
formaXML = False

success = doc.Load(DiretorioRetorno & "\NFe\" & cnpjXML & NotaXML & SerieXML & IIf(TipoXML = 2, "11011101", "00000000") & "-ret.xml")
If success = False Then
    USMsgBox "Não foi possível obter retorno da Sefaz, favor consultar o log de erros.", vbExclamation, "CAPRIND v5.0"
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
    
    USMsgBox (NodeStatus.Text & " - " & NodeDescricao.Text & "."), vbInformation, "CAPRIND v5.0"
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
        If USMsgBox("Deseja visualizar a Danfe?", vbYesNo, "CAPRIND v5.0") = vbYes Then procAbrirNotaPDF "NFe", CnpjNF, txtNota, txtSerie, DiretorioXMLDanfe, False
    End If
End If

ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(lblPaginas.Caption, Len(lblPaginas.Caption) - 5))))
With frmFaturamento_Prod_Serv
    .ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
End With

Exit Sub
tratar_erro:
    If Err.Number = 91 Then
        USMsgBox "Não foi possível ter um retorno da Sefaz, favor tentar mais tarde.", vbExclamation, "CAPRIND v5.0"
    Else
        USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    End If
End Sub

Public Sub ProcStatusNFe()
On Error GoTo tratar_erro

Dim ResultadoNFe As String
Dim StatusNFe As String
Dim protocolo As String
NomeArquivo = txtNota
nfDocumento = "ST" & NomeArquivo

If txtchNFe.Text = "" Then
USMsgBox "Chave de acesso NFe não informada, a consulta não será realizada.", vbCritical, "CAPRIND v5.0"
Exit Sub
End If

If CnpjNF = "" Then
USMsgBox "CNPJ do emitente não informado, a consulta não será realizada.", vbCritical, "CAPRIND v5.0"
Exit Sub
End If


If USMsgBox("Deseja realmente consultar o status da NFe N° " & txtNota.Text & " na SEFAZ?", vbYesNo, "CAPRIND 5.0") = vbNo Then Exit Sub
  txtRetorno.Text = ""

'===============================================================
' Primeiro consulta pelo numero da chave de acesso
'===============================================================
 If txtchNFe.Text <> "" Then
 CnpjNF = ReturnNumbersOnly(CnpjNF)
 
 If NFCe = False Then
    StatusNFe = consultarSituacao(CnpjNF, txtchNFe.Text, tpAmb, "4.00")
    xMotivo = LerDadosJSON(StatusNFe, "retConsSitNFe", "xMotivo", "")
    cStat = LerDadosJSON(StatusNFe, "retConsSitNFe", "cStat", "")
 Else
    StatusNFe = NFCe_consultarSituacao(txtchNFe.Text, tpAmb)
    xMotivo = LerDadosJSON(StatusNFe, "retEvento", "xMotivo", "")
    
    If xMotivo <> "Evento registrado e vinculado a NF-e" Then
    xMotivo = LerDadosJSON(StatusNFe, "nfeProc", "xMotivo", "")
    End If
    cStat = LerDadosJSON(StatusNFe, "nfeProc", "cStat", "")
    nProt = LerDadosJSON(StatusNFe, "nfeProc", "nProt", "")
 End If
 
  If xMotivo = "Autorizado o uso da NF-e" Then
    Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Status = " & cStat & " where id_nota = " & txtID_nota
    Conexao.Execute "Update tbl_dados_nota_fiscal Set int_Status = '1'  where id = " & txtID_nota
  End If
  
If xMotivo = "Evento registrado e vinculado a NF-e" Then
    xMotivo = "Documento cancelado!"
    Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Status = '101' where id_nota = " & txtID_nota
    Conexao.Execute "Update tbl_dados_nota_fiscal Set int_Status = '2'  where id = " & txtID_nota
  End If

If xMotivo = "Cancelamento de NF-e homologado" Then
    'xMotivo = "Documento cancelado!"
    Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Status = '101' where id_nota = " & txtID_nota
    Conexao.Execute "Update tbl_dados_nota_fiscal Set int_Status = '2'  where id = " & txtID_nota
 End If

  
 USMsgBox xMotivo, vbInformation, "CAPRIND v5.0"
 Exit Sub
 End If
 
'===============================================================
' Depois consulta pelo numero NS Nrec
'===============================================================
If txtnsNrec.Text <> "" And cStat = "100" Then
'===============================================================
' Pega o numero do CNPJ do emissor
'===============================================================
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from empresa where Empresa = '" & frmFaturamento_Prod_Serv.txtEmpresa.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
CnpjNF = ReturnNumbersOnly(TBAbrir!CNPJ)
End If
TBAbrir.Close
CnpjNF = ReturnNumbersOnly(CnpjNF)
'===============================================================
' Envia dados para consulta
'===============================================================

If NFCe = False Then
ResultadoNFe = consultarStatusProcessamento(CnpjNF, txtnsNrec.Text, tpAmb)
Else
ResultadoNFe = NFCe_consultarSituacao(txtchNFe, tpAmb)
End If


If ResultadoNFe <> "" Then

'Ler status
If NFCe = False Then 'Se não for NFCe
Status_Nfe = LerDadosJSON(ResultadoNFe, "status", "", "")
cStat = LerDadosJSON(ResultadoNFe, "cStat", "", "")
xMotivo = LerDadosJSON(ResultadoNFe, "xMotivo", "", "")

If cStat = "100" Then
'Lê a chNFe
chNFe = LerDadosJSON(ResultadoNFe, "chNFe", "", "")
'Lê o nProt
nProt = LerDadosJSON(ResultadoNFe, "nProt", "", "")
'Lê o motivo
motivo = LerDadosJSON(ResultadoNFe, "motivo", "", "")
'Lê o nsNRec
nsNRec = LerDadosJSON(ResultadoNFe, "nsNRec", "", "")
Else
'Lê a chNFe
chNFe = LerDadosJSON(ResultadoNFe, "nfeProc", "chNFe", "")
'Lê o nProt
nProt = LerDadosJSON(ResultadoNFe, "nfeProc", "nProt", "")
'Lê o motivo
motivo = LerDadosJSON(ResultadoNFe, "nfeProc", "motivo", "")
'Lê o nsNRec
nsNRec = LerDadosJSON(ResultadoNFe, "nfeProc", "nsNRec", "")
End If

'
'Debug.print ResultadoNFe

'lendo status do JSON recebido da API
Mensagem = "Status envio : " & LerDadosJSON(ResultadoNFe, "status", "", "")
'lendo motivo do JSON recebido da API
Mensagem = Mensagem & vbCrLf & "Motivo : " & LerDadosJSON(ResultadoNFe, "motivo", "", "")
'lendo chave de acesso do JSON recebido da API
Mensagem = Mensagem & vbCrLf & "Chave NFe : " & LerDadosJSON(ResultadoNFe, "chNFe", "", "")
'lendo Data e Hora de Recebimento na Sefaz, retornado no JSON recebido da API
Mensagem = Mensagem & vbCrLf & "Data recbto : " & LerDadosJSON(ResultadoNFe, "dhRecbto", "", "")
'lendo cSat da Sefaz retornado no JSON recebido da API
Mensagem = Mensagem & vbCrLf & "Status nota : " & LerDadosJSON(ResultadoNFe, "cStat", "", "")
'lendo xMotivo da Sefaz retornado no JSON recebido da API
Mensagem = Mensagem & vbCrLf & "Mensagem : " & LerDadosJSON(ResultadoNFe, "xMotivo", "", "")
'lendo nProt(Protocolo de Autorização) retornado no JSON recebido da API
Mensagem = Mensagem & vbCrLf & "Chave de proteção : " & LerDadosJSON(ResultadoNFe, "nProt", "", "")
 Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Status = " & cStat & ", chave_acesso = '" & chNFe & "', nProt ='" & nProt & "', nsNRec = '" & txtnsNrec.Text & "'  where id_nota = " & txtID_nota
Else
 Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Status = " & cStat & ", chave_acesso = '" & chNFe & "', nProt ='0000', nsNRec = '0000'  where id_nota = " & txtID_nota
End If
'============================================================================================
' Grava os dados retornados no banco atualizando o status da NFe
'============================================================================================
' Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Status = " & cStat & ", chave_acesso = '" & chNFe & "', nProt ='" & nProt & "', nsNRec = '" & txtnsNrec.Text & "'  where id_nota = " & txtID_nota
 ProcCarregaListaNota (1)
 'frmFaturamento_Prod_Serv.ProcCarregaListaNota (1)
 'txtchNFe.Text = chNFe
End If
 'Debug.print Mensagem
 txtRetorno.Text = Mensagem
Else
txtRetorno.Text = xMotivo
ProcCarregaListaNota (1)
frmFaturamento_Prod_Serv.ProcCarregaListaNota (1)
'Usmsgbox "Não é possivel consultar o status da nota sem o numero de recibo NS.", vbCritical, "CAPRIND v5.0"
Exit Sub
End If

ProcCarregaListaNota (1)
frmFaturamento_Prod_Serv.ProcCarregaListaNota (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procConsultarNFE()
On Error GoTo SAI
'===============================================
' BUSCA CNPJ DO EMITENTE DA NOTA
'===============================================
If txtnsNrec.Text = "" And NFCe = False Then
USMsgBox "Numero de recibo não localizado, não será possivel a consulta.", vbCritical, "CAPRIND v5.0"
Exit Sub
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from empresa where Empresa = '" & frmFaturamento_Prod_Serv.txtEmpresa.Text & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
CnpjNF = ReturnNumbersOnly(TBAbrir!CNPJ)
End If
TBAbrir.Close
'==============================================
' INFORMA NUMERO DO RECIBO EMITIDO PELA NS TECNOLOGIA
'==============================================
nsNRec = frmFaturamento_Prod_Serv_NFe_NS.txtnsNrec.Text

txtnsNrec.Text = nsNRec
Var = nsNRec
'==============================================
' EFETUA CONSULTA STATUS DA NFE NO SEFAZ
'==============================================
ResultadoNFe = consultarStatusProcessamento(CnpjNF, Var, tpAmb)
'==============================================
' GRAVA DADOS RETORNADOS EM JSON NOS TEXTOS
'==============================================
txtchNFe.Text = LerDadosJSON(ResultadoNFe, "chNFe", "", "")
txtcStat.Text = LerDadosJSON(ResultadoNFe, "cStat", "", "")
txt_nProt.Text = LerDadosJSON(ResultadoNFe, "nProt", "", "")
'==============================================
' ATUALIZA DADOS DA NFE NA TABELA
'==============================================
Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Status = " & txtcStat.Text & ", chave_acesso = '" & txtchNFe.Text & "', nsNRec ='" & txtnsNrec.Text & "' , nProt ='" & txt_nProt & "'  where id_nota = " & txtID_nota
Conexao.Execute "Update tbl_dados_nota_fiscal_NFe Set Enviar_email = '2' where id_nota = " & txtID_nota
Conexao.Execute "Update tbl_dados_nota_fiscal Set Imprimir = '1' where id = " & txtID_nota

ProcCarregaListaNota (1)
frmFaturamento_Prod_Serv.ProcCarregaListaNota (1)
    
Exit Sub
SAI:
    USMsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, "CAPRIND v5.0", titleCTeAPI
End Sub

Public Sub procConsultar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If txtID_nota = 0 Then
    USMsgBox ("Informe a nota fiscal antes de consultar o status."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Acao = "consultar"
If funVerificaMigrate = False Then Exit Sub

If USMsgBox("Deseja consultar esta nota fiscal?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    NomeArquivo = "NF" & txtNota & txtSerie & "CON"
    'procConsultarXML
    TipoXML = 3
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub proc_XML_Consultar()
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
'    objConsultar.childNodes(1).Text = ChaveMigrate
    
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
        objConsulta.childNodes(3).Text = ReturnNumbersOnly(CnpjNF)
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
objDom.Save (DiretorioEnvio & NomeArquivo & ".xml")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub proc_XML_Produtos()
On Error GoTo tratar_erro
Dim NItem As Integer

NItem = 1
'===========================================
'Criar o nó det
'===========================================
'Abre det=================================================================================================
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select N.*, NFE.Documento_importacao, NFE.Numero_adicao, NFE.Numero_sequencial, NFE.Codigo_fabricante, NFE.Data_registro, NFE.Data_desembaraco, NFE.Local_desembaraco, NFE.UF_desembaraco, NFE.Codigo_exportador, NFE.Via_transp, NFE.Valor_AFRMM, NFE.Forma_imp, NFE.Tipo_produto, NFE.Codigo_ANP, NFE.UF_consumo, NFE.Descricao_ANP from tbl_Detalhes_Nota N LEFT JOIN tbl_Detalhes_Nota_NFe NFE ON N.Int_codigo = NFE.ID_item where N.ID_nota = " & txtID_nota & " order by N.Int_codigo", Conexao, adOpenKeyset, adLockReadOnly

Do While TBProduto.EOF = False

Set objDet = objDom.createElement("det")
objinfNFe.appendChild objDet
objDet.setAttribute "nItem", NItem
   

'nó prod dentro de DetItem
Set objProd = objDom.createElement("prod")
objDet.appendChild objProd
'Abre prod==================================================================================================
    objProd.appendChild objDom.createElement("cProd") '0
    'Verifica se é para utilizar o código de referência na DANFE
    Set TBCodigoDesc = CreateObject("adodb.recordset")
    TBCodigoDesc.Open "Select CodRef from tbl_Dados_Nota_Fiscal_NFe where ID_Nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
    If TBCodigoDesc.EOF = False Then
        If TBCodigoDesc!CodRef = False Or IsNull(TBCodigoDesc!CodRef) = True Then
            objProd.getElementsByTagName("cProd").Item(0).Text = Trim(RemoveAccents(TBProduto!int_Cod_Produto))
            Set TBCodigoDesc = CreateObject("adodb.recordset")
            TBCodigoDesc.Open "Select Codigo_ref_desc_DANFE from empresa where codigo = " & TBproducao!ID_empresa, Conexao, adOpenKeyset, adLockReadOnly
            If TBCodigoDesc!Codigo_ref_desc_DANFE = True Then
                CodRef = 2
            Else
                CodRef = 0
            End If
        Else
            objProd.getElementsByTagName("cProd").Item(0).Text = Trim(RemoveAccents(TBProduto!N_referencia))
            CodRef = 1
        End If
    End If
    TBCodigoDesc.Close
    
    objProd.appendChild objDom.createElement("cEAN") '1
    objProd.appendChild objDom.createElement("xProd") '2
    CompLetra = 0
    If IsNull(Trim(TBProduto!N_referencia)) = False And Trim(TBProduto!N_referencia) <> "" And TBProduto!N_referencia <> TBProduto!int_Cod_Produto And CodRef = 2 Then CompLetra = Len(Trim(TBProduto!N_referencia)) + 3
    If IsNull(Trim(TBProduto!Complemento_descricao)) = False And Trim(TBProduto!Complemento_descricao) <> "" Then CompLetra = Len(Trim(TBProduto!Complemento_descricao)) + 3
    If IsNull(Trim(TBProduto!PCCliente)) = False And Trim(TBProduto!PCCliente) <> "" Then CompLetra = CompLetra + Len(Trim(TBProduto!PCCliente)) + 8
    If IsNull(TBProduto!N_item) = False And Trim(TBProduto!N_item) <> "" Then CompLetra = CompLetra + Len(Trim(TBProduto!N_item)) + 11
                
    DescricaoProduto = Left(Trim(TBProduto!Txt_descricao), 120 - CompLetra)
    
    If CodRef = 2 Then
        If IsNull(TBProduto!N_referencia) = False And TBProduto!N_referencia <> "" And TBProduto!N_referencia <> TBProduto!int_Cod_Produto Then
            DescricaoProduto = "(" & TBProduto!N_referencia & ") - " & DescricaoProduto
        Else
            If Len(TBproducao!txt_tipocliente) = 2 Then TipoFiltro = "C" Else TipoFiltro = "F"
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select IA.N_Referencia from item_aplicacoes IA INNER JOIN projproduto P ON IA.codproduto = P.codproduto where P.Desenho = '" & TBProduto!int_Cod_Produto & "' and IA.ID_cliente_forn = " & TBproducao!Id_Int_Cliente & " and IA.Tipo = '" & TipoFiltro & "' and IA.N_Referencia IS NOT NULL and IA.N_Referencia <> '" & TBProduto!int_Cod_Produto & "'", Conexao, adOpenKeyset, adLockReadOnly
            If TBItem.EOF = False Then
                DescricaoProduto = "(" & TBItem!N_referencia & ") - " & DescricaoProduto
            End If
            TBItem.Close
        End If
    End If
    
    If IsNull(Trim(TBProduto!Complemento_descricao)) = False And Trim(TBProduto!Complemento_descricao) <> "" Then DescricaoProduto = DescricaoProduto & " - " & Trim(TBProduto!Complemento_descricao)
    If IsNull(Trim(TBProduto!PCCliente)) = False And Trim(TBProduto!PCCliente) <> "" Then DescricaoProduto = DescricaoProduto & " - Ped. " & Trim(TBProduto!PCCliente)
    If IsNull(Trim(TBProduto!N_item)) = False And Trim(TBProduto!N_item) <> "" Then DescricaoProduto = DescricaoProduto & " - N. item " & Trim(TBProduto!N_item)
    
   'DescricaoProduto = RemoverCaracter(DescricaoProduto)
    DescricaoProduto = RemoveAccents(DescricaoProduto)
    DescricaoProduto = Replace(DescricaoProduto, Chr(10), vbNullString, 1, -1, vbTextCompare)
    DescricaoProduto = Replace(DescricaoProduto, Chr(13), vbNullString, 1, -1, vbTextCompare)
    DescricaoProduto = Trim(DescricaoProduto)
    
    If Right(DescricaoProduto, 1) = Chr(13) Then
     DescricaoProduto = Replace(DescricaoProduto, Chr(13), vbNullString, 1, -1, vbTextCompare)
    End If
    
    If Right(DescricaoProduto, 1) = Chr(10) Then
    'DescricaoProduto = Replace(DescricaoProduto, Chr(10), vbNullString)
     DescricaoProduto = Replace(DescricaoProduto, Chr(10), vbNullString, 1, -1, vbTextCompare)
    End If
    
    If Right(DescricaoProduto, 1) = Chr(13) Then
     DescricaoProduto = Replace(DescricaoProduto, Chr(13), vbNullString)
    End If
    
    DescricaoProduto = Trim(DescricaoProduto)
    
    
    objProd.getElementsByTagName("xProd").Item(0).Text = Trim(RemoveAccents(Left(Trim(DescricaoProduto), 120)))
    
    
   'Informar o codigo d substituição tributária quando tiver
    Set TBControleNF = CreateObject("adodb.recordset")
    TBControleNF.Open "Select IDIntClasse, CEST from tbl_ClassificacaoFiscal where Idclass = " & TBProduto!ID_CF, Conexao, adOpenKeyset, adLockReadOnly
    If TBControleNF.EOF = False Then
        objProd.appendChild objDom.createElement("NCM") '3
        If TBControleNF!IDIntClasse = "0000.00.00" Then
            objProd.getElementsByTagName("NCM").Item(0).Text = "00000000" 'CFOP
        Else
            objProd.getElementsByTagName("NCM").Item(0).Text = ReturnNumbersOnly(TBControleNF!IDIntClasse) 'CFOP
        End If
        If IsNull(TBControleNF!CEST) = False Then
        If Len(ReturnNumbersOnly(TBControleNF!CEST)) = 7 Then
            objProd.appendChild objDom.createElement("CEST") '21
            objProd.getElementsByTagName("CEST").Item(0).Text = ReturnNumbersOnly(TBControleNF!CEST)
        End If
        End If
    End If
   
    
    Set TBControleNF = CreateObject("adodb.recordset")
    TBControleNF.Open "Select id_CFOP, Devolucao from tbl_NaturezaOperacao where IDCountCfop = " & IIf(IsNull(TBProduto!ID_CFOP), 0, TBProduto!ID_CFOP), Conexao, adOpenKeyset, adLockReadOnly
    If TBControleNF.EOF = False Then
        If Len(TBControleNF!ID_CFOP) > 5 Then
            If TBProduto!retorno = True Then
                If Left(TBControleNF!ID_CFOP, 5) = "5.902" Or Left(TBControleNF!ID_CFOP, 5) = "6.902" Or Left(TBControleNF!ID_CFOP, 5) = "5.916" Or Left(TBControleNF!ID_CFOP, 5) = "6.916" Or Left(TBControleNF!ID_CFOP, 5) = "5.925" Or Left(TBControleNF!ID_CFOP, 5) = "6.925" Then
                    CFOP_Produto = ReturnNumbersOnly(Left(TBControleNF!ID_CFOP, 5))
                ElseIf Right(TBControleNF!ID_CFOP, 5) = "5.902" Or Right(TBControleNF!ID_CFOP, 5) = "6.902" Or Right(TBControleNF!ID_CFOP, 5) = "5.916" Or Right(TBControleNF!ID_CFOP, 5) = "6.916" Or Right(TBControleNF!ID_CFOP, 5) = "5.925" Or Right(TBControleNF!ID_CFOP, 5) = "6.925" Then
                    CFOP_Produto = ReturnNumbersOnly(Right(TBControleNF!ID_CFOP, 5))
                End If
            Else
                If Left(TBControleNF!ID_CFOP, 5) <> "5.902" And Left(TBControleNF!ID_CFOP, 5) <> "6.902" And Left(TBControleNF!ID_CFOP, 5) <> "5.916" And Left(TBControleNF!ID_CFOP, 5) <> "6.916" And Left(TBControleNF!ID_CFOP, 5) <> "5.925" And Left(TBControleNF!ID_CFOP, 5) <> "6.925" Then
                    CFOP_Produto = ReturnNumbersOnly(Left(TBControleNF!ID_CFOP, 5))
                ElseIf Right(TBControleNF!ID_CFOP, 5) <> "5.902" And Right(TBControleNF!ID_CFOP, 5) <> "6.902" And Right(TBControleNF!ID_CFOP, 5) <> "5.916" And Right(TBControleNF!ID_CFOP, 5) <> "6.916" And Right(TBControleNF!ID_CFOP, 5) <> "5.925" And Right(TBControleNF!ID_CFOP, 5) <> "6.925" Then
                    CFOP_Produto = ReturnNumbersOnly(Right(TBControleNF!ID_CFOP, 5))
                End If
            End If
        Else
            CFOP_Produto = ReturnNumbersOnly(TBControleNF!ID_CFOP)
        End If
        objProd.appendChild objDom.createElement("CFOP") '4
        objProd.getElementsByTagName("CFOP").Item(0).Text = CFOP_Produto 'CFOP
        
        If TBControleNF!Devolucao = True Then Devolucao = True Else Devolucao = False
    End If
    TBControleNF.Close
    

    objProd.appendChild objDom.createElement("uCom") '5
    objProd.getElementsByTagName("uCom").Item(0).Text = RemoveAccents(TBProduto!Unidade_com)
    
    objProd.appendChild objDom.createElement("qCom") '6
    objProd.getElementsByTagName("qCom").Item(0).Text = Replace(Format(TBProduto!int_Qtd, "0.#000"), ",", ".")
    
    objProd.appendChild objDom.createElement("vUnCom") '7
    objProd.getElementsByTagName("vUnCom").Item(0).Text = Replace(Format(TBProduto!dbl_ValorUnitario, "0.#0000000"), ",", ".")
    
    objProd.appendChild objDom.createElement("vProd") '8
    objProd.getElementsByTagName("vProd").Item(0).Text = Replace(Format(TBProduto!dbl_ValorTotal, "0.#0"), ",", ".")
    
    objProd.appendChild objDom.createElement("cEANTrib") '9
    
    If TBProduto!GTIN = "" Or IsNull(TBProduto!GTIN) = True Then
        objProd.getElementsByTagName("cEAN").Item(0).Text = "SEM GTIN"
        objProd.getElementsByTagName("cEANTrib").Item(0).Text = "SEM GTIN"
    Else
        objProd.getElementsByTagName("cEAN").Item(0).Text = RemoveAccents(TBProduto!GTIN)
        objProd.getElementsByTagName("cEANTrib").Item(0).Text = RemoveAccents(TBProduto!GTIN)
    End If
'==================================================================================================================
'Valores tributos do produto
'==================================================================================================================
Set TBItem = CreateObject("adodb.recordset")
StrSql = "Select * from projproduto where Desenho = '" & TBProduto!int_Cod_Produto & "'"
TBItem.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly

If TBItem.EOF = False Then
    If TBItem!uTrib = "" Or IsNull(TBItem!uTrib) = True Then
        uTrib = RemoveAccents(TBProduto!txt_Unid)
    Else
        uTrib = TBItem!uTrib
    End If
    
    If TBItem!vTrib = "" Or IsNull(TBItem!vTrib) = True Or TBItem!vTrib = 0 Then
        vTrib = TBProduto!dbl_ValorUnitario
    Else
        vTrib = TBItem!vTrib * TBProduto!int_Qtd
    End If
    
    If IsNull(TBItem!vTrib) = False And TBItem!vTrib <> 0 Then
        qTrib = TBProduto!dbl_ValorTotal / vTrib
    Else
        qTrib = TBProduto!int_Qtd
    End If

End If
TBItem.Close

'==================================================================================================================
'Unidade tributada, valor unitario tributado, e quantidade total tributado
'==================================================================================================================
    'Unidade tributada
    objProd.appendChild objDom.createElement("uTrib") '10
    objProd.getElementsByTagName("uTrib").Item(0).Text = uTrib 'RemoveAccents(TBProduto!txt_Unid)
    
    'Peso unitário tributado
    objProd.appendChild objDom.createElement("qTrib") '11
    objProd.getElementsByTagName("qTrib").Item(0).Text = Replace(Format(qTrib, "0.#000"), ",", ".")
    
    'Peso total tributado
    Var1 = TBProduto!int_Qtd
    objProd.appendChild objDom.createElement("vUnTrib") '12
    objProd.getElementsByTagName("vUnTrib").Item(0).Text = Replace(Format(vTrib, "0.#0000000"), ",", ".")
    
'==================================================================================================================
'Se tiver valor no frete acrescenta tag
'==================================================================================================================
    If TBProduto!Valor_frete > 0 And IsNull(TBProduto!Valor_frete) = False Then
    objProd.appendChild objDom.createElement("vFrete") '13
    objProd.getElementsByTagName("vFrete").Item(0).Text = Replace(IIf(IsNull(TBProduto!Valor_frete), "00.00000000", Format(TBProduto!Valor_frete, "#0.#0")), ",", ".")
    End If
   
    If TBProduto!Valor_acessorias > 0 And IsNull(TBProduto!Valor_acessorias) = False Then
    objProd.appendChild objDom.createElement("vOutro") '14
    objProd.getElementsByTagName("vOutro").Item(0).Text = Replace(IIf(IsNull(TBProduto!Valor_acessorias), "00.00000000", Format(TBProduto!Valor_acessorias, "#0.#0")), ",", ".")
    End If

    
    If TBProduto!Valor_seguro > 0 And IsNull(TBProduto!Valor_seguro) = False Then
    objProd.appendChild objDom.createElement("vSeg") '14
    objProd.getElementsByTagName("vSeg").Item(0).Text = Replace(IIf(IsNull(TBProduto!Valor_seguro), "00.00000000", Format(TBProduto!Valor_seguro, "#0.#0")), ",", ".")
    End If
    
    If TBProduto!Valor_desconto > 0 And IsNull(TBProduto!Valor_desconto) = False Or TBProduto!Valor_desconto_SUFRAMA > 0 And IsNull(TBProduto!Valor_desconto_SUFRAMA) = False Then
    ValorDesconto = IIf(IsNull(TBProduto!Valor_desconto), "00.00000000", TBProduto!Valor_desconto) + IIf(IsNull(TBProduto!Valor_desconto_SUFRAMA), 0, TBProduto!Valor_desconto_SUFRAMA)
    ValorDesconto = ValorDesconto '/ Var1
    ValorDesconto = Format(ValorDesconto, "0.#0")
    ValorDesconto = Replace(ValorDesconto, ",", ".")

    objProd.appendChild objDom.createElement("vDesc") '15
    objProd.getElementsByTagName("vDesc").Item(0).Text = ValorDesconto 'Replace(IIf(IsNull(TBProduto!Valor_desconto), "00.00000000", TBProduto!Valor_desconto) + IIf(IsNull(TBProduto!Valor_desconto_SUFRAMA), 0, TBProduto!Valor_desconto_SUFRAMA), ",", ".")
    End If
    
'    If TBProduto!Valor_desconto > 0 And IsNull(TBProduto!Valor_desconto) = False Then
'    objProd.appendChild objDom.createElement("vOutro_item") '16
'    objProd.getElementsByTagName("vOutro_item").Item(0).Text = Replace(IIf(IsNull(TBProduto!Valor_desconto), "00.00000000", Format(TBProduto!Valor_acessorias, "#0.#0")), ",", ".")
'    End If
'==================================================================================================================
'Totais da nota
'==================================================================================================================
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
    
    'objProd.appendChild objDom.createElement("nTipoItem") '18
    'objProd.getElementsByTagName("nTipoItem").Item(0).Text = IIf(IsNull(TBProduto!Tipo_produto), "0", TBProduto!Tipo_produto)
    
    If IsNull(TBProduto!PCCliente) = False And TBProduto!PCCliente <> "" Then
        objProd.appendChild objDom.createElement("xPed")
        objProd.getElementsByTagName("xPed").Item(0).Text = Left(Trim(TBProduto!PCCliente), 15)
    End If
    
    If IsNull(TBProduto!N_item) = False And TBProduto!N_item <> "" Then
        objProd.appendChild objDom.createElement("nItemPed") '20
        objProd.getElementsByTagName("nItemPed").Item(0).Text = ReturnNumbersOnly(TBProduto!N_item)
    End If
   '==========================================================================================================
   ' se for importacao
   '==========================================================================================================
    If IsNull(TBProduto!Documento_importacao) = False And TBProduto!Documento_importacao <> "" And IsNull(TBProduto!Numero_adicao) = False And TBProduto!Numero_adicao <> "" And IsNull(TBProduto!Numero_sequencial) = False And TBProduto!Numero_sequencial <> "" And IsNull(TBProduto!Codigo_fabricante) = False And TBProduto!Codigo_fabricante <> "" Then
        'nó detDI dentro de Prod
        Set objDI = objDom.createElement("DI")
        objProd.appendChild objDI
        'Abre objDetDI==================================================================================================
            'nó DetDIItem dentro de DetDI
'            Set objDetDIItem = objDom.createElement("detDIItem")
'            objDetDI.appendChild objDetDIItem
            'Abre objDetDIItem==================================================================================================
                objDI.appendChild objDom.createElement("nDI") '0
                objDI.childNodes(0).Text = TBProduto!Documento_importacao
                objDI.appendChild objDom.createElement("dDI") '1
                objDI.childNodes(1).Text = Format(TBProduto!Data_registro, "yyyy-mm-dd")
                objDI.appendChild objDom.createElement("xLocDesemb") '2
                objDI.childNodes(2).Text = TBProduto!Local_desembaraco
                objDI.appendChild objDom.createElement("UFDesemb") '3
                objDI.childNodes(3).Text = TBProduto!UF_desembaraco
                objDI.appendChild objDom.createElement("dDesemb") '5
                objDI.childNodes(4).Text = Format(TBProduto!Data_desembaraco, "yyyy-mm-dd")
                objDI.appendChild objDom.createElement("tpViaTransp") '6
                
                objDI.childNodes(5).Text = TBProduto!Via_transp
                
                If TBProduto!Via_transp = 1 And IsNull(TBProduto!Valor_AFRMM) = False Then
                objDI.appendChild objDom.createElement("vAFRMM") '7
                objDI.childNodes(6).Text = Replace(TBProduto!Valor_AFRMM, ",", ".")
                Else
                objDI.appendChild objDom.createElement("vAFRMM") '7
                objDI.childNodes(6).Text = "0"
                End If
                
                objDI.appendChild objDom.createElement("tpIntermedio") '8
                objDI.childNodes(7).Text = IIf(IsNull(TBProduto!Forma_imp), 1, TBProduto!Forma_imp)
'                If TBProduto!Forma_imp <> 1 Then
                objDI.appendChild objDom.createElement("CNPJ") '9
                objDI.childNodes(8).Text = ReturnNumbersOnly(TBproducao!CNPJ)
                objDI.appendChild objDom.createElement("UFTerceiro") '10
                objDI.childNodes(9).Text = TBproducao!UF
                objDI.appendChild objDom.createElement("cExportador") '4
                objDI.childNodes(10).Text = TBProduto!Codigo_exportador
                
 '               End If
            'Fecha objDetDIItem=================================================================================================
            
            'nó detAdicoes dentro de DetDIItem
            Set objadi = objDom.createElement("adi")
            objDI.appendChild objadi
            'Abre objDetAdicoes==================================================================================================
                'nó detAdicoes dentro de DetDI
                'Set objDetAdicoesItem = objDom.createElement("detAdicoesItem")
                'objDetAdicoes.appendChild objDetAdicoesItem
                'Abre objDetAdicoesItem==================================================================================================
                    objadi.appendChild objDom.createElement("nAdicao") '0
                    objadi.childNodes(0).Text = TBProduto!Numero_adicao
                    objadi.appendChild objDom.createElement("nSeqAdic") '1
                    objadi.childNodes(1).Text = TBProduto!Numero_sequencial
                    objadi.appendChild objDom.createElement("cFabricante") '2
                    objadi.childNodes(2).Text = TBProduto!Codigo_fabricante
                'Fecha objDetAdicoesItem=================================================================================================
            'Fecha objDetAdicoes=================================================================================================
        'Fecha objDetDI=================================================================================================
    End If
 '============================================================================
 ' Informações de Combustiveis
 '============================================================================
    If IsNull(TBProduto!Codigo_ANP) = False And TBProduto!Codigo_ANP <> "" Then
        'nó comb dentro de prod
        Set objComb = objDom.createElement("comb")
        objProd.appendChild objComb
        'Abre comb
        '==================================================================================================
            objComb.appendChild objDom.createElement("cProdANP") '0
            objComb.childNodes(0).Text = TBProduto!Codigo_ANP
            objComb.appendChild objDom.createElement("descANP") '2
            objComb.childNodes(1).Text = TBProduto!Descricao_ANP
            objComb.appendChild objDom.createElement("pGLP")
            objComb.childNodes(2).Text = "100.0000"
            objComb.appendChild objDom.createElement("pGNn")
            objComb.childNodes(3).Text = "0.0000"
            objComb.appendChild objDom.createElement("pGNi")
            objComb.childNodes(4).Text = "0.0000"
            objComb.appendChild objDom.createElement("vPart")
            objComb.childNodes(5).Text = "0.01"
            objComb.appendChild objDom.createElement("UFCons") '1
            objComb.childNodes(6).Text = TBProduto!UF_consumo
        'Fecha comb
        '=================================================================================================
    End If
'================================================================================================
'Fecha prod
'================================================================================================
Proc_XML_Impostos
Proc_XML_Imposto_Devolucao
'=====================================================================================================
' Informações adicionais do produto
'=====================================================================================================
If IsNull(TBProduto!Inf_adicionais_prod) = False And TBProduto!Inf_adicionais_prod <> "" Then
    objDet.appendChild objDom.createElement("infAdProd")
    objDet.getElementsByTagName("infAdProd").Item(0).Text = Trim(TBProduto!Inf_adicionais_prod)
End If

TBProduto.MoveNext
NItem = NItem + 1
Loop
'================================================================================================
'Fecha detalhes da nota fiscal
'================================================================================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub Proc_XML_Imposto_Interestadual()



End Sub

Public Sub Proc_XML_Imposto_Devolucao()
On Error GoTo tratar_erro

Set TBDev = CreateObject("adodb.recordset")
TBDev.Open "Select * from tbl_Detalhes_Nota_CST_IPI where id_item = " & TBProduto!Int_codigo, Conexao, adOpenKeyset, adLockReadOnly
If TBDev.EOF = False Then

'Verificar depois
If Devolucao = True And TBDev!vIPIdevolv > 0 Then
    'nó impostoDevol dentro de detItem (G02)
    Set objImpostoDevol = objDom.createElement("impostoDevol")
    objDet.appendChild objImpostoDevol
    'Abre objImpostoDevol==================================================================================================
        objImpostoDevol.appendChild objDom.createElement("pDevol")
        objImpostoDevol.childNodes(0).Text = Replace(Format(TBDev!pIPIdevolv, "0.#0"), ",", ".")
        'nó IPIDevol dentro de impostoDevol
        Set objIPI = objDom.createElement("IPI")
        objImpostoDevol.appendChild objIPI
        'Abre IPIDevol==================================================================================================
            objIPI.appendChild objDom.createElement("vIPIDevol")
            objIPI.childNodes(0).Text = Replace(Format(TBDev!vIPIdevolv, "0.#0"), ",", ".")
        'Fecha IPIDevol=================================================================================================
                
    
    'Fecha objImpostoDevol=================================================================================================
End If
End If
TBDev.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub Proc_XML_IMP_ICMS_SIMPLESNACIONAL()
On Error GoTo tratar_erro

'=================================================================================================
' GRUPO 101
'=================================================================================================
If FimCST = "101" Then

vFimCST = "101"

VarObjeto = "objICMSSN" & vFimCST
VarObjetonome = "ICMSSN" & vFimCST

Set VarObjeto = objDom.createElement(VarObjetonome)
objICMS.appendChild VarObjeto
  VarObjeto.appendChild objDom.createElement("orig")
  VarObjeto.getElementsByTagName("orig").Item(0).Text = TBCST!Origem_mercadoria
  VarObjeto.appendChild objDom.createElement("CSOSN")
  VarObjeto.getElementsByTagName("CSOSN").Item(0).Text = vFimCST
  VarObjeto.appendChild objDom.createElement("pCredSN")
  VarObjeto.getElementsByTagName("pCredSN").Item(0).Text = Replace(Format(TBCST!ICMS_SN, "0.#000"), ",", ".")
  VarObjeto.appendChild objDom.createElement("vCredICMSSN")
  VarObjeto.getElementsByTagName("vCredICMSSN").Item(0).Text = Replace(Format(TBCST!Valor_ICMS_SN, "0.#0"), ",", ".")
 End If
'=================================================================================================
' GRUPO 102
'=================================================================================================
If FimCST = "102" Or FimCST = "103" Or FimCST = "300" Or FimCST = "400" Then

vFimCST = "102"

VarObjeto = "objICMSSN" & vFimCST
VarObjetonome = "ICMSSN" & vFimCST

Set VarObjeto = objDom.createElement(VarObjetonome)
objICMS.appendChild VarObjeto
  VarObjeto.appendChild objDom.createElement("orig")
  VarObjeto.getElementsByTagName("orig").Item(0).Text = TBCST!Origem_mercadoria
  VarObjeto.appendChild objDom.createElement("CSOSN")
  VarObjeto.getElementsByTagName("CSOSN").Item(0).Text = FimCST
End If

'=================================================================================================
' GRUPO 201
'=================================================================================================
If FimCST = "201" Then

vFimCST = "201"

VarObjeto = "objICMSSN" & FimCST
VarObjetonome = "ICMSSN" & FimCST

Set VarObjeto = objDom.createElement(VarObjetonome)
objICMS.appendChild VarObjeto
  VarObjeto.appendChild objDom.createElement("orig")
  VarObjeto.getElementsByTagName("orig").Item(0).Text = TBCST!Origem_mercadoria
  VarObjeto.appendChild objDom.createElement("CSOSN")
  VarObjeto.getElementsByTagName("CSOSN").Item(0).Text = FimCST
  VarObjeto.appendChild objDom.createElement("modBCST")
  VarObjeto.getElementsByTagName("modBCST").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
'  Falta acrescentar
'========================================================
'  VarObjeto.appendChild objDom.createElement("pMVAST")
'  VarObjeto.getElementsByTagName("pMVAST").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
  VarObjeto.appendChild objDom.createElement("pRedBCST")
  VarObjeto.getElementsByTagName("pRedBCST").Item(0).Text = Replace(Format(TBCST!Percentual_reducao_BC_ST, "0.#000"), ",", ".") 'pRedBCST
  VarObjeto.appendChild objDom.createElement("vBCST")
  VarObjeto.getElementsByTagName("vBCST").Item(0).Text = Replace(Format(TBCST!Valor_BC_ST, "0.#0"), ",", ".") 'vBCST
  VarObjeto.appendChild objDom.createElement("pICMSST")
  VarObjeto.getElementsByTagName("pICMSST").Item(0).Text = Replace(Format(TBCST!Aliquota_imposto_ST, "0.#000"), ",", ".") 'pICMSST
  VarObjeto.appendChild objDom.createElement("vICMSST")
  VarObjeto.getElementsByTagName("vICMSST").Item(0).Text = Replace(Format(TBCST!Valor_ICMS_ST, "0.#0"), ",", ".") 'vICMSST_icms
  VarObjeto.appendChild objDom.createElement("pCredSN")
  VarObjeto.getElementsByTagName("pCredSN").Item(0).Text = Replace(Format(TBCST!ICMS_SN, "0.#000"), ",", ".")
  VarObjeto.appendChild objDom.createElement("vCredICMSSN")
  VarObjeto.getElementsByTagName("vCredICMSSN").Item(0).Text = Replace(Format(TBCST!Valor_ICMS_SN, "0.#0"), ",", ".")
'=========================================
' Chama calculo de substituição tributária
'=========================================
  Proc_XML_IMP_SUBST_TRIBUT
'=========================================
End If
'=================================================================================================
' GRUPO 202
'=================================================================================================
If FimCST = "202" Or FimCST = "203" Then

vFimCST = "202"

VarObjeto = "objICMSSN" & vFimCST
VarObjetonome = "ICMSSN" & vFimCST

Set VarObjeto = objDom.createElement(VarObjetonome)
objICMS.appendChild VarObjeto
  VarObjeto.appendChild objDom.createElement("orig")
  VarObjeto.getElementsByTagName("orig").Item(0).Text = TBCST!Origem_mercadoria
  VarObjeto.appendChild objDom.createElement("CSOSN")
  VarObjeto.getElementsByTagName("CSOSN").Item(0).Text = FimCST
  VarObjeto.appendChild objDom.createElement("modBCST")
  VarObjeto.getElementsByTagName("modBCST").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
  VarObjeto.appendChild objDom.createElement("vBCST")
  VarObjeto.getElementsByTagName("vBCST").Item(0).Text = Replace(Format(IIf(IsNull(TBCST!Valor_BC_ST) = False, TBCST!Valor_BC_ST, "00.00"), "0.#0"), ",", ".") 'vBCST
  VarObjeto.appendChild objDom.createElement("pICMSST")
  VarObjeto.getElementsByTagName("pICMSST").Item(0).Text = Replace(Format(TBCST!Aliquota_imposto_ST, "0.#000"), ",", ".") 'pICMSST
  VarObjeto.appendChild objDom.createElement("vICMSST")
  VarObjeto.getElementsByTagName("vICMSST").Item(0).Text = Replace(Format(TBCST!Valor_ICMS_ST, "0.#0"), ",", ".") 'vICMSST_icms
'=========================================
' Chama calculo de substituição tributária
'=========================================
   Proc_XML_IMP_SUBST_TRIBUT
'=========================================
End If

'=================================================================================================
' GRUPO 500
'=================================================================================================
If FimCST = "500" Then

vFimCST = "500"

VarObjeto = "objICMSSN" & FimCST
VarObjetonome = "ICMSSN" & FimCST

Set VarObjeto = objDom.createElement(VarObjetonome)
objICMS.appendChild VarObjeto
  VarObjeto.appendChild objDom.createElement("orig")
  VarObjeto.getElementsByTagName("orig").Item(0).Text = TBCST!Origem_mercadoria
  VarObjeto.appendChild objDom.createElement("CSOSN")
  VarObjeto.getElementsByTagName("CSOSN").Item(0).Text = FimCST
'========================================================
'  Falta acrescentar
'========================================================
'  VarObjeto.appendChild objDom.createElement("vBCSTRet")
'  VarObjeto.getElementsByTagName("vBCSTRet").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'  VarObjeto.appendChild objDom.createElement("vICMSSTRet")
'  VarObjeto.getElementsByTagName("vICMSSTRet").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
End If

'=================================================================================================
' GRUPO 900
'=================================================================================================
If FimCST = "900" Then

vFimCST = "900"

VarObjeto = "objICMSSN" & FimCST
VarObjetonome = "ICMSSN" & FimCST

Set VarObjeto = objDom.createElement(VarObjetonome)
objICMS.appendChild VarObjeto
  VarObjeto.appendChild objDom.createElement("orig")
  VarObjeto.getElementsByTagName("orig").Item(0).Text = TBCST!Origem_mercadoria
  VarObjeto.appendChild objDom.createElement("CSOSN")
  VarObjeto.getElementsByTagName("CSOSN").Item(0).Text = FimCST
  VarObjeto.appendChild objDom.createElement("modBC")
  VarObjeto.getElementsByTagName("modBC").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao) = False, TBCST!Modalidade_determinacao, "0") 'modBC
  VarObjeto.appendChild objDom.createElement("vBC")
  VarObjeto.getElementsByTagName("vBC").Item(0).Text = Replace(Format(TBCST!Valor_BC, "0.#0"), ",", ".")  'vBC
  VarObjeto.appendChild objDom.createElement("pRedBC")
  VarObjeto.getElementsByTagName("pRedBC").Item(0).Text = Replace(Format(TBCST!Percentual_reducao_BC, "0.#000"), ",", ".") 'pRedBC
  VarObjeto.appendChild objDom.createElement("pICMS")
  VarObjeto.getElementsByTagName("pICMS").Item(0).Text = Replace(Format(TBProduto!int_ICMS, "0.#0"), ",", ".")  'pICMS
  VarObjeto.appendChild objDom.createElement("vICMS")
  VarObjeto.getElementsByTagName("vICMS").Item(0).Text = Replace(Format(TBCST!Valor_ICMS, "0.#0"), ",", ".")  'vICMS_icms
  VarObjeto.appendChild objDom.createElement("modBCST")
  VarObjeto.getElementsByTagName("modBCST").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
'  Falta acrescentar
'========================================================
  VarObjeto.appendChild objDom.createElement("pMVAST")
  VarObjeto.getElementsByTagName("pMVAST").Item(0).Text = "0.00" 'IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
  VarObjeto.appendChild objDom.createElement("pRedBCST")
  VarObjeto.getElementsByTagName("pRedBCST").Item(0).Text = Replace(Format(IIf(IsNull(TBCST!Percentual_reducao_BC_ST) = False, TBCST!Percentual_reducao_BC_ST, "00.0000"), "0.#000"), ",", ".") 'pRedBCST
  VarObjeto.appendChild objDom.createElement("vBCST")
  VarObjeto.getElementsByTagName("vBCST").Item(0).Text = Replace(Format(IIf(IsNull(TBCST!Valor_BC_ST) = False, TBCST!Valor_BC_ST, "00.00"), "0.#0"), ",", ".") 'vBCST
  VarObjeto.appendChild objDom.createElement("pICMSST")
  VarObjeto.getElementsByTagName("pICMSST").Item(0).Text = Replace(Format(TBCST!Aliquota_imposto_ST, "0.#000"), ",", ".") 'pICMSST
  VarObjeto.appendChild objDom.createElement("vICMSST")
  VarObjeto.getElementsByTagName("vICMSST").Item(0).Text = Replace(Format(TBCST!Valor_ICMS_ST, "0.#0"), ",", ".") 'vICMSST_icms
  VarObjeto.appendChild objDom.createElement("pCredSN")
  VarObjeto.getElementsByTagName("pCredSN").Item(0).Text = Replace(Format(TBCST!ICMS_SN, "0.#000"), ",", ".")
  VarObjeto.appendChild objDom.createElement("vCredICMSSN")
  VarObjeto.getElementsByTagName("vCredICMSSN").Item(0).Text = Replace(Format(TBCST!Valor_ICMS_SN, "0.#0"), ",", ".")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub Proc_XML_IMP_ICMS()
On Error GoTo tratar_erro

'===============================================================================
' CST DO ICMS
'===============================================================================

 If TBCST.EOF = False Then
'nó objICMS dentro de objImposto
 Set objICMS = objDom.createElement("ICMS")
  objImposto.appendChild objICMS
 Select Case RegimeEmpresa
   '=====================================================
   ' SE O EMITENTE FOR SIMPLES NACIONAL
   '=====================================================
   Case 1: Proc_XML_IMP_ICMS_SIMPLESNACIONAL
   '=====================================================
   ' SE O EMITENTE FOR LUCRO PRESUMIDO
   '=====================================================
   Case 2: Proc_XML_IMP_ICMS_LP_LR
   '=====================================================
   ' SE O EMITENTE FOR LUCRO REAL
   '=====================================================
   Case 3: Proc_XML_IMP_ICMS_LP_LR
   '=====================================================
   ' SE O EMITENTE FOR SIMPLES NACIONAL EXCESSO SUB LIMITE
   '=====================================================
   Case 4: Proc_XML_IMP_ICMS_LP_LR

 End Select

End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub Proc_XML_IMP_ICMS_LP_LR()
On Error GoTo tratar_erro
'=================================================================================================
' GRUPO 00
'=================================================================================================
If FimCST = "00" Then

VarObjeto = "objICMS" & FimCST
VarObjetonome = "ICMS" & FimCST

Set VarObjeto = objDom.createElement(VarObjetonome)
objICMS.appendChild VarObjeto
  VarObjeto.appendChild objDom.createElement("orig")
  VarObjeto.getElementsByTagName("orig").Item(0).Text = TBCST!Origem_mercadoria
  VarObjeto.appendChild objDom.createElement("CST")
  VarObjeto.getElementsByTagName("CST").Item(0).Text = FimCST
  VarObjeto.appendChild objDom.createElement("modBC")
  VarObjeto.getElementsByTagName("modBC").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao) = False, TBCST!Modalidade_determinacao, "0") 'modBC
  VarObjeto.appendChild objDom.createElement("vBC")
  VarObjeto.getElementsByTagName("vBC").Item(0).Text = Replace(Format(TBCST!Valor_BC, "0.#0"), ",", ".")   'vBC
  VarObjeto.appendChild objDom.createElement("pICMS")
  VarObjeto.getElementsByTagName("pICMS").Item(0).Text = Replace(Format(TBProduto!int_ICMS, "0.#0"), ",", ".") 'pICMS
  VarObjeto.appendChild objDom.createElement("vICMS")
  VarObjeto.getElementsByTagName("vICMS").Item(0).Text = Replace(Format(TBCST!Valor_ICMS, "0.#0"), ",", ".")  'vICMS_icms

End If

'=================================================================================================
' GRUPO 10
'=================================================================================================
If FimCST = "10" Then

VarObjeto = "objICMS" & FimCST
VarObjetonome = "ICMS" & FimCST

Set VarObjeto = objDom.createElement(VarObjetonome)
objICMS.appendChild VarObjeto
  VarObjeto.appendChild objDom.createElement("orig")
  VarObjeto.getElementsByTagName("orig").Item(0).Text = TBCST!Origem_mercadoria
  VarObjeto.appendChild objDom.createElement("CST")
  VarObjeto.getElementsByTagName("CST").Item(0).Text = FimCST
  VarObjeto.appendChild objDom.createElement("modBC")
  VarObjeto.getElementsByTagName("modBC").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao) = False, TBCST!Modalidade_determinacao, "0") 'modBC
  VarObjeto.appendChild objDom.createElement("vBC")
  VarObjeto.getElementsByTagName("vBC").Item(0).Text = Replace(Format(TBCST!Valor_BC, "0.#0"), ",", ".")   'vBC
  VarObjeto.appendChild objDom.createElement("pICMS")
  VarObjeto.getElementsByTagName("pICMS").Item(0).Text = Replace(Format(TBProduto!int_ICMS, "0.#0"), ",", ".") 'pICMS
  VarObjeto.appendChild objDom.createElement("vICMS")
  VarObjeto.getElementsByTagName("vICMS").Item(0).Text = Replace(Format(TBCST!Valor_ICMS, "0.#0"), ",", ".")  'vICMS_icms
  VarObjeto.appendChild objDom.createElement("modBCST")
  VarObjeto.getElementsByTagName("modBCST").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
'  Falta acrescentar
'========================================================
'  VarObjeto.appendChild objDom.createElement("pMVAST")
'  VarObjeto.getElementsByTagName("pMVAST").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
  VarObjeto.appendChild objDom.createElement("pRedBCST")
  VarObjeto.getElementsByTagName("pRedBCST").Item(0).Text = Replace(Format(TBCST!Percentual_reducao_BC_ST, "0.#000"), ",", ".") 'pRedBCST
  VarObjeto.appendChild objDom.createElement("vBCST")
  VarObjeto.getElementsByTagName("vBCST").Item(0).Text = Replace(Format(TBCST!Valor_BC_ST, "0.#0"), ",", ".") 'vBCST
  VarObjeto.appendChild objDom.createElement("pICMSST")
  VarObjeto.getElementsByTagName("pICMSST").Item(0).Text = Replace(Format(TBCST!Aliquota_imposto_ST, "0.#000"), ",", ".") 'pICMSST
  VarObjeto.appendChild objDom.createElement("vICMSST")
  VarObjeto.getElementsByTagName("vICMSST").Item(0).Text = Replace(Format(TBCST!Valor_ICMS_ST, "0.#0"), ",", ".") 'vICMSST_icms
'=========================================
' Chama calculo de substituição tributária
'=========================================
    Proc_XML_IMP_SUBST_TRIBUT
'=========================================
End If

'=================================================================================================
' GRUPO 20
'=================================================================================================
If FimCST = "20" Then

VarObjeto = "objICMS" & FimCST
VarObjetonome = "ICMS" & FimCST

Set VarObjeto = objDom.createElement(VarObjetonome)
objICMS.appendChild VarObjeto
  VarObjeto.appendChild objDom.createElement("orig")
  VarObjeto.getElementsByTagName("orig").Item(0).Text = TBCST!Origem_mercadoria
  VarObjeto.appendChild objDom.createElement("CST")
  VarObjeto.getElementsByTagName("CST").Item(0).Text = FimCST
  VarObjeto.appendChild objDom.createElement("modBC")
  VarObjeto.getElementsByTagName("modBC").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao) = False, TBCST!Modalidade_determinacao, "0") 'modBC
  VarObjeto.appendChild objDom.createElement("pRedBC")
  VarObjeto.getElementsByTagName("pRedBC").Item(0).Text = Replace(Format(TBCST!Percentual_reducao_BC, "0.#000"), ",", ".") 'pRedBC
  VarObjeto.appendChild objDom.createElement("vBC")
  VarObjeto.getElementsByTagName("vBC").Item(0).Text = Replace(Format(TBCST!Valor_BC, "0.#0"), ",", ".")   'vBC
  VarObjeto.appendChild objDom.createElement("pICMS")
  VarObjeto.getElementsByTagName("pICMS").Item(0).Text = Replace(Format(TBProduto!int_ICMS, "0.#0"), ",", ".")   'pICMS
  VarObjeto.appendChild objDom.createElement("vICMS")
  VarObjeto.getElementsByTagName("vICMS").Item(0).Text = Replace(Format(TBCST!Valor_ICMS, "0.#0"), ",", ".")  'vICMS_icms

'========================================================
'  Falta acrescentar
'========================================================
'  VarObjeto.appendChild objDom.createElement("vICMSDeson")
'  VarObjeto.getElementsByTagName("vICMSDeson").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'  VarObjeto.appendChild objDom.createElement("motDesICMS")
'  VarObjeto.getElementsByTagName("motDesICMS").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
End If


'=================================================================================================
' GRUPO 30
'=================================================================================================
If FimCST = "30" Then

VarObjeto = "objICMS" & FimCST
VarObjetonome = "ICMS" & FimCST

Set VarObjeto = objDom.createElement(VarObjetonome)
objICMS.appendChild VarObjeto
  VarObjeto.appendChild objDom.createElement("orig")
  VarObjeto.getElementsByTagName("orig").Item(0).Text = TBCST!Origem_mercadoria
  VarObjeto.appendChild objDom.createElement("CST")
  VarObjeto.appendChild objDom.createElement("modBCST")
  VarObjeto.getElementsByTagName("modBCST").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
'  Falta acrescentar
'========================================================
'  VarObjeto.appendChild objDom.createElement("pMVAST")
'  VarObjeto.getElementsByTagName("pMVAST").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'  VarObjeto.appendChild objDom.createElement("pRedBCST")
'  VarObjeto.getElementsByTagName("pRedBCST").Item(0).Text = Replace(Format(TBCST!Percentual_reducao_BC_ST, "0.#0"), ",", ".") 'pRedBCST
'========================================================
 
  VarObjeto.appendChild objDom.createElement("vBCST")
  VarObjeto.getElementsByTagName("vBCST").Item(0).Text = Replace(Format(TBCST!Valor_BC_ST, "0.#0"), ",", ".") 'vBCST
  VarObjeto.appendChild objDom.createElement("pICMSST")
  VarObjeto.getElementsByTagName("pICMSST").Item(0).Text = Replace(Format(TBCST!Aliquota_imposto_ST, "0.#000"), ",", ".") 'pICMSST
  VarObjeto.appendChild objDom.createElement("vICMSST")
  VarObjeto.getElementsByTagName("vICMSST").Item(0).Text = Replace(Format(TBCST!Valor_ICMS_ST, "0.#0"), ",", ".") 'vICMSST_icms

'========================================================
'  Falta acrescentar
'========================================================
'  VarObjeto.appendChild objDom.createElement("vICMSDeson")
'  VarObjeto.getElementsByTagName("vICMSDeson").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'  VarObjeto.appendChild objDom.createElement("motDesICMS")
'  VarObjeto.getElementsByTagName("motDesICMS").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
'=========================================
' Chama calculo de substituição tributária
'=========================================
    Proc_XML_IMP_SUBST_TRIBUT
'=========================================
End If

'=================================================================================================
' GRUPO 40
'=================================================================================================
If FimCST = "40" Or FimCST = "41" Or FimCST = "50" Then
  vFimCST = "40"
  
  VarObjeto = "objICMS" & vFimCST
  VarObjetonome = "ICMS" & vFimCST
  
  Set VarObjeto = objDom.createElement(VarObjetonome)
  objICMS.appendChild VarObjeto
    VarObjeto.appendChild objDom.createElement("orig")
    VarObjeto.getElementsByTagName("orig").Item(0).Text = TBCST!Origem_mercadoria
    VarObjeto.appendChild objDom.createElement("CST")
    VarObjeto.getElementsByTagName("CST").Item(0).Text = FimCST
  '========================================================
  '  Acrescentado em 31-10-2019 (Esplendor)
  '========================================================
  ' Só coloca a tag se tiver Suframa marcado na CFOP
  '========================================================
  If Suframa = True And Desconto_Suframa = True Then
    VarObjeto.appendChild objDom.createElement("vICMSDeson")
    VarObjeto.getElementsByTagName("vICMSDeson").Item(0).Text = IIf(IsNull(TBCST!Valor_ICMS_desonerado), 4, Replace(Format(TBCST!Valor_ICMS_desonerado, "0.#0"), ",", "."))
    VarObjeto.appendChild objDom.createElement("motDesICMS")
    VarObjeto.getElementsByTagName("motDesICMS").Item(0).Text = IIf(IsNull(TBCST!Motivo_ICMS_desonerado), 4, TBCST!Motivo_ICMS_desonerado)
  End If
  '========================================================
End If

'=================================================================================================
' GRUPO 51
'=================================================================================================
If FimCST = "51" Then

VarObjeto = "objICMS" & FimCST
VarObjetonome = "ICMS" & FimCST

Set VarObjeto = objDom.createElement(VarObjetonome)
objICMS.appendChild VarObjeto
  VarObjeto.appendChild objDom.createElement("orig")
  VarObjeto.getElementsByTagName("orig").Item(0).Text = TBCST!Origem_mercadoria
  VarObjeto.appendChild objDom.createElement("CST")
  VarObjeto.getElementsByTagName("CST").Item(0).Text = FimCST
  VarObjeto.appendChild objDom.createElement("modBC")
  VarObjeto.getElementsByTagName("modBC").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao) = False, TBCST!Modalidade_determinacao, "0") 'modBC
  VarObjeto.appendChild objDom.createElement("pRedBC")
  VarObjeto.getElementsByTagName("pRedBC").Item(0).Text = Replace(Format(TBCST!Percentual_reducao_BC, "0.#000"), ",", ".") 'pRedBC
  VarObjeto.appendChild objDom.createElement("vBC")
  VarObjeto.getElementsByTagName("vBC").Item(0).Text = Replace(Format(TBCST!Valor_BC, "0.#0"), ",", ".")  'vBC
  VarObjeto.appendChild objDom.createElement("pICMS")
  VarObjeto.getElementsByTagName("pICMS").Item(0).Text = Replace(Format(TBProduto!int_ICMS, "0.#0"), ",", ".")  'pICMS
  VarObjeto.appendChild objDom.createElement("vICMSOp")
  VarObjeto.getElementsByTagName("vICMSOp").Item(0).Text = Replace(Format(TBCST!Valor_ICMS, "0.#0"), ",", ".")  'vICMSOp
  VarObjeto.appendChild objDom.createElement("pDif")
  VarObjeto.getElementsByTagName("pDif").Item(0).Text = Replace(Format(TBCST!Percentual_ICMS_DIF, "0.#000"), ",", ".")   'pDif
  VarObjeto.appendChild objDom.createElement("vICMSDif")
  VarObjeto.getElementsByTagName("vICMSDif").Item(0).Text = Replace(Format(TBCST!Valor_ICMS_DIF, "0.#0"), ",", ".")  'vICMSDif
  VarObjeto.appendChild objDom.createElement("vICMS")
  VarObjeto.getElementsByTagName("vICMS").Item(0).Text = Replace(Format(TBCST!Valor_ICMS, "0.#0"), ",", ".")  'vICMS_icms

End If

'=================================================================================================
' GRUPO 60
'=================================================================================================
If FimCST = "60" Then

VarObjeto = "objICMS" & FimCST
VarObjetonome = "ICMS" & FimCST

Set VarObjeto = objDom.createElement(VarObjetonome)
objICMS.appendChild VarObjeto
  VarObjeto.appendChild objDom.createElement("orig")
  VarObjeto.getElementsByTagName("orig").Item(0).Text = TBCST!Origem_mercadoria
  VarObjeto.appendChild objDom.createElement("CST")
  VarObjeto.getElementsByTagName("CST").Item(0).Text = FimCST
'========================================================
'  Falta acrescentar
'<vBCSTRet>0.00</vBCSTRet>
'<pST>0.0000</pST>
'<vICMSSubstituto>0.00</vICMSSubstituto>
'<vICMSSTRet>0.00</vICMSSTRet>
'========================================================
''modBCST
  VarObjeto.appendChild objDom.createElement("vBCSTRet")
  VarObjeto.getElementsByTagName("vBCSTRet").Item(0).Text = "0.00" 'Format("0", "0.#0") 'IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
  VarObjeto.appendChild objDom.createElement("pST")
  VarObjeto.getElementsByTagName("pST").Item(0).Text = "0.0000"
  VarObjeto.appendChild objDom.createElement("vICMSSubstituto")
  VarObjeto.getElementsByTagName("vICMSSubstituto").Item(0).Text = "0.00"
  VarObjeto.appendChild objDom.createElement("vICMSSTRet")
  VarObjeto.getElementsByTagName("vICMSSTRet").Item(0).Text = "0.00" 'Format("0", "0.#000") 'IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
'=========================================
' Chama calculo de substituição tributária
'=========================================
 '   Proc_XML_IMP_SUBST_TRIBUT
'=========================================
End If

'=================================================================================================
' GRUPO 70
'=================================================================================================
If FimCST = "70" Then

VarObjeto = "objICMS" & FimCST
VarObjetonome = "ICMS" & FimCST

Set VarObjeto = objDom.createElement(VarObjetonome)
objICMS.appendChild VarObjeto
  VarObjeto.appendChild objDom.createElement("orig")
  VarObjeto.getElementsByTagName("orig").Item(0).Text = TBCST!Origem_mercadoria
  VarObjeto.appendChild objDom.createElement("CST")
  VarObjeto.getElementsByTagName("CST").Item(0).Text = FimCST
  VarObjeto.appendChild objDom.createElement("modBC")
  VarObjeto.getElementsByTagName("modBC").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao) = False, TBCST!Modalidade_determinacao, "0") 'modBC
  VarObjeto.appendChild objDom.createElement("pRedBC")
  VarObjeto.getElementsByTagName("pRedBC").Item(0).Text = Replace(Format(TBCST!Percentual_reducao_BC, "0.#000"), ",", ".") 'pRedBC
  VarObjeto.appendChild objDom.createElement("vBC")
  VarObjeto.getElementsByTagName("vBC").Item(0).Text = Replace(Format(TBCST!Valor_BC, "0.#0"), ",", ".")  'vBC
  VarObjeto.appendChild objDom.createElement("pICMS")
  VarObjeto.getElementsByTagName("pICMS").Item(0).Text = Replace(Format(TBProduto!int_ICMS, "0.#0"), ",", ".")  'pICMS
  VarObjeto.appendChild objDom.createElement("vICMS")
  VarObjeto.getElementsByTagName("vICMS").Item(0).Text = Replace(Format(TBCST!Valor_ICMS, "0.#0"), ",", ".")  'vICMS_icms
  VarObjeto.appendChild objDom.createElement("modBCST")
  VarObjeto.getElementsByTagName("modBCST").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
'  Falta acrescentar
'========================================================
'  VarObjeto.appendChild objDom.createElement("pMVAST")
'  VarObjeto.getElementsByTagName("pMVAST").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
  VarObjeto.appendChild objDom.createElement("pRedBCST")
  VarObjeto.getElementsByTagName("pRedBCST").Item(0).Text = Replace(Format(TBCST!Percentual_reducao_BC_ST, "0.#000"), ",", ".") 'pRedBCST
  VarObjeto.appendChild objDom.createElement("vBCST")
  VarObjeto.getElementsByTagName("vBCST").Item(0).Text = Replace(Format(TBCST!Valor_BC_ST, "0.#0"), ",", ".") 'vBCST
  VarObjeto.appendChild objDom.createElement("pICMSST")
  VarObjeto.getElementsByTagName("pICMSST").Item(0).Text = Replace(Format(TBCST!Aliquota_imposto_ST, "0.#000"), ",", ".") 'pICMSST
  VarObjeto.appendChild objDom.createElement("vICMSST")
  VarObjeto.getElementsByTagName("vICMSST").Item(0).Text = Replace(Format(TBCST!Valor_ICMS_ST, "0.#0"), ",", ".") 'vICMSST_icms
'========================================================
'  Falta acrescentar
'========================================================
'  VarObjeto.appendChild objDom.createElement("vICMSDeson")
'  VarObjeto.getElementsByTagName("vICMSDeson").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'  VarObjeto.appendChild objDom.createElement("motDesICMS")
'  VarObjeto.getElementsByTagName("motDesICMS").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
'=========================================
' Chama calculo de substituição tributária
'=========================================
    Proc_XML_IMP_SUBST_TRIBUT
'='========================================
End If

'=================================================================================================
' GRUPO 90
'=================================================================================================
If FimCST = "90" Then

VarObjeto = "objICMS" & FimCST
VarObjetonome = "ICMS" & FimCST

Set VarObjeto = objDom.createElement(VarObjetonome)
objICMS.appendChild VarObjeto
  VarObjeto.appendChild objDom.createElement("orig")
  VarObjeto.getElementsByTagName("orig").Item(0).Text = TBCST!Origem_mercadoria
  VarObjeto.appendChild objDom.createElement("CST")
  VarObjeto.getElementsByTagName("CST").Item(0).Text = FimCST
  VarObjeto.appendChild objDom.createElement("modBC")
  VarObjeto.getElementsByTagName("modBC").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao) = False, TBCST!Modalidade_determinacao, "0") 'modBC
  VarObjeto.appendChild objDom.createElement("vBC")
  VarObjeto.getElementsByTagName("vBC").Item(0).Text = Replace(Format(TBCST!Valor_BC, "0.#0"), ",", ".")  'vBC
  VarObjeto.appendChild objDom.createElement("pRedBC")
  VarObjeto.getElementsByTagName("pRedBC").Item(0).Text = Replace(Format(TBCST!Percentual_reducao_BC, "0.#000"), ",", ".") 'pRedBC
  VarObjeto.appendChild objDom.createElement("pICMS")
  VarObjeto.getElementsByTagName("pICMS").Item(0).Text = Replace(Format(TBProduto!int_ICMS, "0.#0"), ",", ".")  'pICMS
  VarObjeto.appendChild objDom.createElement("vICMS")
  VarObjeto.getElementsByTagName("vICMS").Item(0).Text = Replace(Format(TBCST!Valor_ICMS, "0.#0"), ",", ".")  'vICMS_icms
  VarObjeto.appendChild objDom.createElement("modBCST")
  VarObjeto.getElementsByTagName("modBCST").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
'  FFalta acrescentar
'========================================================
'  VarObjeto.appendChild objDom.createElement("pMVAST")
'  VarObjeto.getElementsByTagName("pMVAST").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
  VarObjeto.appendChild objDom.createElement("pRedBCST")
  VarObjeto.getElementsByTagName("pRedBCST").Item(0).Text = Replace(Format(TBCST!Percentual_reducao_BC_ST, "0.#000"), ",", ".") 'pRedBCST
  VarObjeto.appendChild objDom.createElement("vBCST")
  VarObjeto.getElementsByTagName("vBCST").Item(0).Text = Replace(Format(TBCST!Valor_BC_ST, "0.#0"), ",", ".") 'vBCST
  VarObjeto.appendChild objDom.createElement("pICMSST")
  VarObjeto.getElementsByTagName("pICMSST").Item(0).Text = Replace(Format(TBCST!Aliquota_imposto_ST, "0.#000"), ",", ".") 'pICMSST
  VarObjeto.appendChild objDom.createElement("vICMSST")
  VarObjeto.getElementsByTagName("vICMSST").Item(0).Text = Replace(Format(TBCST!Valor_ICMS_ST, "0.#0"), ",", ".") 'vICMSST_icms
'========================================================
'  Falta acrescentar
'========================================================
'  VarObjeto.appendChild objDom.createElement("vICMSDeson")
'  VarObjeto.getElementsByTagName("vICMSDeson").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'  VarObjeto.appendChild objDom.createElement("motDesICMS")
'  VarObjeto.getElementsByTagName("motDesICMS").Item(0).Text = IIf(IsNull(TBCST!Modalidade_determinacao_ST), 4, TBCST!Modalidade_determinacao_ST) 'modBCST
'========================================================
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Proc_XML_IMP_SUBST_TRIBUT()
On Error GoTo tratar_erro
'=================================================================================================
' GRUPO SUBSTIUIÇÃO TRIBUTÁRIA
'===============================================================================================
' SE tag: idDest = 2) com Consumidor Final (tag: indFinal = 1) e Não Contribuinte (tag: indIEDest = 9)
' sendo que a Operação não é de prestação de serviços, ou seja, não possui o Grupo de Tributação ISSQN (Imposto Sobre Serviço de Qualquer Natureza) e não foi informado o Grupo do ICMS para a UF de Destino (tag: ICMSUFDest)
'===============================================================================================
If idDest = "2" And indFinal = "1" And indIEDest = "9" Then ' And VarST = True Then
'ICMSUFDest e dentro dele há os campos: ok
    '
'    vBCUFDest: Valor da Base de Cálculo da UF de Destino - ok
'    vBCFCPUFDest: Valor da Base de Cálculo do Fundo de Combate à Pobreza na UF de Destino
'    pFCPUFDest: Alíquota do Fundo de Combate à Pobreza na UF de Destino - ok
'    pICMSUFDest: Alíquota do ICMS da UF de Destino - ok
'    pICMSInter: Alíquota do ICMS Interestadual - ok
'    pICMSInterPart: Alíquota do ICMS Interestadual de Partilha - ok
'    vFCPUFDest: Valor do Fundo de Combate à Pobreza na UF de Destino
'    vICMSUFDest: Valor do ICMS na UF de Destino
'    vICMSUFRemet: Valor do ICMS da UF do Remetente


       'nó objICMSUFDest dentro de objImposto
       Set objICMSUFDest = objDom.createElement("ICMSUFDest")
       objImposto.appendChild objICMSUFDest
       'Abre ICMSUFDest==================================================================================================
       ' Busca dados de impostos
            ProcBuscaTributos IIf(IsNull(TBProduto!ID_CF), 0, TBProduto!ID_CF)
            ProcVerificaRegiao TBproducao!txt_UF, TBproducao!Id_Int_Cliente, TBproducao!txt_Razao_Nome
       '=================================================================================================================
           objICMSUFDest.appendChild objDom.createElement("vBCUFDest")
           objICMSUFDest.getElementsByTagName("vBCUFDest").Item(0).Text = Replace(Format(TBCST!Valor_BC_ICMS_UF_dest, "0.#0"), ",", ".") 'Replace(IIf(IsNull(TBCST!Valor_BC_ICMS_UF_dest), 0, TBCST!Valor_BC_ICMS_UF_dest), ",", ".")
           
           objICMSUFDest.appendChild objDom.createElement("pFCPUFDest") '1
           objICMSUFDest.getElementsByTagName("pFCPUFDest").Item(0).Text = Replace(Format(IIf(IsNull(TBCST!Percentual_FCP), 4, TBCST!Percentual_FCP), "0.#0"), ",", ".")
           
           objICMSUFDest.appendChild objDom.createElement("pICMSUFDest") '2
               Set TBFIltro = CreateObject("adodb.recordset")
               TBFIltro.Open "Select ICMS_interno from regioes where UF = '" & TBproducao!txt_UF & "'", Conexao, adOpenKeyset, adLockOptimistic
               If TBFIltro.EOF = False Then
                   objICMSUFDest.getElementsByTagName("pICMSUFDest").Item(0).Text = IIf(IsNull(TBFIltro!ICMS_interno), "0.00", Replace(Format(TBFIltro!ICMS_interno, "0.#0"), ",", ".")) 'Format(TBFIltro!ICMS_Interno, "0.#0") 'IIf(IsNull(TBFIltro!ICMS_Interno), 0, TBFIltro!ICMS_Interno)
               Else
                   objICMSUFDest.getElementsByTagName("pICMSUFDest").Item(0).Text = "0.00"
               End If
           vICMSUFDest = (TBCST!Valor_BC_ICMS_UF_dest * IIf(IsNull(TBFIltro!ICMS_interno), 0, TBFIltro!ICMS_interno)) / 100
           ttvICMSUFDest = ttvICMSUFDest + vICMSUFDest
               'TBFIltro.Close
           
           'vICMSUFDest = TBCST!Valor_BC_ICMS_UF_dest * TBFIltro!ICMS_Interno
           
           objICMSUFDest.appendChild objDom.createElement("pICMSInter") '3
           strvRegiao = vRegiao(0, 1)
           
           objICMSUFDest.getElementsByTagName("pICMSInter").Item(0).Text = Replace(Format(strvRegiao, "0.#0"), ",", ".")

           objICMSUFDest.appendChild objDom.createElement("pICMSInterPart") '4
           objICMSUFDest.getElementsByTagName("pICMSInterPart").Item(0).Text = Replace(Format(IIf(IsNull(TBCST!Percentual_provisorio), 0, TBCST!Percentual_provisorio), "0.#0"), ",", ".")
           
           objICMSUFDest.appendChild objDom.createElement("vFCPUFDest") '5
           'Debug.print Replace(Format(IIf(IsNull(TBCST!Valor_ICMS_FCP), 0, TBCST!Valor_ICMS_FCP), "0.#0"), ",", ".")
           objICMSUFDest.getElementsByTagName("vFCPUFDest").Item(0).Text = Replace(Format(IIf(IsNull(TBCST!Valor_ICMS_FCP), 0, TBCST!Valor_ICMS_FCP), "0.#0"), ",", ".")
           
           objICMSUFDest.appendChild objDom.createElement("vICMSUFDest") '6
           objICMSUFDest.getElementsByTagName("vICMSUFDest").Item(0).Text = Replace(Format(IIf(IsNull(TBCST!Valor_ICMS_INT_UF_dest), 0, TBCST!Valor_ICMS_INT_UF_dest), "0.#0"), ",", ".")
           
           objICMSUFDest.appendChild objDom.createElement("vICMSUFRemet") '7
           objICMSUFDest.getElementsByTagName("vICMSUFRemet").Item(0).Text = Replace(Format(IIf(IsNull(TBCST!Valor_ICMS_INT_UF_rem), "0.00", TBCST!Valor_ICMS_INT_UF_rem), "0.#0"), ",", ".")
               
       'Fecha ICMSUFDest=================================================================================================
   End If
   TBCST.Close
'VarST = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub Proc_XML_IMP_IPI()
On Error GoTo tratar_erro
'===============================================================================
' Verifica se tem IPI
'===============================================================================
   Set TBIPI = CreateObject("adodb.recordset")
   TBIPI.Open "Select * from tbl_NaturezaOperacao where IDCountCfop = " & TBProduto!ID_CFOP, Conexao, adOpenKeyset, adLockReadOnly
   If TBIPI.EOF = False Then
    TemICMS = TBIPI!Txt_ICMS
    TemIPI = TBIPI!txt_IPI
    TemPIS = TBIPI!TemPIS
    TemCOFINS = TBIPI!TemCOFINS
    End If
   TBIPI.Close

If TemIPI = "SIM" Then
'===============================================================================
' CST DO IPI
'===============================================================================
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
 End If
'=================================================================================================
' GRUPO IPITrib
'=================================================================================================
 If FimCST = "00" Or FimCST = "49" Or FimCST = "50" Or FimCST = "99" Then
       Set objIPITrib = objDom.createElement("IPITrib")
        objIPI.appendChild objIPITrib
        objIPITrib.appendChild objDom.createElement("CST") '0
        objIPITrib.getElementsByTagName("CST").Item(0).Text = FimCST
        objIPITrib.appendChild objDom.createElement("vBC")
        objIPITrib.getElementsByTagName("vBC").Item(0).Text = Replace(Format(TBCST!Valor_BC, "0.00"), ",", ".")
        objIPITrib.appendChild objDom.createElement("pIPI")
        objIPITrib.getElementsByTagName("pIPI").Item(0).Text = Replace(Format(TBProduto!int_IPI, "0.0000"), ",", ".")
       ' objIPITrib.appendChild objDom.createElement("qUnid") '2
       ' objIPITrib.getElementsByTagName("qUnid").Item(0).Text = "0"
       ' objIPITrib.appendChild objDom.createElement("vUnid") '3
       ' objIPITrib.getElementsByTagName("vUnid").Item(0).Text = "0"
        objIPITrib.appendChild objDom.createElement("vIPI")
       If IsNull(TBProduto!dbl_ValorTotal) = True Or TBProduto!dbl_ValorTotal = 0 Then
        objIPITrib.getElementsByTagName("vIPI").Item(0).Text = Replace(Format((TBCST!Valor_BC * TBProduto!int_IPI) / 100, "0.00"), ",", ".")
       Else
      '    If Devolucao = False Then
           objIPITrib.getElementsByTagName("vIPI").Item(0).Text = Replace(Format(TBProduto!dbl_valoripi, "0.00"), ",", ".")
       '   Else
       '     objIPITrib.getElementsByTagName("vIPI").Item(0).Text = "0.00"
       '   End If
       End If
 End If
 
'=================================================================================================
' GRUPO IPINT
'=================================================================================================
   'Verifica o grupo do CST do IPI
 If FimCST = "01" Or FimCST = "02" Or FimCST = "03" Or FimCST = "04" Or FimCST = "51" Or FimCST = "52" Or FimCST = "53" Or FimCST = "54" Or FimCST = "55" Then
        Set objIPINT = objDom.createElement("IPINT")
        objIPI.appendChild objIPINT 'objIPINT
        objIPINT.appendChild objDom.createElement("CST") '0
        objIPINT.getElementsByTagName("CST").Item(0).Text = FimCST
 End If
 
'=========================================================================================

TBCST.Close
End If

End If

'==========================================================================================================
' SE FOR ITEM IMPORTADO
'==========================================================================================================
If IsNull(TBProduto!Local_desembaraco) = False And TBProduto!Local_desembaraco <> "" Then

Set TBCiclo = CreateObject("adodb.recordset")
TBCiclo.Open "Select * from tbl_Detalhes_Nota_NFe where ID_item = " & TBProduto!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
If TBCiclo.EOF = False Then

   'nó II dentro de objImposto
   Set objII = objDom.createElement("II")
   objImposto.appendChild objII
   'Abre objII==================================================================================================
   objII.appendChild objDom.createElement("vBC")
   objII.childNodes(0).Text = Replace(Format(TBCiclo!Valor_BC_importacao, "0.00"), ",", ".")
   objII.appendChild objDom.createElement("vDespAdu")
   objII.childNodes(1).Text = Replace(Format(TBCiclo!Valor_despesas, "0.00"), ",", ".")
   objII.appendChild objDom.createElement("vII")
   objII.childNodes(2).Text = Replace(Format(TBCiclo!Valor_imposto_importacao, "0.00"), ",", ".")
   objII.appendChild objDom.createElement("vIOF")
   objII.childNodes(3).Text = Replace(Format(TBCiclo!Valor_imposto_OperacoesFinanceiras, "0.00"), ",", ".")
   'Fecha objII==================================================================================================
End If
TBCiclo.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub Proc_XML_IMP_PIS()
On Error GoTo tratar_erro

'===========================================================================
'CST DO PIS
'===========================================================================
 If IsNull(TBProduto!CST_PIS) = False And TBProduto!CST_PIS <> "" Then
   FimCST = Right(TBProduto!CST_PIS, 2)
   Set TBCST = CreateObject("adodb.recordset")
   TBCST.Open "Select * from tbl_Detalhes_Nota_CST_PIS where id_item = " & TBProduto!Int_codigo, Conexao, adOpenKeyset, adLockReadOnly
   If TBCST.EOF = False Then
    'nó objPis dentro de objImposto
    Set objPis = objDom.createElement("PIS")
    objImposto.appendChild objPis
'=================================================================================================
' GRUPO PISAliq
'=================================================================================================
 If FimCST = "01" Or FimCST = "02" Then
 Set objPISAliq = objDom.createElement("PISAliq")
      objPis.appendChild objPISAliq
        'Abre Pis==================================================================================================
        objPISAliq.appendChild objDom.createElement("CST")
        objPISAliq.getElementsByTagName("CST").Item(0).Text = FimCST
        objPISAliq.appendChild objDom.createElement("vBC")
        objPISAliq.getElementsByTagName("vBC").Item(0).Text = Replace(Format(TBCST!Valor_BC, "0.#0"), ",", ".")
        objPISAliq.appendChild objDom.createElement("pPIS")
        objPISAliq.getElementsByTagName("pPIS").Item(0).Text = Replace(Format(TBProduto!PIS_Prod, "0.#000"), ",", ".")
        objPISAliq.appendChild objDom.createElement("vPIS")
        objPISAliq.getElementsByTagName("vPIS").Item(0).Text = Replace(Format(TBProduto!Total_PIS_prod, "0.#0"), ",", ".")
End If
'=================================================================================================
' GRUPO PISQtde
'=================================================================================================
 If FimCST = "03" Then
  Set objPISQtde = objDom.createElement("PISQtde")
      objPis.appendChild objPISQtde
        'Abre Pis==================================================================================================
        objPISQtde.appendChild objDom.createElement("CST")
        objPISQtde.getElementsByTagName("CST").Item(0).Text = FimCST
        objPISQtde.appendChild objDom.createElement("qBCProd") 'qBCProd
        objPISQtde.getElementsByTagName("qBCProd").Item(0).Text = Replace(Format(TBProduto!int_Qtd, "0.#0"), ",", ".")
        objPISQtde.appendChild objDom.createElement("vAliqProd")
        objPISQtde.getElementsByTagName("vAliqProd").Item(0).Text = Replace(Format(TBProduto!PIS_Prod, "0.#000"), ",", ".")
        objPISQtde.appendChild objDom.createElement("vPIS")
        objPISQtde.getElementsByTagName("vPIS").Item(0).Text = Replace(Format(TBProduto!Total_PIS_prod, "0.#0"), ",", ".")
End If
'=================================================================================================
' GRUPO PISNT
'=================================================================================================
 If FimCST = "04" Or FimCST = "05" Or FimCST = "06" Or FimCST = "07" Or FimCST = "08" Or FimCST = "09" Then

  Set objPISNT = objDom.createElement("PISNT")
      objPis.appendChild objPISNT
        'Abre Pis==================================================================================================
        objPISNT.appendChild objDom.createElement("CST")
        objPISNT.getElementsByTagName("CST").Item(0).Text = FimCST
        
End If

'=================================================================================================
' GRUPO PISOutr
'=================================================================================================
 If FimCST = "50" Or FimCST = "51" Or FimCST = "52" Or FimCST = "53" Or FimCST = "54" Or FimCST = "55" Or FimCST = "56" Or FimCST = "60" Or FimCST = "61" Or FimCST = "62" Or FimCST = "63" Or FimCST = "64" Or FimCST = "65" Or FimCST = "66" Or FimCST = "67" Or FimCST = "70" Or FimCST = "71" Or FimCST = "72" Or FimCST = "73" Or FimCST = "74" Or FimCST = "75" Or FimCST = "98" Or FimCST = "99" Then
  Set objPISOutr = objDom.createElement("PISOutr")
      objPis.appendChild objPISOutr
        'Abre Pis==================================================================================================
        objPISOutr.appendChild objDom.createElement("CST")
        objPISOutr.getElementsByTagName("CST").Item(0).Text = FimCST
        objPISOutr.appendChild objDom.createElement("vBC")
        objPISOutr.getElementsByTagName("vBC").Item(0).Text = Replace(Format(TBCST!Valor_BC, "0.#0"), ",", ".")
        objPISOutr.appendChild objDom.createElement("pPIS")
        objPISOutr.getElementsByTagName("pPIS").Item(0).Text = Replace(Format(TBProduto!PIS_Prod, "0.#000"), ",", ".")
        objPISOutr.appendChild objDom.createElement("vPIS")
        objPISOutr.getElementsByTagName("vPIS").Item(0).Text = Replace(Format(TBProduto!Total_PIS_prod, "0.#0"), ",", ".")
        'objPISOutr.appendChild objDom.createElement("qBCProd")
        'objPISOutr.getElementsByTagName("qBCProd").Item(0).Text = Replace(Format(TBProduto!int_Qtd, "0.#0"), ",", ".")
        'objPISOutr.appendChild objDom.createElement("vAliqProd")
        'objPISOutr.getElementsByTagName("vAliqProd").Item(0).Text = Replace(TBProduto!PIS_Prod, ",", ".")
 End If
 
 If FimCST = "49" Then
  Set objPISOutr = objDom.createElement("PISOutr")
      objPis.appendChild objPISOutr
        'Abre Pis==================================================================================================
        objPISOutr.appendChild objDom.createElement("CST")
        objPISOutr.getElementsByTagName("CST").Item(0).Text = FimCST
        objPISOutr.appendChild objDom.createElement("vBC")
        objPISOutr.getElementsByTagName("vBC").Item(0).Text = Replace(Format(TBCST!Valor_BC, "0.#0"), ",", ".")
        objPISOutr.appendChild objDom.createElement("pPIS")
        objPISOutr.getElementsByTagName("pPIS").Item(0).Text = Replace(Format(TBProduto!PIS_Prod, "0.#000"), ",", ".")
        objPISOutr.appendChild objDom.createElement("vPIS")
        objPISOutr.getElementsByTagName("vPIS").Item(0).Text = Replace(Format(TBProduto!Total_PIS_prod, "0.#0"), ",", ".")
        'objPISOutr.appendChild objDom.createElement("qBCProd")
        'objPISOutr.getElementsByTagName("qBCProd").Item(0).Text = Replace(Format(TBProduto!int_Qtd, "0.#0"), ",", ".")
        'objPISOutr.appendChild objDom.createElement("vAliqProd")
        'objPISOutr.getElementsByTagName("vAliqProd").Item(0).Text = Replace(TBProduto!PIS_Prod, ",", ".")
 End If
End If

TBCST.Close
End If

'=================================================================================================
'Fecha cadastros do Pis
'================================================================================================
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub Proc_XML_IMP_COFINS()
On Error GoTo tratar_erro

'=================================================================================================
' CST DO COFINS
'=================================================================================================
If IsNull(TBProduto!CST_Cofins) = False And TBProduto!CST_Cofins <> "" Then
  FimCST = Right(TBProduto!CST_Cofins, 2)
  Set TBCST = CreateObject("adodb.recordset")
  TBCST.Open "Select * from tbl_Detalhes_Nota_CST_Cofins where id_item = " & TBProduto!Int_codigo, Conexao, adOpenKeyset, adLockReadOnly
  If TBCST.EOF = False Then
   Set objCofins = objDom.createElement("COFINS")
   objImposto.appendChild objCofins
   
'=================================================================================================
' GRUPO COFINSAliq
'=================================================================================================
  If FimCST = "01" Or FimCST = "02" Then
   Set objCOFINSAliq = objDom.createElement("COFINSAliq")
   objCofins.appendChild objCOFINSAliq
   objCOFINSAliq.appendChild objDom.createElement("CST")
   objCOFINSAliq.getElementsByTagName("CST").Item(0).Text = FimCST
   objCOFINSAliq.appendChild objDom.createElement("vBC")
   objCOFINSAliq.getElementsByTagName("vBC").Item(0).Text = Replace(Format(TBCST!Valor_BC, "0.#0"), ",", ".")
   objCOFINSAliq.appendChild objDom.createElement("pCOFINS")
   objCOFINSAliq.getElementsByTagName("pCOFINS").Item(0).Text = Replace(Format(TBProduto!Cofins_Prod, "0.#000"), ",", ".")
   objCOFINSAliq.appendChild objDom.createElement("vCOFINS")
   objCOFINSAliq.getElementsByTagName("vCOFINS").Item(0).Text = Replace(Format(TBProduto!Total_Cofins_prod, "0.#0"), ",", ".")
  End If
 
'=================================================================================================
' GRUPO COFINSQtde
'=================================================================================================
  If FimCST = "03" Then
   Set objCOFINSQtde = objDom.createElement("COFINSQtde")
   objCofins.appendChild objCOFINSQtde
   objCOFINSQtde.appendChild objDom.createElement("CST")
   objCOFINSQtde.getElementsByTagName("CST").Item(0).Text = FimCST
   objCOFINSQtde.appendChild objDom.createElement("qBCProd")
   objCOFINSQtde.getElementsByTagName("qBCProd").Item(0).Text = Replace(TBProduto!int_Qtd, ",", ".")
   objCOFINSQtde.appendChild objDom.createElement("vAliqProd")
   objCOFINSQtde.getElementsByTagName("vAliqProd").Item(0).Text = Replace(Format(TBProduto!Cofins_Prod, "0.#000"), ",", ".")
   objCOFINSQtde.appendChild objDom.createElement("vCOFINS")
   objCOFINSQtde.getElementsByTagName("vCOFINS").Item(0).Text = Replace(Format(TBProduto!Total_Cofins_prod, "0.#0"), ",", ".")
  End If
  
'=================================================================================================
' GRUPO COFINSNT
'=================================================================================================
  If FimCST = "04" Or FimCST = "05" Or FimCST = "06" Or FimCST = "07" Or FimCST = "08" Or FimCST = "09" Then
   Set objCOFINSNT = objDom.createElement("COFINSNT")
   objCofins.appendChild objCOFINSNT
   objCOFINSNT.appendChild objDom.createElement("CST")
   objCOFINSNT.getElementsByTagName("CST").Item(0).Text = FimCST
  End If
  
'=================================================================================================
' GRUPO COFINSOutr
'=================================================================================================
   If FimCST = "49" Or FimCST = "50" Or FimCST = "51" Or FimCST = "52" Or FimCST = "53" Or FimCST = "54" Or FimCST = "55" Or FimCST = "56" Or FimCST = "60" Or FimCST = "61" Or FimCST = "62" Or FimCST = "63" Or FimCST = "64" Or FimCST = "65" Or FimCST = "66" Or FimCST = "67" Or FimCST = "70" Or FimCST = "71" Or FimCST = "72" Or FimCST = "73" Or FimCST = "74" Or FimCST = "75" Or FimCST = "98" Or FimCST = "99" Then
    Set objCOFINSOutr = objDom.createElement("COFINSOutr")
    objCofins.appendChild objCOFINSOutr
    objCOFINSOutr.appendChild objDom.createElement("CST")
    objCOFINSOutr.getElementsByTagName("CST").Item(0).Text = FimCST
    objCOFINSOutr.appendChild objDom.createElement("vBC")
    objCOFINSOutr.getElementsByTagName("vBC").Item(0).Text = Replace(Format(TBCST!Valor_BC, "0.#0"), ",", ".")
    objCOFINSOutr.appendChild objDom.createElement("pCOFINS")
    objCOFINSOutr.getElementsByTagName("pCOFINS").Item(0).Text = Replace(Format(TBProduto!Cofins_Prod, "0.#000"), ",", ".")
    objCOFINSOutr.appendChild objDom.createElement("vCOFINS")
    objCOFINSOutr.getElementsByTagName("vCOFINS").Item(0).Text = Replace(Format(TBProduto!Total_Cofins_prod, "0.#0"), ",", ".")
    'objCOFINSOutr.appendChild objDom.createElement("qBCProd")
    'objCOFINSOutr.getElementsByTagName("qBCProd").Item(0).Text = Replace(TBProduto!int_Qtd, ",", ".")
    'objCOFINSOutr.appendChild objDom.createElement("vAliqProd")
    'objCOFINSOutr.getElementsByTagName("vAliqProd").Item(0).Text = Replace(TBProduto!Cofins_Prod, ",", ".")
 End If
End If
TBCST.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub Proc_XML_Impostos()
On Error GoTo tratar_erro
'===============================================
' INICIO DOS CADASTROS DOS IMPOSTOS DO PRODUTO
'===============================================
Set objImposto = objDom.createElement("imposto")
objDet.appendChild objImposto
'Abre objImposto================================
If IsNull(TBProduto!Valor_aprox_tributos) = False And TBProduto!Valor_aprox_tributos <> "" And TBProduto!Valor_aprox_tributos <> "0" Then
    objImposto.appendChild objDom.createElement("vTotTrib") '0
    objImposto.getElementsByTagName("vTotTrib").Item(0).Text = Replace(Format(TBProduto!Valor_aprox_tributos, "0.#0"), ",", ".")
End If

If IsNull(TBProduto!txt_CST) = False And TBProduto!txt_CST <> "" Then
    If Len(TBProduto!txt_CST) = 4 Then FimCST = Right(TBProduto!txt_CST, 3) Else FimCST = Right(TBProduto!txt_CST, 2)
    Set TBCST = CreateObject("adodb.recordset")
    TBCST.Open "Select * from tbl_Detalhes_Nota_CST_ICMS where id_item = '" & TBProduto!Int_codigo & "'", Conexao, adOpenKeyset, adLockReadOnly
End If

Proc_XML_IMP_ICMS
Proc_XML_IMP_IPI
Proc_XML_IMP_PIS
Proc_XML_IMP_COFINS

'=====================================================================================
'Se for pra fora do estado, venda pra consumidor final e não for contrinuinte de icms
'=====================================================================================

If idDest = "2" And indFinal = "1" And indIEDest = "9" Then ' And VarST = True Then
    If IsNull(TBProduto!txt_CST) = False And TBProduto!txt_CST <> "" Then
        If Len(TBProduto!txt_CST) = 4 Then FimCST = Right(TBProduto!txt_CST, 3) Else FimCST = Right(TBProduto!txt_CST, 2)
        Set TBCST = CreateObject("adodb.recordset")
        TBCST.Open "Select * from tbl_Detalhes_Nota_CST_ICMS where id_item = '" & TBProduto!Int_codigo & "'", Conexao, adOpenKeyset, adLockReadOnly
    End If
    
    
    Proc_XML_IMP_SUBST_TRIBUT
End If

'==============================================
'FECHA OS IMPOSTOS DO PRODUTO
'==============================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub Proc_Gerar_cNF()
On Error GoTo tratar_erro

Dim NumAleatorio As Double
Dim StrNumeroFormatado As String
NumAleatorio = Int(Rnd * 99999999)
StrNumeroFormatado = LTrim(RTrim(str(NumAleatorio)))
Var3 = String(8 - Len(StrNumeroFormatado), "0") & StrNumeroFormatado

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub proc_XML_Identificacao()
On Error GoTo tratar_erro

'no ide dentro de Envio (A01)
Set objIde = objDom.createElement("ide")
objinfNFe.appendChild objIde

'Abre ide=================================================================================================
    'filhos ide
    
    FamiliaAntiga = RemoveAccents(TBproducao!Cidade) 'Empresa
    
    objIde.appendChild objDom.createElement("cUF")
    objIde.getElementsByTagName("cUF").Item(0).Text = FunVerificaCodUF(FamiliaAntiga, TBproducao!UF)
    chCodUF = FunVerificaCodUF(FamiliaAntiga, TBproducao!UF)
    
    If txtchNFe.Text = "" Then
     Proc_Gerar_cNF
    Else
     Var3 = Right(txtchNFe.Text, 9)
     Var3 = Left(Var3, 8)
    End If
    
    'Debug.print Var3
    objIde.appendChild objDom.createElement("cNF")
    objIde.getElementsByTagName("cNF").Item(0).Text = Var3 'Right(txtNota, 8)
    'chNNfe = FunTamanhoTextoZeroEsq(txtNota, 8)
    chNNfe = Right(txtNota, 9)
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select CFOP.Txt_descricao from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.ID_nota = " & txtID_nota & " order by NFP.Int_codigo", Conexao, adOpenKeyset, adLockReadOnly
    If TBCFOP.EOF = False Then
        objIde.appendChild objDom.createElement("natOp")
        objIde.getElementsByTagName("natOp").Item(0).Text = RemoveAccents(Trim(TBCFOP!Txt_descricao))
    End If
    TBCFOP.Close
    
    objIde.appendChild objDom.createElement("mod")
    Modelo = Left(frmFaturamento_Prod_Serv.Cmb_modelo, 2)
    objIde.getElementsByTagName("mod").Item(0).Text = Modelo
    chModelo = Modelo
    
    objIde.appendChild objDom.createElement("serie")
    objIde.getElementsByTagName("serie").Item(0).Text = txtSerie
    chSerie = txtSerie.Text
    
    objIde.appendChild objDom.createElement("nNF")
    objIde.getElementsByTagName("nNF").Item(0).Text = Format(txtNota, "0")
    objIde.appendChild objDom.createElement("dhEmi")
    
    objIde.getElementsByTagName("dhEmi").Item(0).Text = Format(TBproducao!dt_DataEmissao, "yyyy-mm-dd") & "T" & Left(TBproducao!Hora_emissao, 8) & FunVerifFusoHorario(True)
    chDTEmissao = Format(TBproducao!dt_DataEmissao, "yymm") '& Month(TBproducao!dt_Dataemissao)
    
  If NFCe = False Then
    objIde.appendChild objDom.createElement("dhSaiEnt")
    objIde.getElementsByTagName("dhSaiEnt").Item(0).Text = Format(TBproducao!dt_Saida_Entrada, "yyyy-mm-dd") & "T" & Left(TBproducao!txt_Hora_Saida, 8) & FunVerifFusoHorario(True)
End If

    objIde.appendChild objDom.createElement("tpNF")
    If TBproducao!int_TipoNota = 1 Then
        objIde.getElementsByTagName("tpNF").Item(0).Text = 1
    Else
        objIde.getElementsByTagName("tpNF").Item(0).Text = 0
    End If

    
'=================================================================
' Operação interna ou externa (Exportação
'=================================================================
    objIde.appendChild objDom.createElement("idDest")
    objIde.getElementsByTagName("idDest").Item(0).Text = idDest

    objIde.appendChild objDom.createElement("cMunFG")
    objIde.getElementsByTagName("cMunFG").Item(0).Text = FunVerificaCodMunicipio(FamiliaAntiga, TBproducao!UF)
    objIde.appendChild objDom.createElement("tpImp")

    objIde.getElementsByTagName("tpImp").Item(0).Text = IIf(NFCe = False, 1, 4) 'DANFE 1 = Retrato - 2 = Paisagem no manual
    objIde.appendChild objDom.createElement("tpEmis")
    
    
    objIde.getElementsByTagName("tpEmis").Item(0).Text = TBproducao!Forma_emissao
    chFormaEmissao = TBproducao!Forma_emissao
    
    objIde.appendChild objDom.createElement("tpAmb")
    objIde.getElementsByTagName("tpAmb").Item(0).Text = tpAmb '1-produção 2-Homologação
    objIde.appendChild objDom.createElement("finNFe")
    objIde.getElementsByTagName("finNFe").Item(0).Text = TBproducao!Finalidade_emissao
    objIde.appendChild objDom.createElement("indFinal")
    objIde.getElementsByTagName("indFinal").Item(0).Text = TBproducao!Consumidor_final
    indFinal = TBproducao!Consumidor_final
    objIde.appendChild objDom.createElement("indPres")
    objIde.getElementsByTagName("indPres").Item(0).Text = TBproducao!Presenca_comprador
    objIde.appendChild objDom.createElement("procEmi")
    objIde.getElementsByTagName("procEmi").Item(0).Text = 0
    objIde.appendChild objDom.createElement("verProc")
    objIde.getElementsByTagName("verProc").Item(0).Text = "5.0"
    'objIde.appendChild objDom.createElement("xJust")
    'objIde.childNodes(17).Text = ""
    'objIde.appendChild objDom.createElement("dhCont")
    'objIde.childNodes(18).Text = ""
'    If email <> "" Then
'        objIde.appendChild objDom.createElement("EmailArquivos")
'        objIde.getElementsByTagName("EmailArquivos").Item(0).Text = Trim(email)
'    End If
    'objIde.appendChild objDom.createElement("NumeroPedido")
    'objIde.childNodes(19).Text = ""
    
 'Se for nota de devolução tem que referenciar as notas de devolução
Contador = 0
    'Abre NFRef=================================================================================================
        'filhos NFRef
        Set TBCiclo = CreateObject("adodb.recordset")
        TBCiclo.Open "Select ID_nota_relacionada AS ID from Faturamento_Relacionamento where ID_nota = " & txtID_nota & " group by ID_nota_relacionada", Conexao, adOpenKeyset, adLockReadOnly
        If TBCiclo.EOF = True Then
            Set TBCiclo = CreateObject("adodb.recordset")
            TBCiclo.Open "Select ID_nota AS ID from Faturamento_Relacionamento where ID_nota_relacionada = " & txtID_nota & " group by ID_nota", Conexao, adOpenKeyset, adLockReadOnly
        Else

        
        Do While TBCiclo.EOF = False
            'nó objNFRef dentro de Ide
            VarObjeto = "objNFRef" & Contador
            'VarObjetonome = "objNFRef" & vFimCST
        
            'Abre NFRef=================================================================================================
                Set TBCarteira = CreateObject("adodb.recordset")
                TBCarteira.Open "Select ID, int_NotaFiscal, txt_Municipio, txt_UF, dt_DataEmissao, txt_CNPJ_CPF, Modelo, Serie from tbl_Dados_Nota_Fiscal where ID = " & TBCiclo!ID, Conexao, adOpenKeyset, adLockOptimistic
                If TBCarteira.EOF = False Then
                    Set TBTempo = CreateObject("adodb.recordset")
                    TBTempo.Open "Select Chave_acesso from tbl_Dados_Nota_Fiscal_NFe where ID_nota = " & TBCarteira!ID & " and Chave_acesso IS NOT NULL and Chave_acesso <> N''", Conexao, adOpenKeyset, adLockOptimistic
                    If TBTempo.EOF = False Then
                        Set VarObjeto = objDom.createElement("NFref")
                        objIde.appendChild VarObjeto
                        VarObjeto.appendChild objDom.createElement("refNFe") '0
                        VarObjeto.childNodes(0).Text = TBTempo!Chave_acesso
                    'Else
                        'objNFRef.appendChild objDom.createElement("cUF_refNFE") '1
                        'objNFRef.appendChild objDom.createElement("AAMM") '2
                        'objNFRef.appendChild objDom.createElement("CNPJ") '3
                        'objNFRef.appendChild objDom.createElement("CPF") '4
                        'objNFRef.appendChild objDom.createElement("mod_refNFE") '5
                        'objNFRef.appendChild objDom.createElement("serie_refNFE") '6
                        'objNFRef.appendChild objDom.createElement("IE_refNFP") '7
                        'objNFRef.appendChild objDom.createElement("RefCte") '8
                        'objNFRef.appendChild objDom.createElement("mod_refECF") '9
                        'objNFRef.appendChild objDom.createElement("nECF_refECF") '10
                        'objNFRef.appendChild objDom.createElement("nCOO_refECF") '11
                    
                        'IDpedido = TBCarteira!int_NotaFiscal
                        'FamiliaAntiga = RemoveAccents(TBCarteira!txt_Municipio)
                        'objNFRef.childNodes(1).Text = FunVerificaCodUF(FamiliaAntiga, TBCarteira!txt_UF)
                        'objNFRef.childNodes(2).Text = Format(TBCarteira!dt_DataEmissao, "YYMM")
                        'objNFRef.childNodes(3).Text = ReturnNumbersOnly(TBCarteira!txt_CNPJ_CPF)
                        'objNFRef.childNodes(5).Text = IIf(Left(TBCarteira!Modelo, 2) = "1B", 1, Left(TBCarteira!Modelo, 2))
                        'objNFRef.childNodes(6).Text = TBCarteira!serie
                        'objNFRef.childNodes(7).Text = IDpedido
                    End If
                    TBTempo.Close
                End If
                TBCarteira.Close
            'Fecha NFRef====================================================================================================
            TBCiclo.MoveNext
            Contador = Contador + 1
        Loop
        Contador = 0
'        TBCiclo.Close
    'Fecha NFRef===================================================================
    End If
'Fecha ide===================================================================
'Debug.print objDom.XML

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub proc_XML_Emitente()
On Error GoTo tratar_erro

'no emit dentro de Envio (A01)
'=================================================================================================
' Dados do emitente dentro dos dados do emitente
'=================================================================================================
Set objEmit = objDom.createElement("emit")
objinfNFe.appendChild objEmit
chCNPJ = ReturnNumbersOnly(TBproducao!CNPJ)
    objEmit.appendChild objDom.createElement("CNPJ")
    objEmit.getElementsByTagName("CNPJ").Item(0).Text = ReturnNumbersOnly(TBproducao!CNPJ)
    'objEmit.appendChild objDom.createElement("CPF_emit")
    'objemit.childNodes(1).Text = ""
    objEmit.appendChild objDom.createElement("xNome")
    objEmit.getElementsByTagName("xNome").Item(0).Text = Trim(RemoveAccents(Left(TBproducao!Razao, 60)))
    objEmit.appendChild objDom.createElement("xFant")
    objEmit.getElementsByTagName("xFant").Item(0).Text = Trim(RemoveAccents(TBproducao!Empresa))
'=================================================================================================
' Endereço do emitente dentro dos dados do emitente
'=================================================================================================
    Set objEnderEmit = objDom.createElement("enderEmit")
    objEmit.appendChild objEnderEmit
    'Abre enderEmit=================================================================================================
        objEnderEmit.appendChild objDom.createElement("xLgr")
        FamiliaAntiga = ""
        If IsNull(TBproducao!Tipo_endereco) = False And TBproducao!Tipo_endereco <> "" Then FamiliaAntiga = TBproducao!Tipo_endereco & ": "
        If FamiliaAntiga <> "" Then FamiliaAntiga = FamiliaAntiga & TBproducao!Endereco Else FamiliaAntiga = TBproducao!Endereco
        objEnderEmit.getElementsByTagName("xLgr").Item(0).Text = Trim(RemoveAccents(FamiliaAntiga))
        
        objEnderEmit.appendChild objDom.createElement("nro")
        objEnderEmit.getElementsByTagName("nro").Item(0).Text = IIf(IsNull(TBproducao!numeroEmpresa) = True, 0, TBproducao!numeroEmpresa)
        
        If IsNull(TBproducao!complemento) = False Then
            objEnderEmit.appendChild objDom.createElement("xCpl") '2
            objEnderEmit.getElementsByTagName("xCpl").Item(0).Text = Trim(TBproducao!complemento)
        End If
        
        objEnderEmit.appendChild objDom.createElement("xBairro") '3
        FamiliaAntiga = ""
        If IsNull(TBproducao!Tipo_bairro) = False And TBproducao!Tipo_bairro <> "" Then FamiliaAntiga = TBproducao!Tipo_bairro & ": "
        If FamiliaAntiga <> "" Then FamiliaAntiga = FamiliaAntiga & TBproducao!Bairro Else Bairro = TBproducao!Bairro
        objEnderEmit.getElementsByTagName("xBairro").Item(0).Text = Trim(RemoveAccents(FamiliaAntiga))
        
        objEnderEmit.appendChild objDom.createElement("cMun") '4
        FamiliaAntiga = RemoveAccents(TBproducao!Cidade)
        objEnderEmit.getElementsByTagName("cMun").Item(0).Text = FunVerificaCodMunicipio(FamiliaAntiga, TBproducao!UF)
        
        objEnderEmit.appendChild objDom.createElement("xMun") '5
        objEnderEmit.getElementsByTagName("xMun").Item(0).Text = RemoveAccents(TBproducao!Cidade)
        objEnderEmit.appendChild objDom.createElement("UF") '6
        objEnderEmit.getElementsByTagName("UF").Item(0).Text = TBproducao!UF
        objEnderEmit.appendChild objDom.createElement("CEP") '7
        objEnderEmit.getElementsByTagName("CEP").Item(0).Text = ReturnNumbersOnly(TBproducao!CEP)
        objEnderEmit.appendChild objDom.createElement("cPais") '8
        objEnderEmit.getElementsByTagName("cPais").Item(0).Text = "1058"
        objEnderEmit.appendChild objDom.createElement("xPais") '9
        objEnderEmit.getElementsByTagName("xPais").Item(0).Text = "BRASIL"
        If IsNull(TBproducao!telefone) = False And TBproducao!telefone <> "" Then
            objEnderEmit.appendChild objDom.createElement("fone") '10
            objEnderEmit.getElementsByTagName("fone").Item(0).Text = ReturnNumbersOnly(TBproducao!telefone)
        End If
    If IsNull(TBproducao!IE) = False And TBproducao!IE <> "" Then
        objEmit.appendChild objDom.createElement("IE")
        objEmit.getElementsByTagName("IE").Item(0).Text = IIf(TBproducao!IE = "ISENTO", "ISENTO", Left(ReturnNumbersOnly(TBproducao!IE), 14))
    End If
    
    objEmit.appendChild objDom.createElement("CRT")
    If TBproducao!Simples = True Then
        objEmit.getElementsByTagName("CRT").Item(0).Text = 1
    ElseIf TBproducao!Simples1 = True Then
        objEmit.getElementsByTagName("CRT").Item(0).Text = 2
    Else
        objEmit.getElementsByTagName("CRT").Item(0).Text = 3
    End If
        
        'If IsNull(TBproducao!Email) = False And TBproducao!Email <> "" Then
        '    objEnderEmit.appendChild objDom.createElement("Email") '11
        '    objEnderEmit.getElementsByTagName("Email").Item(0).Text = TBproducao!Email
        'End If
    'Fecha enderEmit================================================================================================
'Fecha emit================================================================================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub proc_XML_Destinatario()
On Error GoTo tratar_erro

'no dest dentro de Envio (A01)
Set objDest = objDom.createElement("dest")
objinfNFe.appendChild objDest
'Abre objDest=================================================================================================


'idEstrangeiro=========Obrigatório?
If TBproducao!txt_UF = "EX" Then
objDest.appendChild objDom.createElement("idEstrangeiro") '2
objDest.getElementsByTagName("idEstrangeiro").Item(0).Text = "ABC1234" 'ReturnNumbersOnly(TBproducao!txt_CNPJ_CPF)
Else
  'If TBproducao!txt_tipocliente = "E" Or Left(TBproducao!txt_tipocliente, 1) = "J" Then
  If Len(ReturnNumbersOnly(Trim(TBproducao!txt_CNPJ_CPF))) = 14 Then
  objDest.appendChild objDom.createElement("CNPJ") '0
  objDest.getElementsByTagName("CNPJ").Item(0).Text = ReturnNumbersOnly(Trim(TBproducao!txt_CNPJ_CPF)) 'CNPJ
  Else
  objDest.appendChild objDom.createElement("CPF") '1
  objDest.getElementsByTagName("CPF").Item(0).Text = ReturnNumbersOnly(Trim(TBproducao!txt_CNPJ_CPF)) 'CPF
  End If
End If

objDest.appendChild objDom.createElement("xNome") '3

If tpAmb = 1 Then
'==================================================================================================
' Se for emitido em ambiente de Normal
'==================================================================================================
objDest.getElementsByTagName("xNome").Item(0).Text = RemoveAccents(Trim(TBproducao!txt_Razao_Nome))
Else
'==================================================================================================
' Se for emitido em ambiente de homologação
'==================================================================================================
objDest.getElementsByTagName("xNome").Item(0).Text = "NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL" 'Usado para testes em homologação
End If
'==================================================================================================
' Endereço do destinatário
'==================================================================================================
Set objEnderDest = objDom.createElement("enderDest")
objDest.appendChild objEnderDest
objEnderDest.appendChild objDom.createElement("xLgr") '4
objEnderDest.getElementsByTagName("xLgr").Item(0).Text = Trim(RemoveAccents(TBproducao!txt_Endereco))
objEnderDest.appendChild objDom.createElement("nro") '0
objEnderDest.getElementsByTagName("nro").Item(0).Text = IIf(IsNull(TBproducao!Numero) = False, TBproducao!Numero, 0)
objEnderDest.appendChild objDom.createElement("xCpl") '2

'========================================================
' Busca complemento do endereço no cadastro da empresa
'========================================================
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select Complemento from Empresa where Codigo = " & TBproducao!Id_Int_Cliente & " And Empresa = '" & TBproducao!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBClientes.EOF = False Then ' Then
  If TBClientes!complemento = "" Or IsNull(TBClientes!complemento) = True Then
  xCpl = "Sem inf."
  Else
  xCpl = Trim(TBClientes!complemento)
  End If
 objEnderDest.getElementsByTagName("xCpl").Item(0).Text = xCpl
End If
TBClientes.Close

'========================================================
' Busca complemento do endereço no cadastro do cliente
'========================================================
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from Clientes where IDCliente = " & TBproducao!Id_Int_Cliente & " And NomeRazao = '" & TBproducao!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBClientes.EOF = False Then ' Then
  If TBClientes!complemento = "" Or IsNull(TBClientes!complemento) = True Then
  xCpl = "Sem inf."
  Else
  xCpl = Trim(TBClientes!complemento)
  End If
 objEnderDest.getElementsByTagName("xCpl").Item(0).Text = xCpl
Else
TBClientes.Close

'========================================================
' Busca complemento do endereço no cadastro do Fornecedor
'========================================================
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select * from Compras_fornecedores where IDCliente = " & TBproducao!Id_Int_Cliente & " And Nome_Razao = '" & TBproducao!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockReadOnly

  If TBClientes.EOF = False Then
  
        If TBClientes!complemento = "" Or IsNull(TBClientes!complemento) = True Then
        xCpl = "Sem inf."
        Else
        xCpl = Trim(TBClientes!complemento)
        End If
     objEnderDest.getElementsByTagName("xCpl").Item(0).Text = xCpl
  End If
End If

TBClientes.Close
'========================================================

objEnderDest.appendChild objDom.createElement("xBairro") '2
objEnderDest.getElementsByTagName("xBairro").Item(0).Text = Trim(RemoveAccents(Trim(TBproducao!txt_Bairro)))

objEnderDest.appendChild objDom.createElement("cMun") '6
objEnderDest.appendChild objDom.createElement("xMun") '7
objEnderDest.appendChild objDom.createElement("UF") '8

If IsNull(TBproducao!txt_UF) = True Or Trim(TBproducao!txt_UF) = "" Or TBproducao!txt_UF = "EX" Or chkOperacaoExterna = 1 Then
objEnderDest.getElementsByTagName("cMun").Item(0).Text = "9999999"
objEnderDest.getElementsByTagName("xMun").Item(0).Text = "EXTERIOR"
objEnderDest.getElementsByTagName("UF").Item(0).Text = "EX"
Else
FamiliaAntiga = RemoveAccents(TBproducao!txt_Municipio)
objEnderDest.getElementsByTagName("cMun").Item(0).Text = FunVerificaCodMunicipio(FamiliaAntiga, TBproducao!txt_UF)

Cidade = TBproducao!txt_Municipio
Cidade = Replace(Cidade, "d'oeste", "Do Oeste")
'Debug.print Cidade

objEnderDest.getElementsByTagName("xMun").Item(0).Text = RemoveAccents(Cidade)
objEnderDest.getElementsByTagName("UF").Item(0).Text = Trim(TBproducao!txt_UF)
End If

If Len(ReturnNumbersOnly(TBproducao!Txt_CEP)) = 8 Then
objEnderDest.appendChild objDom.createElement("CEP") '9
objEnderDest.getElementsByTagName("CEP").Item(0).Text = Left(ReturnNumbersOnly(Trim(TBproducao!Txt_CEP)), 8)
End If

objEnderDest.appendChild objDom.createElement("cPais") '10
objEnderDest.appendChild objDom.createElement("xPais") '5
'=============================================================================
' Verifica se é fornecedor
'=============================================================================
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select Email, Codigo_pais, Pais, Complemento, RG_IM, Nao_contribuinte_ICMS, Pessoa from Compras_fornecedores where IDCliente = " & TBproducao!Id_Int_Cliente & " and Nome_Razao = '" & TBproducao!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
If (TBFornecedor!Nao_contribuinte_ICMS) = True Then Nao_contribuinte_ICMS = "Sim" Else Nao_contribuinte_ICMS = "Não"
objEnderDest.getElementsByTagName("xPais").Item(0).Text = TBFornecedor!Pais
objEnderDest.getElementsByTagName("cPais").Item(0).Text = TBFornecedor!Codigo_pais
Else
'=============================================================================
' Verifica se é cliente
'=============================================================================
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select Email, Codigo_pais, Pais, Complemento, RG_IM, Nao_contribuinte_ICMS, Tipo from Clientes where IDCliente = " & TBproducao!Id_Int_Cliente & " and NomeRazao = '" & TBproducao!txt_Razao_Nome & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
   If (TBFornecedor!Nao_contribuinte_ICMS) = True Then Nao_contribuinte_ICMS = "Sim" Else Nao_contribuinte_ICMS = "Não"
   objEnderDest.getElementsByTagName("xPais").Item(0).Text = TBFornecedor!Pais
   objEnderDest.getElementsByTagName("cPais").Item(0).Text = TBFornecedor!Codigo_pais
'=============================================================================
' Verifica se para a própria empresa
'=============================================================================
Else
   Set TBFornecedor = CreateObject("adodb.recordset")
   TBFornecedor.Open "Select email, Codigo_pais, Pais, Complemento, IM from Empresa where Codigo = " & TBproducao!Id_Int_Cliente, Conexao, adOpenKeyset, adLockOptimistic
   If TBFornecedor.EOF = False Then
       Nao_contribuinte_ICMS = "Não"
       objEnderDest.getElementsByTagName("xPais").Item(0).Text = TBFornecedor!Pais
       objEnderDest.getElementsByTagName("cPais").Item(0).Text = TBFornecedor!Codigo_pais
   End If
End If
End If
TBFornecedor.Close

'Adiciona o telefone
If IsNull(TBproducao!txt_Fone_Fax) = False And TBproducao!txt_Fone_Fax <> "" Then
objEnderDest.appendChild objDom.createElement("fone") '11
objEnderDest.getElementsByTagName("fone").Item(0).Text = Right(ReturnNumbersOnly(TBproducao!txt_Fone_Fax), 14)
End If
                
'================================================================================
'Fecha endereço do Destinatário
'================================================================================
'Aqui verifica se não é contribuinte do icms se for informa a inscrição estadual
'1=Contribuinte ICMS (informar a IE do destinatário);
'2=Contribuinte isento de Inscrição no cadastro de Contribuintes
'9=Não Contribuinte, que pode ou não possuir Inscrição Estadual no Cadastro de Contribuintes do ICMS.

'Nota 1: No caso de NFC-e informar indIEDest=9 e não informar a tag IE do destinatário;
'Nota 2: No caso de operação com o Exterior informar indIEDest=9 e não informar a tag IE do destinatário;
'Nota 3: No caso de Contribuinte Isento de Inscrição (indIEDest=2), não informar a tag IE do destinatário.
'================================================================================

objDest.appendChild objDom.createElement("indIEDest") 'IE
'Se for pessoa juridica
If TBproducao!txt_tipocliente = "E" Or Left(TBproducao!txt_tipocliente, 1) = "J" Then
    If TBproducao!txt_UF = "EX" Then 'Se for exportação
       objDest.getElementsByTagName("indIEDest").Item(0).Text = 9 '9=Não Contribuinte, que pode ou não possuir Inscrição Estadual no Cadastro de Contribuintes do ICMS.
       indIEDest = "9"
    Else ' Emissão normal pessoa juridica
          If Nao_contribuinte_ICMS = "Sim" Then 'Não Contribuinte icms
            objDest.getElementsByTagName("indIEDest").Item(0).Text = 9 '9=Não Contribuinte, que pode ou não possuir Inscrição Estadual no Cadastro de Contribuintes do ICMS.
            indIEDest = "9"
          Else
              If IsNull(TBproducao!txt_IE_Cliente) = True Or TBproducao!txt_IE_Cliente = "ISENTO" Or TBproducao!txt_IE_Cliente = "" Or TBproducao!txt_IE_Cliente = "Isento" Then
                  objDest.getElementsByTagName("indIEDest").Item(0).Text = 2 '2=Contribuinte isento de Inscrição no cadastro de Contribuintes
                  indIEDest = "2"
              Else
                  objDest.getElementsByTagName("indIEDest").Item(0).Text = 1 '1=Contribuinte ICMS (informar a IE do destinatário);
                  objDest.appendChild objDom.createElement("IE")
                  objDest.getElementsByTagName("IE").Item(0).Text = Left(ReturnNumbersOnly(TBproducao!txt_IE_Cliente), 14)
                  indIEDest = "1"
              End If
          End If
    End If
Else ' Se for fisica
       If Nao_contribuinte_ICMS = "Sim" Then 'Não Contribuinte icms
            objDest.getElementsByTagName("indIEDest").Item(0).Text = 9 '9=Não Contribuinte, que pode ou não possuir Inscrição Estadual no Cadastro de Contribuintes do ICMS.
            indIEDest = "9"
       Else
            objDest.getElementsByTagName("indIEDest").Item(0).Text = 2 '2=Contribuinte isento de Inscrição no cadastro de Contribuintes
            indIEDest = "2"
       End If
       
End If

'If Cmb_presenca_comprador.Text = "1 - Operação presencial" Then
'       objDest.getElementsByTagName("indIEDest").Item(0).Text = 1 '9=Não Contribuinte, que pode ou não possuir Inscrição Estadual no Cadastro de Contribuintes do ICMS.
'       indIEDest = "2"
'End If

'Verifica se tem suframa
Set TBClientes = CreateObject("adodb.recordset")
TBClientes.Open "Select CFOP.* from tbl_Detalhes_Nota NFP INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = NFP.ID_CFOP where NFP.ID_nota = " & txtID_nota & " and CFOP.Suframa = 'True'", Conexao, adOpenKeyset, adLockReadOnly
If TBClientes.EOF = False Then
  Set TBClientes = CreateObject("adodb.recordset")
  TBClientes.Open "Select * from Clientes where IDCliente = " & TBproducao!Id_Int_Cliente & " and Suframa is not null", Conexao, adOpenKeyset, adLockReadOnly
  If TBClientes.EOF = False Then
      If TBClientes!Suframa <> "" Then
          objDest.appendChild objDom.createElement("ISUF") '5
          objDest.getElementsByTagName("ISUF").Item(0).Text = Left(ReturnNumbersOnly(TBClientes!Suframa), 9)
      End If
  End If
End If
TBClientes.Close

'Fecha objDest================================================================================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub proc_XML_Totais()
On Error GoTo tratar_erro
Contador3 = 0

Set TBTotaisnota = CreateObject("adodb.recordset")
TBTotaisnota.Open "Select * from tbl_Totais_Nota where ID_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBTotaisnota.EOF = False Then
    'nó total dentro de Enviar (A01)
    Set objTotal = objDom.createElement("total")
    objinfNFe.appendChild objTotal
    'Abre objTotal==================================================================================================
        'nó ICMStot dentro de total
        Set objICMStot = objDom.createElement("ICMSTot")
        objTotal.appendChild objICMStot
        'Abre objICMSTot============================================================================================
            objICMStot.appendChild objDom.createElement("vBC") '0
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Base_ICMS), "00.00", Format(TBTotaisnota!dbl_Base_ICMS, "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
            
            objICMStot.appendChild objDom.createElement("vICMS") '1
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_ICMS), 0, Format(TBTotaisnota!dbl_Valor_ICMS, "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
            
            objICMStot.appendChild objDom.createElement("vICMSDeson") '2
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!Valor_total_desconto_SUFRAMA), "0.00", Format(TBTotaisnota!Valor_total_desconto_SUFRAMA, "0.#0")), ",", ".") 'Novo layout da Sefaz (3.10) - Não é obrigatório
            Contador3 = Contador3 + 1
            
'============================================DIFAL============================================================
' SE tag: idDest = 2) com Consumidor Final (tag: indFinal = 1) e Não Contribuinte (tag: indIEDest = 9)
'=============================================================================================================
If RegimeEmpresa <> 1 Then

If idDest = "2" And indFinal = "1" And indIEDest = "9" Then
          
            objICMStot.appendChild objDom.createElement("vFCPUFDest") '16
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!Valor_total_ICMS_FCP), "0.00", Format(TBTotaisnota!Valor_total_ICMS_FCP, "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
            
            objICMStot.appendChild objDom.createElement("vICMSUFDest") '17
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!Valor_total_ICMS_INT_UF_dest), "0.00", Format(TBTotaisnota!Valor_total_ICMS_INT_UF_dest, "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
            
            objICMStot.appendChild objDom.createElement("vICMSUFRemet") '18
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!Valor_total_ICMS_INT_UF_rem), "0.00", Format(TBTotaisnota!Valor_total_ICMS_INT_UF_rem, "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
            
End If

End If


'============================================================
'Se for substituicao tributaria
'============================================================
    'If VarST = True Then
            objICMStot.appendChild objDom.createElement("vFCP") '20
            objICMStot.childNodes(Contador3).Text = "0.00"
            Contador3 = Contador3 + 1
            
            objICMStot.appendChild objDom.createElement("vBCST") '3
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Base_ICMS_Subst), "0.00", Format(TBTotaisnota!dbl_Base_ICMS_Subst, "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
            
            objICMStot.appendChild objDom.createElement("vST") '4
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_ICMS_Subst), "0.00", Format(TBTotaisnota!dbl_Valor_ICMS_Subst, "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
            
            objICMStot.appendChild objDom.createElement("vFCPST") '22
            objICMStot.childNodes(Contador3).Text = "0.00"
            Contador3 = Contador3 + 1
            
            objICMStot.appendChild objDom.createElement("vFCPSTRet") '23
            objICMStot.childNodes(Contador3).Text = "0.00"
            Contador3 = Contador3 + 1
            
'   End If
'===========================================================
            objICMStot.appendChild objDom.createElement("vProd") '
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Produtos), 0, Format(TBTotaisnota!dbl_Valor_Total_Produtos + IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota_Serv), 0, TBTotaisnota!dbl_Valor_Total_Nota_Serv), "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
                        
            objICMStot.appendChild objDom.createElement("vFrete") '6
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_Frete), "0.00", Format(TBTotaisnota!dbl_Valor_Frete, "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
            
            objICMStot.appendChild objDom.createElement("vSeg") '7
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_Seguro), "0.00", Format(TBTotaisnota!dbl_Valor_Seguro, "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
            
            objICMStot.appendChild objDom.createElement("vDesc") '8
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!Valor_total_desconto), 0, Format(TBTotaisnota!Valor_total_desconto + IIf(IsNull(TBTotaisnota!Valor_total_desconto_SUFRAMA), 0, TBTotaisnota!Valor_total_desconto_SUFRAMA), "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
            
            
            objICMStot.appendChild objDom.createElement("vII") '9
            
            'Se for nota de importação coloca os valores se não for coloca valor =0
            If idDest = 3 Then
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!Valor_total_II), "0.00", Format(TBTotaisnota!Valor_total_II, "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
            
            Else
            objICMStot.childNodes(Contador3).Text = "0.00"
            Contador3 = Contador3 + 1
            
            End If
            
            objICMStot.appendChild objDom.createElement("vIPI") '10
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_Total_IPI), "0.00", Format(TBTotaisnota!dbl_Valor_Total_IPI, "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
            
            
            objICMStot.appendChild objDom.createElement("vIPIDevol") 'Informa IPI devolução
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!Total_IPI_devolv), "0.00", Format(TBTotaisnota!Total_IPI_devolv, "0.#0")), ",", ".") '"0.00"
            Contador3 = Contador3 + 1
            
            objICMStot.appendChild objDom.createElement("vPIS") '11
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!Total_PIS_prod), "0.00", Format(TBTotaisnota!Total_PIS_prod, "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
            
            objICMStot.appendChild objDom.createElement("vCOFINS") '12
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!Total_Cofins_prod), "0.00", Format(TBTotaisnota!Total_Cofins_prod, "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
            
            objICMStot.appendChild objDom.createElement("vOutro") '13
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Desp_Adicionais), "0.00", Format(TBTotaisnota!dbl_Desp_Adicionais, "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
            
            objICMStot.appendChild objDom.createElement("vNF") '14
            objICMStot.childNodes(Contador3).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota), "0.00", Format(TBTotaisnota!dbl_Valor_Total_Nota, "0.#0")), ",", ".")
            Contador3 = Contador3 + 1
            
            objICMStot.appendChild objDom.createElement("vTotTrib") '15
            If IsNull(TBTotaisnota!Valor_total_aprox_tributos) = False And TBTotaisnota!Valor_total_aprox_tributos <> "" And TBTotaisnota!Valor_total_aprox_tributos <> "0" Then
            objICMStot.childNodes(Contador3).Text = Replace(Format(TBTotaisnota!Valor_total_aprox_tributos, "0.#0"), ",", ".")
            Contador3 = Contador3 + 1
            Else
            objICMStot.childNodes(Contador3).Text = "0.00"
            Contador3 = Contador3 + 1
            End If

'
'            'objICMSTot.appendChild objDom.createElement("vIPIDevol") '23
'
'            DAPartilhaICMS = ""
'            If objICMStot.childNodes(17).Text > 0 Then DAPartilhaICMS = "Partilha ICMS operação interestadual consumidor final, disposto na Emenda constitucional 87/2015. Valor ICMS para UF destino (" & TBproducao!txt_UF & "): R$" & Format(objICMStot.childNodes(18).Text, "0.#0") & ". Valor FCP para o destino: R$" & Format(objICMStot.childNodes(17).Text, "0.#0") & ". Valor ICMS UF remetente (" & TBproducao!UF & "): R$" & Format(objICMStot.childNodes(19).Text, "0.#0") & "."
        
        'Fecha objICMSTot=================================================================================================
        
        'nó RetTrib dentro de Total
        'Set objRetTrib = objDom.createElement("retTrib")
        'objTotal.appendChild objRetTrib
        'Abre RetTrib=================================================================================================='Format(, "0.#0")
            'objRetTrib.appendChild objDom.createElement("vRetPIS") '0
            'objRetTrib.childNodes(0).Text = Replace(IIf(IsNull(TBTotaisnota!Total_retencao_PIS), "0.00", Format(TBTotaisnota!Total_retencao_PIS, "0.#0")), ",", ".")
            'objRetTrib.appendChild objDom.createElement("vRetCOFINS_servttlnfe") '1
            'objRetTrib.childNodes(1).Text = Replace(IIf(IsNull(TBTotaisnota!Total_retencao_Cofins), "0.00", Format(TBTotaisnota!Total_retencao_Cofins, "0.#0")), ",", ".")
            'objRetTrib.appendChild objDom.createElement("vRetCSLL") '2
            'objRetTrib.childNodes(2).Text = Replace(IIf(IsNull(TBTotaisnota!Total_CSLL_serv), "0.00", Format(TBTotaisnota!Total_CSLL_serv, "0.#0")), ",", ".")
            'objRetTrib.appendChild objDom.createElement("vBCIRRF") '3
            'objRetTrib.childNodes(3).Text = Replace(IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota_Serv), "0.00", Format(TBTotaisnota!dbl_Valor_Total_Nota_Serv, "0.#0")), ",", ".")
            'objRetTrib.appendChild objDom.createElement("vIRRF") '4
            'objRetTrib.childNodes(4).Text = Replace(IIf(IsNull(TBTotaisnota!Total_IRRF_serv), "0.00", Format(TBTotaisnota!Total_IRRF_serv, "0.#0")), ",", ".")
            ''objRetTrib.appendChild objDom.createElement("vBCRetPrev") '5
            ''objRetTrib.appendChild objDom.createElement("vRetPrev") '6
        'Fecha RetTrib=================================================================================================
    'Fecha objTotal=================================================================================================
End If
TBTotaisnota.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub proc_XML_Transporte()
On Error GoTo tratar_erro

Set TBTransporte = CreateObject("adodb.recordset")
TBTransporte.Open "Select * from tbl_Dados_Transp where ID_Nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBTransporte.EOF = False Then
'==========================================================================================
'Criar nó transp dentro da NFE
'==========================================================================================
Set objTransp = objDom.createElement("transp")
objinfNFe.appendChild objTransp
'Abre transp==================================================================================================
objTransp.appendChild objDom.createElement("modFrete") '0
objTransp.getElementsByTagName("modFrete").Item(0).Text = IIf(IsNull(TBTransporte!txt_Frete_Conta), 0, TBTransporte!txt_Frete_Conta) 'Frete Novo layout da Sefaz (4.0)
'objTransp.appendChild objDom.createElement("balsa") '1
'objTransp.appendChild objDom.createElement("vagao") '2
'==========================================================================================
'Criar nó transporta dentro de transp
'==========================================================================================
If TBTransporte!txt_Frete_Conta <> "9" Then
Set objTransporta = objDom.createElement("transporta")
objTransp.appendChild objTransporta
'==========================================================================================
'Abre Transporta
'==========================================================================================
'Verifica se é tipo fornecedor, cliente, ou a própria empresa que irá transportar a carga
'==========================================================================================
Familiatext = ""
'==========================================================================================
' Se for fornecedor pega o CNPJ e monta o endereço completo
'==========================================================================================
If IsNull(TBTransporte!txt_CNPJ) = False And TBTransporte!txt_CNPJ <> "" Then
    Set TBFornecedor = CreateObject("adodb.recordset")
    TBFornecedor.Open "Select * from Compras_fornecedores where IDCliente = " & TBTransporte!IdIntTransp & " and Nome_Razao = '" & TBTransporte!txt_Razao & "'", Conexao, adOpenKeyset, adLockReadOnly
    If TBFornecedor.EOF = False Then
        If Left(TBFornecedor!Pessoa, 1) = "J" Then
            objTransporta.appendChild objDom.createElement("CNPJ") '0
            objTransporta.getElementsByTagName("CNPJ").Item(0).Text = ReturnNumbersOnly(TBTransporte!txt_CNPJ)
        Else
            objTransporta.appendChild objDom.createElement("CPF") '1
            objTransporta.getElementsByTagName("CPF").Item(0).Text = ReturnNumbersOnly(TBTransporte!txt_CNPJ)
        End If
        If IsNull(TBTransporte!txt_Endereco) = False And TBTransporte!txt_Endereco <> "" Then Familiatext = TBTransporte!txt_Endereco
        If IsNull(TBTransporte!int_numero) = False And TBTransporte!int_numero <> "" Then
            If Familiatext <> "" Then Familiatext = Familiatext & ", " & TBTransporte!int_numero Else Familiatext = TBTransporte!int_numero
        End If
        If IsNull(TBFornecedor!Bairro) = False And TBFornecedor!Bairro <> "" Then
            If Familiatext <> "" Then Familiatext = Familiatext & " - " & TBFornecedor!Bairro Else Familiatext = TBFornecedor!Bairro
        End If
    Else
'==========================================================================================
' Se for cliente pega o CNPJ e monta o endereço completo
'==========================================================================================
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select * from Clientes where IDCliente = " & TBTransporte!IdIntTransp & " and NomeRazao = '" & TBTransporte!txt_Razao & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
'==========================================================================================
' Tipo Juridico com CNPJ
'==========================================================================================
    If Left(TBFornecedor!Tipo, 1) = "J" Then
        objTransporta.appendChild objDom.createElement("CNPJ")
        objTransporta.getElementsByTagName("CNPJ").Item(0).Text = ReturnNumbersOnly(TBTransporte!txt_CNPJ)
'==========================================================================================
' Tipo físico com CPF
'==========================================================================================
        Else
            objTransporta.appendChild objDom.createElement("CPF")
            objTransporta.getElementsByTagName("CPF").Item(0).Text = ReturnNumbersOnly(TBTransporte!txt_CNPJ)
        End If
        If IsNull(TBTransporte!txt_Endereco) = False And TBTransporte!txt_Endereco <> "" Then Familiatext = TBTransporte!txt_Endereco
        If IsNull(TBTransporte!int_numero) = False And TBTransporte!int_numero <> "" Then
            If Familiatext <> "" Then Familiatext = Familiatext & ", " & TBTransporte!int_numero Else Familiatext = TBTransporte!int_numero
        End If
        If IsNull(TBFornecedor!Bairro) = False And TBFornecedor!Bairro <> "" Then
            If Familiatext <> "" Then Familiatext = Familiatext & " - " & TBFornecedor!Bairro Else Familiatext = TBFornecedor!Bairro
        End If
    Else
'==========================================================================================
' Se for a empresa pega o CNPJ e monta o endereço completo
'==========================================================================================
    Set TBFornecedor = CreateObject("adodb.recordset")
    TBFornecedor.Open "Select * from Empresa where Codigo = " & TBTransporte!IdIntTransp, Conexao, adOpenKeyset, adLockOptimistic
    If TBFornecedor.EOF = False Then
        objTransporta.appendChild objDom.createElement("CNPJ")
        objTransporta.getElementsByTagName("CNPJ").Item(0).Text = ReturnNumbersOnly(TBTransporte!txt_CNPJ)
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

'Nome da transportadora
objTransporta.appendChild objDom.createElement("xNome") '2
objTransporta.getElementsByTagName("xNome").Item(0).Text = Trim(RemoveAccents(TBTransporte!txt_Razao))
 
' Inscricao estadual
If IsNull(TBTransporte!txt_IE) = False Then
objTransporta.appendChild objDom.createElement("IE")
objTransporta.getElementsByTagName("IE").Item(0).Text = IIf(TBTransporte!txt_IE = "ISENTO", "ISENTO", Left(ReturnNumbersOnly(TBTransporte!txt_IE), 14))
End If

' Endereço
objTransporta.appendChild objDom.createElement("xEnder") '4
If Familiatext <> "" Then
Familiatext = Left(Familiatext, 60)
Familitext = Trim(Familiatext)
objTransporta.getElementsByTagName("xEnder").Item(0).Text = IIf(IsNull(Familiatext) = False, (RemoveAccents(Trim(Familiatext))), "Exterior")
End If
'Municipio
objTransporta.appendChild objDom.createElement("xMun") '5
objTransporta.getElementsByTagName("xMun").Item(0).Text = RemoveAccents(TBTransporte!txt_Municipio)
objTransporta.appendChild objDom.createElement("UF")
objTransporta.getElementsByTagName("UF").Item(0).Text = TBTransporte!txt_UF
End If
'=============================================================================
'Fecha Transporta
'==============================================================================
'Dados do veículo que irá transportar a carga
'==============================================================================
If TBTransporte!txt_Placa <> "" And IsNull(TBTransporte!txt_Placa) = False And TBproducao!UF = TBproducao!txt_UF Then
'no VeicTransp dentro de Transp (Y01)
Set objVeicTransp = objDom.createElement("veicTransp")
objTransp.appendChild objVeicTransp
'Abre VeicTransp===============================================================
    objVeicTransp.appendChild objDom.createElement("placa")
    objVeicTransp.getElementsByTagName("placa").Item(0).Text = TBTransporte!txt_Placa
    If TBTransporte!txt_UF_Placa <> "" And IsNull(TBTransporte!txt_UF_Placa) = False Then
        objVeicTransp.appendChild objDom.createElement("UF") '1
        objVeicTransp.getElementsByTagName("UF").Item(0).Text = TBTransporte!txt_UF_Placa
    End If
    If IsNull(TBTransporte!Codigo_ANTT) = False Then
        objVeicTransp.appendChild objDom.createElement("RNTC") '2
        objVeicTransp.getElementsByTagName("RNTC").Item(0).Text = "000000" 'IIf(IsNull(TBTransporte!Codigo_ANTT) = False, TBTransporte!Codigo_ANTT, "000000")
    End If
'Fecha VeicTransp
'no Reboque dentro de Transp (Y01)
Set objReboque = objDom.createElement("reboque")
objTransp.appendChild objReboque
'Abre reboque=================================================================
        objReboque.appendChild objDom.createElement("placa")
        objReboque.getElementsByTagName("placa").Item(0).Text = TBTransporte!txt_Placa
        If TBTransporte!txt_UF_Placa <> "" And IsNull(TBTransporte!txt_UF_Placa) = False Then
            objReboque.appendChild objDom.createElement("UF")
            objReboque.getElementsByTagName("UF").Item(0).Text = TBTransporte!txt_UF_Placa
        End If
    'If IsNull(TBTransporte!Codigo_ANTT) = False Or TBTransporte!Codigo_ANTT <> "" Then
    '    objVeicTransp.appendChild objDom.createElement("RNTC") '2
    '    objVeicTransp.getElementsByTagName("RNTC").Item(0).Text = "000000" 'IIf(IsNull(TBTransporte!Codigo_ANTT) = False, TBTransporte!Codigo_ANTT, "000000")
    'End If
'Fecha reboque=================================================================================================
End If
    
'no vol dentro de Transp (Y01)
Set objVol = objDom.createElement("vol")
objTransp.appendChild objVol
'Abre Vol==================================================================================================
      objVol.appendChild objDom.createElement("qVol") '0
      objVol.getElementsByTagName("qVol").Item(0).Text = Replace(IIf(IsNull(TBTransporte!int_Qtd_Transp), 0, TBTransporte!int_Qtd_Transp), ",", ".")
      objVol.appendChild objDom.createElement("esp") '1
      objVol.getElementsByTagName("esp").Item(0).Text = IIf(TBTransporte!txt_Especie = "", "Volume", Trim(TBTransporte!txt_Especie))
      objVol.appendChild objDom.createElement("marca") '2
      objVol.getElementsByTagName("marca").Item(0).Text = IIf(TBTransporte!txt_Marca = "", "Propria", Trim(TBTransporte!txt_Marca))
      objVol.appendChild objDom.createElement("nVol") '3
      objVol.getElementsByTagName("nVol").Item(0).Text = IIf(IsNull(TBTransporte!int_Qtd_Transp), 0, TBTransporte!int_Qtd_Transp)
      objVol.appendChild objDom.createElement("pesoL") '4
      objVol.getElementsByTagName("pesoL").Item(0).Text = Replace(IIf(IsNull(TBTransporte!dbl_Peso_Liquido), "0.000", Format(TBTransporte!dbl_Peso_Liquido, "0.000")), ",", ".")
      objVol.appendChild objDom.createElement("pesoB") '5
      objVol.getElementsByTagName("pesoB").Item(0).Text = Replace(IIf(IsNull(TBTransporte!dbl_Peso_Bruto), "0.000", Format(TBTransporte!dbl_Peso_Bruto, "0.000")), ",", ".")
  'Fecha volItem=================================================================================================
'Fecha Vol=================================================================================================
'Fecha transp=================================================================================================
End If
End If

TBTransporte.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub proc_XML_Adicionais()
On Error GoTo tratar_erro
Dim ReservadoFisco As String

'no infAdic (AB01) dentro de Enviar (A01)
Set objInfAdic = objDom.createElement("infAdic")
objinfNFe.appendChild objInfAdic
'Abre InfAdic====================================================================================================
    
    
    Familiatext = ""
    DadosAdicionaisTexto = ""
    ReservadoFisco = ""
'=======================================================================
' Aqui verifica se tem dados adicionais gravados
' Senão não, não adiciona na nota.
'=======================================================================
    Set TBControleNF = CreateObject("adodb.recordset")
    TBControleNF.Open "Select * from tbl_DadosAdicionais where ID_nota = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
    If TBControleNF.EOF = False Then
        
        If IsNull(TBControleNF!mem_corpo) = False And TBControleNF!mem_corpo <> "" Then
            ReservadoFisco = FunTiraAcentosTexto(Trim(TBControleNF!mem_corpo))
        End If
        If IsNull(TBControleNF!mem_DadosAdicionais) = False And TBControleNF!mem_DadosAdicionais <> "" Then
            DadosAdicionaisTexto = FunTiraAcentosTexto(Trim(TBControleNF!mem_DadosAdicionais))
        End If
        
    End If
    TBControleNF.Close
    
'=======================================================================
' Aqui verifica se tem endereço de entrega
' Senão não, não adiciona na nota.
'=======================================================================
    endereco_entrega = ""
    If TBproducao!DA_entrega = True Then
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select * from clientes_entrega where identrega = " & TBproducao!ID_entrega, Conexao, adOpenKeyset, adLockReadOnly
        If TBClientes.EOF = False Then
        
        Cidade = TBClientes!cidade_entrega
         Cidade = Replace(Cidade, "d'oeste", "Do Oeste")

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
            endereco_entrega = Endereco & " - " & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & " - " & Bairro & " - " & Cidade & " - " & IIf(IsNull(TBClientes!uf_entrega), "", TBClientes!uf_entrega) & " - " & IIf(IsNull(TBClientes!cep_entrega), "", TBClientes!cep_entrega)
        End If
        TBClientes.Close
    End If
    
'=======================================================================
' Aqui verifica se tem endereço de cobrança gravado
' Senão não, não adiciona na nota.
'=======================================================================
    endereco_Cobranca = ""
    If TBproducao!DA_cobranca = True Then
        Set TBClientes = CreateObject("adodb.recordset")
        TBClientes.Open "Select * from clientes_cobranca where idcobranca = " & TBproducao!ID_Cobranca, Conexao, adOpenKeyset, adLockReadOnly
        If TBClientes.EOF = False Then
        
        Cidade = TBClientes!cidade_Cobranca
        Cidade = Replace(Cidade, "d'oeste", "Do Oeste")

            If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
                Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!endereco_Cobranca), "", TBClientes!endereco_Cobranca)
            Else
                Endereco = IIf(IsNull(TBClientes!endereco_Cobranca), "", TBClientes!endereco_Cobranca)
            End If
            If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then
                Bairro = TBClientes!Tipo_bairro & ": " & IIf(IsNull(TBClientes!bairro_Cobranca), "", TBClientes!bairro_Cobranca)
            Else
                Bairro = IIf(IsNull(TBClientes!bairro_Cobranca), "", TBClientes!bairro_Cobranca)
            End If
            endereco_Cobranca = Endereco & " - " & IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero) & " - " & Bairro & " - " & Cidade & " - " & IIf(IsNull(TBClientes!uf_Cobranca), "", TBClientes!uf_Cobranca) & " - " & IIf(IsNull(TBClientes!cep_Cobranca), "", TBClientes!cep_Cobranca)
        End If
        TBClientes.Close
    End If
                    
    If DadosAdicionaisTexto <> "" Or endereco_entrega <> "" Or endereco_Cobranca <> "" Or DAPartilhaICMS <> "" Then
        If DadosAdicionaisTexto <> "" Then Familiatext = DadosAdicionaisTexto
        If endereco_entrega <> "" Then
           If Familiatext <> "" Then Familiatext = Familiatext & "|Endereço de entrega: " & endereco_entrega Else Familiatext = "Endereço de entrega: " & endereco_entrega
        End If
        If endereco_Cobranca <> "" Then
           If Familiatext <> "" Then Familiatext = Familiatext & "|Endereço de cobrança: " & endereco_Cobranca Else Familiatext = "Endereço de cobrança: " & endereco_Cobranca
        End If
        If DAPartilhaICMS <> "" Then
           If Familiatext <> "" Then Familiatext = Familiatext & "|" & DAPartilhaICMS Else Familiatext = DAPartilhaICMS
        End If
    End If
    
'=======================================================================
' Aqui verifica se tem dados adicionais
' Senão não adiciona a tag na nota
'=======================================================================
'Debug.print ReservadoFisco

 If ReservadoFisco <> "" Then
   objInfAdic.appendChild objDom.createElement("infAdFisco") '1
   objInfAdic.getElementsByTagName("infAdFisco").Item(0).Text = FunTiraAcentosTexto(LTrim(Trim(ReservadoFisco)))
 End If

'Debug.print Familiatext

 If Familiatext <> "" Then
   objInfAdic.appendChild objDom.createElement("infCpl") '1
   objInfAdic.getElementsByTagName("infCpl").Item(0).Text = FunTiraAcentosTexto(LTrim(Trim(Familiatext)))
 End If
 
 '=====================================================================
 ' Informações do responsável técnico
 '=====================================================================
 
 Set objinfRespTec = objDom.createElement("infRespTec")
    objinfNFe.appendChild objinfRespTec
    objinfRespTec.appendChild objDom.createElement("CNPJ")
    objinfRespTec.getElementsByTagName("CNPJ").Item(0).Text = "34270461000104"
    
    objinfRespTec.appendChild objDom.createElement("xContato")
    objinfRespTec.getElementsByTagName("xContato").Item(0).Text = "suporte"
    
    objinfRespTec.appendChild objDom.createElement("email")
    objinfRespTec.getElementsByTagName("email").Item(0).Text = "suporte@caprind.com.br"
    
    objinfRespTec.appendChild objDom.createElement("fone")
    objinfRespTec.getElementsByTagName("fone").Item(0).Text = "1933282575"
    
 
'Fecha InfAdic===================================================================================================
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Function funVerificacaoEnviar() As Boolean
On Error GoTo tratar_erro

funVerificacaoEnviar = True

If funVerifLiberacao(True) = False Then
    funVerificacaoEnviar = False
    Exit Function
End If

If NFCe = False Then
'Verifica se a cidade está cadastrada corretamente
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Dados_Nota_Fiscal where ID = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
    FamiliaAntiga = RemoveAccents(TBAbrir!txt_Municipio)
    
    If IsNull(TBAbrir!txt_UF) = False And TBAbrir!txt_UF <> "" And TBAbrir!txt_UF <> "EX" Then
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from CEP where Municipio = '" & FamiliaAntiga & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBFI.EOF = True Then
            USMsgBox ("Não é permitido liberar esta nota fiscal para envio, pois a nota esta com a cidade errada."), vbExclamation, "CAPRIND v5.0"
            funVerificacaoEnviar = False
            TBFI.Close
            Exit Function
        End If
        
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from CEP where Sigla_UF = '" & TBAbrir!txt_UF & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBFI.EOF = True Then
            USMsgBox ("Não é permitido liberar esta nota fiscal para envio, pois a nota esta com o estado errado."), vbExclamation, "CAPRIND v5.0"
            funVerificacaoEnviar = False
            TBFI.Close
            Exit Function
        End If
        
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from CEP where Municipio = '" & FamiliaAntiga & "' and Sigla_UF = '" & TBAbrir!txt_UF & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBFI.EOF = True Then
            USMsgBox ("Não é permitido liberar esta nota fiscal para envio, pois não existe o munícipio " & FamiliaAntiga & " no estado " & UF & " na tabela CEP."), vbExclamation, "CAPRIND v5.0"
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
                USMsgBox ("Não é permitido liberar esta nota fiscal para envio, pois este cliente não tem país cadastrado."), vbExclamation, "CAPRIND v5.0"
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
                USMsgBox ("Não é permitido liberar esta nota fiscal para envio, pois este fornecedor não tem país cadastrado."), vbExclamation, "CAPRIND v5.0"
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
            If USMsgBox("A(s) duplicata(s) ainda não foi(ram) gerada(s), deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then
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
        USMsgBox ("Não é permitido liberar a nota fiscal para envio, pois não foi informado o caminho onde será armazenado os aquivos para envio."), vbExclamation, "CAPRIND v5.0"
        funVerificacaoEnviar = False
        Exit Function
    End If
    'Verificar se o caminho existe
    If DS.FileOrDirExists(TBMaquinas!Caminho_Nfe) = False Then
        USMsgBox ("Não é permitido liberar a nota fiscal para envio, pois não foi encontrado o caminho " & TBMaquinas!Caminho_Nfe & ", onde será armazenado os aquivos para envio."), vbExclamation, "CAPRIND v5.0"
        funVerificacaoEnviar = False
        Exit Function
    End If
End If
TBMaquinas.Close

'Verifica se é nota fiscal de devolução ou complementar e se esta referenciado a nota fiscal
Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select NFE.ID from tbl_Dados_Nota_Fiscal_NFe NFE LEFT JOIN Faturamento_Relacionamento FR ON FR.ID_nota = NFE.ID_nota where NFE.ID_nota = " & txtID_nota & " and NFE.Finalidade_emissao <> 1 and NFE.Finalidade_emissao <> 3 and FR.ID IS NULL", Conexao, adOpenKeyset, adLockReadOnly
If TBMaquinas.EOF = False Then
    USMsgBox ("Não é permitido liberar a nota fiscal para envio, pois não foi feito o relacionamento."), vbExclamation, "CAPRIND v5.0"
    funVerificacaoEnviar = False
    Exit Function
End If
TBMaquinas.Close

'Verifica se o clinte é fisico e esta com cnpj e vice e versa
Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select txt_tipocliente, txt_CNPJ_CPF from tbl_Dados_Nota_Fiscal where ID = " & txtID_nota & " and txt_uf <> 'EX'", Conexao, adOpenKeyset, adLockReadOnly
If TBMaquinas.EOF = False Then
    If Left(TBMaquinas!txt_tipocliente, 1) = "J" And Len(TBMaquinas!txt_CNPJ_CPF) < 14 Then
        USMsgBox ("Não é permitido liberar a nota fiscal para envio, pois o CNPJ do destinatario esta errado."), vbExclamation, "CAPRIND v5.0"
        TBMaquinas.Close
        funVerificacaoEnviar = False
        Exit Function
    ElseIf Left(TBMaquinas!txt_tipocliente, 1) = "F" And Len(TBMaquinas!txt_CNPJ_CPF) > 14 Then
        USMsgBox ("Não é permitido liberar a nota fiscal para envio, pois o CPF do destinatario esta errado."), vbExclamation, "CAPRIND v5.0"
        TBMaquinas.Close
        funVerificacaoEnviar = False
        Exit Function
    End If
End If
TBMaquinas.Close
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Public Sub procMontaEmail()
On Error GoTo tratar_erro

EmailEnvioNFe = ""
EmailCliente = ""
EmailTransportadora = ""
EmailUsuario = ""

If chkTransportadora.Value = 0 And chkUsuario.Value = 0 Then
ProcCarregaEmailCliente
EmailEnvioNFe = EmailCliente
End If

If chkTransportadora.Value = 1 And chkUsuario.Value = 0 Then
ProcCarregaEmailCliente
ProcCarregaEmailTransportadora
EmailEnvioNFe = EmailCliente & ", " & EmailTransportadora
End If

If chkTransportadora.Value = 1 And chkUsuario.Value = 1 Then
ProcCarregaEmailCliente
ProcCarregaEmailTransportadora
ProcCarregaEmailUsuario

If EmailCliente <> "" And EmailTranportadora <> "" And EmailUsuario <> "" Then
EmailEnvioNFe = EmailCliente & ", " & EmailTransportadora & ", " & EmailUsuario
End If

If EmailCliente <> "" And EmailTranportadora <> "" And EmailUsuario = "" Then
EmailEnvioNFe = EmailCliente & ", " & EmailTransportadora
End If

If EmailCliente <> "" And EmailTranportadora = "" And EmailUsuario = "" Then
EmailEnvioNFe = EmailCliente
End If

End If

'Debug.print EmailEnvioNFe

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procCarregaEmpresa()
On Error GoTo tratar_erro

UF_transp = ""
Cidade = ""
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select E.UF, E.Cidade, E.CNPJ, E.Caminho_Nfe, E.Caminho_XMLDanfe, E.Caminho_RetornoNfe, N.Obs from Empresa E INNER JOIN tbl_Dados_Nota_Fiscal N ON E.Codigo = N.ID_empresa where N.ID = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly

If TBFI.EOF = False Then
    UF_transp = IIf(IsNull(TBFI!UF), "", TBFI!UF)
    Cidade = IIf(IsNull(TBFI!Cidade), "", TBFI!Cidade)
    CnpjNF = IIf(IsNull(TBFI!CNPJ), "", TBFI!CNPJ)
    DiretorioEnvio = IIf(IsNull(TBFI!Caminho_Nfe), "", TBFI!Caminho_Nfe)
    DiretorioXMLDanfe = IIf(IsNull(TBFI!Caminho_XMLDanfe), "", TBFI!Caminho_XMLDanfe)
    DiretorioRetorno = IIf(IsNull(TBFI!Caminho_RetornoNfe), "", TBFI!Caminho_RetornoNfe)
    txtMotivo = IIf(IsNull(TBFI!Obs), "", TBFI!Obs)
End If
TBFI.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcImprimir()
On Error GoTo tratar_erro

If txtID_nota = 0 Then
    USMsgBox ("Informe a nota fiscal antes de solicitar impressão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If txtcStat = "" Then
    If USMsgBox("Deseja visualizar uma prévia da Danfe antes da emissão?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            btnCriarXML_Click
            btnPrevia_Click
    End If
Else
ProcCriarPastaDanfe
If NFCe = False Then
    retorno2 = NFeAPI.downloadNFeAndSave(txtchNFe, tpAmb, "P", DiretorioDanfe, True)
Else
    retorno2 = NFCe_downloadESalvar(txtchNFe, tpAmb, DiretorioDanfe, True)
End If
    
End If
   
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
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
USMsgBox "Cadastro de transportadora não encontrado, favor verificar", vbCritical, "CAPRIND V5.0"
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCancelar()
On Error GoTo tratar_erro

If IsInternetOnline = False Then
 USMsgBox "Internet indisponível no momento, tente mais tarde.", vbCritical, "CAPRIND v5.0"
 Exit Sub
End If

If txtID_nota = 0 Then
    USMsgBox ("Informe a nota fiscal antes de cancelar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select status from tbl_Dados_Nota_Fiscal_NFe WHERE ID_nota = " & txtID_nota & " AND status <> 100", Conexao, adOpenKeyset, adLockReadOnly
If TBproducao.EOF = False Then
    USMsgBox ("Só é possível cancelar notas aprovadas."), vbExclamation, "CAPRIND v5.0"
    TBproducao.Close
    Exit Sub
End If

Acao = "cancelar"

Mensagem:
If USMsgBox("Deseja cancelar a nota fiscal N° " & txtNota.Text & " ?", vbYesNo, "CAPRIND") = vbYes Then
frmFaturamento_RetornoSEFAZ.Show
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub proc_XML_Cancelar()
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
'    objEnvioEvento.childNodes(3).Text = ChaveMigrate
    
    'nó Evento
    Set objEvento = objDom.createElement("Evento")
    objEnvioEvento.appendChild objEvento
    'Abre Evento===================================================================================================
        objEvento.appendChild objDom.createElement("NtfCnpjEmissor")
        objEvento.childNodes(0).Text = ReturnNumbersOnly(CnpjNF)
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
                
objDom.Save (DiretorioEnvio & NomeArquivo & ".xml")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Function procVerificaMigrate() As Boolean
On Error GoTo tratar_erro
procVerificaMigrate = False


If DiretorioEnvio = "" Then
    NomeCampo = "o diretório de envio no cadastro da empresa"
    ProcVerificaAcao
    Exit Function
End If

'If DiretorioRetorno = "" Then
'    NomeCampo = "o diretório de retorno no cadastro da empresa"
'    ProcVerificaAcao
'    Exit Function
'End If

If DiretorioXMLDanfe = "" Then
    NomeCampo = "o diretório de XML e Danfe no cadastro da empresa"
    ProcVerificaAcao
    Exit Function
End If

procVerificaMigrate = True

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Function

Public Sub procCancelarTabelas()
On Error GoTo tratar_erro
'Aqui exclui os relacionamentos da nota fiscal se tiver
ProcExcluirRelacionamentoNF txtID_nota
'Aqui exclui arquivo remessa se tiver
ProcExcluirArquivosRemessa txtID_nota

'Aqui exclui os dados do financeiro se tiver
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select int_TipoNota, txt_tipocliente from tbl_Dados_Nota_Fiscal WHERE ID = " & txtID_nota, Conexao, adOpenKeyset, adLockReadOnly
If TBproducao.EOF = False Then
    ProcExcluirContas txtID_nota, IIf(TBproducao!int_TipoNota = 1, True, False), TBproducao!txt_tipocliente
End If
TBproducao.Close

'Aqui exclui os empenhos da nota fiscal
Conexao.Execute "DELETE from ECEV from Estoque_Controle_Empenho_Vendas ECEV INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = ECEV.ID_faturamento where NFP.ID_nota = " & txtID_nota

Conexao.Execute "DELETE FROM tbl_proposta_nota WHERE id_nota = " & ID_nota
frmFaturamento_Prod_Serv.ProcAtualizaDadosPedido ID_nota, True

'======================================================
'Apaga RE do item e Movimentação no estoque
'======================================================
ID_nota = txtID_nota
ApagarMovimentacaoNFe
'======================================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

