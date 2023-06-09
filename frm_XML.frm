VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{8C1279ED-044C-4258-A3E3-0D5514B899FC}#1.44#0"; "ControlesUteis.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_XML 
   Caption         =   "Importar XML da nota fiscal eletr�nica de entrada"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8955
      Left            =   0
      TabIndex        =   1
      Top             =   1020
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   15796
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Nota fiscal eletr�nica"
      TabPicture(0)   =   "frm_XML.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ActiveResize1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Lista de produtos"
      TabPicture(1)   =   "frm_XML.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Lista"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Dados transporte/Inf. adicionais"
      TabPicture(2)   =   "frm_XML.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(2)=   "Frame6"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Fatura (Duplicatas)"
      TabPicture(3)   =   "frm_XML.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).Control(1)=   "ListaDuplicatas"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame7 
         Caption         =   "Dados da cobran�a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1095
         Left            =   -74940
         TabIndex        =   71
         Top             =   420
         Width           =   15180
         Begin DrawSuite2022.USButton btnCriarNota 
            Height          =   885
            Left            =   10440
            TabIndex        =   110
            Top             =   120
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   1561
            DibPicture      =   "frm_XML.frx":0070
            Caption         =   "Criar Nota fiscal"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
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
            PicAlign        =   7
            PicSize         =   5
            PicSizeH        =   32
            PicSizeW        =   32
            Theme           =   3
         End
         Begin ControlesUteis.txt fatnFat 
            Height          =   555
            Left            =   180
            TabIndex        =   72
            Top             =   390
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   979
            Tamanho         =   1200
            Text            =   ""
            Caption         =   "N� fatura"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt fatvDesc 
            Height          =   555
            Left            =   2625
            TabIndex        =   73
            Top             =   390
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   979
            Tamanho         =   1290
            Text            =   ""
            Caption         =   "Desconto"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt fatvLiq 
            Height          =   555
            Left            =   3930
            TabIndex        =   74
            Top             =   390
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   979
            Tamanho         =   1305
            Text            =   ""
            Caption         =   "Valor l�quido"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt fatindPag 
            Height          =   555
            Left            =   5250
            TabIndex        =   75
            Top             =   390
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   979
            Tamanho         =   1905
            Text            =   ""
            Caption         =   "Tipo pagamento"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt fatvOrig 
            Height          =   555
            Left            =   1380
            TabIndex        =   76
            Top             =   390
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   979
            Tamanho         =   1230
            Text            =   ""
            Caption         =   "Origem"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt fattPag 
            Height          =   555
            Left            =   7170
            TabIndex        =   77
            Top             =   390
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   979
            Tamanho         =   1545
            Text            =   ""
            Caption         =   "Forma pagamento"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt fatvPag 
            Height          =   555
            Left            =   8730
            TabIndex        =   78
            Top             =   390
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   979
            Tamanho         =   1545
            Text            =   ""
            Caption         =   "Valor  pagamento"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin DrawSuite2022.USButton btnReceber_estoque 
            Height          =   885
            Left            =   12930
            TabIndex        =   111
            Top             =   120
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   1561
            DibPicture      =   "frm_XML.frx":9B1D
            Caption         =   "Receber estoque"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
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
            PicSize         =   5
            PicSizeH        =   32
            PicSizeW        =   32
            Theme           =   5
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Informa��es complementares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   5535
         Left            =   -74910
         TabIndex        =   69
         Top             =   3360
         Width           =   15210
         Begin VB.TextBox infCpl 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5055
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   70
            Top             =   360
            Width           =   14955
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados dos volumes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1095
         Left            =   -74940
         TabIndex        =   62
         Top             =   2250
         Width           =   15240
         Begin ControlesUteis.txt transpqVol 
            Height          =   555
            Left            =   180
            TabIndex        =   63
            Top             =   390
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   979
            Tamanho         =   1200
            Text            =   ""
            Caption         =   "qVol"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt transpMarca 
            Height          =   555
            Left            =   7635
            TabIndex        =   64
            Top             =   375
            Width           =   3750
            _ExtentX        =   6615
            _ExtentY        =   979
            Tamanho         =   3750
            Text            =   ""
            Caption         =   "Marca"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt transpnVol 
            Height          =   555
            Left            =   11400
            TabIndex        =   65
            Top             =   375
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   979
            Tamanho         =   1035
            Text            =   ""
            Caption         =   "Volumes"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt transpesp 
            Height          =   555
            Left            =   1410
            TabIndex        =   66
            Top             =   390
            Width           =   6210
            _ExtentX        =   10954
            _ExtentY        =   979
            Tamanho         =   6210
            Text            =   ""
            Caption         =   "Esp�cie"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt transppesoL 
            Height          =   555
            Left            =   12450
            TabIndex        =   67
            Top             =   360
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   979
            Tamanho         =   1305
            Text            =   ""
            Caption         =   "Peso l�quido"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt transppesoB 
            Height          =   555
            Left            =   13770
            TabIndex        =   68
            Top             =   360
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   979
            Tamanho         =   1275
            Text            =   ""
            Caption         =   "Peso bruto"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados da transportadora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1755
         Left            =   -74940
         TabIndex        =   55
         Top             =   450
         Width           =   15240
         Begin ControlesUteis.txt transpxNome 
            Height          =   555
            Left            =   180
            TabIndex        =   56
            Top             =   390
            Width           =   7620
            _ExtentX        =   13441
            _ExtentY        =   979
            Tamanho         =   7620
            Text            =   ""
            Caption         =   "Nome raz�o social"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt transpCNPJ 
            Height          =   555
            Left            =   7800
            TabIndex        =   57
            Top             =   390
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   979
            Tamanho         =   1605
            Tipo            =   8
            Text            =   ""
            Caption         =   "CNPJ"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt transpUF 
            Height          =   555
            Left            =   10815
            TabIndex        =   58
            Top             =   390
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   979
            Tamanho         =   330
            Text            =   ""
            Caption         =   "UF"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt transpxMun 
            Height          =   555
            Left            =   11160
            TabIndex        =   59
            Top             =   390
            Width           =   3945
            _ExtentX        =   6959
            _ExtentY        =   979
            Tamanho         =   3945
            Text            =   ""
            Caption         =   "Municipio"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt transpxEnder 
            Height          =   555
            Left            =   210
            TabIndex        =   60
            Top             =   1065
            Width           =   14895
            _ExtentX        =   26273
            _ExtentY        =   979
            Tamanho         =   14895
            Text            =   ""
            Caption         =   "Endere�o"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt transpIE 
            Height          =   555
            Left            =   9420
            TabIndex        =   61
            Top             =   390
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   979
            Tamanho         =   1380
            Text            =   ""
            Caption         =   "Inscri��o estadual"
            Enabled         =   0   'False
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Destinat�rio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1515
         Left            =   60
         TabIndex        =   28
         Top             =   4905
         Width           =   15240
         Begin ControlesUteis.txt dest_xNome 
            Height          =   555
            Left            =   180
            TabIndex        =   29
            Top             =   360
            Width           =   6120
            _ExtentX        =   10795
            _ExtentY        =   979
            Tamanho         =   6120
            Text            =   ""
            Caption         =   "Nome raz�o social"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt dest_CNPJ 
            Height          =   555
            Left            =   6300
            TabIndex        =   30
            Top             =   360
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   979
            Tamanho         =   1605
            Tipo            =   8
            Text            =   ""
            Caption         =   "CNPJ"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt dest_indIEDest 
            Height          =   555
            Left            =   12420
            TabIndex        =   31
            Top             =   930
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   979
            Tamanho         =   2670
            Text            =   ""
            Caption         =   "indIEDest"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt dest_xPais 
            Height          =   555
            Left            =   11280
            TabIndex        =   32
            Top             =   930
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   979
            Tamanho         =   1140
            Text            =   ""
            Caption         =   "Pa�s"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt dest_CEP 
            Height          =   555
            Left            =   10335
            TabIndex        =   33
            Top             =   930
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   979
            Tamanho         =   945
            Tipo            =   9
            Text            =   ""
            Caption         =   "CEP"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt dest_UF 
            Height          =   555
            Left            =   10005
            TabIndex        =   34
            Top             =   930
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   979
            Tamanho         =   330
            Text            =   ""
            Caption         =   "UF"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt dest_xMun 
            Height          =   555
            Left            =   6300
            TabIndex        =   35
            Top             =   930
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   979
            Tamanho         =   3705
            Text            =   ""
            Caption         =   "Municipio"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt dest_nro 
            Height          =   555
            Left            =   210
            TabIndex        =   36
            Top             =   930
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   979
            Tamanho         =   690
            Text            =   ""
            Caption         =   "Numero"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt dest_xLgr 
            Height          =   555
            Left            =   7920
            TabIndex        =   37
            Top             =   360
            Width           =   7185
            _ExtentX        =   12674
            _ExtentY        =   979
            Tamanho         =   7185
            Text            =   ""
            Caption         =   "Logradouro"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt dest_xBairro 
            Height          =   555
            Left            =   2265
            TabIndex        =   38
            Top             =   930
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   979
            Tamanho         =   4035
            Text            =   ""
            Caption         =   "Bairro"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt dest_xCpl 
            Height          =   555
            Left            =   900
            TabIndex        =   54
            Top             =   930
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   979
            Tamanho         =   1365
            Text            =   ""
            Caption         =   "Complemento"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Totaliza��es de valores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2460
         Left            =   90
         TabIndex        =   27
         Top             =   6450
         Width           =   15210
         Begin VB.Frame Frame15 
            Caption         =   "TOTAIS NOTA FISCAL"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   2055
            Left            =   12720
            TabIndex        =   103
            Top             =   360
            Width           =   2415
            Begin ControlesUteis.txt vProdTotal 
               Height          =   555
               Left            =   330
               TabIndex        =   104
               Top             =   300
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   979
               Tamanho         =   1815
               Text            =   ""
               Caption         =   "vProd"
               Locked          =   -1  'True
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
            Begin ControlesUteis.txt vNF 
               Height          =   555
               Left            =   330
               TabIndex        =   105
               Top             =   1470
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   979
               Tamanho         =   1815
               Text            =   ""
               CaptionColor    =   128
               Caption         =   "vNF"
               Locked          =   -1  'True
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
            Begin ControlesUteis.txt vTotTrib 
               Height          =   555
               Left            =   330
               TabIndex        =   106
               Top             =   900
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   979
               Tamanho         =   1815
               Text            =   ""
               Caption         =   "vTotTrib"
               Locked          =   -1  'True
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "FRETE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   2055
            Left            =   10740
            TabIndex        =   92
            Top             =   360
            Width           =   1905
            Begin ControlesUteis.txt vFrete 
               Height          =   555
               Left            =   150
               TabIndex        =   93
               Top             =   270
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vFrete"
               Enabled         =   0   'False
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
            Begin ControlesUteis.txt vSeg 
               Height          =   555
               Left            =   150
               TabIndex        =   98
               Top             =   870
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vSeg"
               Enabled         =   0   'False
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
            Begin ControlesUteis.txt vDesc 
               Height          =   555
               Left            =   150
               TabIndex        =   99
               Top             =   1440
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vDesc"
               Enabled         =   0   'False
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "SUBST. TRIBUT�RIA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   2055
            Left            =   7440
            TabIndex        =   90
            Top             =   360
            Width           =   3255
            Begin ControlesUteis.txt vFCP 
               Height          =   555
               Left            =   1665
               TabIndex        =   94
               Top             =   270
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vFCP"
               Locked          =   -1  'True
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
            Begin ControlesUteis.txt vBCST 
               Height          =   555
               Left            =   210
               TabIndex        =   95
               Top             =   270
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vBCST"
               Locked          =   -1  'True
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
            Begin ControlesUteis.txt vST 
               Height          =   555
               Left            =   210
               TabIndex        =   96
               Top             =   840
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vST"
               Locked          =   -1  'True
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
            Begin ControlesUteis.txt vFCPST 
               Height          =   555
               Left            =   225
               TabIndex        =   97
               Top             =   1410
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vFCPST"
               Locked          =   -1  'True
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
            Begin ControlesUteis.txt vFCPSTRet 
               Height          =   555
               Left            =   1650
               TabIndex        =   100
               Top             =   840
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vFCPSTRet"
               Locked          =   -1  'True
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "PIS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   2055
            Left            =   3690
            TabIndex        =   86
            Top             =   360
            Width           =   1905
            Begin ControlesUteis.txt vPIS 
               Height          =   555
               Left            =   180
               TabIndex        =   91
               Top             =   270
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vPIS"
               Enabled         =   0   'False
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "COFINS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   2055
            Left            =   5640
            TabIndex        =   85
            Top             =   360
            Width           =   1725
            Begin ControlesUteis.txt vCOFINS 
               Height          =   555
               Left            =   150
               TabIndex        =   89
               Top             =   300
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vCOFINS"
               Enabled         =   0   'False
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
            Begin ControlesUteis.txt vII 
               Height          =   555
               Left            =   150
               TabIndex        =   101
               Top             =   870
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vII"
               Enabled         =   0   'False
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
            Begin ControlesUteis.txt vOutro 
               Height          =   555
               Left            =   150
               TabIndex        =   102
               Top             =   1440
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vOutro"
               Enabled         =   0   'False
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "IPI"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   2055
            Left            =   1950
            TabIndex        =   84
            Top             =   360
            Width           =   1695
            Begin ControlesUteis.txt vIPI 
               Height          =   555
               Left            =   150
               TabIndex        =   87
               Top             =   900
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vIPI"
               Enabled         =   0   'False
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
            Begin ControlesUteis.txt vIPIDevol 
               Height          =   555
               Left            =   150
               TabIndex        =   88
               Top             =   270
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vIPIDevol"
               Enabled         =   0   'False
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "ICMS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   2055
            Left            =   120
            TabIndex        =   80
            Top             =   360
            Width           =   1785
            Begin ControlesUteis.txt vICMSDeson 
               Height          =   555
               Left            =   195
               TabIndex        =   81
               Top             =   1440
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vICMSDeson"
               Enabled         =   0   'False
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
            Begin ControlesUteis.txt vICMS 
               Height          =   555
               Left            =   195
               TabIndex        =   82
               Top             =   870
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vICMS"
               Enabled         =   0   'False
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
            Begin ControlesUteis.txt vBC 
               Height          =   555
               Left            =   180
               TabIndex        =   83
               Top             =   300
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   979
               Tamanho         =   1365
               Text            =   ""
               Caption         =   "vBC"
               Enabled         =   0   'False
               Negative        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Emitente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2205
         Left            =   60
         TabIndex        =   14
         Top             =   2655
         Width           =   15240
         Begin ControlesUteis.txt xFant 
            Height          =   555
            Left            =   9150
            TabIndex        =   15
            Top             =   390
            Width           =   5940
            _ExtentX        =   10478
            _ExtentY        =   979
            Tamanho         =   5940
            Text            =   ""
            Caption         =   "Nome fantasia"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt xNome 
            Height          =   555
            Left            =   180
            TabIndex        =   16
            Top             =   390
            Width           =   5010
            _ExtentX        =   8837
            _ExtentY        =   979
            Tamanho         =   5010
            Text            =   ""
            Caption         =   "Nome raz�o social"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt CNPJ 
            Height          =   555
            Left            =   5190
            TabIndex        =   17
            Top             =   390
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   979
            Tamanho         =   1605
            Tipo            =   8
            Text            =   ""
            Caption         =   "CNPJ"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt XML 
            Height          =   555
            Left            =   9480
            TabIndex        =   18
            Top             =   1605
            Width           =   5610
            _ExtentX        =   9895
            _ExtentY        =   979
            Tamanho         =   5610
            Text            =   ""
            Caption         =   "Arquivo XML"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt fone 
            Height          =   555
            Left            =   8100
            TabIndex        =   19
            Top             =   1605
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   979
            Tamanho         =   1380
            Text            =   ""
            Caption         =   "Telefone"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt xPais 
            Height          =   555
            Left            =   5940
            TabIndex        =   20
            Top             =   1605
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   979
            Tamanho         =   1500
            Text            =   ""
            Caption         =   "Pa�s"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt CEP 
            Height          =   555
            Left            =   4995
            TabIndex        =   21
            Top             =   1605
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   979
            Tamanho         =   945
            Tipo            =   9
            Text            =   ""
            Caption         =   "CEP"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt UF 
            Height          =   555
            Left            =   4665
            TabIndex        =   22
            Top             =   1605
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   979
            Tamanho         =   330
            Text            =   ""
            Caption         =   "UF"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt xMun 
            Height          =   555
            Left            =   1020
            TabIndex        =   23
            Top             =   1605
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   979
            Tamanho         =   3645
            Text            =   ""
            Caption         =   "Municipio"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt nro 
            Height          =   555
            Left            =   9495
            TabIndex        =   24
            Top             =   975
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   979
            Tamanho         =   690
            Text            =   ""
            Caption         =   "Numero"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt xLgr 
            Height          =   555
            Left            =   180
            TabIndex        =   25
            Top             =   975
            Width           =   9315
            _ExtentX        =   16431
            _ExtentY        =   979
            Tamanho         =   9315
            Text            =   ""
            Caption         =   "Logradouto"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt xBairro 
            Height          =   555
            Left            =   10185
            TabIndex        =   26
            Top             =   975
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   979
            Tamanho         =   4905
            Text            =   ""
            Caption         =   "Bairro"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt cPais 
            Height          =   555
            Left            =   7440
            TabIndex        =   50
            Top             =   1605
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   979
            Tamanho         =   660
            Text            =   ""
            Caption         =   "C�digo"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt IE 
            Height          =   555
            Left            =   6810
            TabIndex        =   51
            Top             =   390
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   979
            Tamanho         =   1380
            Text            =   ""
            Caption         =   "Inscri��o estadual"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt CRT 
            Height          =   555
            Left            =   8190
            TabIndex        =   52
            Top             =   390
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   979
            Tamanho         =   960
            Text            =   ""
            Caption         =   "CRT"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt cMun 
            Height          =   555
            Left            =   180
            TabIndex        =   53
            Top             =   1605
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   979
            Tamanho         =   840
            Text            =   ""
            Caption         =   "cMun"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Dados principais"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2145
         Left            =   60
         TabIndex        =   3
         Top             =   420
         Width           =   15225
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
            ItemData        =   "frm_XML.frx":B371
            Left            =   10140
            List            =   "frm_XML.frx":B373
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   107
            ToolTipText     =   "Empresa."
            Top             =   570
            Width           =   4935
         End
         Begin ControlesUteis.txt natOp 
            Height          =   555
            Left            =   270
            TabIndex        =   4
            Top             =   960
            Width           =   11400
            _ExtentX        =   20108
            _ExtentY        =   979
            Tamanho         =   11400
            Text            =   ""
            Caption         =   "Natureza de opera��o"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt nNF 
            Height          =   555
            Left            =   270
            TabIndex        =   5
            Top             =   390
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   979
            Tamanho         =   900
            Text            =   ""
            Caption         =   "Nota fiscal"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
         End
         Begin ControlesUteis.txt Serie 
            Height          =   555
            Left            =   1170
            TabIndex        =   6
            Top             =   390
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   979
            Tamanho         =   450
            Text            =   ""
            Caption         =   "S�rie"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt dhEmi 
            Height          =   555
            Left            =   1620
            TabIndex        =   7
            Top             =   390
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   979
            Tamanho         =   1710
            Tipo            =   0
            Text            =   ""
            Caption         =   "Data - hora emiss�o"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt dhSaiEnt 
            Height          =   555
            Left            =   3330
            TabIndex        =   8
            Top             =   390
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   979
            Tamanho         =   1710
            Tipo            =   0
            Text            =   ""
            Caption         =   "Data - hora sa�da"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt nProt 
            Height          =   555
            Left            =   270
            TabIndex        =   9
            Top             =   1545
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   979
            Tamanho         =   1470
            Tipo            =   0
            Text            =   ""
            CaptionColor    =   128
            Caption         =   "Protocolo SEFAZ"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
         End
         Begin ControlesUteis.txt indmod 
            Height          =   555
            Left            =   5040
            TabIndex        =   10
            Top             =   390
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   979
            Tamanho         =   450
            Tipo            =   0
            Text            =   ""
            Caption         =   "Mod."
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt indFinal 
            Height          =   555
            Left            =   5490
            TabIndex        =   11
            Top             =   390
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   979
            Tamanho         =   1935
            Tipo            =   0
            Text            =   ""
            Caption         =   "Consumidor Final"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt indPres 
            Height          =   555
            Left            =   7425
            TabIndex        =   12
            Top             =   390
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   979
            Tamanho         =   1935
            Tipo            =   0
            Text            =   ""
            Caption         =   "Presen�a do comprador"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt finNFe 
            Height          =   555
            Left            =   11670
            TabIndex        =   13
            Top             =   960
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   979
            Tamanho         =   3360
            Tipo            =   0
            Text            =   ""
            Caption         =   "Finalidade emiss�o"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt tpNF 
            Height          =   555
            Left            =   8730
            TabIndex        =   39
            Top             =   1530
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   979
            Tamanho         =   540
            Tipo            =   0
            Text            =   ""
            Caption         =   "tpNF"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt idDest 
            Height          =   555
            Left            =   9270
            TabIndex        =   40
            Top             =   1530
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   979
            Tamanho         =   540
            Tipo            =   0
            Text            =   ""
            Caption         =   "idDest"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt cMunFG 
            Height          =   555
            Left            =   9810
            TabIndex        =   41
            Top             =   1530
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   979
            Tamanho         =   900
            Tipo            =   0
            Text            =   ""
            Caption         =   "cMunFG"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt tpImp 
            Height          =   555
            Left            =   10710
            TabIndex        =   42
            Top             =   1530
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   979
            Tamanho         =   540
            Tipo            =   0
            Text            =   ""
            Caption         =   "tpImp"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt tpEmis 
            Height          =   555
            Left            =   11250
            TabIndex        =   43
            Top             =   1530
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   979
            Tamanho         =   540
            Tipo            =   0
            Text            =   ""
            Caption         =   "tpEmis"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt cDV 
            Height          =   555
            Left            =   11790
            TabIndex        =   44
            Top             =   1530
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   979
            Tamanho         =   540
            Tipo            =   0
            Text            =   ""
            Caption         =   "cDV"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt tpAmb 
            Height          =   555
            Left            =   12330
            TabIndex        =   45
            Top             =   1530
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   979
            Tamanho         =   540
            Tipo            =   0
            Text            =   ""
            Caption         =   "tpAmb"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt procEmi 
            Height          =   555
            Left            =   12870
            TabIndex        =   46
            Top             =   1530
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   979
            Tamanho         =   690
            Tipo            =   0
            Text            =   ""
            Caption         =   "procEmi"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt verProc 
            Height          =   555
            Left            =   13560
            TabIndex        =   47
            Top             =   1530
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   979
            Tamanho         =   1470
            Tipo            =   0
            Text            =   ""
            Caption         =   "verProc"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt cUF 
            Height          =   555
            Left            =   9360
            TabIndex        =   48
            Top             =   390
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   979
            Tamanho         =   750
            Tipo            =   0
            Text            =   ""
            Caption         =   "cUF"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
         End
         Begin ControlesUteis.txt chNF 
            Height          =   555
            Left            =   1740
            TabIndex        =   49
            Top             =   1530
            Width           =   4260
            _ExtentX        =   7514
            _ExtentY        =   979
            Tamanho         =   4260
            Text            =   ""
            CaptionColor    =   128
            Caption         =   "Chave de acesso da nota fiscal eletr�nica"
            Locked          =   -1  'True
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
         End
         Begin ControlesUteis.txt xMotivo 
            Height          =   555
            Left            =   6000
            TabIndex        =   109
            Top             =   1530
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   979
            Tamanho         =   2730
            Tipo            =   0
            Text            =   ""
            CaptionColor    =   128
            Caption         =   "Status SEFAZ"
            Negative        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   128
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empresa*"
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
            Left            =   12150
            TabIndex        =   108
            Top             =   360
            Width           =   840
         End
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   8880
         Left            =   -74970
         TabIndex        =   2
         Top             =   330
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   15663
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
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
         NumItems        =   18
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "Item"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "D"
            Text            =   "C�digo"
            Object.Width           =   2435
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "D"
            Text            =   "Descri��o"
            Object.Width           =   7409
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "NCM"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "CFOP"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "CST"
            Object.Width           =   1236
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "(Ipi)"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "(Pis)"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Object.Tag             =   "T"
            Text            =   "(Cofins)"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "UN"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Valor Unit."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Qtd."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Text            =   "Vlr. Total"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   13
            Text            =   "ICMS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   14
            Text            =   "vlr_ICMS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   15
            Text            =   "IPI"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   16
            Text            =   "vlr_IPI"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Object.Width           =   2540
         EndProperty
      End
      Begin ActiveResizeCtl.ActiveResize ActiveResize1 
         Left            =   60
         Top             =   390
         _ExtentX        =   847
         _ExtentY        =   847
         Resolution      =   99
         ScreenHeight    =   768
         ScreenWidth     =   1366
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
      Begin MSComctlLib.ListView ListaDuplicatas 
         Height          =   7320
         Left            =   -74910
         TabIndex        =   79
         Top             =   1530
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   12912
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "N� Duplicata"
            Object.Width           =   1501
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "D"
            Text            =   "Vencimento"
            Object.Width           =   2435
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "D"
            Text            =   "Valor"
            Object.Width           =   7409
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1350
      Top             =   1845
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15465
      _ExtentX        =   27279
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
      ButtonCaption1  =   "Importar XML"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Importar XML da nota fiscal de terceiros (F2)"
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
      ButtonWidth1    =   72
      ButtonHeight1   =   21
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
      ButtonLeft2     =   76
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
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
      ButtonLeft3     =   80
      ButtonTop3      =   2
      ButtonWidth3    =   36
      ButtonHeight3   =   21
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
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
      ButtonLeft4     =   118
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonEnabled5  =   0   'False
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   146
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   3300
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frm_XML.frx":B375
         Count           =   1
      End
   End
End
Attribute VB_Name = "frm_XML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Public PosicaoBase As Integer
Public lngPosicaoInicial As Long
Public lngPosicaoFinal As Long
Public lngPosicaoAuxiliar As Long

Public i As Integer
Public n As Long
Public lLinha As Integer

Public Function ProcCarregacampo(V1 As String, V2 As String, V3 As Integer)
On Error GoTo tratar_erro
    
        lngPosicaoInicial = InStr(IIf(PosicaoBase > 0, PosicaoBase, 1), strarquivo, V1, 1)
        lngPosicaoFinal = InStr(IIf(PosicaoBase > 0, PosicaoBase, 1), strarquivo, V2, 1)
        
    If lngPosicaoFinal > 0 And lngPosicaoInicial > 0 Then
        If lngPosicaoFinal > lngPosicaoInicial Then
            ProcCarregacampo = Mid(strarquivo, lngPosicaoInicial + V3, (lngPosicaoFinal - (lngPosicaoInicial + V3)))
            PosicaoBase = lngPosicaoFinal
            'Debug.print PosicaoBase
        End If
    End If
    
Exit Function
tratar_erro:
    MsgBox ("Descri��o do erro : " + Error()), vbCritical
    Exit Function
End Function

Public Sub ProcCriarNotaXML()
On Error GoTo tratar_erro

'If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

ID_nota = 0
Acao = "emitir a nota"
If Cmb_empresa = "" Then
    NomeCampo = "a empresa"
    ProcVerificaAcao
    Cmb_empresa1.SetFocus
    Exit Sub
End If

'Verifica se tem algum produto/servi�o recebido para o pedido
If Lista.ListItems.Count = 0 Then
    USMsgBox ("� necess�rio receber o(s) produto(s) antes de emitir a nota."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If


'Cria a nota fiscal
'If nNF.Text <> "" Then TextoFiltro = " and Serie = '" & Serie & "'" Else TextoFiltro = ""
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_Dados_Nota_Fiscal where int_NotaFiscal = '" & nNF.Text & "' and txt_CNPJ_CPF = '" & CNPJ.Text & "' and int_TipoNota = 2", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
End If

    TBGravar!TabelaSN = 0
    TBGravar!Regime = FunVerifRegimeEmpresa(Cmb_empresa.ItemData(Cmb_empresa.ListIndex))
    TBGravar!pedido_interno = False
    TBGravar!DtValidacaoOF = Now
    TBGravar!RespValidacaoOF = pubUsuario
    TBGravar!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    TBGravar!int_NotaFiscal = nNF.Text
    TBGravar!Serie = Serie.Text
    TBGravar!int_TipoNota = "2"
    TBGravar!TipoNF = "M1"
    TBGravar!dt_DataEmissao = dhEmi.Text
    TBGravar!dt_Saida_Entrada = dhSaiEnt.Text
    TBGravar!Hora_emissao = Format(dhEmi.Text, "hh:mm")
    TBGravar!Modelo = indmod.Text
    TBGravar!DtValidacao = Date
    TBGravar!RespValidacao = pubUsuario

    Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select * from clientes where CPF_CNPJ = '" & CNPJ.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False Then
    TBGravar!Id_Int_Cliente = TBClientes!IDCliente
    TBGravar!txt_Razao_Nome = TBClientes!NomeRazao
        
        If IsNull(TBClientes!Tipo_endereco) = False And TBClientes!Tipo_endereco <> "" Then
            Endereco = TBClientes!Tipo_endereco & ": " & IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
        Else
            Endereco = IIf(IsNull(TBClientes!Endereco), "", TBClientes!Endereco)
        End If
        TBGravar!txt_Endereco = Endereco
        TBGravar!Numero = IIf(IsNull(TBClientes!Numero), "", TBClientes!Numero)
        If IsNull(TBClientes!Tipo_bairro) = False And TBClientes!Tipo_bairro <> "" Then
            Bairro = TBClientes!Tipo_bairro & ": " & IIf(IsNull(TBClientes!Bairro), "", TBClientes!Bairro)
        Else
            Bairro = IIf(IsNull(TBClientes!Bairro), "", TBClientes!Bairro)
        End If
        TBGravar!txt_Bairro = Bairro
        
            TBGravar!txt_tipocliente = IIf(IsNull(TBClientes!Tipo), "", TBClientes!Tipo)
            If TBClientes!Tipo = "JP" Or TBClientes!Tipo = "JR" Then TBGravar!txt_IE_Cliente = IIf(IsNull(TBClientes!RG_IE), "", TBClientes!RG_IE)
            TBGravar!txt_UF = IIf(IsNull(TBClientes!UF), "", TBClientes!UF)
            TBGravar!txt_Fone_Fax = IIf(IsNull(TBClientes!Tel01), "", TBClientes!Tel01)
            If TBClientes!chkSuframa = True Then Suframa = True Else Suframa = False
'            TBGravar!txt_UF = IIf(IsNull(TBClientes!Estado), "", TBClientes!Estado)
            TBGravar!txt_Fone_Fax = IIf(IsNull(TBClientes!Tel01), "", TBClientes!Tel01)
            Suframa = False
        If TBClientes!idTipoEmpresa = 1 Then TBGravar!txt_CNPJ_CPF = IIf(IsNull(TBClientes!CPF_CNPJ), "", TBClientes!CPF_CNPJ)
        TBGravar!Txt_CEP = IIf(IsNull(TBClientes!CEP), "", TBClientes!CEP)
        TBGravar!txt_Municipio = IIf(IsNull(TBClientes!Cidade), "", TBClientes!Cidade)
    End If
    
    TBGravar!txt_Hora_Saida = Format(dhSaiEnt.Text, "hh:mm")
    TBGravar!Int_status = "1"
    TBGravar!Aplicacao = "T"
    TBGravar.Update
    
'Else
'USMsgBox "Aten��o, n�os ser� possivel executar a opera��o de cria��o pois j� existe cadastro da nota fiscal " & nNF.Text & " do cliente " & xNome.Text & " no banco de dados.", vbCritical, "CAPRIND v5.0"
'Exit Sub
'End If

ID_nota = TBGravar!ID
TBGravar.Close
USMsgBox "Dados da nota fiscal cadastrados com sucesso!", vbInformation, "CAPRIND v5.0"
'===========================================================================================
'Gravar chave de acesso
'===========================================================================================
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_Dados_Nota_Fiscal_NFe where ID_nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
End If
    TBGravar!ID_nota = ID_nota
    TBGravar!Chave_acesso = chNF.Text
    TBGravar!Finalidade_emissao = "1" 'finNFe.Text
    
    Select Case indFinal.Text
        Case "N�o" '"0"
        TBGravar!Consumidor_final = "0"
        Case "Consumidor final" '"1"
        TBGravar!Consumidor_final = "1"
    End Select
       
       
    Select Case indPres.Text
        Case "N�o se aplica" '"0"
        TBGravar!Presenca_comprador = "0"
        Case "Opera��o presencial" '"1"
         TBGravar!Presenca_comprador = "1"
        Case "Opera��o n�o presencial, pela Internet" '"2"
         TBGravar!Presenca_comprador = "2"
        Case "Opera��o n�o presencial, Teleatendimento;" '"3"
         TBGravar!Presenca_comprador = "3"
        Case "NFC-e em opera��o com entrega em domic�lio;" '"4"
         TBGravar!Presenca_comprador = "4"
        Case "Opera��o presencial, fora do estabelecimento" '"5"
         TBGravar!Presenca_comprador = "5"
        Case "Opera��o n�o presencial, outros." '"9"
         TBGravar!Presenca_comprador = "9"
    End Select
       
    TBGravar.Update
'End If
TBGravar.Close

'===========================================================================================
'Cadastra lista de produtos
'===========================================================================================

Contador = Lista.ListItems.Count
        Do While Contador > 0
           cProd = Lista.ListItems.Item(Contador).ListSubItems(1).Text
           uCom = Lista.ListItems.Item(Contador).ListSubItems(9).Text
           vUnCom = Lista.ListItems.Item(Contador).ListSubItems(10).Text
           qCom = Lista.ListItems.Item(Contador).ListSubItems(11).Text
           CFOP = Lista.ListItems.Item(Contador).ListSubItems(4).Text
           NCM = Lista.ListItems.Item(Contador).ListSubItems(3).Text
           ICMSCST = Lista.ListItems.Item(Contador).ListSubItems(5).Text
          IPICST = Lista.ListItems.Item(Contador).ListSubItems(6).Text
          PISCST = Lista.ListItems.Item(Contador).ListSubItems(7).Text
          COFINSCST = Lista.ListItems.Item(Contador).ListSubItems(8).Text
          ICMSpICMS = Lista.ListItems.Item(Contador).ListSubItems(13).Text
          ICMSvICMS = Lista.ListItems.Item(Contador).ListSubItems(14).Text
          IPIpIPI = Lista.ListItems.Item(Contador).ListSubItems(15).Text
          IPIvIPI = Lista.ListItems.Item(Contador).ListSubItems(16).Text
          
'======================================================================
' Verifica se tem o item cadastrado como c�digo de referencia
'======================================================================
Codproduto = ""
    Set TBComponente = CreateObject("adodb.recordset")
    TBComponente.Open "Select * from item_aplicacoes where n_referencia = '" & cProd & "'", Conexao, adOpenKeyset, adLockOptimistic
      If TBComponente.EOF = False Then
      Codproduto = TBComponente!Codproduto
'================================================================
      Else
      Codproduto = 0
      cCod = Contador
frmFaturamento_Importacao.txtdescricao.Text = cDesc

frmFaturamento_Importacao.Show 1

'If Cod_produto = "" Then Exit Sub

 TBItem.Open "Select * from projProduto where Desenho = '" & Cod_produto & "'", Conexao, adOpenKeyset, adLockOptimistic
   If TBItem.EOF = False Then
   cProd = TBItem!Codproduto
   cCod = TBItem!Desenho
   cDesc = TBItem!Descricao
   End If
End If
'================================================================
      
           Set TBItem = CreateObject("adodb.recordset")
             TBItem.Open "Select * from projProduto where codproduto = " & Codproduto & "", Conexao, adOpenKeyset, adLockOptimistic
                If TBItem.EOF = False Then
                 Set TBAbrir = CreateObject("adodb.recordset")
                     TBAbrir.Open "Select * from tbl_Detalhes_Nota where int_Cod_Produto = '" & cProd & "' and id_nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = True Then TBAbrir.AddNew
                          TBAbrir!Tipo = "P"
                          TBAbrir!int_Cod_Produto = TBItem!Desenho
                          TBAbrir!N_referencia = cProd
                          TBAbrir!int_NotaFiscal = nNF.Text
                          TBAbrir!ID_nota = ID_nota
                          TBAbrir!int_Qtd = Replace(qCom, ".", ",") * FunVerificaTabelaConversaoUnidade(TBItem!Unidade, TBItem!Unidade_com)
                          TBAbrir!Saldo = TBAbrir!int_Qtd
                          TBAbrir!Codproduto = TBItem!Codproduto
                          TBAbrir!Txt_descricao = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
                          
                          Set TBCFOP = CreateObject("adodb.recordset")
                          CFOP = Format(CFOP, "@.@@@")
                            TBCFOP.Open "Select * from tbl_NaturezaOperacao where id_CFOP = '" & CFOP & "'", Conexao, adOpenKeyset, adLockOptimistic
                             If TBCFOP.EOF = False Then
                             TBAbrir!ID_CFOP = IIf(IsNull(TBCFOP!IDCountCfop), 0, TBCFOP!IDCountCfop)
                             End If
                             TBCFOP.Close
                          
                          Set TBAliquota = CreateObject("adodb.recordset")
                          NCM = Format(NCM, "@@@@.@@.@@")
                            TBAliquota.Open "Select * from tbl_ClassificacaoFiscal where IDIntClasse = '" & NCM & "'", Conexao, adOpenKeyset, adLockOptimistic
                             If TBAliquota.EOF = False Then
                             TBAbrir!ID_CF = IIf(IsNull(TBAliquota!Idclass), 0, TBAliquota!Idclass)
                             End If
                             TBAliquota.Close
                          
                          TBAbrir!txt_Unid = uCom
                          TBAbrir!Unidade_com = uCom
                          TBAbrir!Familia = IIf(IsNull(TBItem!Classe), "", TBItem!Classe)
                          TBAbrir!dbl_ValorUnitario = vUnCom
                          TBAbrir!dbl_ValorTotal = Format(TBAbrir!dbl_ValorUnitario * TBAbrir!int_Qtd, "###,##0.00")
                          
                          TBAbrir!int_ICMS = ICMSpICMS
                          TBAbrir!ICMS_SN = ICMSpICMS
                          
                          TBAbrir!int_IPI = IPIpIPI
                          TBAbrir!dbl_valoripi = IPIvIPI
                          
                          TBAbrir!txt_CST = ICMSCST
                          TBAbrir!CST_IPI = IPICST
                          TBAbrir!CST_PIS = PISCST
                          TBAbrir!CST_Cofins = COFINSCST
                          
                          TBAbrir.Update
                          TBAbrir.Close
                          TBItem.Close
                      End If
            Contador = Contador - 1
        Loop
USMsgBox "Produtos da nota fiscal cadastrados com sucesso!", vbInformation, "CAPRIND v5.0"

'========================================================================================================
' Cadastrar totais da nota fiscal
'========================================================================================================
 Set TBTotaisnota = CreateObject("adodb.recordset")
   TBTotaisnota.Open "Select * from tbl_Totais_Nota where ID_Nota = " & ID_nota, Conexao, adOpenKeyset, adLockOptimistic
    If TBTotaisnota.EOF = True Then
    TBTotaisnota.AddNew
    TBTotaisnota!ID_nota = ID_nota
    TBTotaisnota!dbl_Base_ICMS = vBC.Text
    TBTotaisnota!dbl_Valor_ICMS = vICMS.Text
    TBTotaisnota!dbl_Base_ICMS_Subst = vBCST.Text
    TBTotaisnota!dbl_Valor_ICMS_Subst = vST.Text
    TBTotaisnota!dbl_Valor_Total_Produtos = vProdTotal.Text
    TBTotaisnota!dbl_Valor_Frete = vFrete.Text
    TBTotaisnota!dbl_Valor_Seguro = vSeg.Text
    TBTotaisnota!dbl_Desp_Adicionais = vOutro.Text
    TBTotaisnota!dbl_Valor_Total_IPI = vIPI.Text
    TBTotaisnota!dbl_Valor_Total_Nota = vNF.Text
    TBTotaisnota.Update
    End If
    TBTotaisnota.Close

USMsgBox "Nota fiscal criada com sucesso!", vbInformation, "CAPRIND v5.0"

If USMsgBox("Deseja efetuar a entrada no estoque do(s) produto(s) agora?", vbYesNo, "CAPRIND v5.0") = vbYes Then
ProcEntradaProdutoEstoque
End If
   


Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Function ProcImportarXML(strCaminho) As Long
On Error GoTo tratar_erro
PosicaoBase = 0

    Lista.ListItems.Clear
    lLinha = 1
    XML.Text = strCaminho
    ' Ler arquivo XML
    n = FreeFile()
    Open strCaminho For Input As #n
    strarquivo = Input(LOF(n), n)
'    txtxml.Text = Replace(strarquivo, "﻿", "")
    strarquivo = Replace(strarquivo, "﻿", "")
    Close #n
    
    infNFe = ProcCarregacampo("<infNFe", "/infNFe>", Len("<infNFe"))
    'Debug.print Left$(infNFe, 4)
    If Left$(infNFe, 4) = " Id=" Then
    infNFe = Left$(infNFe, 52)
    infNFe = Right$(infNFe, 44)
    Else
    infNFe = Left$(infNFe, 66)
    infNFe = Right$(infNFe, 44)
    End If
    chNF.Text = infNFe
    
'Dados da nota fiscal
    V1 = "ide"
    PosicaoBase = InStr(1, strarquivo, V1, 1)

    'Dados da nota fiscal
    cUF.Text = ProcCarregacampo("<cUF>", "</cUF>", Len("<cUF>"))
    
    natOp.Text = UCase(ProcCarregacampo("<natOp>", "</natOp>", Len("<natOp>")))
    indmod.Text = UCase(ProcCarregacampo("<mod>", "</mod>", Len("<mod>")))
    Serie.Text = ProcCarregacampo("<serie>", "</serie>", Len("<serie>"))
    nNF.Text = ProcCarregacampo("<nNF>", "</nNF>", Len("<nNF>"))
    nNF.Text = FunTamanhoTextoZeroEsq(ReturnNumbersOnly(nNF.Text), 9)
    

   
    dhEmi.Text = ProcCarregacampo("<dhEmi>", "</dhEmi>", Len("<dhEmi>"))
    dhEmi.Text = Replace(dhEmi.Text, "T", " ")
    dhEmi.Text = Left$(dhEmi.Text, 19)
    dhEmi.Text = Format(dhEmi.Text, "General date")
   
    dhSaiEnt.Text = ProcCarregacampo("<dhSaiEnt>", "</dhSaiEnt>", Len("<dhSaiEnt>"))
    dhSaiEnt.Text = Replace(dhSaiEnt.Text, "T", " ")
    dhSaiEnt.Text = Left$(dhSaiEnt.Text, 19)
    dhSaiEnt.Text = Format(dhSaiEnt.Text, "General date")
    
    tpNF.Text = UCase(ProcCarregacampo("<tpNF>", "</tpNF>", Len("<tpNF>")))
    idDest.Text = UCase(ProcCarregacampo("<idDest>", "</idDest>", Len("<idDest>")))
    cMunFG.Text = UCase(ProcCarregacampo("<cMunFG>", "</cMunFG>", Len("<cMunFG>")))
    
    tpImp.Text = UCase(ProcCarregacampo("<tpImp>", "</tpImp>", Len("<tpImp>")))
    tpEmis.Text = UCase(ProcCarregacampo("<tpEmis>", "</tpEmis>", Len("<tpEmis>")))
    'cMunFG.Text = UCase(ProcCarregacampo("<cMunFG>", "</cMunFG>", Len("<cMunFG>")))
    cDV.Text = UCase(ProcCarregacampo("<cDV>", "</cDV>", Len("<cDV>")))
    tpAmb.Text = UCase(ProcCarregacampo("<tpAmb>", "</tpAmb>", Len("<tpAmb>")))
    
    finNFe.Text = UCase(ProcCarregacampo("<finNFe>", "</finNFe>", Len("<finNFe>")))
    
    Select Case finNFe.Text
        Case "1"
        finNFe.Text = "NF-e normal"
        Case "2"
        finNFe.Text = "NF-e complementar"
        Case "3"
        finNFe.Text = "NF-e de ajuste"
        Case "4"
        finNFe.Text = "Devolu��o/Retorno"
    End Select
   
    
    indFinal.Text = UCase(ProcCarregacampo("<indFinal>", "</indFinal>", Len("<indFinal>")))
    Select Case indFinal.Text
        Case "0"
        indFinal.Text = "N�o"
        Case "1"
        indFinal.Text = "Consumidor final"
    End Select
    
    indPres.Text = UCase(ProcCarregacampo("<indPres>", "</indPres>", Len("<indPres>")))
    
    procEmi.Text = UCase(ProcCarregacampo("<procEmi>", "</procEmi>", Len("<procEmi>")))
    verProc.Text = UCase(ProcCarregacampo("<verProc>", "</verProc>", Len("<verProc>")))
    
    
    Select Case indPres.Text
        Case "0"
        indPres.Text = "N�o se aplica" ' (por exemplo, para a Nota Fiscal complementar ou de ajuste);
        Case "1"
        indPres.Text = "Opera��o presencial"
        Case "2"
        indPres.Text = "Opera��o n�o presencial, pela Internet"
        Case "3"
        indPres.Text = "Opera��o n�o presencial, Teleatendimento;"
        Case "4"
        indPres.Text = "NFC-e em opera��o com entrega em domic�lio;"
        Case "5"
        indPres.Text = "Opera��o presencial, fora do estabelecimento"
        Case "9"
        indPres.Text = "Opera��o n�o presencial, outros."
    End Select
    
    'Dados do emitente
    'CNPJ.Text = LerDadosXML(strarquivo, "SignatureValue", "")
    CNPJ.Text = ProcCarregacampo("<CNPJ>", "</CNPJ>", Len("<CNPJ>"))
    xNome.Text = UCase(ProcCarregacampo("<xNome>", "</xNome>", Len("<xNome>")))
    xFant.Text = UCase(ProcCarregacampo("<xFant>", "</xFant>", Len("<xFant>")))
    
    'Endere�o emitente
    xLgr.Text = UCase(ProcCarregacampo("<xLgr>", "</xLgr>", Len("<xLgr>")))
    nro.Text = ProcCarregacampo("<nro>", "</nro>", Len("<nro>"))
    xBairro.Text = UCase(ProcCarregacampo("<xBairro>", "</xBairro>", Len("<xBairro>")))
    cMun.Text = UCase(ProcCarregacampo("<cMun>", "</cMun>", Len("<cMun>")))
    xMun.Text = UCase(ProcCarregacampo("<xMun>", "</xMun>", Len("<xMun>")))
    UF.Text = UCase(ProcCarregacampo("<UF>", "</UF>", Len("<UF>")))
    CEP.Text = ProcCarregacampo("<CEP>", "</CEP>", Len("<CEP>"))
    cPais.Text = UCase(ProcCarregacampo("<cPais>", "</cPais>", Len("<cPais>")))
    xPais.Text = UCase(ProcCarregacampo("<xPais>", "</xPais>", Len("<xPais>")))
    Var1 = "fone"
    fone.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    IE.Text = UCase(ProcCarregacampo("<IE>", "</IE>", Len("<IE>")))
    CRT.Text = UCase(ProcCarregacampo("<CRT>", "</CRT>", Len("<CRT>")))

    'Dados do Destinat�rio
    dest_CNPJ.Text = ProcCarregacampo("<CNPJ>", "</CNPJ>", Len("<CNPJ>"))
    dest_xNome.Text = UCase(ProcCarregacampo("<xNome>", "</xNome>", Len("<xNome>")))
    
    'Endere�o Destinatario
    dest_xLgr.Text = UCase(ProcCarregacampo("<xLgr>", "</xLgr>", Len("<xLgr>")))
    dest_nro.Text = ProcCarregacampo("<nro>", "</nro>", Len("<nro>"))
    dest_xCpl.Text = UCase(ProcCarregacampo("<xCpl>", "</xCpl>", Len("<xCpl>")))
    dest_xBairro.Text = UCase(ProcCarregacampo("<xBairro>", "</xBairro>", Len("<xBairro>")))
    dest_xMun.Text = UCase(ProcCarregacampo("<xMun>", "</xMun>", Len("<xMun>")))
    dest_UF.Text = UCase(ProcCarregacampo("<UF>", "</UF>", Len("<UF>")))
    dest_CEP.Text = ProcCarregacampo("<CEP>", "</CEP>", Len("<CEP>"))
    dest_xPais.Text = UCase(ProcCarregacampo("<xPais>", "</xPais>", Len("<xPais>")))
    dest_indIEDest.Text = UCase(ProcCarregacampo("<indIEDest>", "</indIEDest>", Len("<indIEDest>")))
    
    Select Case dest_indIEDest.Text
      Case "1": dest_indIEDest.Text = "1 - Contribuinte ICMS (informar a IE do destinat�rio)"
      Case "2": dest_indIEDest.Text = "2 - Contribuinte isento de Inscri��o no cadastro de Contribuintes"
      Case "9": dest_indIEDest.Text = "9 - N�o Contribuinte, que pode ou n�o possuir Inscri��o Estadual no Cadastro de Contribuintes do ICMS."
    End Select
    '=====================================
    'Carrega Dados lista de produtos
    '=====================================
    
    V1 = "prod"
    PosicaoBase = InStr(IIf(lngPosicaoFinal > 0, lngPosicaoFinal, 1), strarquivo, V1, 1)
    
Inicio:
    
    If PosicaoBase > 0 Then

    Dim cProd As String, xProd As String, NCM As String, CFOP As String, uCom As String, qCom As String, vUnCom As String, vProd As String, orig As String, ICMS As String, v_ICMS As String, IPI As String, v_IPI As String ', CST As String
    
    Var1 = "cProd"
    cProd = UCase(ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
    Var1 = "xProd"
    xProd = UCase(ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
    Var1 = "NCM"
    NCM = UCase(ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
    Var1 = "CFOP"
    CFOP = UCase(ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
    Var1 = "uCom"
    uCom = UCase(ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
    Var1 = "qCom"
    qCom = UCase(ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
    Var1 = "vUnCom"
    vUnCom = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    Var1 = "vProd"
    vProd = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

'Carrega impostos do produto
    V1 = "imposto"
    PosicaoBase = InStr(IIf(lngPosicaoFinal > 0, lngPosicaoFinal, 1), strarquivo, V1, 1)
'Debug.print PosicaoBase
If PosicaoBase > 0 Then
    Var1 = "orig"
    orig = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
        
    Var1 = "CSOSN"
    
    CST = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    If CST = "" Then
        Var1 = "CST"
        CST = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    End If
    
    orig = orig & CST

    Var1 = "CST"
    V1 = "IPI"
    V2 = "PIS"
    PosicaoFinal = InStr(lngPosicaoFinal, strarquivo, V2, 1)
    PosicaoBase = InStr(lngPosicaoFinal, strarquivo, V1, 1)
    
    If PosicaoBase < PosicaoFinal Then
    CSTIPI = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    End If
    
    V1 = "PIS"
    PosicaoBase = InStr(lngPosicaoFinal, strarquivo, V1, 1)
    CSTPIS = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
    V1 = "COFINS"
    PosicaoBase = InStr(lngPosicaoFinal, strarquivo, V1, 1)
    CSTCOFINS = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
   
If cProd <> "" Then

If IsNumeric(vUnCom) Then
    vUnCom = Replace(vUnCom, ".", ",")
    vUnCom = "R$ " & vUnCom
End If

If IsNumeric(vProd) Then
    vProd = Replace(vProd, ".", ",")
    vProd = "R$ " & vProd
End If

    
    ValorTotal = 0

        With Lista.ListItems
            .Add , , lLinha
            .Item(.Count).SubItems(1) = cProd
            .Item(.Count).SubItems(2) = xProd
            .Item(.Count).SubItems(3) = NCM
            .Item(.Count).SubItems(4) = CFOP
            .Item(.Count).SubItems(5) = orig
            .Item(.Count).SubItems(6) = CSTIPI
            .Item(.Count).SubItems(7) = CSTPIS
            .Item(.Count).SubItems(8) = CSTCOFINS
            .Item(.Count).SubItems(9) = uCom
            .Item(.Count).SubItems(10) = vUnCom
            .Item(.Count).SubItems(11) = qCom
            .Item(.Count).SubItems(12) = vProd
            .Item(.Count).SubItems(13) = "0,00"
            .Item(.Count).SubItems(14) = "0,00"
            .Item(.Count).SubItems(15) = "0,00"
            .Item(.Count).SubItems(16) = "0,00"
        End With
    
    lLinha = lLinha + 1
GoTo Inicio
End If

End If
End If

'Carregatotais da nota
    V1 = "total"
    PosicaoBase = InStr(1, strarquivo, V1, 1)

    Var1 = "vBC"
    vBC.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vBC.Text = Replace(vBC.Text, ".", ",")
    vBC.Text = Format(vBC.Text, "###,##0.00")

    Var1 = "vICMS"
    vICMS.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vICMS.Text = Replace(vICMS.Text, ".", ",")
    vICMS.Text = Format(vICMS.Text, "###,##0.00")
    
    Var1 = "vICMSDeson"
    vICMSDeson.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vICMSDeson.Text = Replace(vICMSDeson.Text, ".", ",")
    vICMSDeson.Text = Format(vICMSDeson.Text, "###,##0.00")
    
    Var1 = "vFCP"
    vFCP.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vFCP.Text = Replace(vFCP.Text, ".", ",")
    vFCP.Text = Format(vFCP.Text, "###,##0.00")
    
    Var1 = "vBCST"
    vBCST.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vBCST.Text = Replace(vBCST.Text, ".", ",")
    vBCST.Text = Format(vBCST.Text, "###,##0.00")
    
    Var1 = "vST"
    vST.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vST.Text = Replace(vST.Text, ".", ",")
    vST.Text = Format(vST.Text, "###,##0.00")
    
    Var1 = "vFCPST"
    vFCPST.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vFCPST.Text = Replace(vFCPST.Text, ".", ",")
    vFCPST.Text = Format(vFCPST.Text, "###,##0.00")
    
    Var1 = "vFCPSTRet"
    vFCPSTRet.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vFCPSTRet.Text = Replace(vFCPSTRet.Text, ".", ",")
    vFCPSTRet.Text = Format(vFCPSTRet.Text, "###,##0.00")
    
    
    Var1 = "vProd"
    vProdTotal.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vProdTotal.Text = Replace(vProdTotal.Text, ".", ",")
    vProdTotal.Text = Format(vProdTotal.Text, "###,##0.00")
    
    
    Var1 = "vFrete"
    vFrete.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vFrete.Text = Replace(vFrete.Text, ".", ",")
    vFrete.Text = Format(vFrete.Text, "###,##0.00")
    
    
    Var1 = "vSeg"
    vSeg.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vSeg.Text = Replace(vSeg.Text, ".", ",")
    vSeg.Text = Format(vSeg.Text, "###,##0.00")

    Var1 = "vDesc"
    vDesc.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vDesc.Text = Replace(vDesc.Text, ".", ",")
    vDesc.Text = Format(vDesc.Text, "###,##0.00")
    
    Var1 = "vII"
    vII.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vII.Text = Replace(vII.Text, ".", ",")
    vII.Text = Format(vII.Text, "###,##0.00")
    
    Var1 = "vIPI"
    vIPI.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vIPI.Text = Replace(vIPI.Text, ".", ",")
    vIPI.Text = Format(vIPI.Text, "###,##0.00")
    
    Var1 = "vIPIDevol"
    vIPIDevol.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vIPIDevol.Text = Replace(vIPIDevol.Text, ".", ",")
    vIPIDevol.Text = Format(vIPIDevol.Text, "###,##0.00")
    
    Var1 = "vPIS"
    vPIS.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vPIS.Text = Replace(vPIS.Text, ".", ",")
    vPIS.Text = Format(vPIS.Text, "###,##0.00")
    
    
    Var1 = "vCOFINS"
    vCOFINS.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vCOFINS.Text = Replace(vCOFINS.Text, ".", ",")
    vCOFINS.Text = Format(vCOFINS.Text, "###,##0.00")
    
    Var1 = "vOutro"
    vOutro.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vOutro.Text = Replace(vOutro.Text, ".", ",")
    vOutro.Text = Format(vOutro.Text, "###,##0.00")
    
    Var1 = "vNF"
    vNF.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vNF.Text = Replace(vNF.Text, ".", ",")
    vNF.Text = Format(vNF.Text, "###,##0.00")
    
    Var1 = "vTotTrib"
    vTotTrib.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    vTotTrib.Text = Replace(vTotTrib.Text, ".", ",")
    vTotTrib.Text = Format(vTotTrib.Text, "###,##0.00")
   
    
    'Carrega dados transporte
    
    Var1 = "CNPJ"
    transpCNPJ.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
  
    Var1 = "xNome"
    transpxNome.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
  
    Var1 = "IE"
    transpIE.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
  
    Var1 = "xEnder"
    transpxEnder.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
 
    Var1 = "xMun"
    transpxMun.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
    Var1 = "UF"
    transpUF.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

    Var1 = "qVol"
    transpqVol.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    transpqVol.Text = Replace(transpqVol.Text, ".", ",")
    transpqVol.Text = Format(transpqVol.Text, "###,##0.00")

    Var1 = "esp"
    transpesp.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

    Var1 = "marca"
    transpMarca.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
    Var1 = "nVol"
    transpnVol.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

    Var1 = "pesoL"
    transppesoL.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    transppesoL.Text = Replace(transppesoL.Text, ".", ",")
    transppesoL.Text = Format(transppesoL.Text, "###,##0.00")

    Var1 = "pesoB"
    transppesoB.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    transppesoB.Text = Replace(transppesoB.Text, ".", ",")
    transppesoB.Text = Format(transppesoB.Text, "###,##0.00")
    
'Carregar a fatura da Nfe

    Var1 = "nFat"
    fatnFat.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

    Var1 = "vOrig"
    fatvOrig.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

    Var1 = "vDesc"
    fatvDesc.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    fatvDesc.Text = Replace(fatvDesc.Text, ".", ",")
    fatvDesc.Text = Format(fatvDesc.Text, "###,##0.00")
  
    Var1 = "vLiq"
    fatvLiq.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    fatvLiq.Text = Replace(fatvLiq.Text, ".", ",")
    fatvLiq.Text = Format(fatvLiq.Text, "###,##0.00")
    
    
'Carregar lista de duplicatas
ListaDuplicatas.ListItems.Clear
'PosicaoBase = 0
'    V1 = "dup"
'    PosicaoBase = InStr(IIf(lngPosicaoFinal > 0, lngPosicaoFinal, 1), strarquivo, V1, 1)
    
    
Inicio2:
    If PosicaoBase > 0 Then
    
    Var1 = "nDup"
    nDup = UCase(ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
    
    Var1 = "dVenc"
    dVenc = UCase(ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
  
    Var1 = "vDup"
    vDup = UCase(ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">")))
    vDup = Replace(vDup, ".", ",")
    vDup = Format(vDup, "###,##0.00")
    
If nDup <> "" Then

          With ListaDuplicatas.ListItems
            .Add , , nDup
            .Item(.Count).SubItems(1) = dVenc
            .Item(.Count).SubItems(2) = vDup
        End With
    
    lLinha = lLinha + 1
GoTo Inicio2
End If
End If

     Var1 = "indPag"
    fatindPag.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
    Var1 = "tPag"
    fattPag.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    
    Var1 = "vPag"
    fatvPag.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))
    fatvPag.Text = Replace(fatvPag.Text, ".", ",")
    fatvPag.Text = Format(fatvPag.Text, "###,##0.00")
    
    
'Carregar dados adicionais
    Var1 = "infCpl"
    infCpl.Text = ProcCarregacampo("<" & Var1 & ">", "</" & Var1 & ">", Len("<" & Var1 & ">"))

'Carregar numero do protocolo de recebimento do SEFAZ
    nProt.Text = ProcCarregacampo("<nProt>", "</nProt>", Len("<nProt>"))

'Carregar status de recebimento do SEFAZ
    xMotivo.Text = ProcCarregacampo("<xMotivo>", "</xMotivo>", Len("<xMotivo>"))
    USMsgBox "Importa��o do XML efetuada com sucesso!", vbInformation, "CAPRIND v5.0"
    
    If USMsgBox("Deseja criar a nota fiscal eletr�nica agora?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    ProcCriarNotaXML
    End If

Exit Function
tratar_erro:
    MsgBox ("Descri��o do erro : " + Error()), vbCritical
    Exit Function
End Function

Sub ProcEnviaDados()
On Error GoTo tratar_erro

'Grava na tabela Estoque_Controle
TBGravar!ID_empresa = Cmb_empresa1.ItemData(Cmb_empresa1.ListIndex)
TBGravar!status = "CONSIGNA��O RECEBIDA"
TBGravar!emissaonf = txtemissao.Value
TBGravar!Consignacao = True
TBGravar!Ref = Cmb_cod_ref
TBGravar!LOTE = txtnotafiscal.Text
TBGravar!Desenho = txtdesenho.Text
TBGravar!Descricao = txtdesc.Text
TBGravar!peso_unit = txtpeso.Text
TBGravar!descricaotecnica = txtdesctecnica.Text
TBGravar!Data = txtData
TBGravar!estoque_real = Format(txtQtde.Text, "###.##0.000")
TBGravar!estoque_real_PC = IIf(Txt_qtde_PC = "", Null, Format(Txt_qtde_PC, "###.##0.000"))
TBGravar!estoque_venda = Format(txtQtde.Text, "###.##0.000")
TBGravar!Qtde = Format(txtQtde.Text, "###.##0.000")
TBGravar!Corrida = txtcorrida.Text
TBGravar!Certificado = txtCertificado.Text
TBGravar!Classe = txtfamilia.Text
TBGravar!Un = txtUN.Text
TBGravar!NF = txtnotafiscal.Text
TBGravar!Serie = Txt_serie
TBGravar!ID_Cliente = txtid_cliente.Text
TBGravar!Cliente = txtCliente.Text
TBGravar!Tipodest_NFcons = Txt_tipodest
TBGravar!valor_unitario = Format(txtVlr_unit, "###.##0.00000")
TBGravar!Valor_total = Format(txtVlr_total, "###.##0.00")
TBGravar!local_armaz = cmbLocal_armaz

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEntradaProdutoEstoque()
On Error GoTo tratar_erro
Dim ID_Cliente As Integer
Dim Nome_Razao As String

'===========================================================
' Localiza dados do cliente da nota fiscal
'===========================================================
 Set TBClientes = CreateObject("adodb.recordset")
 TBClientes.Open "Select * from clientes where CPF_CNPJ = '" & CNPJ.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
 If TBClientes.EOF = False Then
  ID_Cliente = TBClientes!IDCliente
  Nome_Razao = TBClientes!NomeRazao
 End If
 TBClientes.Close
'===========================================================
' Inicio da entrada do item no estoque
'===========================================================
' Verifica quandtidade de itens na lista de entrada
'===========================================================
Contador = Lista.ListItems.Count
'===========================================================
Do While Contador > 0
'===========================================================
' Atribui valores as variaveis do produto
'===========================================================
 cProd = Lista.ListItems.Item(Contador).ListSubItems(1).Text
 xProd = Lista.ListItems.Item(Contador).ListSubItems(2).Text
 uCom = Lista.ListItems.Item(Contador).ListSubItems(9).Text
 vUnCom = Lista.ListItems.Item(Contador).ListSubItems(10).Text
 qCom = Lista.ListItems.Item(Contador).ListSubItems(11).Text
 qCom = Replace(qCom, ".", ",")
'===========================================================
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Estoque_controle where Lote = '" & nNF.Text & "' and desenho = '" & cProd & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    Evento = "Novo"
End If
'===============================================================
' Localiza o codigo de referencia
'===============================================================
Set TBReferencia = CreateObject("adodb.recordset")
TBReferencia.Open "Select * from Item_aplicacoes where n_referencia = '" & cProd & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBReferencia.EOF = False Then
Codproduto = TBReferencia!Codproduto
End If
TBReferencia.Close

'==============================================================
' Localiza familia do produto
'==============================================================
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select * from Projproduto where codproduto = " & Codproduto & "", Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
  TBGravar!Classe = TBProduto!Classe
  TBGravar!peso_unit = TBProduto!PBruto
  TBGravar!Desenho = TBProduto!Desenho
  TBGravar!Descricao = TBProduto!Descricao
  TBGravar!descricaotecnica = TBProduto!descricaotecnica
End If
'===========================================================
'Grava lote do produto na tabela Estoque_Controle
'===========================================================
  TBGravar!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
  TBGravar!status = "CONSIGNA��O RECEBIDA"
  TBGravar!emissaonf = dhEmi.Text
  TBGravar!Consignacao = True
  TBGravar!Ref = Cmb_cod_ref
  TBGravar!LOTE = nNF.Text
  TBGravar!Data = dhEmi.Text
  TBGravar!estoque_real = Format(qCom, "###.##0.000")
  TBGravar!estoque_real_PC = IIf(qCom = "", Null, Format(qCom, "###.##0.000"))
  TBGravar!estoque_venda = Format(qCom, "###.##0.000")
  TBGravar!Qtde = Format(qCom, "###.##0.000")
  TBGravar!Corrida = "0"
  TBGravar!Certificado = "0"
  TBGravar!Un = uCom
  TBGravar!NF = nNF.Text
  TBGravar!Serie = Serie.Text
  TBGravar!ID_Cliente = ID_Cliente
  TBGravar!Cliente = Nome_Razao
  'TBGravar!Tipodest_NFcons = Txt_tipodest
  TBGravar!valor_unitario = Format(vUnCom, "###.##0.00000")
  TBGravar!Valor_total = Format((vUnCom * qCom), "###.##0.00")
  TBGravar!local_armaz = "INDUSTRIALIZA��O" 'cmbLocal_armaz
  TBGravar!Tipodest_NFcons = "C"
  TBGravar!Ref = cCod
TBGravar.Update
'===========================================================
' Grava a movimenta��o de entrada no estoque
'===========================================================
Set TBEstoque = CreateObject("adodb.recordset")
   TBEstoque.Open "Select * from estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
   TBEstoque.AddNew
   TBEstoque!Destino = "Interno"
   TBEstoque!Terceiros = False
   TBEstoque!IDEstoque = TBGravar!IDEstoque
   TBEstoque!Operacao = "ENTRADA_NOTA_FISCAL_CONSIGNA��O"
   TBEstoque!Desenho = TBProduto!Desenho
   TBEstoque!Documento = nNF.Text
   TBEstoque!LOTE = nNF.Text
   TBEstoque!Descricao = TBProduto!Descricao
   TBEstoque!DtEmissao = dhEmi.Text
   TBEstoque!Entrada = Format(qCom, "###.##0.000")
   TBEstoque!Entrada_PC = Format(qCom, "###.##0.000")
   TBEstoque!Responsavel = pubUsuario
   TBEstoque!Cliente = ID_Cliente
   TBEstoque!Data = Date
   TBEstoque!VlrUnit = Format(vUnCom, "###.##0.00000")
   TBEstoque!vlrTotal = Format((vUnCom * qCom), "###.##0.00")
   TBEstoque!Obs = "Entrada por importa��o de XML"
   TBEstoque.Update
   TBEstoque.Close
   Contador = Contador - 1
TBProduto.Close
Loop
''==================================
'Modulo = "Estoque/Recebimento/Consigna��o"
'ID_documento = TBGravar!IDestoque
'Documento = "Nota fiscal: " & txtnotafiscal & " - Emitente: " & txtcliente
'Documento1 = "C�digo interno: " & txtdesenho
'ProcGravaEvento
''==================================
TBGravar.Close
USMsgBox "Entrada de produto(s) da nota fiscal n� " & nNF.Text & " no estoque com sucesso!", vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descri��o do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirXML()
On Error GoTo tratar_erro
    CommonDialog1.Filter = "Arquivo XML (*.xml)|*.xml"
    CommonDialog1.ShowOpen
    strCaminho = CommonDialog1.filename
    If strCaminho = "" Then Exit Sub
If USMsgBox("Deseja realmente importar o XML " & strCaminho & "", vbYesNo, "CAPRIND v5.0") = vbYes Then
    ProcImportarXML (strCaminho)
    PosicaoBase = 1
Else
  USMsgBox "Importa��o cancelada com sucesso!", vbInformation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    MsgBox ("Descri��o do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub btnCriarNota_Click()
On Error GoTo tratar_erro

ProcCriarNotaXML

Exit Sub
tratar_erro:
    MsgBox ("Descri��o do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub btnReceber_estoque_Click()
On Error GoTo tratar_erro

ProcEntradaProdutoEstoque

Exit Sub
tratar_erro:
    MsgBox ("Descri��o do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 5, True
ProcRemoveObjetosResize Me
SSTab1.Tab = 0
ProcCarregaComboEmpresa Cmb_empresa, False

Exit Sub
tratar_erro:
    MsgBox ("Descri��o do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcAbrirXML
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descri��o do erro : " + Error()), vbCritical
    Exit Sub
End Sub

