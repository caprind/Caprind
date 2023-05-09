VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frm_Instituicoes2 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Financeiro - Instituições"
   ClientHeight    =   10035
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   15360
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Instituicoes2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
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
   Begin VB.ComboBox Cmb_empresa 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "frm_Instituicoes2.frx":1042
      Left            =   2715
      List            =   "frm_Instituicoes2.frx":1044
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1695
      Width           =   3330
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10035
      Left            =   -30
      TabIndex        =   82
      Top             =   0
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   17701
      _Version        =   393216
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
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
      TabPicture(0)   =   "frm_Instituicoes2.frx":1046
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame14"
      Tab(0).Control(1)=   "Frame18"
      Tab(0).Control(2)=   "Frame19"
      Tab(0).Control(3)=   "Frame13"
      Tab(0).Control(4)=   "Frame20"
      Tab(0).Control(5)=   "USImageList1"
      Tab(0).Control(6)=   "txtCodBanco"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(8)=   "USToolBar1"
      Tab(0).Control(9)=   "lst_Instituicoes"
      Tab(0).Control(10)=   "Frame5"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Movimentação financeira"
      TabPicture(1)   =   "frm_Instituicoes2.frx":1062
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "USImageList2"
      Tab(1).Control(1)=   "USToolBar2"
      Tab(1).Control(2)=   "SSTab3"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Extrato"
      TabPicture(2)   =   "frm_Instituicoes2.frx":107E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame41"
      Tab(2).Control(1)=   "FrameFiltro"
      Tab(2).Control(2)=   "Lst_extrato"
      Tab(2).Control(3)=   "USImageList3"
      Tab(2).Control(4)=   "USToolBar3"
      Tab(2).Control(5)=   "PBLista(4)"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Cheques emitidos"
      TabPicture(3)   =   "frm_Instituicoes2.frx":109A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "PBLista(5)"
      Tab(3).Control(1)=   "USImageList4"
      Tab(3).Control(2)=   "Frame6"
      Tab(3).Control(3)=   "Frame7"
      Tab(3).Control(4)=   "SSTab2"
      Tab(3).Control(5)=   "USToolBar4"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Cheques recebidos"
      TabPicture(4)   =   "frm_Instituicoes2.frx":10B6
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).Control(1)=   "Lista_cheque"
      Tab(4).Control(2)=   "USImageList5"
      Tab(4).Control(3)=   "USToolBar5"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Carteira de títulos"
      TabPicture(5)   =   "frm_Instituicoes2.frx":10D2
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "lst_Duplicata"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame21"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "FramePesquisa"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      Begin VB.Frame Frame14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Local para armazenamento do arquivo remessa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -69750
         TabIndex        =   233
         Top             =   5445
         Width           =   5235
         Begin VB.CommandButton cmdLocal 
            Caption         =   "..."
            Height          =   315
            Left            =   4500
            TabIndex        =   235
            TabStop         =   0   'False
            ToolTipText     =   "Abrirl local de armazenamento dos arquivos de remessa."
            Top             =   270
            Width           =   570
         End
         Begin VB.TextBox Txtlocal 
            Enabled         =   0   'False
            Height          =   315
            Left            =   165
            Locked          =   -1  'True
            TabIndex        =   234
            TabStop         =   0   'False
            Top             =   270
            Width           =   4320
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Arquivo de configuração da carteira"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69750
         TabIndex        =   230
         Top             =   4800
         Width           =   5235
         Begin VB.CommandButton cmdArquivo 
            Caption         =   "..."
            Height          =   315
            Left            =   4500
            TabIndex        =   232
            TabStop         =   0   'False
            ToolTipText     =   "Abrirl local de armazenamento do arquivo de configuração da carteira."
            Top             =   270
            Width           =   570
         End
         Begin VB.TextBox txtcarteiraconf 
            Enabled         =   0   'False
            Height          =   315
            Left            =   165
            Locked          =   -1  'True
            TabIndex        =   231
            TabStop         =   0   'False
            Top             =   270
            Width           =   4320
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Assunto email"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -69750
         TabIndex        =   228
         Top             =   4080
         Width           =   5235
         Begin VB.TextBox txtAssunto 
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
            Left            =   150
            TabIndex        =   229
            Text            =   "Boleto Sistema Caprind"
            ToolTipText     =   "Assunto para email á ser enviado."
            Top             =   270
            Width           =   4935
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Instruções á serem enviadas para o banco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   -74940
         TabIndex        =   216
         Top             =   4080
         Width           =   5190
         Begin VB.CommandButton cmdSalvarInstrucoes 
            Caption         =   "Salvar"
            Height          =   285
            Left            =   4020
            TabIndex        =   222
            Top             =   540
            Width           =   915
         End
         Begin VB.TextBox Txtpercentual_juros 
            Alignment       =   2  'Centralizar
            Height          =   285
            Left            =   240
            TabIndex        =   221
            Text            =   "0,20"
            ToolTipText     =   "Percentual dos juros a serem aplicados por dia de atraso."
            Top             =   540
            Width           =   885
         End
         Begin VB.TextBox Txtdias_protesto 
            Alignment       =   2  'Centralizar
            Height          =   285
            Left            =   2835
            TabIndex        =   220
            Text            =   "30"
            ToolTipText     =   "Numero de dias do prazo antes do título ser protestado."
            Top             =   540
            Width           =   1185
         End
         Begin VB.TextBox Txtinstrucoes 
            Height          =   780
            Left            =   225
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   219
            Text            =   "frm_Instituicoes2.frx":10EE
            ToolTipText     =   "Instruções para o banco."
            Top             =   1080
            Width           =   4755
         End
         Begin VB.TextBox Txtpercentual_desconto 
            Alignment       =   2  'Centralizar
            Height          =   285
            Left            =   1140
            TabIndex        =   218
            Text            =   "0,00"
            ToolTipText     =   "Percentual de desconto a ser aplicado por dia de antecipação."
            Top             =   540
            Width           =   915
         End
         Begin VB.TextBox Txtpercentual_multa 
            Alignment       =   2  'Centralizar
            Height          =   285
            Left            =   2070
            TabIndex        =   217
            Text            =   "10,00"
            ToolTipText     =   "Percentual da multa a ser aplicado sobre o valor total do boleto."
            Top             =   540
            Width           =   750
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "% Juros"
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
            Index           =   64
            Left            =   420
            TabIndex        =   227
            Top             =   330
            Width           =   600
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Dias protesto"
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
            Index           =   65
            Left            =   2940
            TabIndex        =   226
            Top             =   330
            Width           =   960
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "% Multa"
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
            Index           =   66
            Left            =   2160
            TabIndex        =   225
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "% Desconto"
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
            Index           =   67
            Left            =   1200
            TabIndex        =   224
            Top             =   330
            Width           =   885
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Instruções para o banco"
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
            Index           =   68
            Left            =   390
            TabIndex        =   223
            Top             =   900
            Width           =   1755
         End
      End
      Begin VB.Frame Frame20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Observação"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   -64500
         TabIndex        =   214
         Top             =   4080
         Width           =   4785
         Begin VB.TextBox txtobs 
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
            Height          =   1575
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   215
            Text            =   "frm_Instituicoes2.frx":110D
            ToolTipText     =   "Observação."
            Top             =   300
            Width           =   4515
         End
      End
      Begin VB.Frame FramePesquisa 
         Caption         =   "Filtrar carteira de títulos"
         Height          =   870
         Left            =   30
         TabIndex        =   209
         Top             =   390
         Width           =   15210
         Begin DrawSuite2014.USButton CmdAprocessar 
            Height          =   495
            Left            =   12420
            TabIndex        =   240
            Top             =   270
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   873
            BorderColor     =   4960354
            BorderColorDisabled=   13160660
            BorderColorDown =   4210752
            BorderColorOver =   49152
            Caption         =   "Á processar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
            ForeColorDown   =   16777215
            ForeColorOver   =   16777215
            GradientColor1  =   4960354
            GradientColor2  =   4960354
            GradientColor3  =   4960354
            GradientColor4  =   4960354
            GradientColorDisabled1=   14215660
            GradientColorDisabled2=   14215660
            GradientColorDisabled3=   14215660
            GradientColorDisabled4=   14215660
            GradientColorDown1=   32768
            GradientColorDown2=   32768
            GradientColorDown3=   32768
            GradientColorDown4=   32768
            GradientColorOver1=   49152
            GradientColorOver2=   49152
            GradientColorOver3=   49152
            GradientColorOver4=   49152
            Theme           =   3
         End
         Begin VB.ComboBox cmbCliente 
            Height          =   330
            Left            =   5280
            TabIndex        =   210
            ToolTipText     =   "Escolha um cliente para pesquisa."
            Top             =   360
            Width           =   6105
         End
         Begin MSComCtl2.DTPicker DTFim 
            Height          =   315
            Left            =   3090
            TabIndex        =   211
            ToolTipText     =   "Data de vencimento final para pesquisa."
            Top             =   360
            Width           =   1185
            _ExtentX        =   2090
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
            Format          =   133365761
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker DTINI 
            Height          =   315
            Left            =   1425
            TabIndex        =   212
            ToolTipText     =   "Data de vencimento de início para pesquisa."
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   133365763
            CurrentDate     =   39057
         End
         Begin DrawSuite2014.USButton cmdProcessados 
            Height          =   495
            Left            =   13740
            TabIndex        =   241
            Top             =   270
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   873
            BorderColor     =   5263559
            BorderColorDisabled=   13160660
            BorderColorDown =   4013465
            BorderColorOver =   4408288
            Caption         =   "Processados"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            Theme           =   4
         End
         Begin VB.Label Label6 
            Caption         =   "do cliente:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4440
            TabIndex        =   239
            Top             =   450
            Width           =   795
         End
         Begin VB.Label Label5 
            Caption         =   "á"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2790
            TabIndex        =   238
            Top             =   450
            Width           =   105
         End
         Begin VB.Label Label3 
            Caption         =   "Vencimento de : "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   237
            Top             =   420
            Width           =   1215
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Processar títulos (Ações)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   30
         TabIndex        =   204
         Top             =   9270
         Width           =   15180
         Begin VB.CheckBox chkRemessa 
            Caption         =   "Gerar arquivo remessa"
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
            Height          =   210
            Left            =   285
            TabIndex        =   208
            Top             =   360
            Width           =   2115
         End
         Begin VB.CheckBox chkEmailcopia 
            Caption         =   "Enviar-me cópia"
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
            Height          =   210
            Left            =   7875
            TabIndex        =   207
            ToolTipText     =   "Enviar uma cópia do boleto para meu email."
            Top             =   390
            Width           =   1485
         End
         Begin VB.CheckBox chkImprimir 
            Caption         =   "Visualizar boleto(s) para impressão"
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
            Height          =   210
            Left            =   2415
            TabIndex        =   206
            ToolTipText     =   "Visualizar boleto para impressão."
            Top             =   390
            Width           =   2985
         End
         Begin VB.CheckBox chkEmail 
            Caption         =   "Enviar boleto(s) por email"
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
            Height          =   210
            Left            =   5430
            TabIndex        =   205
            ToolTipText     =   "Enviar boleto por email para o cliente."
            Top             =   390
            Width           =   2535
         End
         Begin DrawSuite2014.USButton cmdProcessar 
            Height          =   495
            Left            =   12660
            TabIndex        =   242
            Top             =   180
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   873
            BorderColor     =   0
            BorderColorDisabled=   13160660
            BorderColorDown =   4210752
            BorderColorOver =   8421504
            Caption         =   "&Processar titulo(s)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
            ForeColorDown   =   16777215
            ForeColorOver   =   16777215
            GradientColor1  =   0
            GradientColor2  =   0
            GradientColor3  =   0
            GradientColor4  =   0
            GradientColorDisabled1=   13160660
            GradientColorDisabled2=   13160660
            GradientColorDisabled3=   13160660
            GradientColorDisabled4=   13160660
            GradientColorDown1=   4210752
            GradientColorDown2=   4210752
            GradientColorDown3=   4210752
            GradientColorDown4=   4210752
            GradientColorOver1=   8421504
            GradientColorOver2=   8421504
            GradientColorOver3=   8421504
            GradientColorOver4=   8421504
            Theme           =   6
         End
      End
      Begin DrawSuite2014.USProgressBar PBLista 
         Height          =   255
         Index           =   5
         Left            =   -74925
         TabIndex        =   189
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
         SearchText      =   "Atualizando..."
         Value           =   0
      End
      Begin DrawSuite2014.USImageList USImageList1 
         Left            =   -64650
         Top             =   540
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frm_Instituicoes2.frx":110F
         Count           =   1
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'Nenhum
         Caption         =   "Operação da lista"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -74925
         TabIndex        =   140
         Top             =   9525
         Width           =   15195
         Begin VB.ComboBox Cmb_opcao_lista_recebidos 
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
            ItemData        =   "frm_Instituicoes2.frx":8965
            Left            =   13080
            List            =   "frm_Instituicoes2.frx":8972
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   60
            Width           =   1965
         End
         Begin DrawSuite2014.USProgressBar PBLista 
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   190
            Top             =   90
            Width           =   11535
            _ExtentX        =   20346
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
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
            Index           =   57
            Left            =   11730
            TabIndex        =   167
            Top             =   113
            Width           =   1260
         End
      End
      Begin DrawSuite2014.USImageList USImageList4 
         Left            =   -61380
         Top             =   570
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frm_Instituicoes2.frx":89A0
         Count           =   1
      End
      Begin VB.TextBox txtCodBanco 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -71160
         Locked          =   -1  'True
         TabIndex        =   131
         TabStop         =   0   'False
         ToolTipText     =   "Código."
         Top             =   6540
         Visible         =   0   'False
         Width           =   1365
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
         Height          =   765
         Left            =   -74925
         TabIndex        =   85
         Top             =   8940
         Width           =   15195
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Operação da lista"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   755
            Left            =   12870
            TabIndex        =   139
            Top             =   10
            Width           =   2310
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
               ItemData        =   "frm_Instituicoes2.frx":F9A1
               Left            =   180
               List            =   "frm_Instituicoes2.frx":F9AE
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   79
               TabStop         =   0   'False
               Top             =   270
               Width           =   1965
            End
         End
         Begin VB.TextBox Txt_valor_total 
            Alignment       =   1  'Alinhar à Direita
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   10560
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   78
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor total."
            Top             =   310
            Width           =   1665
         End
         Begin VB.TextBox Txt_valor_cancelado 
            Alignment       =   1  'Alinhar à Direita
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   8460
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   77
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor de cheque(s) cancelado(s)."
            Top             =   310
            Width           =   1665
         End
         Begin VB.TextBox Txt_valor_ativo 
            Alignment       =   1  'Alinhar à Direita
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   6360
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   76
            TabStop         =   0   'False
            Text            =   "0,00"
            ToolTipText     =   "Valor de cheque(s) ativo(s)."
            Top             =   310
            Width           =   1665
         End
         Begin VB.TextBox Txt_qtde_total 
            Alignment       =   1  'Alinhar à Direita
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   4350
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   75
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "Quantidade total de cheques."
            Top             =   310
            Width           =   1665
         End
         Begin VB.TextBox Txt_qtde_cancelado 
            Alignment       =   1  'Alinhar à Direita
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2265
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   74
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "Quantidade de cheque(s) cancelado(s)."
            Top             =   310
            Width           =   1665
         End
         Begin VB.TextBox Txt_qtde_ativo 
            Alignment       =   1  'Alinhar à Direita
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   73
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "Quantidade de cheque(s) ativo(s)."
            Top             =   310
            Width           =   1665
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Alinhar à Direita
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Valor total"
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
            Index           =   56
            Left            =   10950
            TabIndex        =   95
            Top             =   120
            Width           =   885
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Alinhar à Direita
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "+"
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
            Left            =   10275
            TabIndex        =   94
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Alinhar à Direita
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Valor cancelado"
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
            Index           =   55
            Left            =   8625
            TabIndex        =   93
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Alinhar à Direita
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "+"
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
            Left            =   8175
            TabIndex        =   92
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Alinhar à Direita
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Valor ativo"
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
            Index           =   54
            Left            =   6735
            TabIndex        =   91
            Top             =   120
            Width           =   915
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Alinhar à Direita
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Qtde. total"
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
            Index           =   53
            Left            =   4732
            TabIndex        =   90
            Top             =   120
            Width           =   900
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Alinhar à Direita
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
            Left            =   4072
            TabIndex        =   89
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Alinhar à Direita
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Qtde. cancelado"
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
            Index           =   52
            Left            =   2422
            TabIndex        =   88
            Top             =   120
            Width           =   1350
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Alinhar à Direita
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "+"
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
            Left            =   1987
            TabIndex        =   87
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Alinhar à Direita
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Qtde. ativo"
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
            Index           =   51
            Left            =   547
            TabIndex        =   86
            Top             =   120
            Width           =   930
         End
      End
      Begin VB.Frame Frame7 
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
         Height          =   1815
         Left            =   -74925
         TabIndex        =   104
         Top             =   1320
         Width           =   15195
         Begin VB.TextBox Txt_favorecido 
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
            TabIndex        =   71
            ToolTipText     =   "Nome do favorecido."
            Top             =   390
            Width           =   14835
         End
         Begin VB.TextBox txtobscheque 
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
            Height          =   675
            Left            =   180
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   72
            ToolTipText     =   "Verso do cheque."
            Top             =   990
            Width           =   14835
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Favorecido"
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
            Index           =   50
            Left            =   7200
            TabIndex        =   106
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Verso do cheque"
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
            Index           =   50
            Left            =   6997
            TabIndex        =   105
            Top             =   780
            Width           =   1200
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   6885
         Left            =   -75000
         TabIndex        =   84
         Top             =   3150
         Width           =   15300
         _ExtentX        =   26988
         _ExtentY        =   12144
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
         TabCaption(0)   =   "Ativos"
         TabPicture(0)   =   "frm_Instituicoes2.frx":F9E5
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Lst_cheque"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Cancelados"
         TabPicture(1)   =   "frm_Instituicoes2.frx":FA01
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Lst_cheque1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin MSComctlLib.ListView Lst_cheque 
            Height          =   5430
            Left            =   -74925
            TabIndex        =   69
            Top             =   345
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   9578
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
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
               Alignment       =   2
               SubItemIndex    =   1
               Object.Tag             =   "D"
               Text            =   "Data"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Cheque"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Fornecedor"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Valor"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Observações"
               Object.Width           =   10063
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Object.Tag             =   "T"
               Text            =   "Compensado"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSComctlLib.ListView Lst_cheque1 
            Height          =   5430
            Left            =   75
            TabIndex        =   70
            Top             =   345
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   9578
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
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
               Alignment       =   2
               SubItemIndex    =   1
               Object.Tag             =   "D"
               Text            =   "Data"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Cheque"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Fornecedor"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Valor"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Observações"
               Object.Width           =   7488
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Object.Tag             =   "D"
               Text            =   "Dt. cancelamento"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Tag             =   "T"
               Text            =   "Responsável"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Tag             =   "T"
               Text            =   "Motivo"
               Object.Width           =   7011
            EndProperty
         End
      End
      Begin VB.Frame Frame41 
         BackColor       =   &H00E0E0E0&
         Height          =   645
         Left            =   -70785
         TabIndex        =   129
         Top             =   1310
         Width           =   11055
         Begin VB.TextBox TxtHistoricoExtrato 
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
            Height          =   325
            Left            =   2145
            Locked          =   -1  'True
            TabIndex        =   67
            TabStop         =   0   'False
            ToolTipText     =   "Histórico do lançamento."
            Top             =   195
            Width           =   8715
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Histórico do lançamento:*"
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
            Index           =   49
            Left            =   195
            TabIndex        =   130
            Top             =   210
            Width           =   1860
         End
      End
      Begin VB.Frame FrameFiltro 
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
         Height          =   645
         Left            =   -74925
         TabIndex        =   111
         Top             =   1310
         Width           =   4125
         Begin MSComCtl2.DTPicker msk_fltInicio 
            Height          =   315
            Left            =   1110
            TabIndex        =   65
            ToolTipText     =   "Data início para pesquisa."
            Top             =   210
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Format          =   133169155
            CurrentDate     =   39473
         End
         Begin MSComCtl2.DTPicker msk_fltFim 
            Height          =   315
            Left            =   2670
            TabIndex        =   66
            ToolTipText     =   "Data final para pesquisa."
            Top             =   210
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Format          =   133169153
            CurrentDate     =   39473
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparente
            Caption         =   "Período de :"
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
            Height          =   255
            Index           =   47
            Left            =   180
            TabIndex        =   113
            Top             =   210
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Centralizar
            BackStyle       =   0  'Transparente
            Caption         =   "à"
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
            Height          =   255
            Index           =   48
            Left            =   2400
            TabIndex        =   112
            Top             =   300
            Width           =   255
         End
      End
      Begin MSComctlLib.ListView Lista_cheque 
         Height          =   8235
         Left            =   -74925
         TabIndex        =   80
         Top             =   1320
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   14526
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Cheque"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Cliente"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Observações"
            Object.Width           =   10063
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Compensado"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView Lst_extrato 
         Height          =   7745
         Left            =   -74925
         TabIndex        =   68
         Top             =   1965
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   13653
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "Id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Histórico do lançamento"
            Object.Width           =   19764
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
            Text            =   "Saldo"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados da Instituição financeira"
         Enabled         =   0   'False
         Height          =   2790
         Left            =   -74940
         TabIndex        =   83
         Top             =   1290
         Width           =   15225
         Begin VB.TextBox Txt_IDBanco 
            Alignment       =   2  'Centralizar
            Height          =   315
            Left            =   2640
            TabIndex        =   236
            Top             =   1020
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.PictureBox Logo_Banco 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'Nenhum
            ForeColor       =   &H80000008&
            Height          =   750
            Left            =   300
            Picture         =   "frm_Instituicoes2.frx":FA1D
            ScaleHeight     =   750
            ScaleWidth      =   1500
            TabIndex        =   198
            Top             =   735
            Width           =   1500
         End
         Begin VB.TextBox txtNomecedente 
            Alignment       =   2  'Centralizar
            Enabled         =   0   'False
            Height          =   315
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   201
            TabStop         =   0   'False
            Top             =   2370
            Width           =   6150
         End
         Begin VB.ComboBox cmbCarteira 
            Height          =   330
            Left            =   7920
            TabIndex        =   200
            Top             =   2370
            Width           =   7245
         End
         Begin VB.TextBox txtStatus 
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   13665
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Status."
            Top             =   390
            Width           =   1455
         End
         Begin VB.TextBox txtDtValidacao 
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9600
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Data e hora da validação."
            Top             =   390
            Width           =   1590
         End
         Begin VB.TextBox txtRespValidacao 
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   11205
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pela validação."
            Top             =   390
            Width           =   2445
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
            Left            =   11670
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   20
            ToolTipText     =   "Centro de custo."
            Top             =   1665
            Width           =   3480
         End
         Begin VB.TextBox Txt_nome_agencia 
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   10620
            MaxLength       =   255
            TabIndex        =   12
            ToolTipText     =   "Nome da agência."
            Top             =   1025
            Width           =   4500
         End
         Begin VB.TextBox Txt_codigo_cedente1 
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
            Left            =   210
            MaxLength       =   50
            TabIndex        =   11
            ToolTipText     =   "Código do cedente/convênio registrado."
            Top             =   2370
            Width           =   1550
         End
         Begin VB.TextBox Txt_codigo_cedente 
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
            Left            =   8160
            MaxLength       =   50
            TabIndex        =   10
            ToolTipText     =   "Código do cedente/convênio."
            Top             =   5220
            Width           =   1565
         End
         Begin VB.TextBox txtResponsavel 
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   6930
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Responsável pelo cadastro."
            Top             =   390
            Width           =   2655
         End
         Begin VB.TextBox txtData1 
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   6015
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Data do cadastro."
            Top             =   390
            Width           =   900
         End
         Begin VB.TextBox txtAgencia 
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
            Left            =   8325
            MaxLength       =   50
            TabIndex        =   7
            ToolTipText     =   "Número da agência."
            Top             =   1025
            Width           =   930
         End
         Begin VB.TextBox txtNBanco 
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
            Left            =   2640
            TabIndex        =   6
            ToolTipText     =   "Número do banco."
            Top             =   1025
            Width           =   930
         End
         Begin VB.ComboBox cmbFamilia 
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
            ItemData        =   "frm_Instituicoes2.frx":134F7
            Left            =   6810
            List            =   "frm_Instituicoes2.frx":134F9
            Sorted          =   -1  'True
            TabIndex        =   16
            ToolTipText     =   "Família."
            Top             =   1660
            Width           =   1515
         End
         Begin VB.TextBox txtgerente 
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2610
            MaxLength       =   50
            TabIndex        =   13
            ToolTipText     =   "Nome do gerente."
            Top             =   1660
            Width           =   2865
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
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3585
            MaxLength       =   50
            TabIndex        =   9
            ToolTipText     =   "Descrição do banco."
            Top             =   1025
            Width           =   4715
         End
         Begin VB.TextBox txtsaldo 
            Alignment       =   2  'Centralizar
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   8340
            TabIndex        =   17
            ToolTipText     =   "Saldo bancário atual."
            Top             =   1660
            Width           =   1155
         End
         Begin VB.TextBox txtConta 
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
            Left            =   9270
            MaxLength       =   20
            TabIndex        =   8
            ToolTipText     =   "Número da conta."
            Top             =   1025
            Width           =   1350
         End
         Begin VB.TextBox txtFone 
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
            Left            =   5490
            MaxLength       =   30
            TabIndex        =   14
            ToolTipText     =   "Número do telefone."
            Top             =   1660
            Width           =   1305
         End
         Begin VB.TextBox txtFAX 
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
            Left            =   6810
            MaxLength       =   30
            TabIndex        =   15
            ToolTipText     =   "Número do fax."
            Top             =   5235
            Width           =   1155
         End
         Begin VB.TextBox txtLimite 
            Alignment       =   2  'Centralizar
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   9510
            TabIndex        =   18
            ToolTipText     =   "Limite para desconto."
            Top             =   1660
            Width           =   1125
         End
         Begin VB.TextBox txtUtilizado 
            Alignment       =   2  'Centralizar
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   10650
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Limite utilizado."
            Top             =   1660
            Width           =   1005
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   1605
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   199
            TabStop         =   0   'False
            Top             =   390
            Width           =   2400
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Carteira"
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
            Index           =   63
            Left            =   11205
            TabIndex        =   203
            Top             =   2160
            Width           =   585
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Nome cedente"
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
            Index           =   6
            Left            =   4328
            TabIndex        =   202
            Top             =   2160
            Width           =   1035
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Limite desc*"
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
            Index           =   22
            Left            =   9585
            TabIndex        =   183
            Top             =   1470
            Width           =   1065
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Limite util."
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
            Index           =   21
            Left            =   10710
            TabIndex        =   182
            Top             =   1470
            Width           =   885
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Fone"
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
            Index           =   20
            Left            =   5955
            TabIndex        =   181
            Top             =   1470
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Fax"
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
            Index           =   19
            Left            =   7245
            TabIndex        =   180
            Top             =   5040
            Width           =   270
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Família"
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
            Index           =   18
            Left            =   7320
            TabIndex        =   179
            Top             =   1470
            Width           =   480
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Centro de custo"
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
            Index           =   17
            Left            =   12788
            TabIndex        =   178
            Top             =   1470
            Width           =   1155
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Agência*"
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
            Index           =   16
            Left            =   8460
            TabIndex        =   177
            Top             =   810
            Width           =   660
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Conta*"
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
            Index           =   15
            Left            =   9675
            TabIndex        =   176
            Top             =   810
            Width           =   525
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Descrição*"
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
            Index           =   14
            Left            =   5550
            TabIndex        =   175
            Top             =   810
            Width           =   780
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Cód. cedente"
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
            Index           =   13
            Left            =   8460
            TabIndex        =   174
            Top             =   5010
            Width           =   975
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Código cedente"
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
            Index           =   12
            Left            =   423
            TabIndex        =   173
            Top             =   2160
            Width           =   1125
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Nome da agência"
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
            Index           =   11
            Left            =   12225
            TabIndex        =   172
            Top             =   810
            Width           =   1230
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
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
            Index           =   10
            Left            =   7800
            TabIndex        =   171
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Data/hora validação"
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
            Left            =   9750
            TabIndex        =   170
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
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
            Index           =   7
            Left            =   11475
            TabIndex        =   169
            Top             =   180
            Width           =   1980
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Status"
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
            Left            =   14115
            TabIndex        =   168
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
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
            Index           =   8
            Left            =   6270
            TabIndex        =   158
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
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
            Index           =   7
            Left            =   3990
            TabIndex        =   157
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Saldo atual*"
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
            Index           =   1
            Left            =   8475
            TabIndex        =   156
            Top             =   1470
            Width           =   1050
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Nº banco*"
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
            Left            =   2730
            TabIndex        =   154
            Top             =   810
            Width           =   750
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Gerente"
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
            Left            =   3810
            TabIndex        =   155
            Top             =   1470
            Width           =   585
         End
      End
      Begin DrawSuite2014.USImageList USImageList2 
         Left            =   -67650
         Top             =   600
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frm_Instituicoes2.frx":134FB
         Count           =   1
      End
      Begin DrawSuite2014.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   135
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   13
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   33
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Filtrar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Filtrar (F2)"
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
         ButtonLeft2     =   37
         ButtonTop2      =   2
         ButtonWidth2    =   36
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Salvar"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Salvar (F3)"
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
         ButtonLeft3     =   75
         ButtonTop3      =   2
         ButtonWidth3    =   38
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Excluir"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Excluir (F4)"
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
         ButtonLeft4     =   115
         ButtonTop4      =   2
         ButtonWidth4    =   39
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Relatório"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Relatório (F5)"
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
         ButtonLeft5     =   156
         ButtonTop5      =   2
         ButtonWidth5    =   51
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Anterior"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Registro anterior."
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
         ButtonLeft6     =   209
         ButtonTop6      =   2
         ButtonWidth6    =   47
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Próximo"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Próximo registro."
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
         ButtonLeft7     =   258
         ButtonTop7      =   2
         ButtonWidth7    =   46
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Copiar"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Copiar (F7)"
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
         ButtonLeft8     =   306
         ButtonTop8      =   2
         ButtonWidth8    =   39
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Atualizar"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Utilizado pelo administrador do sistema."
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
         ButtonLeft9     =   347
         ButtonTop9      =   2
         ButtonWidth9    =   50
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonAlignment10=   2
         ButtonType10    =   1
         ButtonStyle10   =   -1
         BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState10   =   -1
         ButtonLeft10    =   399
         ButtonTop10     =   4
         ButtonWidth10   =   2
         ButtonHeight10  =   54
         ButtonCaption11 =   "Ajuda"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Ajuda (F1)"
         ButtonKey11     =   "11"
         ButtonAlignment11=   2
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft11    =   403
         ButtonTop11     =   2
         ButtonWidth11   =   36
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonCaption12 =   "Sair"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Sair (Esc)"
         ButtonKey12     =   "12"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   441
         ButtonTop12     =   2
         ButtonWidth12   =   26
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonKey13     =   "13"
         ButtonAlignment13=   2
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState13   =   5
         ButtonLeft13    =   469
         ButtonTop13     =   2
         ButtonWidth13   =   24
         ButtonHeight13  =   24
         ButtonUseMaskColor13=   0   'False
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   8745
         Left            =   -75000
         TabIndex        =   96
         Top             =   1320
         Width           =   15600
         _ExtentX        =   27517
         _ExtentY        =   15425
         _Version        =   393216
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
         TabCaption(0)   =   "Depósito/Transferência"
         TabPicture(0)   =   "frm_Instituicoes2.frx":1AD43
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1(62)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "PBLista(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lst_transferencias"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "frm_filtro"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtid"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Txt_vlr_total_deptran"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Saque"
         TabPicture(1)   =   "frm_Instituicoes2.frx":1AD5F
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "LblValortotal"
         Tab(1).Control(1)=   "Frame10"
         Tab(1).Control(2)=   "Frame9"
         Tab(1).Control(3)=   "PBLista(2)"
         Tab(1).Control(4)=   "Lst_Contas"
         Tab(1).Control(5)=   "Lst_saque"
         Tab(1).Control(6)=   "Frame8"
         Tab(1).Control(7)=   "Frame11"
         Tab(1).Control(8)=   "Txt_id_saque"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "Tarifas"
         TabPicture(2)   =   "frm_Instituicoes2.frx":1AD7B
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame12"
         Tab(2).Control(1)=   "Txt_id_tarifa"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Frame4"
         Tab(2).Control(3)=   "Lst_tarifa"
         Tab(2).ControlCount=   4
         Begin VB.TextBox Txt_vlr_total_deptran 
            Alignment       =   1  'Alinhar à Direita
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   13710
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "Valor total pago."
            Top             =   8370
            Width           =   1560
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'Nenhum
            Caption         =   "Operação da lista"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -74925
            TabIndex        =   161
            Top             =   8280
            Width           =   15195
            Begin VB.TextBox Txt_valor_total_tarifas 
               Alignment       =   1  'Alinhar à Direita
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   10500
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   163
               TabStop         =   0   'False
               ToolTipText     =   "Valor total pago."
               Top             =   60
               Width           =   1560
            End
            Begin VB.TextBox Txt_valor_total_tarifas1 
               Alignment       =   1  'Alinhar à Direita
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   13470
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   162
               TabStop         =   0   'False
               ToolTipText     =   "Valor total recebido."
               Top             =   60
               Width           =   1560
            End
            Begin DrawSuite2014.USProgressBar PBLista 
               Height          =   255
               Index           =   3
               Left            =   0
               TabIndex        =   166
               Top             =   60
               Width           =   9045
               _ExtentX        =   15954
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
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparente
               Caption         =   "Vlr. total pago :"
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
               Index           =   45
               Left            =   9150
               TabIndex        =   165
               Top             =   60
               Width           =   2175
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparente
               Caption         =   "Vlr. total rec. :"
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
               Index           =   46
               Left            =   12210
               TabIndex        =   164
               Top             =   60
               Width           =   2175
               WordWrap        =   -1  'True
            End
         End
         Begin VB.TextBox Txt_id_tarifa 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   -73410
            Locked          =   -1  'True
            TabIndex        =   150
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   2250
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Txt_id_saque 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   -74010
            Locked          =   -1  'True
            TabIndex        =   149
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   2250
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtid 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   148
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Frame Frame4 
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
            Height          =   1410
            Left            =   -74925
            TabIndex        =   141
            Top             =   330
            Width           =   15195
            Begin VB.ComboBox Cmb_operacao 
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
               ItemData        =   "frm_Instituicoes2.frx":1AD97
               Left            =   5100
               List            =   "frm_Instituicoes2.frx":1ADA1
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   52
               ToolTipText     =   "Operação."
               Top             =   360
               Width           =   1065
            End
            Begin VB.CommandButton Cmd_forma 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   11070
               Picture         =   "frm_Instituicoes2.frx":1ADB6
               Style           =   1  'Graphical
               TabIndex        =   56
               ToolTipText     =   "Localizar forma da baixa."
               Top             =   360
               Width           =   315
            End
            Begin VB.CommandButton Cmd_localizar_tipo_dcto 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   7230
               Picture         =   "frm_Instituicoes2.frx":1AEB8
               Style           =   1  'Graphical
               TabIndex        =   54
               ToolTipText     =   "Localizar tipo do documento."
               Top             =   360
               Width           =   315
            End
            Begin VB.ComboBox Cmb_tipo 
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
               ItemData        =   "frm_Instituicoes2.frx":1AFBA
               Left            =   6180
               List            =   "frm_Instituicoes2.frx":1AFBC
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   53
               ToolTipText     =   "Tipo do documento."
               Top             =   360
               Width           =   1065
            End
            Begin VB.ComboBox cmb_forma1 
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
               ItemData        =   "frm_Instituicoes2.frx":1AFBE
               Left            =   7620
               List            =   "frm_Instituicoes2.frx":1AFC0
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   55
               ToolTipText     =   "Forma da baixa."
               Top             =   360
               Width           =   3465
            End
            Begin VB.CommandButton Cmd_PC 
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   13095
               Picture         =   "frm_Instituicoes2.frx":1AFC2
               Style           =   1  'Graphical
               TabIndex        =   62
               ToolTipText     =   "Abrir formulário para cadastro de plano de contas."
               Top             =   945
               Width           =   315
            End
            Begin VB.CommandButton Cmd_localizar_PC 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   12780
               Picture         =   "frm_Instituicoes2.frx":1B0A4
               Style           =   1  'Graphical
               TabIndex        =   61
               ToolTipText     =   "Localizar plano de contas."
               Top             =   945
               Width           =   315
            End
            Begin VB.TextBox Txt_ID_PC 
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
               Left            =   135
               MaxLength       =   255
               TabIndex        =   58
               Text            =   "0"
               ToolTipText     =   "ID PC."
               Top             =   945
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.TextBox Txt_descricao_PC 
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
               Left            =   2030
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   60
               TabStop         =   0   'False
               ToolTipText     =   "Descrição."
               Top             =   945
               Width           =   10755
            End
            Begin VB.TextBox Txt_valor1 
               Alignment       =   1  'Alinhar à Direita
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
               Left            =   13485
               TabIndex        =   63
               ToolTipText     =   "Valor."
               Top             =   945
               Width           =   1515
            End
            Begin VB.TextBox txtResponsavel3 
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
               Left            =   1360
               Locked          =   -1  'True
               TabIndex        =   51
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   360
               Width           =   3720
            End
            Begin VB.TextBox txtObsFluxo2 
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
               Left            =   11460
               MaxLength       =   255
               TabIndex        =   57
               Text            =   "Tarifa"
               ToolTipText     =   "Histórico do lançamento."
               Top             =   360
               Width           =   3540
            End
            Begin MSComCtl2.DTPicker txtdata3 
               Height          =   315
               Left            =   135
               TabIndex        =   50
               ToolTipText     =   "Data da movimentação."
               Top             =   360
               Width           =   1215
               _ExtentX        =   2143
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
               Format          =   224591873
               CurrentDate     =   39057
            End
            Begin VB.TextBox Txt_codigo_PC 
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
               Left            =   135
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   59
               TabStop         =   0   'False
               ToolTipText     =   "Código."
               Top             =   945
               Width           =   1875
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparente
               Caption         =   "Operação*"
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
               Index           =   38
               Left            =   5235
               TabIndex        =   153
               Top             =   165
               Width           =   795
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparente
               Caption         =   "Tipo docto.*"
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
               Index           =   39
               Left            =   6262
               TabIndex        =   152
               Top             =   165
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H80000009&
               BackStyle       =   0  'Transparente
               Caption         =   "Forma da baixa*"
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
               Index           =   40
               Left            =   8752
               TabIndex        =   151
               Top             =   150
               Width           =   1200
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparente
               Caption         =   "Código*"
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
               Index           =   42
               Left            =   780
               TabIndex        =   147
               Top             =   750
               Width           =   585
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparente
               Caption         =   "Descrição*"
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
               Index           =   43
               Left            =   7017
               TabIndex        =   146
               Top             =   750
               Width           =   780
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparente
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
               Index           =   36
               Left            =   570
               TabIndex        =   145
               Top             =   165
               Width           =   345
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparente
               Caption         =   "Valor*"
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
               Index           =   44
               Left            =   14017
               TabIndex        =   144
               Top             =   750
               Width           =   450
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparente
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
               Index           =   37
               Left            =   2763
               TabIndex        =   143
               Top             =   165
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparente
               Caption         =   "Histórico do lançamento"
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
               Index           =   41
               Left            =   12375
               TabIndex        =   142
               Top             =   165
               Width           =   1710
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Saldos do(s) saque(s)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   -74925
            TabIndex        =   119
            Top             =   7545
            Width           =   5025
            Begin VB.Frame Frame17 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Total"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   90
               TabIndex        =   124
               Top             =   210
               Width           =   1600
               Begin VB.TextBox TxtDisponivel 
                  Alignment       =   2  'Centralizar
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   0  'Nenhum
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   225
                  Left            =   60
                  Locked          =   -1  'True
                  MaxLength       =   14
                  TabIndex        =   125
                  TabStop         =   0   'False
                  Text            =   "0,00"
                  Top             =   240
                  Width           =   1485
               End
            End
            Begin VB.Frame Frame16 
               BackColor       =   &H00E0E0E0&
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
               Height          =   555
               Left            =   3300
               TabIndex        =   122
               Top             =   210
               Width           =   1600
               Begin VB.TextBox TxtSaldoSaque 
                  Alignment       =   2  'Centralizar
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   0  'Nenhum
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   225
                  Left            =   30
                  Locked          =   -1  'True
                  MaxLength       =   14
                  TabIndex        =   123
                  TabStop         =   0   'False
                  Text            =   "0,00"
                  Top             =   240
                  Width           =   1515
               End
            End
            Begin VB.Frame Frame15 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Utilizado"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   1695
               TabIndex        =   120
               Top             =   210
               Width           =   1600
               Begin VB.TextBox TxtValorSaqueUtilizado 
                  Alignment       =   2  'Centralizar
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   0  'Nenhum
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   225
                  Left            =   60
                  Locked          =   -1  'True
                  MaxLength       =   14
                  TabIndex        =   121
                  TabStop         =   0   'False
                  Text            =   "0,00"
                  Top             =   240
                  Width           =   1485
               End
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparente
               Caption         =   "Disponível"
               Height          =   195
               Left            =   420
               TabIndex        =   126
               Top             =   300
               Width           =   720
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
            Height          =   810
            Left            =   -74925
            TabIndex        =   107
            Top             =   330
            Width           =   15195
            Begin VB.TextBox txtObsFluxo1 
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
               Left            =   7020
               MaxLength       =   255
               TabIndex        =   46
               Text            =   "Saque"
               ToolTipText     =   "Histórico do lançamento."
               Top             =   360
               Width           =   6455
            End
            Begin VB.TextBox txtResponsavel2 
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
               Left            =   1360
               Locked          =   -1  'True
               TabIndex        =   45
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   360
               Width           =   5640
            End
            Begin VB.TextBox Txt_valor 
               Alignment       =   1  'Alinhar à Direita
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
               Left            =   13485
               TabIndex        =   47
               ToolTipText     =   "Valor."
               Top             =   360
               Width           =   1515
            End
            Begin MSComCtl2.DTPicker txtdata2 
               Height          =   315
               Left            =   135
               TabIndex        =   44
               ToolTipText     =   "Data da movimentação."
               Top             =   360
               Width           =   1215
               _ExtentX        =   2143
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
               Format          =   224591873
               CurrentDate     =   39057
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparente
               Caption         =   "Histórico do lançamento"
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
               Left            =   9392
               TabIndex        =   133
               Top             =   165
               Width           =   1710
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparente
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
               Index           =   32
               Left            =   3723
               TabIndex        =   110
               Top             =   165
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparente
               Caption         =   "Valor*"
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
               Index           =   33
               Left            =   14017
               TabIndex        =   109
               Top             =   165
               Width           =   450
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparente
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
               Index           =   31
               Left            =   570
               TabIndex        =   108
               Top             =   165
               Width           =   345
            End
         End
         Begin VB.Frame frm_filtro 
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
            Height          =   2580
            Left            =   75
            TabIndex        =   97
            Top             =   330
            Width           =   15195
            Begin VB.TextBox Txt_ID_PC_instituicao_rec 
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
               MaxLength       =   255
               TabIndex        =   196
               Text            =   "0"
               ToolTipText     =   "ID PC."
               Top             =   2115
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.TextBox Txt_descricao_PC_instituicao_rec 
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
               Left            =   2070
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   39
               TabStop         =   0   'False
               ToolTipText     =   "Descrição."
               Top             =   2115
               Width           =   12285
            End
            Begin VB.CommandButton Cmd_localizar_PC_instituicao_rec 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   14355
               Picture         =   "frm_Instituicoes2.frx":1B1A6
               Style           =   1  'Graphical
               TabIndex        =   40
               ToolTipText     =   "Localizar plano de contas."
               Top             =   2115
               Width           =   315
            End
            Begin VB.CommandButton Cmd_PC_instituicao_rec 
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   14695
               Picture         =   "frm_Instituicoes2.frx":1B2A8
               Style           =   1  'Graphical
               TabIndex        =   41
               ToolTipText     =   "Abrir formulário para cadastro de plano de contas."
               Top             =   2115
               Width           =   315
            End
            Begin VB.TextBox Txt_descricao_PC_instituicao 
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
               Left            =   2070
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   35
               TabStop         =   0   'False
               ToolTipText     =   "Descrição."
               Top             =   1545
               Width           =   12285
            End
            Begin VB.TextBox Txt_ID_PC_instituicao 
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
               MaxLength       =   255
               TabIndex        =   191
               Text            =   "0"
               ToolTipText     =   "ID PC."
               Top             =   1545
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.CommandButton Cmd_localizar_PC_instituicao 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   14355
               Picture         =   "frm_Instituicoes2.frx":1B38A
               Style           =   1  'Graphical
               TabIndex        =   36
               ToolTipText     =   "Localizar plano de contas."
               Top             =   1545
               Width           =   315
            End
            Begin VB.CommandButton Cmd_PC_instituicao 
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   14695
               Picture         =   "frm_Instituicoes2.frx":1B48C
               Style           =   1  'Graphical
               TabIndex        =   37
               ToolTipText     =   "Abrir formulário para cadastro de plano de contas."
               Top             =   1545
               Width           =   315
            End
            Begin VB.TextBox txtObsFluxo 
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
               Height          =   325
               Left            =   6420
               MaxLength       =   255
               TabIndex        =   32
               ToolTipText     =   "Histórico do lançamento."
               Top             =   945
               Width           =   4920
            End
            Begin VB.TextBox TxtHistDepTranf 
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
               Height          =   325
               Left            =   1525
               Locked          =   -1  'True
               TabIndex        =   31
               TabStop         =   0   'False
               ToolTipText     =   "Histórico padrão do lançamento."
               Top             =   945
               Width           =   4880
            End
            Begin VB.TextBox txtfavorecido 
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
               Height          =   325
               Left            =   6960
               TabIndex        =   28
               ToolTipText     =   "Nome do favorecido."
               Top             =   360
               Width           =   4425
            End
            Begin VB.TextBox txtCheque 
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
               Height          =   325
               Left            =   5430
               MaxLength       =   20
               TabIndex        =   27
               ToolTipText     =   "Número do documento."
               Top             =   360
               Width           =   1515
            End
            Begin VB.ComboBox cmb_forma 
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
               ItemData        =   "frm_Instituicoes2.frx":1B56E
               Left            =   4035
               List            =   "frm_Instituicoes2.frx":1B570
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   26
               ToolTipText     =   "Forma da movimentação."
               Top             =   360
               Width           =   1375
            End
            Begin VB.TextBox mskvalor 
               Alignment       =   1  'Alinhar à Direita
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
               Height          =   325
               Left            =   180
               MaxLength       =   15
               TabIndex        =   30
               ToolTipText     =   "Valor."
               Top             =   945
               Width           =   1335
            End
            Begin VB.ComboBox cmbrecebedor 
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
               Left            =   11395
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   29
               ToolTipText     =   "Instituição bancária recebedora."
               Top             =   360
               Width           =   3615
            End
            Begin VB.TextBox txtResponsavel1 
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
               Height          =   325
               Left            =   11350
               Locked          =   -1  'True
               TabIndex        =   33
               TabStop         =   0   'False
               ToolTipText     =   "Responsável pelo cadastro."
               Top             =   945
               Width           =   3660
            End
            Begin MSComCtl2.DTPicker txtdata 
               Height          =   330
               Left            =   2820
               TabIndex        =   25
               ToolTipText     =   "Data da movimentação."
               Top             =   360
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   582
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
               Format          =   224657409
               CurrentDate     =   39057
            End
            Begin VB.Frame Frame33 
               BackColor       =   &H00E0E0E0&
               Height          =   430
               Left            =   180
               TabIndex        =   127
               Top             =   240
               Width           =   2595
               Begin VB.OptionButton OptTransferencia 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Transferência*"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Left            =   1140
                  TabIndex        =   24
                  Top             =   180
                  Width           =   1395
               End
               Begin VB.OptionButton OptDeposito 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Depósito*"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Left            =   120
                  TabIndex        =   23
                  Top             =   180
                  Width           =   1005
               End
            End
            Begin VB.TextBox Txt_codigo_PC_instituicao 
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
               MaxLength       =   255
               TabIndex        =   34
               TabStop         =   0   'False
               ToolTipText     =   "Código."
               Top             =   1545
               Width           =   1875
            End
            Begin VB.TextBox Txt_codigo_PC_instituicao_rec 
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
               MaxLength       =   255
               TabIndex        =   38
               TabStop         =   0   'False
               ToolTipText     =   "Código."
               Top             =   2115
               Width           =   1875
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparente
               Caption         =   "Descrição da conta contábil da instituição recebedora*"
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
               Index           =   61
               Left            =   6210
               TabIndex        =   195
               Top             =   1920
               Width           =   7905
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparente
               Caption         =   "Código*"
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
               Index           =   60
               Left            =   825
               TabIndex        =   194
               Top             =   1920
               Width           =   585
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparente
               Caption         =   "Descrição da conta contábil da instituição*"
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
               Index           =   59
               Left            =   6615
               TabIndex        =   193
               Top             =   1350
               Width           =   6705
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparente
               Caption         =   "Código*"
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
               Index           =   58
               Left            =   825
               TabIndex        =   192
               Top             =   1350
               Width           =   585
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparente
               Caption         =   "Forma da movim.*"
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
               Index           =   30
               Left            =   4062
               TabIndex        =   186
               Top             =   165
               Width           =   1320
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparente
               Caption         =   "Histórico do lançamento"
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
               Index           =   28
               Left            =   8025
               TabIndex        =   132
               Top             =   750
               Width           =   1710
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparente
               Caption         =   "Histórico padrão do lançamento"
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
               Index           =   27
               Left            =   2833
               TabIndex        =   128
               Top             =   750
               Width           =   2265
            End
            Begin VB.Label LblDocumento 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparente
               Caption         =   "N° do documento"
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
               Left            =   5565
               TabIndex        =   103
               Top             =   165
               Width           =   1245
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H80000009&
               BackStyle       =   0  'Transparente
               Caption         =   "Favorecido"
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
               Index           =   24
               Left            =   8865
               TabIndex        =   102
               Top             =   165
               Width           =   810
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparente
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
               Index           =   23
               Left            =   3255
               TabIndex        =   101
               Top             =   165
               Width           =   345
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparente
               Caption         =   "Valor *"
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
               Index           =   26
               Left            =   660
               TabIndex        =   100
               Top             =   750
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparente
               Caption         =   "Instituição bancária recebedora*"
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
               Index           =   25
               Left            =   12017
               TabIndex        =   99
               Top             =   165
               Width           =   2370
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparente
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
               Index           =   29
               Left            =   12723
               TabIndex        =   98
               Top             =   750
               Width           =   915
            End
         End
         Begin MSComctlLib.ListView lst_transferencias 
            Height          =   5370
            Left            =   75
            TabIndex        =   42
            Top             =   2925
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   9472
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
               Alignment       =   2
               SubItemIndex    =   1
               Object.Tag             =   "D"
               Text            =   "Data"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Responsável"
               Object.Width           =   4057
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Tipo"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Banco remetente"
               Object.Width           =   7588
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Banco recebedor"
               Object.Width           =   7588
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Valor"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Tag             =   "N"
               Text            =   "id_banco_rem"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Object.Tag             =   "N"
               Text            =   "id_banco_rec"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView Lst_saque 
            Height          =   5970
            Left            =   -74925
            TabIndex        =   48
            Top             =   1560
            Width           =   5025
            _ExtentX        =   8864
            _ExtentY        =   10530
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Tag             =   "N"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Object.Tag             =   "D"
               Text            =   "Data"
               Object.Width           =   1588
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Object.Tag             =   "N"
               Text            =   "Valor"
               Object.Width           =   1941
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Object.Tag             =   "N"
               Text            =   "Utilizado"
               Object.Width           =   2058
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Object.Tag             =   "N"
               Text            =   "Saldo"
               Object.Width           =   2058
            EndProperty
         End
         Begin MSComctlLib.ListView Lst_Contas 
            Height          =   6840
            Left            =   -69885
            TabIndex        =   49
            Top             =   1560
            Width           =   10140
            _ExtentX        =   17886
            _ExtentY        =   12065
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
               Text            =   "Id"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Object.Tag             =   "D"
               Text            =   "Data"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Fornecedor"
               Object.Width           =   8643
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
               Text            =   "Pago"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Object.Tag             =   "N"
               Text            =   "Vlr. antecipação"
               Object.Width           =   2381
            EndProperty
         End
         Begin MSComctlLib.ListView Lst_tarifa 
            Height          =   6510
            Left            =   -74925
            TabIndex        =   64
            Top             =   1755
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   11483
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
               Alignment       =   2
               SubItemIndex    =   1
               Object.Tag             =   "D"
               Text            =   "Data"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Tag             =   "T"
               Text            =   "Responsável"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Tag             =   "T"
               Text            =   "Operação"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Tag             =   "T"
               Text            =   "Código"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Tag             =   "T"
               Text            =   "Descrição"
               Object.Width           =   11827
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Object.Tag             =   "N"
               Text            =   "Valor"
               Object.Width           =   2117
            EndProperty
         End
         Begin DrawSuite2014.USProgressBar PBLista 
            Height          =   255
            Index           =   1
            Left            =   75
            TabIndex        =   185
            Top             =   8355
            Width           =   12615
            _ExtentX        =   22251
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
         Begin DrawSuite2014.USProgressBar PBLista 
            Height          =   255
            Index           =   2
            Left            =   -74925
            TabIndex        =   187
            Top             =   8415
            Width           =   12855
            _ExtentX        =   22675
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
         Begin VB.Frame Frame9 
            BackColor       =   &H00E0E0E0&
            Height          =   405
            Left            =   -74925
            TabIndex        =   114
            Top             =   1145
            Width           =   5025
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparente
               Caption         =   "Lista de saque(s) efetuado(s)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   34
               Left            =   1230
               TabIndex        =   115
               Top             =   150
               Width           =   2130
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00E0E0E0&
            Height          =   405
            Left            =   -69900
            TabIndex        =   116
            Top             =   1145
            Width           =   10160
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparente
               Caption         =   "Lista de contas relacionadas com saque"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   35
               Left            =   3780
               TabIndex        =   117
               Top             =   150
               Width           =   2865
            End
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparente
            Caption         =   "Vlr. total :"
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
            Index           =   62
            Left            =   12810
            TabIndex        =   197
            Top             =   8370
            Width           =   2175
            WordWrap        =   -1  'True
         End
         Begin VB.Label LblValortotal 
            Alignment       =   1  'Alinhar à Direita
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
            Caption         =   "Valor total pago = 0.000,00"
            Height          =   210
            Left            =   -61845
            TabIndex        =   118
            Top             =   8430
            Width           =   2100
         End
      End
      Begin DrawSuite2014.USImageList USImageList3 
         Left            =   -67725
         Top             =   570
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frm_Instituicoes2.frx":1B572
         Count           =   1
      End
      Begin DrawSuite2014.USToolBar USToolBar3 
         Height          =   975
         Left            =   -74925
         TabIndex        =   136
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   11
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
         ButtonLeft3     =   80
         ButtonTop3      =   2
         ButtonWidth3    =   51
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Anterior"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Registro anterior."
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
         ButtonLeft4     =   133
         ButtonTop4      =   2
         ButtonWidth4    =   47
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Próximo"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Próximo registro."
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
         ButtonLeft5     =   182
         ButtonTop5      =   2
         ButtonWidth5    =   46
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Visualizar"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Visualizar conta(s) da movimentação (F7)"
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
         ButtonLeft6     =   230
         ButtonTop6      =   2
         ButtonWidth6    =   52
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Atualizar"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Utilizado pelo administrador do sistema."
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
         ButtonLeft7     =   284
         ButtonTop7      =   2
         ButtonWidth7    =   50
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonAlignment8=   2
         ButtonType8     =   1
         ButtonStyle8    =   -1
         BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState8    =   -1
         ButtonLeft8     =   336
         ButtonTop8      =   4
         ButtonWidth8    =   2
         ButtonHeight8   =   54
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
         ButtonLeft9     =   340
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
         ButtonLeft10    =   378
         ButtonTop10     =   2
         ButtonWidth10   =   26
         ButtonHeight10  =   21
         ButtonUseMaskColor10=   0   'False
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonKey11     =   "11"
         ButtonAlignment11=   2
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState11   =   5
         ButtonLeft11    =   406
         ButtonTop11     =   2
         ButtonWidth11   =   24
         ButtonHeight11  =   24
         ButtonUseMaskColor11=   0   'False
      End
      Begin DrawSuite2014.USToolBar USToolBar4 
         Height          =   975
         Left            =   -74925
         TabIndex        =   137
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   13
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
         ButtonCaption3  =   "Excluir/cancelar"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Excluir/cancelar (F4)"
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
         ButtonWidth3    =   83
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
         ButtonLeft4     =   165
         ButtonTop4      =   2
         ButtonWidth4    =   51
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Anterior"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Registro anterior."
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
         ButtonLeft5     =   218
         ButtonTop5      =   2
         ButtonWidth5    =   47
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Próximo"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Próximo registro."
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
         ButtonLeft6     =   267
         ButtonTop6      =   2
         ButtonWidth6    =   46
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Cópia de cheque"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Emitir cópia de cheque (F6)"
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
         ButtonLeft7     =   315
         ButtonTop7      =   2
         ButtonWidth7    =   88
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Compensar"
         ButtonEnabled8  =   0   'False
         ButtonToolTipText8=   "Compensar cheque (F7)"
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
         ButtonLeft8     =   405
         ButtonTop8      =   2
         ButtonWidth8    =   62
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Cancelar compensação"
         ButtonEnabled9  =   0   'False
         ButtonToolTipText9=   "Cancelar compensação do cheque (F8)"
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
         ButtonLeft9     =   469
         ButtonTop9      =   2
         ButtonWidth9    =   118
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonAlignment10=   2
         ButtonType10    =   1
         ButtonStyle10   =   -1
         BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState10   =   -1
         ButtonLeft10    =   589
         ButtonTop10     =   4
         ButtonWidth10   =   2
         ButtonHeight10  =   54
         ButtonCaption11 =   "Ajuda"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Ajuda (F1)"
         ButtonKey11     =   "11"
         ButtonAlignment11=   2
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft11    =   593
         ButtonTop11     =   2
         ButtonWidth11   =   36
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonCaption12 =   "Sair"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Sair (Esc)"
         ButtonKey12     =   "12"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   631
         ButtonTop12     =   2
         ButtonWidth12   =   26
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonEnabled13 =   0   'False
         ButtonIconSize13=   32
         ButtonKey13     =   "13"
         ButtonAlignment13=   2
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState13   =   5
         ButtonLeft13    =   659
         ButtonTop13     =   2
         ButtonWidth13   =   24
         ButtonHeight13  =   24
         ButtonUseMaskColor13=   0   'False
      End
      Begin DrawSuite2014.USImageList USImageList5 
         Left            =   -67845
         Top             =   660
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frm_Instituicoes2.frx":216EB
         Count           =   1
      End
      Begin DrawSuite2014.USToolBar USToolBar5 
         Height          =   975
         Left            =   -74925
         TabIndex        =   138
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
         ButtonLeft2     =   40
         ButtonTop2      =   2
         ButtonWidth2    =   39
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Anterior"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Registro anterior."
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
         ButtonLeft3     =   81
         ButtonTop3      =   2
         ButtonWidth3    =   47
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Próximo"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Próximo registro."
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
         ButtonLeft4     =   130
         ButtonTop4      =   2
         ButtonWidth4    =   46
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Compensar"
         ButtonEnabled5  =   0   'False
         ButtonToolTipText5=   "Compensar cheque (F7)"
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
         ButtonLeft5     =   178
         ButtonTop5      =   2
         ButtonWidth5    =   62
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Cancelar compensação"
         ButtonEnabled6  =   0   'False
         ButtonToolTipText6=   "Cancelar compensação do cheque (F8)"
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
         ButtonLeft6     =   242
         ButtonTop6      =   2
         ButtonWidth6    =   118
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
         ButtonLeft7     =   362
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
         ButtonLeft8     =   366
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
         ButtonLeft9     =   404
         ButtonTop9      =   2
         ButtonWidth9    =   26
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonKey10     =   "10"
         ButtonAlignment10=   2
         BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState10   =   5
         ButtonLeft10    =   432
         ButtonTop10     =   2
         ButtonWidth10   =   24
         ButtonHeight10  =   24
         ButtonUseMaskColor10=   0   'False
      End
      Begin DrawSuite2014.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74925
         TabIndex        =   134
         Top             =   330
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   1720
         ButtonCount     =   13
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   33
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Filtrar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Filtrar (F2)"
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
         ButtonLeft2     =   37
         ButtonTop2      =   2
         ButtonWidth2    =   36
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Salvar"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Salvar (F3)"
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
         ButtonLeft3     =   75
         ButtonTop3      =   2
         ButtonWidth3    =   38
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Excluir"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Excluir (F4)"
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
         ButtonLeft4     =   115
         ButtonTop4      =   2
         ButtonWidth4    =   39
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonCaption5  =   "Anterior"
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonToolTipText5=   "Registro anterior."
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
         ButtonLeft5     =   156
         ButtonTop5      =   2
         ButtonWidth5    =   47
         ButtonHeight5   =   21
         ButtonUseMaskColor5=   0   'False
         ButtonCaption6  =   "Próximo"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Próximo registro."
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
         ButtonLeft6     =   205
         ButtonTop6      =   2
         ButtonWidth6    =   46
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Status"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Status (F7)"
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
         ButtonLeft7     =   253
         ButtonTop7      =   2
         ButtonWidth7    =   39
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonCaption8  =   "Validação"
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonToolTipText8=   "Validar/Cancelar validação (F8)"
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
         ButtonLeft8     =   294
         ButtonTop8      =   2
         ButtonWidth8    =   53
         ButtonHeight8   =   21
         ButtonUseMaskColor8=   0   'False
         ButtonCaption9  =   "Atualizar"
         ButtonEnabled9  =   0   'False
         ButtonIconSize9 =   32
         ButtonToolTipText9=   "Utilizado pelo administrador do sistema."
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
         ButtonLeft9     =   349
         ButtonTop9      =   2
         ButtonWidth9    =   50
         ButtonHeight9   =   21
         ButtonUseMaskColor9=   0   'False
         ButtonEnabled10 =   0   'False
         ButtonIconSize10=   32
         ButtonAlignment10=   2
         ButtonType10    =   1
         ButtonStyle10   =   -1
         BeginProperty ButtonFont10 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState10   =   -1
         ButtonLeft10    =   401
         ButtonTop10     =   4
         ButtonWidth10   =   2
         ButtonHeight10  =   54
         ButtonCaption11 =   "Ajuda"
         ButtonEnabled11 =   0   'False
         ButtonIconSize11=   32
         ButtonToolTipText11=   "Ajuda (F1)"
         ButtonKey11     =   "11"
         ButtonAlignment11=   2
         BeginProperty ButtonFont11 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft11    =   405
         ButtonTop11     =   2
         ButtonWidth11   =   36
         ButtonHeight11  =   21
         ButtonUseMaskColor11=   0   'False
         ButtonCaption12 =   "Sair"
         ButtonEnabled12 =   0   'False
         ButtonIconSize12=   32
         ButtonToolTipText12=   "Sair (Esc)"
         ButtonKey12     =   "12"
         ButtonAlignment12=   2
         BeginProperty ButtonFont12 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft12    =   443
         ButtonTop12     =   2
         ButtonWidth12   =   26
         ButtonHeight12  =   21
         ButtonUseMaskColor12=   0   'False
         ButtonEnabled13 =   0   'False
         ButtonKey13     =   "13"
         BeginProperty ButtonFont13 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState13   =   5
         ButtonLeft13    =   471
         ButtonTop13     =   2
         ButtonWidth13   =   24
         ButtonHeight13  =   24
      End
      Begin MSComctlLib.ListView lst_Instituicoes 
         Height          =   3435
         Left            =   -74925
         TabIndex        =   21
         Top             =   6120
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   6059
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Object.Tag             =   "T"
            Text            =   "Empresa"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "N"
            Text            =   "Banco"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Agência"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Conta"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   11033
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "IDempresa"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "Validada"
            Object.Width           =   1499
         EndProperty
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'Nenhum
         Caption         =   "Operação da lista"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -74925
         TabIndex        =   159
         Top             =   9525
         Width           =   15195
         Begin VB.ComboBox cmb_Opcao_Lista_Instituicao 
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
            ItemData        =   "frm_Instituicoes2.frx":26ADD
            Left            =   13080
            List            =   "frm_Instituicoes2.frx":26AEA
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   60
            Width           =   1965
         End
         Begin DrawSuite2014.USProgressBar PBLista 
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   184
            Top             =   90
            Width           =   11535
            _ExtentX        =   20346
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparente
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
            Index           =   4
            Left            =   11730
            TabIndex        =   160
            Top             =   113
            Width           =   1260
         End
      End
      Begin DrawSuite2014.USProgressBar PBLista 
         Height          =   255
         Index           =   4
         Left            =   -74925
         TabIndex        =   188
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
         SearchText      =   "Atualizando..."
         Value           =   0
      End
      Begin MSComctlLib.ListView lst_Duplicata 
         Height          =   8055
         Left            =   30
         TabIndex        =   213
         Top             =   1260
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   14208
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Nota fiscal"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Sacado (Cliente)"
            Object.Width           =   10231
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Vencimento"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Parcela"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "N"
            Text            =   "Nosso número"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Remessa"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Env.Financeiro"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Env. email?"
            Object.Width           =   2293
         EndProperty
      End
   End
End
Attribute VB_Name = "frm_Instituicoes2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Banco                                          As Boolean 'OK
Public Novo_Banco1                                      As Boolean 'OK
Public Novo_Banco2                                      As Boolean 'OK
Public Novo_Banco3                                      As Boolean 'OK
Public Instituicao_Localizar_Transf                     As String 'OK
Public Instituicao_Localizar_Saque                      As String 'OK
Public Instituicao_Localizar_Tarifa                     As String 'OK
Public StrSql_Instituicoes_Localizar                    As String 'OK
Public StrSql_Instituicoes_Localizar_Cheque             As String 'OK
Public StrSql_Instituicoes_Localizar_Cheque_Cancelados  As String 'OK
Public StrSql_Instituicoes_Localizar_Cheque_Recebidos   As String 'OK
Public FormulaRel_Instituicao                           As String 'OK
Public FormulaRel_Instituicao1                          As String 'OK
Dim Total2                                              As Double 'OK
Dim VlrCheque                                           As Double 'OK
Dim Vlrconta                                            As Double 'OK
Dim Saldo                                               As Double 'OK
Dim SaldoAlterado                                       As Double 'OK
Public StrSql_Contas_Pagar_Cheque                       As String 'OK
Public Cheques_Emitidos                                 As Boolean 'OK
Dim TBLISTA_Instituicao                                 As ADODB.Recordset 'OK


Public Sub ProcPassadadosEmailCopiaParaCobrebemX1()
On Error GoTo tratar_erro

    If chkEmailcopia.Value = 1 Then
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.CopiaReply = True
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailReply.Endereco = EmailCopia
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailReply.Nome = txtAssunto.Text
    End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcPassadadosEmailParaCobreBemX1()
On Error GoTo tratar_erro


    CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.LayoutBoleto = "Padrao"
    
    'Utilizar para sair o endereço em outro campo
    
    CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.LayoutBoleto = "PadraoReciboPersonalizado"
    CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.HTMLReciboPersonalizado = TxtHTLM
    CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.MargemSuperior = 0
    CobreBemX1.LocalPagamento = "Preferencialmente nas Casas Lotéricas até o valor limite"

    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select EE.*, E.Empresa from Empresa E INNER JOIN Empresa_email EE ON EE.ID_empresa = E.Codigo where EE.ID_empresa = " & txtCodcedente & " and EE.Aplicacao = 'F'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
       'Início da configuração dos dados do Cedente para envio de boletos por email
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.SMTP.Servidor = TBAbrir!Servidor_SMTP ' Trocar para apontar para o seu servidor SMTP
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.SMTP.Porta = TBAbrir!Porta
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.SMTP.Usuario = TBAbrir!Usuario 'utilizar esta propriedade para acesso a servidores SMTP seguros
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.SMTP.Senha = TBAbrir!Senha 'utilizar esta propriedade para acesso a servidores SMTP seguros
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.URLImagensCodigoBarras = "http://www.bptob.com/imagenscbe/"
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.URLLogotipo = "http://www.thisf.com.br/banners/BannerCBE.gif"
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.Assunto = "Boleto Caprind" 'Assunto_email
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailFrom.Endereco = TBAbrir!Email
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailFrom.Nome = TBAbrir!Nome
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.FormaEnvio = IIf(Tipo_Documento = "PDF", feeSMTPMensagemBoletoPDFAnexo, feeSMTPBoletoHTML)
    End If

'Logotipo do cedente na parte superior do boleto
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Logotipo from Empresa where Codigo = " & txtCodcedente & " and Logotipo <> 'Null'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If TBAbrir!Logotipo <> "" Then CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.ArquivoLogotipo = TBAbrir!Logotipo
End If
TBAbrir.Close
       
CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.CaminhoImagensCodigoBarras = Localrel & "\Imagens\Bancos\"

'Utilize o parâmetro abaixo para efetuar ajustes na impressão do boleto subindo ou descendo o mesmo na folha de papel
'Os valores devem ser informados em milímetros e quanto maior o valor mais para baixo será iniciado o boleto
'Se este parâmetro não for passado será assumido o valor 15 que é o indicado para a maioria das impressoras Jato de Tinta }
CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.MargemSuperior = 3

'A próxima linha é utilizada para exibir um texto do lado direito do logotipo nos boletos impressos ou enviados por email
'CobreBemX1.PadroesBoleto.IdentificacaoCedente =

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcPassaDadosContaCorrenteParaCobreBemX1(Carteira As String, Carteira1 As String, Codigocedente As String, ID_empresa As Integer, EmitirBoleto As Boolean, Assunto_email As String)
On Error GoTo tratar_erro

Diretorio = Txtlocal
ContaCorrente = txtContacorrente

If txtNBanco.Text = "053" Then
        
    Select Case Len(ContaCorrente)
        Case 1: ContaCorrenteBol = "00000-" & ContaCorrente
        Case 2: ContaCorrenteBol = "0000" & Left(ContaCorrente, 1) & "-" & Right(ContaCorrente, 1)
        Case 3: ContaCorrenteBol = "000" & Left(ContaCorrente, 2) & "-" & Right(ContaCorrente, 1)
        Case 4: ContaCorrenteBol = "00" & Left(ContaCorrente, 3) & "-" & Right(ContaCorrente, 1)
        Case 5: ContaCorrenteBol = "0" & Left(ContaCorrente, 4) & "-" & Right(ContaCorrente, 1)
        Case Is >= 6: ContaCorrenteBol = Left(ContaCorrente, 5) & "-" & Right(ContaCorrente, 1)
    End Select
End If

If txtNBanco.Text = "001" Then
        
            Select Case Len(txtAgencia)
                Case 1: AgenciaBol = "0000-" & txtAgencia
                Case 2: AgenciaBol = "000" & Left(txtAgencia, 1) & "-" & Right(txtAgencia, 1)
                Case 3: AgenciaBol = "00" & Left(txtAgencia, 2) & "-" & Right(txtAgencia, 1)
                Case 4: AgenciaBol = "0" & Left(txtAgencia, 3) & "-" & Right(txtAgencia, 1)
                Case Is >= 5: AgenciaBol = Left(txtAgencia, 4) & "-" & Right(txtAgencia, 1)
            End Select
            Select Case Len(ContaCorrente)
                Case 1: ContaCorrenteBol = "00000000-" & ContaCorrente
                Case 2: ContaCorrenteBol = "0000000" & Left(ContaCorrente, 1) & "-" & Right(ContaCorrente, 1)
                Case 3: ContaCorrenteBol = "000000" & Left(ContaCorrente, 2) & "-" & Right(ContaCorrente, 1)
                Case 4: ContaCorrenteBol = "00000" & Left(ContaCorrente, 3) & "-" & Right(ContaCorrente, 1)
                Case 5: ContaCorrenteBol = "0000" & Left(ContaCorrente, 4) & "-" & Right(ContaCorrente, 1)
                Case 6: ContaCorrenteBol = "000" & Left(ContaCorrente, 5) & "-" & Right(ContaCorrente, 1)
                Case 7: ContaCorrenteBol = "00" & Left(ContaCorrente, 6) & "-" & Right(ContaCorrente, 1)
                Case 8: ContaCorrenteBol = "0" & Left(ContaCorrente, 7) & "-" & Right(ContaCorrente, 1)
                Case Is >= 9: ContaCorrenteBol = Left(ContaCorrente, 8) & "-" & Right(ContaCorrente, 1)
            End Select
End If
Debug.Print ContaCorrenteBol
Debug.Print AgenciaBol

ArquivoLicensa = txtcarteiraconf
CobreBemX1.ArquivoLicenca = Localrel & "\Boletos\Carteiras\" & ArquivoLicensa

If CobreBemX1.ArquivoLicenca = "" Then
    txtcarteiraconf.Text = ""
    MsgBox "Arquivo " & ArquivoLicensa & " de configuração da carteira " & cmbCarteira & " do banco " & cmbBanco & ", não foi encontrado na pasta de carteiras do banco, favor verificar antes de prosseguir", vbExclamation
    chkRemessa.Value = 0
    chkRemessa.Enabled = False
    chkEmail.Value = 0
    chkEmail.Enabled = False
    chkEmailcopia.Value = 0
    chkEmailcopia.Enabled = False
    chkImprimir.Value = 0
    chkImprimir.Enabled = False
    Exit Sub
End If

CobreBemX1.CodigoAgencia = AgenciaBol
CobreBemX1.NumeroContaCorrente = ContaCorrenteBol
CobreBemX1.Codigocedente = "035778"
CobreBemX1.OutroDadoConfiguracao1 = "019"
CobreBemX1.OutroDadoConfiguracao2 = "019"

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcPassaDadosBoletosParaCobreBemX1()
On Error GoTo tratar_erro
Dim Boleto As Object
Dim Email As Object

Set TBAbrir = CreateObject("adodb.recordset")
Set TBContas = CreateObject("adodb.recordset")

CobreBemX1.DocumentosCobranca.Clear
   
    With frm_Boleto.lst_Duplicata
        For InitFor = 1 To .ListItems.Count
        Contador = .ListItems.Count
        Contador = 100 / Contador
            If .ListItems.Item(InitFor).Checked = True Then
                TBContas.Open "Select * from tbl_contas_receber where IdIntConta = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    ProcCarregadadosSacado (TBContas!Nome_Razao)
                    TBAbrir.Open "Select * from tbl_Detalhes_Recebimento where IdContaReceber = " & TBContas!IDintconta & "", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                                            
                                Set Boleto = CobreBemX1.DocumentosCobranca.Add
                                Boleto.TipoDocumentoCobranca = Especie
                                Boleto.NumeroDocumento = Right(TBContas!NFiscal, 6) & "/" & Left(TBContas!Parcela, 3)
                                Boleto.NomeSacado = TBContas!Nome_Razao
                                Boleto.CNPJSacado = StrCnpj
                                Boleto.EnderecoSacado = StrEndereco
                                Boleto.BairroSacado = StrBairro
                                Boleto.CidadeSacado = StrCidade
                                Boleto.EstadoSacado = StrEstado
                                Boleto.CepSacado = StrCEP
                                Boleto.DataDocumento = Format$(Date, "dd/mm/yyyy")
                                Boleto.DataVencimento = TBAbrir!dt_Vencimento
                                Boleto.DataProcessamento = Format$(Date, "dd/mm/yyyy")
                                Boleto.ValorDocumento = TBAbrir!dbl_valor
                                Boleto.PercentualJurosDiaAtraso = Txtpercentual_juros
                                Boleto.PercentualMultaAtraso = Txtpercentual_multa
                                Boleto.PercentualDesconto = Txtpercentual_desconto
                                Boleto.ValorOutrosAcrescimos = IIf(IsNull(TBAbrir!Acrescimos), 0, TBAbrir!Acrescimos)
                                Boleto.PadroesBoleto.Demonstrativo = Right(TBContas!NFiscal, 6) & "/" & Left(TBContas!Parcela, 3)
                                Boleto.PadroesBoleto.InstrucoesCaixa = Txtinstrucoes
                                Boleto.ControleProcessamentoDocumento.Imprime = scpExecutar
                                Boleto.ControleProcessamentoDocumento.EnviaEmail = scpExecutar
                                Boleto.ControleProcessamentoDocumento.GravaRemessa = scpExecutar
                                ProcGerarNossoNumero
                                Boleto.NossoNumero = Var
                                
                            'Passa dados boleto para o detalhe recebimento
                                TBAbrir!Seq_remessa = IIf(Seq > 0, Seq, 1)
                                TBAbrir!Nosso_numero = Var
                                TBAbrir!Juros = Txtpercentual_juros
                                TBAbrir!Multa = Txtpercentual_multa
                                TBAbrir!Desconto = Txtpercentual_desconto
                                TBAbrir!Instrucoes = Txtinstrucoes
                                TBAbrir!dias_protesto = Txtdias_protesto
                                TBAbrir!Carteira = cmbCarteira
                                TBAbrir!Data_emissao = Date
                                TBAbrir!Vencimento_boleto = Boleto.DataVencimento
                                TBAbrir!Valor_boleto = Boleto.ValorDocumento
                                TBAbrir!Numero_documento = Right(TBContas!NFiscal, 6) & "/" & Left(TBContas!Parcela, 3)
                                TBContas!txt_ndocumento = Right(TBContas!NFiscal, 6) & "/" & Left(TBContas!Parcela, 3)
                                TBContas.Update
                            'Envia email
                                    If chkEmail.Value = 1 Then
                                            Set TBFIltro = CreateObject("adodb.recordset")
                                                TBFIltro.Open "Select * from Clientes_Contatos where IDCliente = " & TBContas!IDCliente & " and Enviar_boleto = 'TRUE'", Conexao, adOpenKeyset, adLockOptimistic
                                             If TBFIltro.EOF = False Then
                                                Do While TBFIltro.EOF = False
                                                    Set Email = Boleto.EnderecosEmailSacado.Add
                                                        Email.Nome = Boleto.NomeSacado
                                                        Email.Endereco = TBFIltro!Email
                                                        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.FormaEnvio = IIf(Tipo_Documento = "PDF", feeSMTPMensagemBoletoPDFAnexo, feeSMTPBoletoHTML)
                                                TBFIltro.MoveNext
                                                Loop
                                             Else
                                                 Email.Endereco = "caprind@caprind.com.br"
                                             End If
                                        TBAbrir!Enviado = True
                                    End If
                                        TBAbrir!data_envio = Date
                                TBAbrir.Update
                                'Imprimir boleto(s)
                                    If chkImprimir.Value = 1 Then
                                        CobreBemX1.ImprimeBoletos
                                    End If

                            'Passa dados boleto para o N_Boleto
                               Set TBBoleto = CreateObject("adodb.recordset")
                                    TBBoleto.Open "Select * from tbl_Detalhes_Recebimento_Nboletos where IDContaReceber = " & TBContas!IDintconta & " and Nosso_numero = '" & Var & "'", Conexao, adOpenKeyset, adLockOptimistic
                                    If TBBoleto.EOF = True Then TBBoleto.AddNew
                                    TBBoleto!data = Date
                                    TBBoleto!Responsavel = "Emissor boleto Caprind" 'pubUsuario
                                    TBBoleto!IdContaReceber = TBContas!IDintconta
                                    TBBoleto!Nosso_numero = Var
                                    TBBoleto!ID_nota = TBContas!ID_nota
                                    TBBoleto.Update
                                    TBBoleto.Close
                                TBContas.Close
                        End If
                        TBAbrir.Close
                End If
            End If
        Next InitFor
    End With
    
   

If chkEmail.Value = 1 Then
    CobreBemX1.EnviaBoletosPorEmail
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaDadosBoletosParaCobreBemX1()
On Error GoTo tratar_erro
     
Set TBAbrir = CreateObject("adodb.recordset")
Set TBContas = CreateObject("adodb.recordset")

CobreBemX1.DocumentosCobranca.Clear

   
    With frm_Boleto.lst_Duplicata
        For InitFor = 1 To .ListItems.Count
        Contador = .ListItems.Count
        Contador = 100 / Contador
            If .ListItems.Item(InitFor).Checked = True Then
                TBContas.Open "Select * from tbl_contas_receber where IdIntConta = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    ProcCarregadadosSacado (TBContas!Nome_Razao)
                    TBAbrir.Open "Select * from tbl_Detalhes_Recebimento where IdContaReceber = " & TBContas!IDintconta & "", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            If chkAtualizar.Value = 1 Then
                                CobreBemX1.ArquivoRemessa.Arquivo = .ListItems(InitFor).ListSubItems.Item(6).Text
                            End If
                            'frm_Atualiza_boleto.Show 1
                
                            ProcGerarNossoNumero
                            
                               Set Boleto = CobreBemX1.DocumentosCobranca.Add
                                Boleto.TipoDocumentoCobranca = Especie
                                Boleto.NumeroDocumento = Right(TBContas!NFiscal, 6) & "/" & Left(TBContas!Parcela, 3)
                                Boleto.NomeSacado = TBContas!Nome_Razao
                                Boleto.CNPJSacado = StrCnpj
                                Boleto.EnderecoSacado = StrEndereco
                                Boleto.BairroSacado = StrBairro
                                Boleto.CidadeSacado = StrCidade
                                Boleto.EstadoSacado = StrEstado
                                Boleto.CepSacado = StrCEP
                                Boleto.DataDocumento = Format$(Date, "dd/mm/yyyy")
                                Boleto.DataVencimento = IIf((chkAtualizar.Value = 1), NovoVencimento, TBAbrir!dt_Vencimento)
                                Boleto.DataProcessamento = Format$(Date, "dd/mm/yyyy")
                                Boleto.ValorDocumento = IIf((chkAtualizar.Value = 1), VA, TBAbrir!dbl_valor)
                                Boleto.PercentualJurosDiaAtraso = Txtpercentual_juros
                                Boleto.PercentualMultaAtraso = Txtpercentual_multa
                                Boleto.PercentualDesconto = Txtpercentual_desconto
                                Boleto.ValorOutrosAcrescimos = IIf(IsNull(TBAbrir!Acrescimos), 0, TBAbrir!Acrescimos)
                                Boleto.PadroesBoleto.Demonstrativo = Right(TBContas!NFiscal, 6) & "/" & Left(TBContas!Parcela, 3)
                                Boleto.PadroesBoleto.InstrucoesCaixa = Txtinstrucoes
                                Boleto.ControleProcessamentoDocumento.Imprime = scpExecutar
                                Boleto.ControleProcessamentoDocumento.EnviaEmail = scpExecutar
                                Boleto.ControleProcessamentoDocumento.GravaRemessa = scpExecutar
                                Boleto.NossoNumero = Var
                            'Passa dados boleto para o detalhe recebimento
                                If chkAtualizar.Value = 0 Then TBAbrir!Seq_remessa = IIf(Seq > 0, Seq, 1)
                                TBAbrir!Nosso_numero = Var
                                TBAbrir!Juros = Txtpercentual_juros
                                TBAbrir!Multa = Txtpercentual_multa
                                TBAbrir!Desconto = Txtpercentual_desconto
                                TBAbrir!Instrucoes = Txtinstrucoes
                                TBAbrir!dias_protesto = Txtdias_protesto
                                TBAbrir!Carteira = cmbCarteira
                                TBAbrir!Data_emissao = Date
                                TBAbrir!Vencimento_boleto = Boleto.DataVencimento
                                TBAbrir!Valor_boleto = Boleto.ValorDocumento
                                TBContas!txt_ndocumento = Right(TBContas!NFiscal, 6) & "/" & Left(TBContas!Parcela, 3)
                                TBContas.Update
                                 'Envia email
                                    If chkEmail.Value = 1 Then
                                            Set TBFIltro = CreateObject("adodb.recordset")
                                                TBFIltro.Open "Select * from Clientes_Contatos where IDCliente = " & TBContas!IDCliente & " and Enviar_boleto = 'TRUE'", Conexao, adOpenKeyset, adLockOptimistic
                                             If TBFIltro.EOF = False Then
                                                Do While TBFIltro.EOF = False
                                                    Set Email = Boleto.EnderecosEmailSacado.Add
                                                        Email.Nome = Boleto.NomeSacado
                                                        Email.Endereco = TBFIltro!Email
                                                TBFIltro.MoveNext
                                                Loop
                                             Else
                                                 Email.Endereco = "caprind@caprind.com.br"
                                             End If
                                        TBAbrir!Enviado = True
                                    End If
                                        TBAbrir!data_envio = Date
                                TBAbrir.Update
                                'Imprimir boleto(s)
                                    If chkImprimir.Value = 1 Then
                                        CobreBemX1.ImprimeBoletos
                                    End If

                            'Passa dados boleto para o N_Boleto
                               Set TBBoleto = CreateObject("adodb.recordset")
                                    TBBoleto.Open "Select * from tbl_Detalhes_Recebimento_Nboletos where IDContaReceber = " & TBContas!IDintconta & " and Nosso_numero = '" & Var & "'", Conexao, adOpenKeyset, adLockOptimistic
                                    If TBBoleto.EOF = True Then TBBoleto.AddNew
                                    TBBoleto!data = Date
                                    TBBoleto!Responsavel = "Emissor boleto Caprind" 'pubUsuario
                                    TBBoleto!IdContaReceber = TBContas!IDintconta
                                    TBBoleto!Nosso_numero = Var
                                    TBBoleto!ID_nota = TBContas!ID_nota
                                    TBBoleto.Update
                                    TBBoleto.Close
                                TBContas.Close
                        End If
                        TBAbrir.Close
                End If
            End If
        Next InitFor
    End With
    
   

If chkEmail.Value = 1 Then
    CobreBemX1.EnviaBoletosPorEmail
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=bxzFUe4ntt4&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=39&feature=plcp")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub chkRemessa_Click()
On Error GoTo tratar_erro

    If chkRemessa.Value = 1 Then
    cmdProcessar.Enabled = chkRemessa.Value
    chkEmail.Enabled = True
    chkImprimir.Enabled = True
    chkImprimir.Value = 0
    End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub chkEmail_Click()
On Error GoTo tratar_erro

If chkEmail.Value = 1 Then
ProcPassadadosEmailParaCobreBemX1
chkEmailcopia.Enabled = chkEmail.Value
cmdProcessar.Enabled = chkEmail.Value
Else
chkEmailcopia.Enabled = chkEmail.Value
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub chkEmailcopia_Click()
On Error GoTo tratar_erro

    If chkEmailcopia.Value = 1 Then
    ProcPassadadosEmailCopiaParaCobrebemX1
    End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub chkImprimir_Click()
On Error GoTo tratar_erro

If chkImprimir.Value = 1 Then
    cmdProcessar.Enabled = chkImprimir.Value
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

ProcLimpaCampos

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmb_forma_Click()
On Error GoTo tratar_erro

LblDocumento.Caption = "N° do documento"
txtCheque = ""
txtCheque.Locked = True
txtCheque.TabStop = False
txtfavorecido = ""
txtfavorecido.Locked = True
txtfavorecido.TabStop = False
TxtHistDepTranf = ""

TextoFiltro = ""
If txtNBanco <> "" Then
    If cmb_forma = "TEV" Then Texto = "=" Else Texto = "<>"
    TextoFiltro = " and int_Nbanco " & Texto & " " & txtNBanco
End If
Select Case cmb_forma.Text
    Case "DOC":
        LblDocumento.Caption = "N° do DOC*"
        txtCheque.Locked = False
        txtCheque.TabStop = True
        txtfavorecido.Locked = True
        txtfavorecido.TabStop = False
        If txtCodBanco <> "" Then ProcCarregaComboBancoFinanceiro cmbrecebedor, "ID <> " & txtCodBanco & " " & TextoFiltro & " and txt_Descricao <> 'Null' and Bloqueado <> 'True' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False
        Exit Sub
    Case "TED":
        LblDocumento.Caption = "N° do TED*"
        txtCheque.Locked = False
        txtCheque.TabStop = True
        txtfavorecido.Locked = True
        txtfavorecido.TabStop = False
        If txtCodBanco <> "" Then ProcCarregaComboBancoFinanceiro cmbrecebedor, "ID <> " & txtCodBanco & " " & TextoFiltro & " and txt_Descricao <> 'Null' and Bloqueado <> 'True' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False
        Exit Sub
    Case "TEV":
        LblDocumento.Caption = "N° do TEV*"
        txtCheque.Locked = False
        txtCheque.TabStop = True
        txtfavorecido.Locked = True
        txtfavorecido.TabStop = False
        If txtCodBanco <> "" Then ProcCarregaComboBancoFinanceiro cmbrecebedor, "ID <> " & txtCodBanco & " " & TextoFiltro & " and txt_Descricao <> 'Null' and Bloqueado <> 'True' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False
        Exit Sub
    Case "CHEQUE":
        LblDocumento.Caption = "N° do cheque*"
        txtCheque.Locked = False
        txtCheque.TabStop = True
        txtfavorecido.Locked = False
        txtfavorecido.TabStop = True
        If txtCodBanco <> "" Then ProcCarregaComboBancoFinanceiro cmbrecebedor, "ID <> " & txtCodBanco & " and txt_Descricao <> 'Null' and Bloqueado <> 'True' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False
    Case "Dinheiro":
        TxtHistDepTranf = "Depósito"
        If txtCodBanco <> "" Then ProcCarregaComboBancoFinanceiro cmbrecebedor, "ID <> " & txtCodBanco & " and txt_Descricao <> 'Null' and Bloqueado <> 'True' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), False
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_Click()
On Error GoTo tratar_erro

With Lst_cheque
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With
With Lst_cheque1
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar4
    Select Case Cmb_opcao_lista
        Case "Excluir/cancelar":
            .ButtonState(3) = 0
            .ButtonState(8) = 5
            .ButtonState(9) = 5
        Case "Compensar":
            .ButtonState(3) = 5
            .ButtonState(8) = 0
            .ButtonState(9) = 5
        Case "Cancelar compensação":
            .ButtonState(3) = 5
            .ButtonState(8) = 5
            .ButtonState(9) = 0
    End Select
    .Refresh
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmb_Opcao_Lista_Instituicao_Click()
On Error GoTo tratar_erro

With lst_Instituicoes
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar1
    If cmb_Opcao_Lista_Instituicao = "Excluir" Then
        .ButtonState(4) = 0
        .ButtonState(7) = 5
        .ButtonState(8) = 5
    ElseIf cmb_Opcao_Lista_Instituicao = "Status" Then
            .ButtonState(4) = 5
            .ButtonState(7) = 0
            .ButtonState(8) = 5
        Else
            .ButtonState(4) = 5
            .ButtonState(7) = 5
            .ButtonState(8) = 0
    End If
    .Refresh
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmb_opcao_lista_recebidos_Click()
On Error GoTo tratar_erro

With Lista_cheque
    For InitFor = 1 To .ListItems.Count
        .ListItems.Item(InitFor).Checked = False
    Next InitFor
End With

With USToolBar5
    Select Case Cmb_opcao_lista_recebidos
        Case "Excluir":
            .ButtonState(2) = 0
            .ButtonState(5) = 5
            .ButtonState(6) = 5
        Case "Compensar":
            .ButtonState(2) = 5
            .ButtonState(5) = 0
            .ButtonState(6) = 5
        Case "Cancelar compensação":
            .ButtonState(2) = 5
            .ButtonState(5) = 5
            .ButtonState(6) = 0
    End Select
    .Refresh
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmb_operacao_Click()
On Error GoTo tratar_erro

ProcCarregaTipoDocumento
ProcCarregaComboForma
Txt_ID_PC = 0
Txt_codigo_PC = ""
Txt_descricao_PC = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCancelarCompensacaoChequeEmitido()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
Cheque = ""
Cheque1 = ""
Permitido = False
With Lst_cheque
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If MsgBox("Deseja realmente cancelar a compensação deste(s) cheque(s)?", vbQuestion + vbYesNo) = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select ID, ID_empresa, Txt_descricao from tbl_Instituicoes WHERE ID = " & txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Cheque = "Cheque n. " & .ListItems(InitFor).ListSubItems(2)
                If Cheque <> Cheque1 Then
                    '==================================
                    Modulo = "Financeiro/Instituições"
                    Evento = "Cancelar compensação do cheque emitido"
                    ID_documento = .ListItems(InitFor)
                    Documento = "Cheque nº: " & .ListItems(InitFor).ListSubItems(2) & " - Instituição bancária: " & TBFI!Txt_descricao
                    Documento1 = ""
                    ProcGravaEvento
                    '==================================
                                    
                    Conexao.Execute "Update tbl_Fluxo_de_caixa Set Bloqueado = 'True' where Operacao = 'Débito' and Instituicao = '" & txtDescricao & "' and ID_empresa = " & TBFI!ID_empresa & " and Descricao = '" & Cheque & "'"
                    Conexao.Execute "Update tbl_ContasPagar Set Data_movimentacao = DataBaixa where Banco = '" & txtDescricao & "' and ID_empresa = " & TBFI!ID_empresa & " and NDoctoBaixa = '" & Cheque & "'"
                    
                    Set TBContas = CreateObject("adodb.recordset")
                    TBContas.Open "Select NDoctoBaixa from tbl_ContasPagar where IdIntConta = " & .ListItems(InitFor) & " and Status = 'DEPÓSITO EM CHEQUE'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBContas.EOF = False Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select IDFluxo, IDFluxo_rec from tbl_instituicoes_transf where NDoctoBaixa = '" & TBContas!NDoctoBaixa & "' and id_banco_rem = " & TBFI!ID & " and FormaBaixa = 'CHEQUE'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            'Corrige saldo do banco recebedor
                            Set TBFluxo = CreateObject("adodb.recordset")
                            TBFluxo.Open "Select ID_empresa, Instituicao, Valor from tbl_Fluxo_de_caixa where IDFluxo = " & TBAbrir!IDFluxo_Rec, Conexao, adOpenKeyset, adLockOptimistic
                            If TBFluxo.EOF = False Then
                                Conexao.Execute "Update tbl_Fluxo_de_caixa Set Bloqueado = 'True' where Operacao = 'Crédito' and Instituicao = '" & TBFluxo!Instituicao & "' and ID_empresa = " & TBFluxo!ID_empresa & " and Descricao = '" & Cheque & "'"
                                Set TBProduto = CreateObject("adodb.recordset")
                                TBProduto.Open "Select Saldo from tbl_instituicoes where txt_descricao = '" & TBFluxo!Instituicao & "' and ID_empresa = " & TBFluxo!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                                If TBProduto.EOF = False Then
                                    TBProduto!Saldo = Format(TBProduto!Saldo - TBFluxo!valor, "###,##0.00")
                                    TBProduto.Update
                                End If
                                TBProduto.Close
                            End If
                            TBFluxo.Close
                            
                            'Corrige saldo do banco remetente
                            Set TBFluxo = CreateObject("adodb.recordset")
                            TBFluxo.Open "Select ID_empresa, Instituicao, Valor from tbl_Fluxo_de_caixa where IDFluxo = " & TBAbrir!IDFluxo, Conexao, adOpenKeyset, adLockOptimistic
                            If TBFluxo.EOF = False Then
                                Set TBProduto = CreateObject("adodb.recordset")
                                TBProduto.Open "Select Saldo from tbl_instituicoes where txt_descricao = '" & TBFluxo!Instituicao & "' and ID_empresa = " & TBFluxo!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                                If TBProduto.EOF = False Then
                                    TBProduto!Saldo = Format(TBProduto!Saldo + TBFluxo!valor, "###,##0.00")
                                    txtsaldo = Format(TBProduto!Saldo, "###,##0.00")
                                    TBProduto.Update
                                End If
                                TBProduto.Close
                            End If
                            TBFluxo.Close
                        End If
                        TBAbrir.Close
                    Else
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select Saldo from tbl_instituicoes where txt_Descricao = '" & TBFI!Txt_descricao & "' and ID_empresa = " & TBFI!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                        If TBGravar.EOF = False Then
                            Cheque = "Cheque n. " & .ListItems(InitFor).ListSubItems(2)
                            Set TBFluxo = CreateObject("adodb.recordset")
                            TBFluxo.Open "Select Valor from tbl_Fluxo_de_caixa where Operacao = 'Débito' and Instituicao = '" & TBFI!Txt_descricao & "' and Descricao = '" & Cheque & "'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFluxo.EOF = False Then
                                TBGravar!Saldo = Format(TBGravar!Saldo + TBFluxo!valor, "###,##0.00")
                                txtsaldo = Format(TBGravar!Saldo, "###,##0.00")
                            End If
                            TBFluxo.Close
                            TBGravar.Update
                        End If
                        TBGravar.Close
                    End If
                    TBContas.Close
                End If
                Cheque1 = "Cheque n. " & .ListItems(InitFor).ListSubItems(2)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    MsgBox ("Informe o(s) cheque(s) antes de cancelar a compensação."), vbExclamation
Else
    MsgBox ("Compensação do(s) cheque(s) cancelada(s) com sucesso."), vbInformation
    ProcCarregaListaCheque
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCancelarCompensacaoChequeRecebido()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
Cheque = ""
Cheque1 = ""
Permitido = False
With Lista_cheque
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If MsgBox("Deseja realmente cancelar a compensação deste(s) cheque(s)?", vbQuestion + vbYesNo) = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from tbl_Instituicoes WHERE ID = " & txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Cheque = "Cheque n. " & .ListItems(InitFor).ListSubItems(2)
                If Cheque <> Cheque1 Then
                    '==================================
                    Modulo = "Financeiro/Instituições"
                    Evento = "Cancelar compensação do cheque recebido"
                    ID_documento = .ListItems(InitFor)
                    Documento = "Cheque nº: " & .ListItems(InitFor).ListSubItems(2) & " - Instituição bancária: " & TBFI!Txt_descricao
                    Documento1 = ""
                    ProcGravaEvento
                    '==================================
                                    
                    Conexao.Execute "Update tbl_Fluxo_de_caixa Set Bloqueado = 'True' where Operacao = 'Crédito' and Instituicao = '" & TBFI!Txt_descricao & "' and Descricao = '" & Cheque & "'"
                    Conexao.Execute "Update tbl_contas_receber Set Data_movimentacao = Data_pagamento where Banco = '" & txtDescricao & "' and NDoctoBaixa = '" & Cheque & "'"
                    
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from tbl_instituicoes where txt_Descricao = '" & TBFI!Txt_descricao & "' and ID_empresa = " & TBFI!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                    If TBGravar.EOF = False Then
                        Cheque = "Cheque n. " & .ListItems(InitFor).ListSubItems(2)
                        Set TBFluxo = CreateObject("adodb.recordset")
                        TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where Operacao = 'Crédito' and Instituicao = '" & TBFI!Txt_descricao & "' and Descricao = '" & Cheque & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFluxo.EOF = False Then
                            TBGravar!Saldo = Format(TBGravar!Saldo - TBFluxo!valor, "###,##0.00")
                            txtsaldo = Format(TBGravar!Saldo, "###,##0.00")
                        End If
                        TBFluxo.Close
                        TBGravar.Update
                    End If
                    TBGravar.Close
                End If
                Cheque1 = "Cheque n. " & .ListItems(InitFor).ListSubItems(2)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    MsgBox ("Informe o(s) cheque(s) antes de cancelar a compensação."), vbExclamation
Else
    MsgBox ("Compensação do(s) cheque(s) cancelada(s) com sucesso."), vbInformation
    ProcCarregaListaCheque
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCompensarChequeEmitido()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
Permitido = False
With Lst_cheque
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then Permitido = True
    Next InitFor
End With
If Permitido = False Then
    MsgBox ("Informe o(s) cheque(es) antes de compensar."), vbExclamation
    Exit Sub
End If
Cheques_Emitidos = True
frm_Instituicoes2_compensar_cheque.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCompensarChequeRecebido()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
Permitido = False
With Lista_cheque
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then Permitido = True
    Next InitFor
End With
If Permitido = False Then
    MsgBox ("Informe o(s) cheque(es) antes de compensar."), vbExclamation
    Exit Sub
End If
Cheques_Emitidos = False
frm_Instituicoes2_compensar_cheque.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCopiaChequeEmitido()
On Error GoTo tratar_erro

If txtDtValidacao = "" Then
    MsgBox "Não é possivel copiar o cheque, pois a instituição ainda não foi validada.", vbExclamation
    Exit Sub
End If
If txtStatus = "Bloqueada" Then
    MsgBox "Não é possivel copiar o cheque, pois a instituição esta bloqueada.", vbExclamation
    Exit Sub
End If
frm_Instituicoes2_menu_impressao_copia_cheque.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAnterior()
On Error GoTo tratar_erro

If txtCodBanco = "" Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Instituicoes order by txt_Descricao", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.BOF = False Then
    TBAbrir.Find ("Id = " & txtCodBanco)
    TBAbrir.MovePrevious
    If TBAbrir.BOF = False Then
        txtCodBanco = TBAbrir!ID
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from tbl_Instituicoes where Id = " & txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcCarregaDados
    Else
        MsgBox ("Fim dos cadastros de instituições bancária."), vbInformation
    End If
End If
Novo_Banco1 = False
Novo_Banco2 = False
Novo_Banco3 = False

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcFiltrarExtrato()
On Error GoTo tratar_erro

TotalCredito = 0
TotalDebito = 0
Lst_extrato.ListItems.Clear
ProcAtualizaSaldoBancario

'Verifica saldo inicial
Valor_total = txtsaldo.Text
Datafim = IIf(Date < msk_fltFim, msk_fltFim, Date)

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(Valor) as Valor, Data, Operacao from tbl_Fluxo_de_caixa where Instituicao = '" & txtDescricao.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and bloqueado = 'False' and (Data) Between '" & Format(msk_fltInicio, "Short Date") & "' And '" & Format(Datafim, "Short Date") & "' and (Operacao = 'Crédito' or Operacao = 'Débito') Group by Data, Operacao Order by Data", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Dataini = TBAbrir!data
    Do While TBAbrir.EOF = False
        Datafim = TBAbrir!data
        If TBAbrir!Operacao = "Crédito" Then
            TotalCredito = TotalCredito + IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
        Else
            TotalDebito = TotalDebito + IIf(IsNull(TBAbrir!valor), 0, TBAbrir!valor)
        End If
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close
Saldo_Anterior = Valor_total - TotalCredito
Saldo_Anterior = Saldo_Anterior + TotalDebito

'Gravar data inicial para pesquisa e saldo inicial
Conexao.Execute "DELETE from tbl_Fluxo_de_caixa_saldos"
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_Fluxo_de_caixa_saldos", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
TBGravar!DataInicial = msk_fltInicio.Value
TBGravar!SaldoInicial = Saldo_Anterior
TBGravar.Update
TBGravar.Close

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_Fluxo_de_caixa where Instituicao = '" & txtDescricao & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and bloqueado = 'False' and (Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and (Operacao = 'Crédito' or Operacao = 'Débito') order by Data, Hora, IDFluxo", Conexao, adOpenKeyset, adLockOptimistic
ProcCarregaListaExtrato

'Gravar data final para pesquisa e saldo final
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_Fluxo_de_caixa_saldos", Conexao, adOpenKeyset, adLockOptimistic
TBGravar!DataFinal = msk_fltFim.Value
TBGravar!SaldoFinal = Saldo_Anterior
TBGravar.Update
TBGravar.Close
   
FormulaRel_Instituicao1 = "{tbl_Fluxo_de_caixa.Instituicao} = '" & txtDescricao & "' and {tbl_Fluxo_de_caixa.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {tbl_Fluxo_de_caixa.bloqueado} = False and {tbl_Fluxo_de_caixa.Data}>=Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {tbl_Fluxo_de_caixa.Data}<= Date(" & _
                                    Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ") and ({tbl_Fluxo_de_caixa.Operacao} = 'Crédito' or {tbl_Fluxo_de_caixa.Operacao} = 'Débito')"
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcSalvarExtrato()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
If txtDtValidacao = "" Then
    MsgBox "Não é possivel alterar o extrato, pois a instituição ainda não foi validada.", vbExclamation
    Exit Sub
End If
If txtStatus = "Bloqueada" Then
    MsgBox "Não é possivel alterar o extrato, pois a instituição esta bloqueada.", vbExclamation
    Exit Sub
End If
Acao = "salvar"
If TxtHistoricoExtrato = "" Then
    NomeCampo = "o histórico do lançamento"
    ProcVerificaAcao
    TxtHistoricoExtrato.SetFocus
    Exit Sub
End If
If Lst_extrato.ListItems.Count > 0 And Lst_extrato.SelectedItem <> "" Then
    Set TBFluxo = CreateObject("adodb.recordset")
    TBFluxo.Open "Select * from Tbl_Fluxo_de_Caixa where IDFluxo = " & Lst_extrato.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBFluxo.EOF = False Then
        TBFluxo!Obs = TxtHistoricoExtrato.Text
        TBFluxo.Update
        MsgBox ("Alteração efetuada com sucesso."), vbInformation
        '==================================
        Modulo = "Financeiro/Instituições"
        Evento = "Alterar histórico do lançamento"
        ID_documento = Lst_extrato.SelectedItem
        Documento = "Instituição bancária: " & txtDescricao
        Documento1 = "ID do lançamento: & " & Lst_extrato.SelectedItem & " - Data do lançamento: " & Lst_extrato.SelectedItem.ListSubItems(1)
        ProcGravaEvento
        '==================================
        ProcFiltrarExtrato
    End If
    TBFluxo.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcProximo()
On Error GoTo tratar_erro

If txtCodBanco = "" Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Instituicoes order by txt_Descricao", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.BOF = False Then
    TBAbrir.Find ("Id = " & txtCodBanco)
    TBAbrir.MoveNext
    If TBAbrir.EOF = False Then
        txtCodBanco = TBAbrir!ID
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from tbl_Instituicoes where Id = " & txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
        ProcLimpaCampos
        ProcCarregaDados
    Else
        MsgBox ("Fim dos cadastros de instituições bancária."), vbInformation
    End If
End If
Novo_Banco1 = False
Novo_Banco2 = False
Novo_Banco3 = False

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_forma_Click()
On Error GoTo tratar_erro

If Cmb_operacao = "" Then Exit Sub
Financeiro_Contas_Pagar = False
Financeiro_Forma_Pgto_Pagar = False
Financeiro_Contas_Receber = False
Financeiro_Forma_Pgto_Receber = False
frmContas_Forma_Pagamento.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_localizar_PC_Click()
On Error GoTo tratar_erro

ProcAtualizaVariaveisCC
Sit_REG = 3
frmproj_produto_PC.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_localizar_PC_instituicao_Click()
On Error GoTo tratar_erro

ProcAtualizaVariaveisCC
Sit_REG = 1
frmproj_produto_PC.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_localizar_PC_instituicao_rec_Click()
On Error GoTo tratar_erro

ProcAtualizaVariaveisCC
Sit_REG = 2
frmproj_produto_PC.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaVariaveisCC()
On Error GoTo tratar_erro

Plano_contas_produtos = False
Plano_contas_familias = False
Plano_centro_de_custo = False
Plano_instituicao = True
Plano_opcoesgerais = False
Plano_Faturamento = False
Financeiro_Contas_Pagar = False
Financeiro_Contas_Pagas = False
Financeiro_Contas_Receber = False
Financeiro_Contas_Recebidas = False
Plano_PCP = False

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_localizar_tipo_dcto_Click()
On Error GoTo tratar_erro

If Cmb_operacao = "" Then Exit Sub
Financeiro_Contas_Pagar = False
Financeiro_Contas_Receber = False
Clientes = False
Compras_Fornecedores = False
frmContas_Tipo_Dcto.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_PC_Click()
On Error GoTo tratar_erro

frmFinanceiro_familia.Show

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_PC_instituicao_Click()
On Error GoTo tratar_erro

frmFinanceiro_familia.Show

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_PC_instituicao_rec_Click()
On Error GoTo tratar_erro

frmFinanceiro_familia.Show

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub CmdAprocessar_Click()
On Error GoTo tratar_erro
Dataini = DTINI.Value
Datafim = DTFim.Value
'----------------
If txtDescricao.Text <> "" And cmbCliente.Text <> "" Then

    StrSql = "SELECT TOP (100) PERCENT dbo.tbl_Detalhes_Recebimento.Enviado,dbo.tbl_Detalhes_Recebimento.Data_envio,dbo.tbl_Detalhes_Recebimento.seq_remessa, dbo.tbl_Detalhes_Recebimento.IDContaReceber,dbo.tbl_Detalhes_Recebimento.txt_Cond_Recebimento, dbo.tbl_Detalhes_Recebimento.Id," _
    & "dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco,dbo.tbl_Detalhes_Recebimento.dt_Vencimento," _
    & "dbo.tbl_Detalhes_Recebimento.txt_tipoPagto, dbo.tbl_Detalhes_Recebimento.dbl_Valor," _
    & "dbo.tbl_Detalhes_Recebimento.int_NotaFiscal,dbo.tbl_Detalhes_Recebimento.txt_parcela, dbo.tbl_Detalhes_Recebimento.Nosso_numero, dbo.tbl_Detalhes_Recebimento.Carteira, dbo.tbl_Detalhes_Recebimento.Data_emissao,dbo.tbl_contas_receber.Nome_Razao FROM dbo.tbl_Detalhes_Recebimento" _
    & " INNER JOIN dbo.tbl_contas_receber ON dbo.tbl_Detalhes_Recebimento.IDContaReceber = dbo.tbl_contas_receber.IDIntconta" _
    & " WHERE (dbo.tbl_contas_receber.Nome_Razao = '" & cmbCliente & "') AND (dbo.tbl_Detalhes_Recebimento.txt_tipoPagto = N'BOLETO') AND (dbo.tbl_Detalhes_Recebimento.Nosso_numero IS NULL) AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento >= '" & Dataini & "') AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento <= '" & Datafim & "') AND (dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco = '" & txtDescricao.Text & "') ORDER BY dbo.tbl_Detalhes_Recebimento.dt_Vencimento"
    ProcCarregaListaDuplicatas
Else
    StrSql = "SELECT TOP (100) PERCENT dbo.tbl_Detalhes_Recebimento.Enviado,dbo.tbl_Detalhes_Recebimento.Data_envio,dbo.tbl_Detalhes_Recebimento.seq_remessa, dbo.tbl_Detalhes_Recebimento.IDContaReceber,dbo.tbl_Detalhes_Recebimento.txt_Cond_Recebimento, dbo.tbl_Detalhes_Recebimento.Id," _
    & "dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco,dbo.tbl_Detalhes_Recebimento.dt_Vencimento," _
    & "dbo.tbl_Detalhes_Recebimento.txt_tipoPagto, dbo.tbl_Detalhes_Recebimento.dbl_Valor," _
    & "dbo.tbl_Detalhes_Recebimento.int_NotaFiscal,dbo.tbl_Detalhes_Recebimento.txt_parcela, dbo.tbl_Detalhes_Recebimento.Nosso_numero, dbo.tbl_Detalhes_Recebimento.Carteira, dbo.tbl_Detalhes_Recebimento.Data_emissao,dbo.tbl_contas_receber.Nome_Razao FROM dbo.tbl_Detalhes_Recebimento" _
    & " INNER JOIN dbo.tbl_contas_receber ON dbo.tbl_Detalhes_Recebimento.IDContaReceber = dbo.tbl_contas_receber.IDIntconta" _
    & " WHERE (dbo.tbl_Detalhes_Recebimento.txt_tipoPagto = N'BOLETO') AND (dbo.tbl_Detalhes_Recebimento.Nosso_numero IS NULL) AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento >= '" & Dataini & "') AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento <= '" & Datafim & "') AND (dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco = '" & txtDescricao.Text & "') ORDER BY dbo.tbl_Detalhes_Recebimento.dt_Vencimento"
    ProcCarregaListaDuplicatas
End If

chkRemessa.Visible = True

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub CmdProcessados_Click()
On Error GoTo tratar_erro
'---------------------------------
If txtDescricao <> "" And cmbCliente.Text <> "" Then
    StrSql = "SELECT TOP (100) PERCENT dbo.tbl_Detalhes_Recebimento.Enviado,dbo.tbl_Detalhes_Recebimento.Data_envio,dbo.tbl_Detalhes_Recebimento.seq_remessa,dbo.tbl_Detalhes_Recebimento.IDContaReceber,dbo.tbl_Detalhes_Recebimento.txt_Cond_Recebimento, dbo.tbl_Detalhes_Recebimento.Id," _
    & "dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco,dbo.tbl_Detalhes_Recebimento.dt_Vencimento," _
    & "dbo.tbl_Detalhes_Recebimento.txt_tipoPagto, dbo.tbl_Detalhes_Recebimento.dbl_Valor," _
    & "dbo.tbl_Detalhes_Recebimento.int_NotaFiscal,dbo.tbl_Detalhes_Recebimento.txt_parcela, dbo.tbl_Detalhes_Recebimento.Nosso_numero, dbo.tbl_Detalhes_Recebimento.Carteira, dbo.tbl_Detalhes_Recebimento.Data_emissao,dbo.tbl_contas_receber.Nome_Razao,dbo.tbl_contas_receber.Vencimento FROM dbo.tbl_Detalhes_Recebimento" _
    & " INNER JOIN dbo.tbl_contas_receber ON dbo.tbl_Detalhes_Recebimento.IDContaReceber = dbo.tbl_contas_receber.IDIntconta" _
    & " WHERE (dbo.tbl_contas_receber.Nome_Razao = '" & cmbCliente & "') AND (dbo.tbl_Detalhes_Recebimento.txt_tipoPagto = N'BOLETO') AND  (NOT(dbo.tbl_Detalhes_Recebimento.Nosso_numero IS NULL)) AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento >= '" & DTINI & "') AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento <= '" & DTFim & "') AND (dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco = '" & txtDescricao.Text & "') ORDER BY dbo.tbl_Detalhes_Recebimento.dt_Vencimento"
    ProcCarregaListaDuplicatas
Else

    StrSql = "SELECT TOP (100) PERCENT dbo.tbl_Detalhes_Recebimento.Enviado,dbo.tbl_Detalhes_Recebimento.Data_envio,dbo.tbl_Detalhes_Recebimento.seq_remessa,dbo.tbl_Detalhes_Recebimento.IDContaReceber,dbo.tbl_Detalhes_Recebimento.txt_Cond_Recebimento, dbo.tbl_Detalhes_Recebimento.Id," _
    & "dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco,dbo.tbl_Detalhes_Recebimento.dt_Vencimento," _
    & "dbo.tbl_Detalhes_Recebimento.txt_tipoPagto, dbo.tbl_Detalhes_Recebimento.dbl_Valor," _
    & "dbo.tbl_Detalhes_Recebimento.int_NotaFiscal,dbo.tbl_Detalhes_Recebimento.txt_parcela, dbo.tbl_Detalhes_Recebimento.Nosso_numero, dbo.tbl_Detalhes_Recebimento.Carteira, dbo.tbl_Detalhes_Recebimento.Data_emissao,dbo.tbl_contas_receber.Nome_Razao,dbo.tbl_contas_receber.Vencimento FROM dbo.tbl_Detalhes_Recebimento" _
    & " INNER JOIN dbo.tbl_contas_receber ON dbo.tbl_Detalhes_Recebimento.IDContaReceber = dbo.tbl_contas_receber.IDIntconta" _
    & " WHERE (dbo.tbl_Detalhes_Recebimento.txt_tipoPagto = N'BOLETO') AND  (NOT(dbo.tbl_Detalhes_Recebimento.Nosso_numero IS NULL)) AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento >= '" & DTINI & "') AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento <= '" & DTFim & "') AND (dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco = '" & txtDescricao.Text & "') ORDER BY dbo.tbl_Detalhes_Recebimento.dt_Vencimento"
    ProcCarregaListaDuplicatas
End If

chkRemessa.Visible = False
chkRemessa.Value = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdSalvarInstrucoes_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente salvar as instruções do boleto?", vbYesNo, "CAPRIND V5.0") = vbYes Then
    Set TBBoleto = CreateObject("adodb.recordset")
         TBBoleto.Open "Select * from tbl_Instituicoes_Instrucoes_Boleto where ID_Instituicao = " & Txt_IDBanco & "", Conexao, adOpenKeyset, adLockOptimistic
         If TBBoleto.EOF = True Then TBBoleto.AddNew
         TBBoleto!ID_instituicao = Txt_IDBanco
         TBBoleto!Juros = Txtpercentual_juros
         TBBoleto!Desconto = Txtpercentual_desconto
         TBBoleto!Multa = Txtpercentual_multa
         TBBoleto!dias_protesto = Txtdias_protesto
         TBBoleto!Instrucoes_protesto = Txtinstrucoes
         TBBoleto!AssuntoEmail = txtAssunto
         TBBoleto.Update
         TBBoleto.Close
USMsgBox "Dados salvos com sucesso!", vbInformation, "CAPRIND V5.0"
End If
ProcCarregaInstrucoesBoleto

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case SSTab1.Tab
    Case 0:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo
            Case vbKeyF2: ProcFiltrar
            Case vbKeyF3: ProcSalvar
            Case vbKeyF4: If cmb_Opcao_Lista_Instituicao = "Excluir" Then ProcExcluir
            Case vbKeyF7: If cmb_Opcao_Lista_Instituicao = "Status" Then ProcStatus
            Case vbKeyF8: If cmb_Opcao_Lista_Instituicao = "Validação" Then ProcValidarRegistros lst_Instituicoes, "Financeiro/Instituições"
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoMovimentacao
            Case vbKeyF2: ProcLocalizarMovimentacao
            Case vbKeyF3: ProcSalvarMovimentacao
            Case vbKeyF4: ProcExcluirMovimentacao
            Case vbKeyF5: ProcImprimirMovimentacao
            Case vbKeyF7: ProcCopiarTarifa
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 2:
        Select Case KeyCode
            Case vbKeyF2: ProcFiltrarExtrato
            Case vbKeyF3: ProcSalvarExtrato
            Case vbKeyF5: ProcImprimirExtrato
            Case vbKeyF7: ProcVisualizarContas
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 3:
        Select Case KeyCode
            Case vbKeyF2: ProcFiltrarChequeEmitido
            Case vbKeyF3: ProcSalvarChequeEmitido
            Case vbKeyF4: If Cmb_opcao_lista = "Excluir/cancelar" Then ProcExcluirChequeEmitido
            Case vbKeyF5: ProcImprimirChequeEmitido
            Case vbKeyF6: ProcCopiaChequeEmitido
            Case vbKeyF7: If Cmb_opcao_lista = "Compensar" Then ProcCompensarChequeEmitido
            Case vbKeyF8: If Cmb_opcao_lista = "Cancelar compensação" Then ProcCancelarCompensacaoChequeEmitido
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    Case 4:
        Select Case KeyCode
            Case vbKeyF2: ProcFiltrarChequeRecebido
            Case vbKeyF4:  If Cmb_opcao_lista_recebidos = "Excluir" Then ProcExcluirChequeRecebido
            Case vbKeyF7: If Cmb_opcao_lista_recebidos = "Compensar" Then ProcCompensarChequeRecebido
            Case vbKeyF8: If Cmb_opcao_lista_recebidos = "Cancelar compensação" Then ProcCancelarCompensacaoChequeRecebido
            Case vbKeyF1: ProcAjuda
            Case vbKeyEscape: ProcSair
        End Select
    
End Select
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 12, True
ProcCarregaToolBar2 Me, 15192, 13, True
ProcCarregaToolBar3 Me, 15192, 11, True
ProcCarregaToolBar4 Me, 15192, 13, True
ProcCarregaToolBar5 Me, 15192, 10, True

cmb_Opcao_Lista_Instituicao = "Validação"

With USToolBar2
    .ButtonState(8) = 5
    .Refresh
End With

Formulario = "Financeiro/Instituições"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
SSTab3.Tab = 0
SSTab2.Tab = 0
cmbFamilia.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select txt_familia from tbl_instituicoes where txt_familia <> 'Null' group by Txt_familia", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        cmbFamilia.AddItem TBLISTA!Txt_familia
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close

txtdata.Value = Date
txtdata2.Value = Date
txtdata3.Value = Date
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
ProcCarregaComboEmpresa Cmb_empresa, False
Cmb_opcao_lista = "Compensar"
Cmb_opcao_lista_recebidos = "Compensar"

StrSql_Instituicoes_Localizar = "Select I.ID, I.ID_empresa, E.Empresa, I.int_NBanco, I.txt_Agencia, I.txt_conta, I.Txt_descricao, I.DtValidacao from tbl_instituicoes I INNER JOIN Empresa E ON E.Codigo = I.ID_empresa where I.Bloqueado = 0 order by I.txt_Descricao"
ProcCarregaLista

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcCarregaComboForma()
On Error GoTo tratar_erro

If Cmb_operacao = "Débito" Then ProcCarregaComboFormaPgtoRcbto cmb_forma1, "Tipo = 'P'" Else ProcCarregaComboFormaPgtoRcbto cmb_forma1, "Tipo = 'R'"
If Txt_id_tarifa <> "" Then ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcCarregaTipoDocumento()
On Error GoTo tratar_erro

If Cmb_operacao = "Débito" Then ProcCarregaComboTipoDocto Cmb_tipo, "Tipo = 'P'" Else ProcCarregaComboTipoDocto Cmb_tipo, "Tipo = 'R'"
If Txt_id_tarifa <> "" Then ProcCarregaCamposCombo

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCarregaCamposCombo()
On Error GoTo tratar_erro

cmb_forma1.ListIndex = -1
Cmb_tipo.ListIndex = -1
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select IDintconta, Tipo from tbl_instituicoes_transf where id_transf = " & IIf(Txt_id_tarifa = "", 0, Txt_id_tarifa), Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Set TBTempo = CreateObject("adodb.recordset")
    If TBFI!Tipo = "P" Then
        TBTempo.Open "Select FormaBaixa, Class_conta from tbl_ContasPagar where IdIntConta = " & TBFI!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
        If TBTempo.EOF = False Then
            If IsNull(TBTempo!FormaBaixa) = False And TBTempo!FormaBaixa <> "" Then cmb_forma1 = TBTempo!FormaBaixa
            If IsNull(TBTempo!Class_conta) = False And TBTempo!Class_conta <> "" Then Cmb_tipo = TBTempo!Class_conta
        End If
    Else
        TBTempo.Open "Select FormaBaixa, Tipo_doc from tbl_Contas_receber where IdIntConta = " & TBFI!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
        If TBTempo.EOF = False Then
            If IsNull(TBTempo!FormaBaixa) = False And TBTempo!FormaBaixa <> "" Then cmb_forma1 = TBTempo!FormaBaixa
            If IsNull(TBTempo!Tipo_doc) = False And TBTempo!Tipo_doc <> "" Then Cmb_tipo = TBTempo!Tipo_doc
        End If
    End If
End If

1:
    TBFI.Close

Exit Sub
tratar_erro:
    If Err.Number = "383" Then GoTo 1
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcCarregaLista()
On Error GoTo tratar_erro

lst_Instituicoes.ListItems.Clear
If StrSql_Instituicoes_Localizar = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSql_Instituicoes_Localizar, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista(0).Min = 0
    PBLista(0).Max = TBLISTA.RecordCount
    PBLista(0).Value = 1
    Contador = 0
    With lst_Instituicoes.ListItems
        Do While TBLISTA.EOF = False
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Empresa), "", TBLISTA!Empresa)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!int_NBanco), "", TBLISTA!int_NBanco)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!txt_Agencia), "", TBLISTA!txt_Agencia)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!txt_conta), "", TBLISTA!txt_conta)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Txt_descricao), "", TBLISTA!Txt_descricao)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!ID_empresa), 0, TBLISTA!ID_empresa)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!DtValidacao) Or TBLISTA!DtValidacao = "", "Não", "Sim")
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista(0).Value = Contador
        Loop
    End With
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Financeiro/Instituições"
Direitos
ProcLimpaVariaveisPrincipais
If txtCodBanco <> "" Then
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select * from tbl_Instituicoes where Id = " & txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        txtsaldo = IIf(IsNull(TBLISTA!Saldo), "0,00", Format(TBLISTA!Saldo, "###,##0.00"))
        txtLimite = IIf(IsNull(TBLISTA!Limite_desconto), "0,00", Format(TBLISTA!Limite_desconto, "###,##0.00"))
        txtUtilizado = IIf(IsNull(TBLISTA!Limite_utilizado), "0,00", Format(TBLISTA!Limite_utilizado, "###,##0.00"))
    End If
    TBLISTA.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizarMovimentacao()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362I" Then
    If MsgBox("Deseja realmente atualizar o saldo dos saques?", vbQuestion + vbYesNo) = vbYes Then
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from tbl_instituicoes_transf where Tipo = 'S' order by data_transf", Conexao, adOpenKeyset, adLockOptimistic
        If TBLISTA.EOF = False Then
            PBLista(2).Min = 0
            PBLista(2).Max = TBLISTA.RecordCount
            PBLista(2).Value = 1
            Contador = 0
            Do While TBLISTA.EOF = False
                'Verifica se o saldo do saque é maior que zero
                Valor_total = 0
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Sum(Valor_utilizado) as Valor_Total from tbl_ContasPagar_Saque where IDSaque = " & TBLISTA!id_transf, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Valor_total = IIf(IsNull(TBAbrir!Valor_total), 0, TBAbrir!Valor_total)
                End If
                TBAbrir.Close
                
                TBLISTA!Saldo = TBLISTA!valor_transf - Valor_total
                
                TBLISTA.MoveNext
                Contador = Contador + 1
                PBLista(2).Value = Contador
            Loop
        End If
        TBLISTA.Close
        MsgBox ("Atualização efetuada com sucesso."), vbInformation
        '==================================
        Modulo = "Financeiro/Instituições"
        Evento = "Atualizar1"
        ID_documento = 0
        Documento = ""
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362I1" Then
    If MsgBox("Deseja realmente atualizar o limite utilizado para desconto e a instituição bancária utilizada nas contas?", vbQuestion + vbYesNo) = vbYes Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from tbl_Instituicoes order by txt_Descricao", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            PBLista(0).Min = 0
            PBLista(0).Max = TBAbrir.RecordCount
            PBLista(0).Value = 1
            Contador = 0
            Do While TBAbrir.EOF = False
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select Sum(tbl_contas_receber.Valor) as Valor from tbl_contas_receber INNER JOIN troca_titulo on tbl_contas_receber.Idtrocatitulo = troca_titulo.ID where troca_titulo.Local_troca = '" & TBAbrir!Txt_descricao & "' and tbl_contas_receber.ID_empresa = " & TBAbrir!ID_empresa & " and tbl_contas_receber.status = 'DUPLICATA DESCONTADA EM ABERTO' and tbl_contas_receber.Logsit = 'N'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    valor = IIf(IsNull(TBFI!valor), 0, TBFI!valor)
                End If
                TBFI.Close
                
                TBAbrir!Limite_utilizado = valor
                TBAbrir.Update
                TBAbrir.MoveNext
                Contador = Contador + 1
                PBLista(0).Value = Contador
            Loop
        End If
        TBAbrir.Close
        
        'Atualiza bancos nas contas, transferências e saques
        'Recebidas
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select * from tbl_contas_receber where LogSit = 'S' order by IDFluxo", Conexao, adOpenKeyset, adLockOptimistic
        If TBContas.EOF = False Then
            Do While TBContas.EOF = False
                Set TBFluxo = CreateObject("adodb.recordset")
                TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & TBContas!IDFluxo, Conexao, adOpenKeyset, adLockOptimistic
                If TBFluxo.EOF = False Then
                    TBContas!Banco = TBFluxo!Instituicao
                    TBContas.Update
                End If
                TBFluxo.Close
                TBContas.MoveNext
            Loop
        End If
        TBContas.Close
        
        'Pagas
        Set TBContas = CreateObject("adodb.recordset")
         TBContas.Open "Select * from tbl_ContasPagar where LogSit = 'S' order by IDFluxo", Conexao, adOpenKeyset, adLockOptimistic
        If TBContas.EOF = False Then
            Do While TBContas.EOF = False
                Set TBFluxo = CreateObject("adodb.recordset")
                TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & TBContas!IDFluxo, Conexao, adOpenKeyset, adLockOptimistic
                If TBFluxo.EOF = False Then
                    TBContas!Banco = TBFluxo!Instituicao
                    TBContas.Update
                End If
                TBFluxo.Close
                TBContas.MoveNext
            Loop
        End If
        TBContas.Close
        
        'Transferencias e depositos
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select * from tbl_instituicoes_transf where Tipo <> 'S' order by id_transf", Conexao, adOpenKeyset, adLockOptimistic
        If TBContas.EOF = False Then
            Do While TBContas.EOF = False
                Select Case TBContas!FormaBaixa
                    Case "CHEQUE": Texto = "Cheque n. " & TBContas!NDoctoBaixa
                    Case "DOC": Texto = "Doc n. " & TBContas!NDoctoBaixa
                    Case "TED": Texto = "Ted n. " & TBContas!NDoctoBaixa
                    Case "TEV": Texto = "Tev n. " & TBContas!NDoctoBaixa
                End Select
                Set TBFluxo = CreateObject("adodb.recordset")
                TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where idintconta = " & TBContas!id_transf & " and Descricao = '" & Texto & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFluxo.EOF = False Then
                    If TBFluxo!Operacao = "Crédito" Then TBContas!banco_recebedor = TBFluxo!Instituicao Else TBContas!banco_remetente = TBFluxo!Instituicao
                    TBContas.Update
                End If
                TBFluxo.Close
                TBContas.MoveNext
            Loop
        End If
        TBContas.Close
        
        'Saques
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select * from tbl_instituicoes_transf where Tipo = 'S' order by IDFluxo", Conexao, adOpenKeyset, adLockOptimistic
        If TBContas.EOF = False Then
            Do While TBContas.EOF = False
                Set TBFluxo = CreateObject("adodb.recordset")
                TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & TBContas!IDFluxo, Conexao, adOpenKeyset, adLockOptimistic
                If TBFluxo.EOF = False Then
                    TBContas!banco_remetente = TBFluxo!Instituicao
                    TBContas.Update
                End If
                TBFluxo.Close
                TBContas.MoveNext
            Loop
        End If
        TBContas.Close
        
        MsgBox ("Atualização efetuada com sucesso."), vbInformation
        '==================================
        Modulo = "Financeiro/Instituições"
        Evento = "Atualizar"
        ID_documento = 0
        Documento = ""
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizarExtrato()
On Error GoTo tratar_erro

If InputBox("Informe a senha para liberar.") = "280362I2" Then
    If MsgBox("Deseja realmente atualizar os históricos de lançamentos?", vbQuestion + vbYesNo) = vbYes Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from tbl_Fluxo_de_caixa where status = 'S' order by IDFluxo", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            PBLista(4).Min = 0
            PBLista(4).Max = TBAbrir.RecordCount
            PBLista(4).Value = 1
            Contador = 0
            Do While TBAbrir.EOF = False
                TBAbrir!Obs = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
                TBAbrir.Update
                TBAbrir.MoveNext
                Contador = Contador + 1
                PBLista(4).Value = Contador
            Loop
        End If
        TBAbrir.Close
        MsgBox ("Atualização efetuada com sucesso."), vbInformation
        '==================================
        Modulo = "Financeiro/Instituições"
        Evento = "Atualizar2"
        ID_documento = 0
        Documento = ""
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With lst_Instituicoes
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If MsgBox("Deseja realmente excluir esta(s) instituição(ões) bancária(s)?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from tbl_Instituicoes WHERE ID = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Financeiro/Instituições"
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                Documento = "Instituição bancária: " & TBFI!Txt_descricao
                Documento1 = ""
                ProcGravaEvento
                '==================================
                Conexao.Execute "DELETE FROM tbl_Instituicoes WHERE ID = " & .ListItems(InitFor)
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    MsgBox ("Informe a(s) instituição(ões) bancária(s) antes de excluir."), vbExclamation
Else
    MsgBox ("Instituição(ões) bancária(s) excluída(s) com sucesso."), vbInformation
    ProcLimpaCampos
    ProcCarregaLista
    Frame2.Enabled = False
    ProcLimparTudo
    Novo_Banco = False
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcExcluirMovimentacao()
On Error GoTo tratar_erro

If Excluir = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
Select Case SSTab3.Tab
    Case 0:
        With lst_transferencias
            For InitFor = 1 To .ListItems.Count
                If .ListItems.Item(InitFor).Checked = True Then
                    If Permitido = False Then
                        If MsgBox("Deseja realmente excluir esta(s) movimentação(ões) financeira(s)?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
                    End If
                    Permitido = True
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from tbl_instituicoes_transf WHERE id_transf = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        '==================================
                        Modulo = "Financeiro/Instituições"
                        Evento = "Excluir movimentação financeira"
                        ID_documento = .ListItems(InitFor)
                        Documento = "Instituição bancária: " & txtDescricao
                        Documento1 = "Data: " & Format(TBFI!data_transf, "dd/mm/yy") & " - Valor: " & Format(TBFI!valor_transf, "###,##0.00")
                        ProcGravaEvento
                        '==================================
                        
                        'Exclui cheque criado na tabela de contas a pagar e receber
                        If TBFI!Tipo = "D" And TBFI!FormaBaixa = "CHEQUE" Then
                            Conexao.Execute "DELETE from tbl_ContasPagar where NDoctoBaixa = '" & TBFI!NDoctoBaixa & "' and Banco = '" & txtDescricao & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
                            Conexao.Execute "DELETE from tbl_Contas_receber where NDoctoBaixa = '" & TBFI!NDoctoBaixa & "' and Banco = '" & txtDescricao & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
                        End If
                                            
                        'Fluxo de Caixa
                        Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & TBFI!IDFluxo_Rec
                        Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & TBFI!IDFluxo
                        
                        'Conta contábil
                        Conexao.Execute "DELETE from familia_financeiro where idconta = " & .ListItems(InitFor) & " and Deposito_transf = 'True'"
                        
                        If IsNull(TBFI!banco_recebedor) = False And TBFI!banco_recebedor <> "" Then ProcAtualizaSaldosExcluir
                        Conexao.Execute "DELETE from tbl_instituicoes_transf where id_transf = " & .ListItems(InitFor)
                        
                    End If
                    TBFI.Close
                End If
            Next InitFor
        End With
        If Permitido = False Then
            MsgBox ("Informe a(s) movimentação(ões) financeira(s) antes de excluir."), vbExclamation
        Else
            MsgBox ("Movimentação(ões) financeira(s) excluída(s) com sucesso."), vbInformation
            ProcLimpaCamposTransf
            ProcCarregaListaTransf
            frm_filtro.Enabled = False
            Novo_Banco1 = False
        End If
    Case 1:
        With Lst_saque
            For InitFor = 1 To .ListItems.Count
                If .ListItems.Item(InitFor).Checked = True Then
                    If Permitido = False Then
                        If MsgBox("Deseja realmente excluir este(s) saque(s)?", vbQuestion + vbYesNo) = vbYes Then GoTo 2 Else Exit Sub
                    End If
2:
                    Permitido = True
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from tbl_instituicoes_transf WHERE id_transf = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        '==================================
                        Modulo = "Financeiro/Instituições"
                        Evento = "Excluir saque"
                        ID_documento = .ListItems(InitFor)
                        Documento = "Instituição bancária: " & txtDescricao
                        Documento1 = "Data: " & Format(TBFI!data_transf, "dd/mm/yy") & " - Valor: " & Format(TBFI!valor_transf, "###,##0.00")
                        ProcGravaEvento
                        '==================================
                        
                        'Verif. se existem contas paga com o saque
                        Set TBFIltro = CreateObject("adodb.recordset")
                        TBFIltro.Open "Select tbl_ContasPagar.* from tbl_ContasPagar inner join tbl_ContasPagar_Saque on tbl_ContasPagar.idintconta = tbl_ContasPagar_Saque.idintconta where tbl_ContasPagar_Saque.IDSaque = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBFIltro.EOF = False Then
                            Do While TBFIltro.EOF = False
                                'Verifica se a conta paga parcial já está liquidada
                                If IsNull(TBFIltro!tituloref) = True Or TBFIltro!tituloref = "" Then
                                    ReferenciaConta = 0
                                Else
                                    ReferenciaConta = TBFIltro!tituloref
                                End If
                                Set TBContas = CreateObject("adodb.recordset")
                                TBContas.Open "Select * from tbl_contaspagar where idintconta = " & ReferenciaConta & " and parcial = 'True' and tituloref <> '" & TBFIltro!IDintconta & "'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBContas.EOF = False Then
                                    ProcSomaRecompra
                                    Set TBCorretiva = CreateObject("adodb.recordset")
                                    TBCorretiva.Open "Select * from tbl_contaspagar where idintconta = " & TBFIltro!tituloref, Conexao, adOpenKeyset, adLockOptimistic
                                    If TBCorretiva.EOF = False Then
                                        ValorParcial = TBFIltro!ValorPago
                                        Pendente = TBCorretiva!dbl_valorpagto
                                        TBCorretiva!dbl_valorpagto = (Pendente + ValorParcial)
                                        
                                        Set TBAbrir = CreateObject("adodb.recordset")
                                        TBAbrir.Open "Select * from tbl_contaspagar where tituloref = '" & TBFIltro!tituloref & "' and idintconta <> " & TBFIltro!tituloref, Conexao, adOpenKeyset, adLockOptimistic
                                        If TBAbrir.EOF = False Then
                                            TBCorretiva!status = "TÍTULO PAGO PARCIAL"
                                        Else
                                            TBCorretiva!status = "TÍTULO EM ABERTO"
                                            TBCorretiva!Parcial = False
                                            TBCorretiva!pagoparcial = 0
                                            TBCorretiva!ValorPendente = 0
                                            TBCorretiva!tituloref = ""
                                            TBCorretiva!valorprincipal = 0
                                        End If
                                        TBAbrir.Close
                                          
                                        'Fluxo de Caixa
                                        Set TBFluxo = CreateObject("adodb.recordset")
                                        TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                                        If TBFluxo.EOF = True Then TBFluxo.AddNew
                                        TBFluxo!Operacao = "À Debitar"
                                        TBFluxo!data = TBCorretiva!dt_Pagamento
                                        TBFluxo!valor = TBCorretiva!dbl_valorpagto
                                        TBFluxo!Descricao = TBCorretiva!Txt_fornecedor
                                        TBFluxo!status = "N"
                                        TBFluxo!int_NotaFiscal = TBCorretiva!txt_ndocumento
                                        TBCorretiva!IDFluxo = TBFluxo!IDFluxo
                                        TBFluxo!Instituicao = Null
                                        TBFluxo!Hora = Null
                                        TBFluxo!Cheque = 0
                                        TBFluxo!Bloqueado = False
                                        TBFluxo.Update
                                        TBFluxo.Close
                                    End If
                                    TBCorretiva.Update
                                    TBCorretiva.Close
                                    
                                    'Exclui conta paga parcial/Fluxo de caixa
                                    Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & TBFIltro!IDFluxo
                                    Conexao.Execute "DELETE from tbl_contaspagar where IdIntConta = " & TBFIltro!IDintconta
                                    Conexao.Execute "DELETE from tbl_contas_antecipacao where ID_Conta = " & TBFIltro!IDintconta & " and tipo = 'P'"
                                Else
                                    ProcSomaRecompra
                                    Set TBCorretiva = CreateObject("adodb.recordset")
                                    TBCorretiva.Open "Select * from tbl_contaspagar where IdIntConta = " & TBFIltro!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
                                    If TBCorretiva.EOF = False Then
                                        Set TBAbrir = CreateObject("adodb.recordset")
                                        TBAbrir.Open "Select * from tbl_contaspagar where tituloref = '" & IIf(IsNull(TBFIltro!tituloref), 0, TBFIltro!tituloref) & "'", Conexao, adOpenKeyset, adLockOptimistic
                                        If TBAbrir.EOF = False Then
                                            TBCorretiva!status = "TÍTULO PAGO PARCIAL"
                                        Else
                                            TBCorretiva!status = "TÍTULO EM ABERTO"
                                            TBCorretiva!Parcial = False
                                            TBCorretiva!pagoparcial = 0
                                            TBCorretiva!ValorPendente = 0
                                            TBCorretiva!tituloref = ""
                                            TBCorretiva!valorprincipal = 0
                                        End If
                                        TBAbrir.Close
                                        
                                        Conexao.Execute "DELETE from tbl_contas_antecipacao where ID_Conta = " & TBFIltro!IDintconta & " and tipo = 'P'"
                                        
                                        'Fluxo de Caixa
                                        Set TBFluxo = CreateObject("adodb.recordset")
                                        TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where idfluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                                        If TBFluxo.EOF = True Then TBFluxo.AddNew
                                        TBFluxo!Operacao = "À Debitar"
                                        TBFluxo!data = TBCorretiva!dt_Pagamento
                                        TBFluxo!valor = TBCorretiva!dbl_valorpagto
                                        TBFluxo!Descricao = TBCorretiva!Txt_fornecedor
                                        TBFluxo!status = "N"
                                        TBFluxo!int_NotaFiscal = TBCorretiva!txt_ndocumento
                                        TBCorretiva!IDFluxo = TBFluxo!IDFluxo
                                        TBFluxo!Instituicao = Null
                                        TBFluxo!Hora = Null
                                        TBFluxo!Cheque = 0
                                        TBFluxo!Bloqueado = False
                                        TBFluxo.Update
                                        TBFluxo.Close
                                    
                                        TBCorretiva!Logsit = "N"
                                        TBCorretiva!DataBaixa = Null
                                        TBCorretiva!Data_movimentacao = Null
                                        TBCorretiva!Bom_para = Null
                                        TBCorretiva!ValorPago = 0
                                        TBCorretiva!NDoctoBaixa = ""
                                        TBCorretiva!Banco = ""
                                        TBCorretiva!Obs = ""
                                        TBCorretiva!Favorecido = ""
                                        TBCorretiva!Obscheque = ""
                                        TBCorretiva!Dias_atraso = 0
                                        TBCorretiva!Juros = 0
                                        TBCorretiva!Juros_valor = 0
                                        TBCorretiva!Multa = 0
                                        TBCorretiva!Multa_valor = 0
                                        TBCorretiva!Desconto = 0
                                        TBCorretiva!Desconto_valor = 0
                                        TBCorretiva.Update
                                    End If
                                    TBCorretiva.Close
                                End If
                                TBContas.Close
                                TBFIltro.MoveNext
                            Loop
                        End If
                        TBFIltro.Close
                        
                        Set TBSaldo = CreateObject("adodb.recordset")
                        TBSaldo.Open "Select saldo from tbl_Instituicoes where ID = " & txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
                        If TBSaldo.EOF = False Then
                            TBSaldo!Saldo = TBSaldo!Saldo + TBFI!valor_transf
                            txtsaldo = Format(TBSaldo!Saldo, "###,##0.00")
                            TBSaldo.Update
                        End If
                        TBSaldo.Close
                        
                        'Exclui contas relacionadas com o saque, saque do banco e do fluxo de caixa
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from tbl_instituicoes_transf where id_transf = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                        
                            'Contas relacionadas
                            Conexao.Execute "DELETE from tbl_ContasPagar_Saque where IDSaque = " & TBAbrir!id_transf
                            
                            'Fluxo de Caixa
                            Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & TBAbrir!IDFluxo
                            
                            TBAbrir.Delete
                        End If
                        TBAbrir.Close
                    End If
                    TBFI.Close
                End If
            Next InitFor
        End With
        If Permitido = False Then
            MsgBox ("Informe o(s) saque(s) antes de excluir."), vbExclamation
        Else
            MsgBox ("Saque(s) excluído(s) com sucesso."), vbInformation
            ProcLimpaCamposSaque
            Lst_Contas.ListItems.Clear
            LblValortotal.Caption = "Valor total pago = 0,00"
            ProcCarregaListaSaque
            Frame8.Enabled = False
            Novo_Banco2 = False
        End If
    Case 2:
        With Lst_tarifa
            For InitFor = 1 To .ListItems.Count
                If .ListItems.Item(InitFor).Checked = True Then
                    If Permitido = False Then
                        If MsgBox("Deseja realmente excluir esta(s) tarifa(s)?", vbQuestion + vbYesNo) = vbYes Then GoTo 3 Else Exit Sub
                    End If
3:
                    Permitido = True
                    
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from tbl_instituicoes_transf where id_transf = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        If TBFI!Tipo = "P" Then
                            ProcExcluirTarifaPag TBFI!IDintconta
                            OperacaoTexto = "Débito"
                        Else
                            ProcExcluirTarifaRec TBFI!IDintconta
                            OperacaoTexto = "Crédito"
                        End If
                        '==================================
                        Modulo = "Financeiro/Instituições"
                        Evento = "Excluir tarifa"
                        ID_documento = .ListItems(InitFor)
                        Documento = "Instituição bancária: " & txtDescricao
                        Documento1 = "ID da conta: " & TBFI!IDintconta & " - Data: " & Format(TBFI!data_transf, "dd/mm/yy") & " - Operação: " & OperacaoTexto & " - Valor: " & Format(TBFI!valor_transf, "###,##0.00")
                        ProcGravaEvento
                        '==================================
                        
                        TBFI.Delete
                    End If
                    TBFI.Close
                End If
            Next InitFor
        End With
        If Permitido = False Then
            MsgBox ("Informe a(s) tarifa(s) antes de excluir."), vbExclamation
        Else
            MsgBox ("Tarifa(s) excluída(s) com sucesso."), vbInformation
            ProcLimpaCamposTarifa
            Lst_tarifa.ListItems.Clear
            ProcCarregaListaTarifa
            Frame4.Enabled = False
            Novo_Banco3 = False
        End If
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcExcluirTarifaPag(IDintconta As Long)
On Error GoTo tratar_erro

Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * from tbl_ContasPagar where IdIntConta = " & IDintconta, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    Set TBSaldo = CreateObject("adodb.recordset")
    TBSaldo.Open "Select saldo from tbl_Instituicoes where ID = " & txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
    If TBSaldo.EOF = False Then
        TBSaldo!Saldo = TBSaldo!Saldo + TBContas!ValorPago
        txtsaldo = Format(TBSaldo!Saldo, "###,##0.00")
        TBSaldo.Update
    End If
    TBSaldo.Close
    
    Conexao.Execute "DELETE from familia_financeiro where idconta = " & IDintconta & " and tipoconta = 'P' and Deposito_transf = 'False'"
    Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & TBContas!IDFluxo
    
    TBContas.Delete
End If
TBContas.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcExcluirTarifaRec(IDintconta As Long)
On Error GoTo tratar_erro

Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * from tbl_contas_receber where IdIntConta = " & IDintconta, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    Set TBSaldo = CreateObject("adodb.recordset")
    TBSaldo.Open "Select saldo from tbl_Instituicoes where ID = " & txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
    If TBSaldo.EOF = False Then
        TBSaldo!Saldo = TBSaldo!Saldo - TBContas!valortitulorecebido
        txtsaldo = Format(TBSaldo!Saldo, "###,##0.00")
        TBSaldo.Update
    End If
    TBSaldo.Close
    
    Conexao.Execute "DELETE from familia_financeiro where idconta = " & IDintconta & " and tipoconta = 'R' and Deposito_transf = 'False'"
    Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & TBContas!IDFluxo
    
    TBContas.Delete
End If
TBContas.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcSomaRecompra()
On Error GoTo tratar_erro

'Soma valor de recompra no bordero
Set TBProduto = CreateObject("adodb.recordset")
TBProduto.Open "Select troca_titulo_valores.IDduplicata, troca_titulo_valores.valor_enviado FROM troca_titulo_valores INNER JOIN tbl_ContasPagar ON troca_titulo_valores.n_conta = tbl_ContasPagar.idcontareceber where tbl_ContasPagar.IdIntConta = " & TBFIltro!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Vlrtotalrecompra from troca_titulo where id = " & TBProduto!IDduplicata, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        TBAbrir!Vlrtotalrecompra = TBAbrir!Vlrtotalrecompra + TBFIltro!ValorPago
        TBAbrir.Update
    End If
    TBAbrir.Close
End If
TBProduto.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaSaldosExcluir()
On Error GoTo tratar_erro

If TBFI!Tipo <> "D" Or TBFI!Tipo = "D" And TBFI!FormaBaixa <> "CHEQUE" Then
    Set TBSaldo = CreateObject("adodb.recordset")
    TBSaldo.Open "Select saldo from tbl_Instituicoes where ID = " & txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
    'Atualiza saldo do banco remetente
    If TBSaldo.EOF = False Then
        TBSaldo!Saldo = (TBSaldo!Saldo + TBFI!valor_transf)
        TBSaldo.Update
        txtsaldo = Format(TBSaldo!Saldo, "###,##0.00")
    End If
    
    'Atualiza saldo do banco recebedor
    Set TBSaldo = CreateObject("adodb.recordset")
    TBSaldo.Open "Select saldo from tbl_Instituicoes where ID = " & TBFI!id_banco_rec, Conexao, adOpenKeyset, adLockOptimistic
    If TBSaldo.EOF = False Then
        TBSaldo!Saldo = (TBSaldo!Saldo - TBFI!valor_transf)
        TBSaldo.Update
    End If
    TBSaldo.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcLimpaCamposTransf()
On Error GoTo tratar_erro

txtid = 0
txtdata.Value = Date
txtResponsavel1 = pubUsuario

OptDeposito.Value = True
Tipo = "D"
cmb_forma.Clear
cmb_forma.AddItem "Dinheiro"
cmb_forma.AddItem "CHEQUE"

OptTransferencia.Value = False
txtCheque = ""
cmbrecebedor.Clear
txtfavorecido = ""
mskvalor = ""
TxtHistDepTranf = ""
txtObsFluxo = ""
Txt_ID_PC_instituicao = 0
Txt_codigo_PC_instituicao = ""
Txt_descricao_PC_instituicao = ""
Txt_ID_PC_instituicao_rec = 0
Txt_codigo_PC_instituicao_rec = ""
Txt_descricao_PC_instituicao_rec = ""
CodigoLista1 = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcLimpaCamposSaque()
On Error GoTo tratar_erro

Txt_id_saque = 0
txtdata2.Value = Date
txtResponsavel2 = pubUsuario
txtObsFluxo1 = "Saque"
Txt_valor = ""
CodigoLista2 = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcLimpaCamposTarifa()
On Error GoTo tratar_erro

Txt_id_tarifa = 0
txtdata3.Value = Date
txtResponsavel3 = pubUsuario
Cmb_operacao.ListIndex = -1
Cmb_tipo.ListIndex = -1
cmb_forma1.ListIndex = -1
txtObsFluxo2 = "Tarifa"
Txt_ID_PC = 0
Txt_codigo_PC = ""
Txt_descricao_PC = ""
Txt_valor1 = ""
CodigoLista3 = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcExcluirChequeEmitido()
On Error GoTo tratar_erro

If Excluir = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
If SSTab2.Tab = 0 Then
    Permitido = False
    With Lst_cheque
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then Permitido = True
        Next InitFor
    End With
    If Permitido = False Then
        MsgBox ("Informe o(s) cheque(es) antes de cancelar."), vbExclamation
        Exit Sub
    End If
    frm_Instituicoes2_cancelar_cheque.Show 1
Else
    Permitido = False
    With Lst_cheque1
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                If Permitido = False Then
                    If MsgBox("Deseja realmente excluir este(s) cheque(s) cancelado(s)?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
                End If
                Permitido = True
                '==================================
                Modulo = "Financeiro/Instituições"
                Evento = "Excluir cheque emitido"
                ID_documento = .ListItems(InitFor)
                Documento = "Cheque nº: " & .ListItems(InitFor).ListSubItems(2) & " - Instituição bancária: " & txtDescricao
                Documento1 = ""
                ProcGravaEvento
                '==================================
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from tbl_ContasPagar where IdIntConta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Conexao.Execute "DELETE from Cheques_Cancelados where ID_conta = " & TBAbrir!IDintconta
                    TBAbrir.Delete
                End If
                TBAbrir.Close
            End If
        Next InitFor
    End With
    If Permitido = False Then
        MsgBox ("Informe o(s) cheque(s) cancelado(s) antes de excluir."), vbExclamation
    Else
        MsgBox ("Cheque(s) cancelado(s) excluído(s) com sucesso."), vbInformation
        ProcCarregaListaCheque
        Frame7.Enabled = False
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcExcluirChequeRecebido()
On Error GoTo tratar_erro

If Excluir = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If

Permitido = False
With Lista_cheque
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If MsgBox("Deseja realmente excluir este(s) cheque(s)?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
            Permitido = True
            '==================================
            Modulo = "Financeiro/Instituições"
            Evento = "Excluir cheque recebido"
            ID_documento = .ListItems(InitFor)
            Documento = "Cheque nº: " & .ListItems(InitFor).ListSubItems(2) & " - Instituição bancária: " & txtDescricao
            Documento1 = ""
            ProcGravaEvento
            '==================================
            
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select * from tbl_contas_receber where idintconta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = False Then
                If IsNull(TBFIltro!tituloref) = True Or TBFIltro!tituloref = "" Then tituloref = 0 Else tituloref = TBFIltro!tituloref
                
                'Verifica se a conta paga parcial já está liquidada
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from tbl_contas_receber where idintconta = " & tituloref & " and parcial = 'True' and tituloref <> '" & TBFIltro!IDintconta & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    Set TBCorretiva = CreateObject("adodb.recordset")
                    TBCorretiva.Open "Select * from tbl_contas_receber where idintconta = " & TBFIltro!tituloref, Conexao, adOpenKeyset, adLockOptimistic
                    If TBCorretiva.EOF = False Then
                        ValorParcial = TBFIltro!valortitulorecebido
                        Pendente = TBCorretiva!valor
                        TBCorretiva!valor = (Pendente + ValorParcial)
                        
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from tbl_contas_receber where tituloref = '" & TBFIltro!tituloref & "' and idintconta <> " & TBFIltro!tituloref & " and idintconta <> " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            If TBCorretiva!Bloqueado = False Then TBCorretiva!status = "TÍTULO RECEBIDO PARCIAL"
                        Else
                            If TBCorretiva!Bloqueado = False Then TBCorretiva!status = "TÍTULO EM ABERTO"
                            TBCorretiva!Parcial = False
                            TBCorretiva!RecebidoParcial = 0
                            TBCorretiva!ValorPendente = 0
                            TBCorretiva!tituloref = ""
                            TBCorretiva!valorprincipal = 0
                        End If
                        TBAbrir.Close
                        
                        'Fluxo de Caixa
                        Cheque = "Cheque n. " & .ListItems(InitFor).ListSubItems(2)
                        Set TBFluxo = CreateObject("adodb.recordset")
                        TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where Operacao = 'Crédito' and Instituicao = '" & txtDescricao & "' and Descricao = '" & Cheque & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFluxo.EOF = False Then
                            TBFluxo!valor = TBFluxo!valor - .ListItems(InitFor).ListSubItems(4)
                            TBFluxo.Update
                            If TBFluxo!valor = 0 Then TBFluxo.Delete
                        End If
                        TBFluxo.Close
                        
                        Set TBFluxo = CreateObject("adodb.recordset")
                        TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                        If TBFluxo.EOF = True Then TBFluxo.AddNew
                        TBFluxo!Operacao = "À Creditar"
                        TBFluxo!data = TBCorretiva!Vencimento
                        TBFluxo!valor = TBCorretiva!valor
                        TBFluxo!Descricao = TBCorretiva!Nome_Razao
                        TBFluxo!status = "N"
                        TBFluxo!int_NotaFiscal = TBCorretiva!txt_ndocumento
                        TBCorretiva!IDFluxo = TBFluxo!IDFluxo
                        TBFluxo!Instituicao = Null
                        TBFluxo!Hora = Null
                        TBFluxo!Cheque = 0
                        TBFluxo!Bloqueado = False
                        TBFluxo.Update
                        TBFluxo.Close
                    End If
                    TBCorretiva.Update
                    TBCorretiva.Close
                    
                    Set TBFamilia = CreateObject("adodb.recordset")
                    TBFamilia.Open "select * from familia_financeiro where idconta = " & tituloref & " and tipoconta = 'R' order by ID_PC", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFamilia.EOF = False Then
                        Do While TBFamilia.EOF = False
                            Set TBCiclo = CreateObject("adodb.recordset")
                            TBCiclo.Open "Select * from familia_financeiro where IDConta = " & .ListItems(InitFor) & " and ID_PC = " & TBFamilia!ID_PC & " and tipoconta = 'R'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBCiclo.EOF = False Then
                                TBFamilia!valor = TBFamilia!valor + ValorParcial
                                TBFamilia.Update
                                TBCiclo.Delete
                            End If
                            TBCiclo.Close
                            TBFamilia.MoveNext
                        Loop
                    End If
                    TBFamilia.Close
                    
                    Set TBCorretiva = CreateObject("adodb.recordset")
                    TBCorretiva.Open "Select * from tbl_contas_receber where idintconta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBCorretiva.EOF = False Then
                        'Fluxo de Caixa
                        Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo)
                    
                        TBCorretiva.Delete
                    End If
                    TBCorretiva.Close
                Else
                    Set TBCorretiva = CreateObject("adodb.recordset")
                    TBCorretiva.Open "Select * from tbl_contas_receber where idintconta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBCorretiva.EOF = False Then
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from tbl_contas_receber where tituloref = '" & tituloref & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            If TBCorretiva!Bloqueado = False Then TBCorretiva!status = "TÍTULO RECEBIDO PARCIAL"
                        Else
                            If TBCorretiva!Bloqueado = False Then TBCorretiva!status = "TÍTULO EM ABERTO"
                            TBCorretiva!Parcial = False
                            TBCorretiva!RecebidoParcial = 0
                            TBCorretiva!ValorPendente = 0
                            TBCorretiva!tituloref = ""
                            TBCorretiva!valorprincipal = 0
                        End If
                        TBAbrir.Close
                                       
                        'Fluxo de Caixa
                        Cheque = "Cheque n. " & .ListItems(InitFor).ListSubItems(2)
                        Set TBFluxo = CreateObject("adodb.recordset")
                        TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where Operacao = 'Crédito' and Instituicao = '" & txtDescricao & "' and Descricao = '" & Cheque & "'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFluxo.EOF = False Then
                            TBFluxo!valor = TBFluxo!valor - .ListItems(InitFor).ListSubItems(4)
                            TBFluxo.Update
                            If TBFluxo!valor = 0 Then TBFluxo.Delete
                        End If
                        TBFluxo.Close
                        
                        Set TBFluxo = CreateObject("adodb.recordset")
                        TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                        If TBFluxo.EOF = True Then TBFluxo.AddNew
                        TBFluxo!Operacao = "À Creditar"
                        TBFluxo!data = TBCorretiva!Vencimento
                        TBFluxo!valor = TBCorretiva!valor
                        TBFluxo!Descricao = TBCorretiva!Nome_Razao
                        TBFluxo!status = "N"
                        TBFluxo!int_NotaFiscal = TBCorretiva!txt_ndocumento
                        TBCorretiva!IDFluxo = TBFluxo!IDFluxo
                        TBFluxo!Instituicao = Null
                        TBFluxo!Hora = Null
                        TBFluxo!Cheque = 0
                        TBFluxo!Bloqueado = False
                        TBFluxo.Update
                        TBFluxo.Close
                                    
                        TBCorretiva!Logsit = "N"
                        TBCorretiva!Data_pagamento = Null
                        TBCorretiva!Data_movimentacao = Null
                        TBCorretiva!valortitulorecebido = 0
                        TBCorretiva!NDoctoBaixa = ""
                        TBCorretiva!Banco = ""
                        TBCorretiva!Obs = ""
                        TBCorretiva!Dias_atraso = 0
                        TBCorretiva!Juros = 0
                        TBCorretiva!Juros_valor = 0
                        TBCorretiva!Multa = 0
                        TBCorretiva!Multa_valor = 0
                        TBCorretiva!Desconto = 0
                        TBCorretiva!Desconto_valor = 0
                        TBCorretiva.Update
                        
                        Conexao.Execute "DELETE from familia_financeiro where IDconta = " & TBCorretiva!IDintconta & " and Pago_recebido = 'True' and tipoconta = 'R' and Deposito_transf = 'False'"
                        Conexao.Execute "Update familia_financeiro Set Pago_recebido = 'False' where idconta = " & TBCorretiva!IDintconta & " and tipoconta = 'R'"
                        
                    End If
                    TBCorretiva.Close
                End If
                TBFI.Close
            End If
            TBFIltro.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    MsgBox ("Informe o(s) cheque(s) antes de excluir."), vbExclamation
Else
    MsgBox ("Cheque(s) excluído(s) com sucesso."), vbInformation
    ProcCarregaListaCheque
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcImprimirMovimentacao()
On Error GoTo tratar_erro
  
If lst_transferencias.ListItems.Count = 0 And Lst_saque.ListItems.Count = 0 Then Exit Sub
NomeRel = "Instituicoes_Movimentacao_Financeira.rpt"
ProcImprimirRel FormulaRel_Instituicao, ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcImprimirExtrato()
On Error GoTo tratar_erro

If Lst_extrato.ListItems.Count = 0 Then Exit Sub
frm_Instituicoes2_extrato_menuimpressao.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcImprimirChequeEmitido()
On Error GoTo tratar_erro

If Lst_cheque.ListItems.Count = 0 And Lst_cheque1.ListItems.Count = 0 Then Exit Sub
frm_Instituicoes2_menu_impressao_cheque.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcLocalizarMovimentacao()
On Error GoTo tratar_erro

frm_filtrotransferencia.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcFiltrarChequeEmitido()
On Error GoTo tratar_erro

Cheques_Emitidos = True
Lista_cheque.ListItems.Clear
frm_Instituicoes2_localizar_cheque.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcFiltrarChequeRecebido()
On Error GoTo tratar_erro

Cheques_Emitidos = False
Lst_cheque.ListItems.Clear
Lst_cheque1.ListItems.Clear
frm_Instituicoes2_localizar_cheque.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If Incluir = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
ProcLimpaCampos
Novo_Banco = True
Frame2.Enabled = True
txtNBanco.SetFocus
ProcLimparTudo

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcLimparTudo()
On Error GoTo tratar_erro

frm_filtro.Enabled = False
Frame8.Enabled = False
Frame4.Enabled = False
Frame7.Enabled = False
ProcLimpaCamposTransf
ProcLimpaCamposSaque
ProcLimpaCamposTarifa
lst_transferencias.ListItems.Clear
Lst_saque.ListItems.Clear
Lst_tarifa.ListItems.Clear
Txt_valor_total_tarifas = "0,00"
Txt_valor_total_tarifas1 = "0,00"
Lst_Contas.ListItems.Clear
Lst_extrato.ListItems.Clear
Txt_favorecido = ""
txtobscheque = ""
Lst_cheque.ListItems.Clear
Lst_cheque1.ListItems.Clear
Lista_cheque.ListItems.Clear
Txt_qtde_ativo = 0
Txt_qtde_cancelado = 0
Txt_qtde_total = 0
Txt_valor_ativo = "0,00"
Txt_valor_cancelado = "0,00"
Txt_valor_total = "0,00"
Novo_Banco1 = False
Novo_Banco2 = False
Novo_Banco3 = False
Instituicao_Localizar_Transf = ""
Instituicao_Localizar_Saque = ""
Instituicao_Localizar_Tarifa = ""
StrSql_Instituicoes_Localizar_Cheque = ""
StrSql_Instituicoes_Localizar_Cheque_Cancelados = ""
StrSql_Instituicoes_Localizar_Cheque_Recebidos = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcNovoMovimentacao()
On Error GoTo tratar_erro

If Incluir = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
If txtDtValidacao = "" Then
    MsgBox "Não é possivel criar nova movimentação, pois a instituição ainda não foi validada.", vbExclamation
    Exit Sub
End If
If txtStatus = "Bloqueada" Then
    MsgBox "Não é possivel criar nova movimentação, pois a instituição esta bloqueada.", vbExclamation
    Exit Sub
End If
Select Case SSTab3.Tab
    Case 0:
        ProcLimpaCamposTransf
        frm_filtro.Enabled = True
        txtdata.SetFocus
        Novo_Banco1 = True
    Case 1:
        ProcLimpaCamposSaque
        Frame8.Enabled = True
        txtdata2.SetFocus
        Novo_Banco2 = True
    Case 2:
        ProcLimpaCamposTarifa
        Frame4.Enabled = True
        txtdata3.SetFocus
        Novo_Banco3 = True
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro
    
If Novo_Banco = True Then
    If MsgBox("A instituição bancária ainda não foi salva, deseja salvar antes de fechar o módulo?", vbYesNo + vbQuestion) = vbYes Then
        ProcSalvar
        If Novo_Banco = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
If Novo_Banco1 = True Or Novo_Banco2 = True Or Novo_Banco3 = True Then
    If Novo_Banco1 = True Then
        OperacaoTexto = "A movimentação financeira ainda não foi salva"
    ElseIf Novo_Banco2 = True Then
            OperacaoTexto = "O saque ainda não foi salvo"
        Else
            OperacaoTexto = "A tarifa ainda não foi salva"
    End If
    If MsgBox(OperacaoTexto & ", deseja salvar antes de fechar o módulo?", vbYesNo + vbQuestion) = vbYes Then
        ProcSalvarMovimentacao
        If Novo_Banco1 = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Banco = False
Novo_Banco1 = False
Novo_Banco2 = False
Novo_Banco3 = False
Conexao.Execute "DELETE from Cheques_Relatorios"
Unload Me

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
If Frame2.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If txtNBanco = "" Then
    NomeCampo = "o número do banco"
    ProcVerificaAcao
    txtNBanco.SetFocus
    Exit Sub
End If
If txtAgencia = "" Then
    NomeCampo = "a agencia"
    ProcVerificaAcao
    txtAgencia.SetFocus
    Exit Sub
End If
If txtConta = "" Then
    NomeCampo = "a conta"
    ProcVerificaAcao
    txtConta.SetFocus
    Exit Sub
End If

If txtDescricao = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    txtDescricao.SetFocus
    Exit Sub
Else
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select id from tbl_instituicoes where id <> " & IIf(txtCodBanco = "", 0, txtCodBanco) & " and txt_Descricao = '" & txtDescricao & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        MsgBox "Já existe esta instituição bancária cadastrada para esta empresa.", vbExclamation
        Exit Sub
    End If
    TBAbrir.Close
End If

If txtsaldo = "" Then
    NomeCampo = "o saldo"
    ProcVerificaAcao
    txtsaldo.SetFocus
    Exit Sub
End If
If txtLimite = "" Then
    NomeCampo = "o limite para desconto de duplicata"
    ProcVerificaAcao
    txtLimite.SetFocus
    Exit Sub
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_instituicoes where id = " & IIf(txtCodBanco = "", 0, txtCodBanco), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If FunVerifValidacaoRegistro("alterar", txtDtValidacao, "mesma", "instituição", False) = False Then Exit Sub
    If txtDescricao <> TBAbrir!Txt_descricao Or txtAgencia <> TBAbrir!txt_Agencia Or txtConta <> TBAbrir!txt_conta Then
        Conexao.Execute "Update tbl_contas_receber Set Banco = '" & txtDescricao & "' where Banco = '" & TBAbrir!Txt_descricao & "'"
        Conexao.Execute "Update tbl_contas_receber Set Nome_Razao = '" & txtDescricao & "' where idcliente = " & IIf(txtCodBanco = "", 0, txtCodBanco) & " and Tipo = 'IN'"
        Conexao.Execute "Update tbl_ContasPagar Set Banco = '" & txtDescricao & "' where Banco = '" & TBAbrir!Txt_descricao & "'"
        Conexao.Execute "Update tbl_ContasPagar Set txt_Fornecedor = '" & txtDescricao & "' where int_codforn = " & IIf(txtCodBanco = "", 0, txtCodBanco) & " and Tipo = 'IN'"
        Conexao.Execute "Update tbl_Detalhes_Recebimento Set txt_Portador_Banco = '" & txtDescricao & "', txt_Agencia = '" & txtAgencia & "', txt_Conta = '" & txtConta & "' where txt_Portador_Banco = '" & TBAbrir!Txt_descricao & "'"
        Conexao.Execute "Update tbl_Fluxo_de_caixa Set Instituicao = '" & txtDescricao & "' where Instituicao = '" & TBAbrir!Txt_descricao & "'"
        Conexao.Execute "Update troca_titulo Set Banco_recebedor = '" & txtDescricao & "' where Banco_recebedor = '" & TBAbrir!Txt_descricao & "'"
        Conexao.Execute "Update tbl_instituicoes_transf Set Banco_remetente = '" & txtDescricao & "' where id_banco_rem = " & TBAbrir!ID
        Conexao.Execute "Update tbl_instituicoes_transf Set Banco_recebedor = '" & txtDescricao & "' where id_banco_rec = " & TBAbrir!ID
    End If
Else
    TBAbrir.AddNew
    TBAbrir!Bloqueado = False
End If
If txtData1 = "" Then TBAbrir!data = Date Else TBAbrir!data = txtData1
If txtResponsavel = "" Then TBAbrir!Responsavel = pubUsuario Else TBAbrir!Responsavel = txtResponsavel
TBAbrir!Txt_familia = cmbFamilia
TBAbrir!Txt_descricao = txtDescricao.Text
TBAbrir!int_NBanco = IIf(txtNBanco = "", Null, txtNBanco)
TBAbrir!txt_Agencia = txtAgencia
TBAbrir!codigo_cedente = Txt_codigo_cedente
TBAbrir!Codigo_cedente_registrado = Txt_codigo_cedente1
TBAbrir!Nome_agencia = Txt_nome_agencia
TBAbrir!txt_conta = txtConta
TBAbrir!txt_gerente = txtgerente.Text
TBAbrir!txt_fone = txtFone
TBAbrir!Txt_fax = txtFAX
TBAbrir!Saldo = txtsaldo.Text
If Cmb_centro <> "" Then TBAbrir!ID_CC = Cmb_centro.ItemData(Cmb_centro.ListIndex) Else TBAbrir!ID_CC = Null
TBAbrir!Txt_obs = txtobs.Text
TBAbrir!Limite_desconto = txtLimite.Text
TBAbrir!Limite_utilizado = txtUtilizado.Text
TBAbrir!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBAbrir.Update
txtCodBanco = TBAbrir!ID
TBAbrir.Close
If Novo_Banco = True Then
    MsgBox ("Nova instituição bancária cadastrada com sucesso."), vbInformation
    Evento = "Novo"
    StrSql_Instituicoes_Localizar = "Select I.ID, I.ID_empresa, E.Empresa, I.int_NBanco, I.txt_Agencia, I.txt_conta, I.Txt_descricao, I.DtValidacao from tbl_Instituicoes I INNER JOIN Empresa E ON E.Codigo = I.ID_empresa where I.id = " & txtCodBanco
    ProcCarregaLista
Else
    MsgBox ("Alteração efetuada com sucesso."), vbInformation
    Evento = "Alterar"
    ProcCarregaLista
    If CodigoLista <> 0 And lst_Instituicoes.ListItems.Count <> 0 Then
        lst_Instituicoes.SelectedItem = lst_Instituicoes.ListItems(CodigoLista)
        lst_Instituicoes.SetFocus
    End If
End If

'==================================
Modulo = "Financeiro/Instituições"
ID_documento = txtCodBanco
Documento = "Instituição bancária: " & txtDescricao
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Banco = False

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcSalvarMovimentacao()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
If txtDtValidacao = "" Then
    MsgBox "Não é possivel alterar a movimentação, pois a instituição ainda não foi validada.", vbExclamation
    Exit Sub
End If
If txtStatus = "Bloqueada" Then
    MsgBox "Não é possivel alterar movimentação, pois a instituição esta bloqueada.", vbExclamation
    Exit Sub
End If
Acao = "salvar"
Select Case SSTab3.Tab
    Case 0:
        If frm_filtro.Enabled = False Then
            ProcVerificaSalvar
            Exit Sub
        End If
        If OptDeposito.Value = False And OptTransferencia.Value = False Then
            NomeCampo = "se é depósito ou transferência"
            ProcVerificaAcao
            Exit Sub
        End If
        If cmb_forma.Text = "" Then
            NomeCampo = "a forma da movimentação"
            ProcVerificaAcao
            cmb_forma.SetFocus
            Exit Sub
        End If
        If txtCheque = "" And (cmb_forma = "CHEQUE" Or cmb_forma = "DOC" Or cmb_forma = "TED" Or cmb_forma = "TEV") Then
            Select Case cmb_forma
                Case "CHEQUE": NomeCampo = "o número do cheque"
                Case "DOC": NomeCampo = "o número do DOC"
                Case "TED": NomeCampo = "o número do TED"
                Case "TEV": NomeCampo = "o número do TEV"
            End Select
            ProcVerificaAcao
            txtCheque.SetFocus
            Exit Sub
        End If
        If cmbrecebedor.Text = "" Then
            NomeCampo = "a instituição bancária recebedora"
            ProcVerificaAcao
            cmbrecebedor.SetFocus
            Exit Sub
        End If
        If mskvalor.Text = "" Then
            NomeCampo = "o valor movimentado"
            ProcVerificaAcao
            mskvalor.SetFocus
            Exit Sub
        End If
        If Txt_ID_PC_instituicao = 0 Then
            NomeCampo = "a conta contábil da instituição"
            ProcVerificaAcao
            Cmd_localizar_PC_instituicao.SetFocus
            Exit Sub
        End If
        If Txt_ID_PC_instituicao_rec = 0 Then
            NomeCampo = "a conta contábil da instituição recebedora"
            ProcVerificaAcao
            Cmd_localizar_PC_instituicao_rec.SetFocus
            Exit Sub
        End If
        ProcTransferir
    Case 1:
        If Frame8.Enabled = False Then
            ProcVerificaSalvar
            Exit Sub
        End If
        If Txt_valor.Text = "" Then
            NomeCampo = "o valor"
            ProcVerificaAcao
            Txt_valor.SetFocus
            Exit Sub
        End If
        ProcSaque
    Case 2:
        If Frame4.Enabled = False Then
            ProcVerificaSalvar
            Exit Sub
        End If
        If Cmb_operacao = "" Then
            NomeCampo = "a operação"
            ProcVerificaAcao
            Cmb_operacao.SetFocus
            Exit Sub
        End If
        If Cmb_tipo = "" Then
            NomeCampo = "o tipo do documento"
            ProcVerificaAcao
            Cmb_tipo.SetFocus
            Exit Sub
        End If
        If cmb_forma1 = "" Then
            NomeCampo = "a forma da baixa"
            ProcVerificaAcao
            cmb_forma1.SetFocus
            Exit Sub
        End If
        If Txt_ID_PC = 0 Then
            NomeCampo = "a conta contábil"
            ProcVerificaAcao
            Cmd_localizar_PC.SetFocus
            Exit Sub
        End If
        If Txt_valor1.Text = "" Then
            NomeCampo = "o valor"
            ProcVerificaAcao
            Txt_valor1.SetFocus
            Exit Sub
        End If
        ProcTarifa
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcTransferir()
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_instituicoes_transf where id_transf = " & txtid, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    If cmb_forma = "CHEQUE" Then
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select * from tbl_instituicoes_transf where id_transf = " & lst_transferencias.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            Cheque = "Cheque n. " & TBFIltro!NDoctoBaixa
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_Fluxo_de_caixa where Instituicao = '" & txtDescricao & "' and Descricao = '" & Cheque & "' and Bloqueado = 'False'", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                MsgBox ("Não é permitido alterar este depósito em cheque, pois o mesmo já está compensado."), vbExclamation
                TBAbrir.Close
                Exit Sub
            End If
            TBAbrir.Close
        End If
        TBFIltro.Close
    End If
    
    If txtCheque <> TBGravar!NDoctoBaixa Then
        Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & TBGravar!IDFluxo_Rec
        Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & TBGravar!IDFluxo
    End If
Else
    TBGravar.AddNew
End If
TBGravar!id_banco_rem = txtCodBanco
If txtResponsavel1 = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = txtResponsavel1
TBGravar!data_transf = txtdata.Value
TBGravar!banco_remetente = txtDescricao
TBGravar!Tipo = Tipo
TBGravar!FormaBaixa = cmb_forma

If Novo_Banco1 = True Then ProcAtualizaSaldos 0, mskvalor, False Else ProcAtualizaSaldos TBGravar!valor_transf, mskvalor, IIf(TBGravar!id_banco_rec <> cmbrecebedor.ItemData(cmbrecebedor.ListIndex), True, False)

TBGravar!valor_transf = mskvalor.Text
TBGravar!NDoctoBaixa = txtCheque
TBGravar!id_banco_rec = cmbrecebedor.ItemData(cmbrecebedor.ListIndex)
TBGravar!banco_recebedor = cmbrecebedor.Text
TBGravar.Update
txtid = TBGravar!id_transf

'Cria cheque na tabela de contas a receber
If Tipo = "D" And cmb_forma = "CHEQUE" Then
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from tbl_Contas_receber where NDoctoBaixa = '" & txtCheque & "' and Banco = '" & txtDescricao & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = True Then TBContas.AddNew
    TBContas!Logsit = Null
    
    'Verifica nome do cliente no cadastro da empresa
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Empresa where Codigo = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        TBContas!Nome_Razao = IIf(IsNull(TBAbrir!Razao), "", TBAbrir!Razao)
    End If
    TBAbrir.Close
    
    TBContas!FormaBaixa = "CHEQUE"
    TBContas!Data_pagamento = txtdata.Value
    TBContas!Data_movimentacao = txtdata.Value
    TBContas!valortitulorecebido = mskvalor.Text
    TBContas!NDoctoBaixa = txtCheque
    TBContas!Banco = cmbrecebedor
    TBContas!status = "DEPÓSITO EM CHEQUE"
    TBContas!resprec = pubUsuario
    TBContas!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    TBContas.Update
    TBContas.Close
End If

'Fluxo de Caixa
Set TBFluxo = CreateObject("adodb.recordset")
TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBGravar!IDFluxo_Rec), 0, TBGravar!IDFluxo_Rec), Conexao, adOpenKeyset, adLockOptimistic
If TBFluxo.EOF = True Then TBFluxo.AddNew
TBFluxo!IDintconta = txtid
TBFluxo!Operacao = "Crédito"
TBFluxo!data = txtdata.Value
TBFluxo!valor = mskvalor
If Tipo = "T" Then
    Select Case cmb_forma
        Case "DOC": TBFluxo!Descricao = "Doc n. " & txtCheque
        Case "TED": TBFluxo!Descricao = "Ted n. " & txtCheque
        Case "TEV": TBFluxo!Descricao = "Tev n. " & txtCheque
    End Select
Else
    If cmb_forma = "Dinheiro" Then
        TBFluxo!Descricao = "Depósito"
    Else
        TBFluxo!Descricao = "Cheque n. " & txtCheque
    End If
End If
TBFluxo!Instituicao = cmbrecebedor
TBFluxo!status = "S"
TBFluxo!Hora = Format(Now, "hh:mm:ss")
TBFluxo!Obs = IIf(txtObsFluxo = "", TBFluxo!Descricao, txtObsFluxo)
If txtCheque <> "" Then TBFluxo!Cheque = txtCheque
If cmb_forma = "CHEQUE" Then TBFluxo!Bloqueado = True Else TBFluxo!Bloqueado = False
TBFluxo!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBFluxo.Update
Conexao.Execute "UPDATE tbl_instituicoes_transf Set IDFluxo_rec = " & TBFluxo!IDFluxo & " where id_transf = " & txtid
TBFluxo.Close

Contador = 0
Do While Contador <> 9999995
    Contador = Contador + 1
Loop

'Cria cheque na tabela de contas a pagar
If Tipo = "D" And cmb_forma = "CHEQUE" Then
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select * from tbl_ContasPagar where NDoctoBaixa = '" & txtCheque & "' and Banco = '" & txtDescricao & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = True Then TBContas.AddNew
    TBContas!Logsit = Null
    TBContas!Txt_fornecedor = cmbrecebedor
    TBContas!FormaBaixa = "CHEQUE"
    TBContas!DataBaixa = txtdata.Value
    TBContas!Data_movimentacao = txtdata.Value
    TBContas!ValorPago = mskvalor.Text
    TBContas!NDoctoBaixa = txtCheque
    TBContas!Banco = txtDescricao
    TBContas!Favorecido = txtfavorecido
    TBContas!status = "DEPÓSITO EM CHEQUE"
    TBContas!resppag = pubUsuario
    TBContas!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
    TBContas.Update
    TBContas.Close
End If

Set TBFluxo = CreateObject("adodb.recordset")
TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBGravar!IDFluxo), 0, TBGravar!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
If TBFluxo.EOF = True Then TBFluxo.AddNew
TBFluxo!IDintconta = txtid
TBFluxo!Operacao = "Débito"
TBFluxo!data = txtdata.Value
TBFluxo!valor = mskvalor
If Tipo = "T" Then
    Select Case cmb_forma
        Case "DOC": TBFluxo!Descricao = "Doc n. " & txtCheque
        Case "TED": TBFluxo!Descricao = "Ted n. " & txtCheque
        Case "TEV": TBFluxo!Descricao = "Tev n. " & txtCheque
    End Select
Else
    If cmb_forma = "Dinheiro" Then
        TBFluxo!Descricao = "Depósito"
    Else
        TBFluxo!Descricao = "Cheque n. " & txtCheque
    End If
End If
TBFluxo!Instituicao = txtDescricao
TBFluxo!status = "S"
TBFluxo!Hora = Format(Now, "hh:mm:ss")
TBFluxo!Obs = IIf(txtObsFluxo = "", TBFluxo!Descricao, txtObsFluxo)
If txtCheque <> "" Then TBFluxo!Cheque = txtCheque
If cmb_forma = "CHEQUE" Then TBFluxo!Bloqueado = True Else TBFluxo!Bloqueado = False
TBFluxo!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBFluxo.Update
Conexao.Execute "UPDATE tbl_instituicoes_transf Set IDFluxo = " & TBFluxo!IDFluxo & " where id_transf = " & txtid
TBFluxo.Close
TBGravar.Close

'Cria conta contábil
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Familia_financeiro where IDConta = " & txtid & " and TipoConta = 'P' and Deposito_transf = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!IDConta = txtid
TBGravar!TipoConta = "P"
TBGravar!valor = mskvalor
TBGravar!Pago_recebido = True
TBGravar!ID_PC = Txt_ID_PC_instituicao
TBGravar!Deposito_transf = True
TBGravar.Update

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Familia_financeiro where IDConta = " & txtid & " and TipoConta = 'R' and Deposito_transf = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!IDConta = txtid
TBGravar!TipoConta = "R"
TBGravar!valor = mskvalor
TBGravar!Pago_recebido = True
TBGravar!ID_PC = Txt_ID_PC_instituicao_rec
TBGravar!Deposito_transf = True
TBGravar.Update

If Novo_Banco1 = True Then
    MsgBox ("Nova movimentação financeira cadastrada com sucesso."), vbInformation
    Evento = "Nova movimentação financeira"
    Instituicao_Localizar_Transf = "Select * from tbl_instituicoes_transf where id_transf = " & txtid
    ProcCarregaListaTransf
Else
    MsgBox ("Alteração efetuada com sucesso."), vbInformation
    Evento = "Alterar movimentação financeira"
    ProcCarregaListaTransf
    If CodigoLista1 <> 0 And lst_transferencias.ListItems.Count <> 0 Then
        lst_transferencias.SelectedItem = lst_transferencias.ListItems(CodigoLista1)
        lst_transferencias.SetFocus
    End If
End If
'==================================
Modulo = "Financeiro/Instituições"
ID_documento = txtid
Documento = "Instituição bancária: " & txtDescricao
Documento1 = "Data: " & txtdata.Value & " - Valor: " & mskvalor
ProcGravaEvento
'==================================
Novo_Banco1 = False

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcSaque()
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_instituicoes_transf where id_transf = " & Txt_id_saque, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
ProcAtualizaSaldosSaque
TBGravar!id_banco_rem = txtCodBanco
If txtResponsavel2 = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = txtResponsavel2
TBGravar!data_transf = txtdata2.Value
TBGravar!banco_remetente = txtDescricao
TBGravar!valor_transf = Txt_valor
TBGravar!Saldo = Txt_valor
TBGravar!Tipo = "S"
TBGravar.Update
Txt_id_saque = TBGravar!id_transf

'Fluxo de Caixa
Set TBFluxo = CreateObject("adodb.recordset")
TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBGravar!IDFluxo), 0, TBGravar!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
If TBFluxo.EOF = True Then TBFluxo.AddNew
TBFluxo!IDintconta = Txt_id_saque
TBFluxo!Operacao = "Débito"
TBFluxo!data = txtdata2.Value
TBFluxo!valor = Txt_valor
TBFluxo!Descricao = "Saque"
TBFluxo!Instituicao = txtDescricao
TBFluxo!status = "S"
TBFluxo!Hora = Format(Now, "hh:mm:ss")
TBFluxo!Obs = IIf(txtObsFluxo1 = "", TBFluxo!Descricao, txtObsFluxo1)
TBFluxo!Bloqueado = False
TBFluxo!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBFluxo.Update
Conexao.Execute "Update tbl_instituicoes_transf Set IDFluxo = " & TBFluxo!IDFluxo & " where id_transf = " & Txt_id_saque
TBFluxo.Close

TBGravar.Close
If Novo_Banco2 = True Then
    MsgBox ("Novo saque cadastrado com sucesso."), vbInformation
    Evento = "Novo saque"
    Instituicao_Localizar_Saque = "Select * from tbl_instituicoes_transf where id_transf = " & Txt_id_saque
    ProcCarregaListaSaque
Else
    MsgBox ("Alteração efetuada com sucesso."), vbInformation
    Evento = "Alterar saque"
    ProcCarregaListaSaque
    If CodigoLista2 <> 0 And Lst_saque.ListItems.Count <> 0 Then
        Lst_saque.SelectedItem = Lst_saque.ListItems(CodigoLista2)
        Lst_saque.SetFocus
    End If
End If
'==================================
Modulo = "Financeiro/Instituições"
ID_documento = Txt_id_saque
Documento = "Instituição bancária: " & txtDescricao
Documento1 = "Data: " & txtdata2 & " - Valor: " & Txt_valor
ProcGravaEvento
'==================================
Novo_Banco2 = False

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcTarifa()
On Error GoTo tratar_erro

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from tbl_instituicoes_transf where id_transf = " & IIf(Txt_id_tarifa = "", 0, Txt_id_tarifa), Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = True Then
    TBFI.AddNew
    If Cmb_operacao = "Crédito" Then ProcCriaTarifaRec "", False Else ProcCriaTarifaPag "", False
Else
    If Cmb_operacao = "Crédito" Then
        If TBFI!Tipo = "P" Then
            ProcExcluirTarifaPag TBFI!IDintconta
            ProcCriaTarifaRec "", False
        Else
            ProcCriaTarifaRec "where IdIntConta = " & TBFI!IDintconta, False
        End If
    Else
        If TBFI!Tipo = "R" Then
            ProcExcluirTarifaRec TBFI!IDintconta
            ProcCriaTarifaPag "", False
        Else
            ProcCriaTarifaPag "where IdIntConta = " & TBFI!IDintconta, False
        End If
    End If
End If
TBFI!id_banco_rem = txtCodBanco
TBFI!banco_remetente = txtDescricao
TBFI!Responsavel = txtResponsavel3
TBFI!data_transf = txtdata3
If Cmb_operacao = "Crédito" Then
    TBFI!Tipo = "R"
    NomeTabela = "tbl_contas_receber"
Else
    TBFI!Tipo = "P"
    NomeTabela = "tbl_ContasPagar"
End If
TBFI!FormaBaixa = cmb_forma1
TBFI!valor_transf = Txt_valor1
TBFI.Update
Txt_id_tarifa = TBFI!id_transf
Conexao.Execute "UPDATE IT set IT.IDFluxo = C.IDFluxo from tbl_instituicoes_transf IT INNER JOIN " & NomeTabela & " C ON IT.IDintconta = C.IDintconta where IT.id_transf = " & Txt_id_tarifa
TBFI.Close

If Novo_Banco3 = True Then
    MsgBox ("Nova tarifa cadastrada com sucesso."), vbInformation
    Evento = "Nova tarifa"
    Instituicao_Localizar_Tarifa = "Select * from tbl_instituicoes_transf where id_transf = " & IIf(Txt_id_tarifa = "", 0, Txt_id_tarifa)
    ProcCarregaListaTarifa
Else
    MsgBox ("Alteração efetuada com sucesso."), vbInformation
    Evento = "Alterar tarifa"
    ProcCarregaListaTarifa
    If CodigoLista3 <> 0 And Lst_tarifa.ListItems.Count <> 0 Then
        Lst_tarifa.SelectedItem = Lst_tarifa.ListItems(CodigoLista3)
        Lst_tarifa.SetFocus
    End If
End If
'==================================
Modulo = "Financeiro/Instituições"
ID_documento = Txt_id_tarifa
Documento = "Instituição bancária: " & txtDescricao
Documento1 = "Data: " & txtdata3 & " - Valor: " & Txt_valor1
ProcGravaEvento
'==================================
Novo_Banco3 = False

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCriaTarifaPag(TextoFiltro As String, Copiar As Boolean)
On Error GoTo tratar_erro

Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * from tbl_contaspagar " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TextoFiltro = "" Or TBContas.EOF = True Then TBContas.AddNew
ProcAtualizaSaldosTarifa
TBContas!Parcial = False
TBContas!impresso = False
TBContas!Bloqueado = False
TBContas!Logsit = "S"
TBContas!Despesas_NF = False
TBContas!Antecipacao = False
TBContas!Devolucao = False
TBContas!Data_transacao = txtdata3.Value
TBContas!Dt_emissao = txtdata3.Value
TBContas!dt_Pagamento = txtdata3.Value
TBContas!DataBaixa = txtdata3.Value
TBContas!Data_movimentacao = txtdata3.Value
TBContas!dbl_valorpagto = Txt_valor1
TBContas!ValorPago = Txt_valor1
TBContas!Banco = txtDescricao
TBContas!FormaBaixa = cmb_forma1
TBContas!Tipo = "IN"
TBContas!int_codforn = txtCodBanco
TBContas!Txt_fornecedor = txtDescricao
TBContas!Class_conta = Cmb_tipo.Text
If Copiar = True Then
    TBContas!Responsavel = pubUsuario
    TBContas!resppag = pubUsuario
Else
    TBContas!Responsavel = IIf(txtResponsavel3 = "", pubUsuario, txtResponsavel3)
    TBContas!resppag = IIf(txtResponsavel3 = "", pubUsuario, txtResponsavel3)
End If
TBContas!txt_Parcela = "001/001"
TBContas!status = "TÍTULO LIQUIDADO"
TBContas!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

TBContas.Update
TBFI!IDintconta = TBContas!IDintconta

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from familia_financeiro where IDConta = " & TBContas!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!Pago_recebido = True
TBGravar!ID_PC = Txt_ID_PC
TBGravar!IDConta = TBContas!IDintconta
TBGravar!valor = Txt_valor1
TBGravar!TipoConta = "P"
TBGravar.Update

'Fluxo de Caixa
Set TBFluxo = CreateObject("adodb.recordset")
TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBContas!IDFluxo), 0, TBContas!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
If TBFluxo.EOF = True Then TBFluxo.AddNew
TBFluxo!IDintconta = TBContas!IDintconta
TBFluxo!Operacao = "Débito"
TBFluxo!data = txtdata3.Value
TBFluxo!valor = Txt_valor1
TBFluxo!Descricao = "Tarifa"
TBFluxo!Instituicao = txtDescricao
TBFluxo!status = "S"
TBFluxo!Hora = Format(Now, "hh:mm:ss")
TBFluxo!Obs = IIf(txtObsFluxo2 = "", TBFluxo!Descricao, txtObsFluxo2)
TBFluxo!Bloqueado = False
TBFluxo!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBFluxo.Update
Conexao.Execute "UPDATE tbl_contaspagar set IDFluxo = " & TBFluxo!IDFluxo & " where IdIntConta = " & TBContas!IDintconta
TBFluxo.Close

TBContas.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCriaTarifaRec(TextoFiltro As String, Copiar As Boolean)
On Error GoTo tratar_erro

Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * from tbl_contas_receber " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TextoFiltro = "" Or TBContas.EOF = True Then TBContas.AddNew
ProcAtualizaSaldosTarifa
TBContas!Parcial = False
TBContas!titulodesc = False
TBContas!Bloqueado = False
TBContas!Logsit = "S"
TBContas!Antecipacao = False
TBContas!Devolucao = False
TBContas!Data_transacao = txtdata3.Value
TBContas!emissao = txtdata3.Value
TBContas!Vencimento = txtdata3.Value
TBContas!Data_pagamento = txtdata3.Value
TBContas!Data_movimentacao = txtdata3.Value
TBContas!valor = Txt_valor1
TBContas!valortitulorecebido = Txt_valor1
TBContas!Banco = txtDescricao
TBContas!FormaBaixa = cmb_forma1
TBContas!Tipo = "IN"
TBContas!IDCliente = txtCodBanco
TBContas!Nome_Razao = txtDescricao
TBContas!Tipo_doc = Cmb_tipo.Text
If Copiar = True Then
    TBContas!Responsavel = pubUsuario
    TBContas!resprec = pubUsuario
Else
    TBContas!Responsavel = IIf(txtResponsavel3 = "", pubUsuario, txtResponsavel3)
    TBContas!resprec = IIf(txtResponsavel3 = "", pubUsuario, txtResponsavel3)
End If
TBContas!Parcela = "001/001"
TBContas!status = "TÍTULO LIQUIDADO"
TBContas!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)

TBContas.Update
TBFI!IDintconta = TBContas!IDintconta

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from familia_financeiro where IDConta = " & TBContas!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!Pago_recebido = True
TBGravar!ID_PC = Txt_ID_PC
TBGravar!IDConta = TBContas!IDintconta
TBGravar!valor = Txt_valor1
TBGravar!TipoConta = "R"
TBGravar.Update

'Fluxo de Caixa
Set TBFluxo = CreateObject("adodb.recordset")
TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBContas!IDFluxo), 0, TBContas!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
If TBFluxo.EOF = True Then TBFluxo.AddNew
TBFluxo!IDintconta = TBContas!IDintconta
TBFluxo!Operacao = "Crédito"
TBFluxo!data = txtdata3.Value
TBFluxo!valor = Txt_valor1
TBFluxo!Descricao = "Tarifa"
TBFluxo!Instituicao = txtDescricao
TBFluxo!status = "S"
TBFluxo!Hora = Format(Now, "hh:mm:ss")
TBFluxo!Obs = IIf(txtObsFluxo2 = "", TBFluxo!Descricao, txtObsFluxo2)
TBFluxo!Bloqueado = False
TBFluxo!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBFluxo.Update
Conexao.Execute "UPDATE tbl_contas_receber set IDFluxo = " & TBFluxo!IDFluxo & " where IdIntConta = " & TBContas!IDintconta
TBFluxo.Close

TBContas.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCopiarTarifa()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
If Txt_id_tarifa = 0 Then
    MsgBox ("Informe a tarifa antes de copiar."), vbExclamation
    Exit Sub
End If
If MsgBox("Deseja realmente copiar esta tarifa?", vbYesNo + vbQuestion) = vbYes Then
    Novo_Banco3 = True
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from tbl_instituicoes_transf", Conexao, adOpenKeyset, adLockOptimistic
    TBFI.AddNew
    If Cmb_operacao = "Crédito" Then ProcCriaTarifaRec "", True Else ProcCriaTarifaPag "", True
    TBFI!id_banco_rem = txtCodBanco
    TBFI!banco_remetente = txtDescricao
    TBFI!Responsavel = pubUsuario
    TBFI!data_transf = txtdata3
    If Cmb_operacao = "Crédito" Then
        TBFI!Tipo = "R"
        NomeTabela = "tbl_contas_receber"
    Else
        TBFI!Tipo = "P"
        NomeTabela = "tbl_ContasPagar"
    End If
    TBFI!valor_transf = Txt_valor1
    TBFI.Update
    Txt_id_tarifa = TBFI!id_transf
    Conexao.Execute "UPDATE IT set IT.IDFluxo = C.IDFluxo from tbl_instituicoes_transf IT INNER JOIN " & NomeTabela & " C ON IT.IDintconta = C.IDintconta where IT.id_transf = " & Txt_id_tarifa
    TBFI.Close
    
    MsgBox ("Tarifa copiada com sucesso."), vbInformation
    '==================================
    Modulo = "Financeiro/Instituições"
    Evento = "Nova tarifa"
    ID_documento = Txt_id_tarifa
    Documento = "Instituição bancária: " & txtDescricao
    Documento1 = "Data: " & txtdata3 & " - Valor: " & Txt_valor1
    ProcGravaEvento
    '==================================
    Instituicao_Localizar_Tarifa = "Select * from tbl_instituicoes_transf where id_transf = " & IIf(Txt_id_tarifa = "", 0, Txt_id_tarifa)
    ProcCarregaListaTarifa
    Novo_Banco3 = False
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaSaldos(VlrTransfAnt As Double, VlrTransfNovo As Double, BancoRecAlterado As Boolean)
On Error GoTo tratar_erro

Valor1 = IIf(IsNumeric(txtsaldo) = True, txtsaldo, 0)
Valor2 = VlrTransfNovo
Valor3 = Valor1 - Valor2
If Novo_Banco1 = True Then
    If cmb_forma <> "CHEQUE" Then
        'Atualiza saldo do banco remetente
        Set TBSaldo = CreateObject("adodb.recordset")
        TBSaldo.Open "Select saldo from tbl_Instituicoes where ID = " & txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
        If TBSaldo.EOF = False Then
            TBSaldo!Saldo = Valor3
            txtsaldo = Format(TBSaldo!Saldo, "###,##0.00")
            TBSaldo.Update
        End If
        TBSaldo.Close
        
        'Atualiza saldo do banco recebedor
        Set TBSaldo = CreateObject("adodb.recordset")
        TBSaldo.Open "Select saldo from tbl_Instituicoes where ID = " & cmbrecebedor.ItemData(cmbrecebedor.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
        If TBSaldo.EOF = False Then
            Valor1 = TBSaldo!Saldo
            Valor3 = Valor1 + Valor2
            TBSaldo!Saldo = Valor3
            TBSaldo.Update
        End If
        TBSaldo.Close
    End If
Else
    Permitido = False
    If cmb_forma = "CHEQUE" Then
        'Verifica se o cheque já foi compensado
        Cheque = "Cheque n. " & txtCheque
        Set TBFIltro = CreateObject("adodb.recordset")
        TBFIltro.Open "Select * from tbl_Fluxo_de_caixa where Instituicao = '" & txtDescricao & "' and Descricao = '" & Cheque & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBFIltro.EOF = False Then
            If TBFIltro!Bloqueado = True Then Permitido = False Else Permitido = True
        End If
        TBFIltro.Close
    Else
        Permitido = True
    End If
    If Permitido = True Then
        'Atualiza saldo do banco remetente
        Set TBSaldo = CreateObject("adodb.recordset")
        TBSaldo.Open "Select saldo from tbl_Instituicoes where ID = " & txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
        If TBSaldo.EOF = False Then
            TBSaldo!Saldo = (TBSaldo!Saldo + VlrTransfAnt) - VlrTransfNovo
            txtsaldo = Format(TBSaldo!Saldo, "###,##0.00")
            TBSaldo.Update
        End If
        TBSaldo.Close
        
        'Atualiza saldo do banco recebedor
        If BancoRecAlterado = True Then
            Set TBSaldo = CreateObject("adodb.recordset")
            TBSaldo.Open "Select saldo from tbl_Instituicoes where ID = " & TBGravar!id_banco_rec, Conexao, adOpenKeyset, adLockOptimistic
            If TBSaldo.EOF = False Then
                TBSaldo!Saldo = (TBSaldo!Saldo - VlrTransfAnt)
                TBSaldo.Update
            End If

            Set TBSaldo = CreateObject("adodb.recordset")
            TBSaldo.Open "Select saldo from tbl_Instituicoes where ID = " & cmbrecebedor.ItemData(cmbrecebedor.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
            If TBSaldo.EOF = False Then
                TBSaldo!Saldo = TBSaldo!Saldo + VlrTransfNovo
                TBSaldo.Update
            End If
            TBSaldo.Close
        Else
            Set TBSaldo = CreateObject("adodb.recordset")
            TBSaldo.Open "Select saldo from tbl_Instituicoes where ID = " & cmbrecebedor.ItemData(cmbrecebedor.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
            If TBSaldo.EOF = False Then
                TBSaldo!Saldo = (TBSaldo!Saldo - VlrTransfAnt) + VlrTransfNovo
                TBSaldo.Update
            End If
            TBSaldo.Close
        End If
    End If
End If
Valor1 = 0
Valor2 = 0
Valor3 = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaSaldosSaque()
On Error GoTo tratar_erro

If Novo_Banco2 = True Then
    Set TBSaldo = CreateObject("adodb.recordset")
    TBSaldo.Open "Select saldo from tbl_Instituicoes where ID = " & txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
    If TBSaldo.EOF = False Then
        TBSaldo!Saldo = TBSaldo!Saldo - Txt_valor
        txtsaldo = Format(TBSaldo!Saldo, "###,##0.00")
        TBSaldo.Update
    End If
    TBSaldo.Close
Else
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select valor_transf from tbl_instituicoes_transf where id_transf = " & Txt_id_saque, Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        Set TBSaldo = CreateObject("adodb.recordset")
        TBSaldo.Open "Select saldo from tbl_Instituicoes where ID = " & txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
        If TBSaldo.EOF = False Then
            TBSaldo!Saldo = (TBSaldo!Saldo + TBFIltro!valor_transf) - Txt_valor
            txtsaldo = Format(TBSaldo!Saldo, "###,##0.00")
            TBSaldo.Update
        End If
        TBSaldo.Close
    End If
    TBFIltro.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaSaldosTarifa()
On Error GoTo tratar_erro

If Novo_Banco3 = True Or (Cmb_operacao = "Débito" And TBFI!Tipo = "R" Or Cmb_operacao = "Crédito" And TBFI!Tipo = "P") Then
    Set TBSaldo = CreateObject("adodb.recordset")
    TBSaldo.Open "Select saldo from tbl_Instituicoes where ID = " & txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
    If TBSaldo.EOF = False Then
        If Cmb_operacao = "Débito" Then TBSaldo!Saldo = TBSaldo!Saldo - Txt_valor1 Else TBSaldo!Saldo = TBSaldo!Saldo + Txt_valor1
        txtsaldo = Format(TBSaldo!Saldo, "###,##0.00")
        TBSaldo.Update
    End If
    TBSaldo.Close
Else
    If Cmb_operacao = "Débito" Then TextoFiltro = "valorpago from tbl_ContasPagar where idintconta = " & TBFI!IDintconta Else TextoFiltro = "valortitulorecebido from tbl_contas_receber where idintconta = " & TBFI!IDintconta
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        Set TBSaldo = CreateObject("adodb.recordset")
        TBSaldo.Open "Select saldo from tbl_Instituicoes where ID = " & txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
        If TBSaldo.EOF = False Then
            If Cmb_operacao = "Débito" Then TBSaldo!Saldo = (TBSaldo!Saldo + TBFIltro!ValorPago) - Txt_valor1 Else TBSaldo!Saldo = (TBSaldo!Saldo - TBFIltro!valortitulorecebido) + Txt_valor1
            txtsaldo = Format(TBSaldo!Saldo, "###,##0.00")
            TBSaldo.Update
        End If
        TBSaldo.Close
    End If
    TBFIltro.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcSalvarChequeEmitido()
On Error GoTo tratar_erro

If Frame7.Enabled = False Then
    MsgBox ("Informe o cheque antes de salvar."), vbExclamation
    Exit Sub
End If
If txtDtValidacao = "" Then
    MsgBox "Não é possivel alterar o cheque, pois a instituição ainda não foi validada.", vbExclamation
    Exit Sub
End If
If txtStatus = "Bloqueada" Then
    MsgBox "Não é possivel alterar o cheque, pois a instituição esta bloqueada.", vbExclamation
    Exit Sub
End If
If SSTab2.Tab = 0 Then
    Conexao.Execute "Update tbl_ContasPagar Set Obscheque = '" & txtobscheque & "', Favorecido = '" & Txt_favorecido & "' where IdIntConta = " & Lst_cheque.SelectedItem
    ID_documento = Lst_cheque.SelectedItem
    Documento = "Cheque nº: " & Lst_cheque.SelectedItem.ListSubItems(3)
Else
    Conexao.Execute "Update tbl_ContasPagar Set obscheque = '" & txtobscheque & "', Favorecido = '" & Txt_favorecido & "' where IdIntConta = " & Lst_cheque1.SelectedItem
    ID_documento = Lst_cheque1.SelectedItem
    Documento = "Cheque nº: " & Lst_cheque1.SelectedItem.ListSubItems(3)
End If
MsgBox ("Alteração efetuada com sucesso."), vbInformation
'==================================
Modulo = "Financeiro/Instituições"
Evento = "Alterar dados do cheque emitido"
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Label2_DblClick(Index As Integer)
On Error GoTo tratar_erro

If TxtHistoricoExtrato <> "" Then
    If InputBox("Informe a senha para esconder este lançamento no extrato.") = "280362BLOQ" Then
        Conexao.Execute "UPDATE Tbl_Fluxo_de_Caixa Set Bloqueado = 1 where IDFluxo = " & Lst_extrato.SelectedItem
        MsgBox ("Operação realizada com sucesso."), vbInformation
        '==================================
        Modulo = "Financeiro/Instituições"
        Evento = "Esconder lançamento no extrato"
        ID_documento = Lst_extrato.SelectedItem
        Documento = "Instituição bancária: " & txtDescricao
        Documento1 = "ID do lançamento: & " & Lst_extrato.SelectedItem & " - Data do lançamento: " & Lst_extrato.SelectedItem.ListSubItems(1)
        ProcGravaEvento
        '==================================
        TxtHistoricoExtrato = ""
        ProcFiltrarExtrato
    End If
End If
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub



Private Sub Lista_cheque_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista_cheque
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If txtDtValidacao = "" Then GoTo Proximo
                If txtStatus = "Bloqueada" Then GoTo Proximo
                
                If Cmb_opcao_lista_recebidos = "Excluir" Or Cmb_opcao_lista_recebidos = "Compensar" Then
                    If .ListItems.Item(InitFor).ListSubItems(6) = "Sim" Then GoTo Proximo
                Else
                    If .ListItems.Item(InitFor).ListSubItems(6) = "Não" Then GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista_cheque, ColumnHeader
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lista_cheque_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista_cheque
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If Cmb_opcao_lista = "Excluir/cancelar" Then
                NomeCampo = "cancelar este"
            ElseIf Cmb_opcao_lista = "Compensar" Then
                NomeCampo = "compensar este"
            Else
                NomeCampo = "cancelar compensação desde"
            End If
        
            If txtDtValidacao = "" Then
                MsgBox "Não é possivel " & NomeCampo & " cheque, pois a instituição ainda não foi validada.", vbExclamation
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If txtStatus = "Bloqueada" Then
                MsgBox "Não é possivel " & NomeCampo & " cheque, pois a instituição esta bloqueada.", vbExclamation
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If Cmb_opcao_lista_recebidos = "Excluir" Or Cmb_opcao_lista_recebidos = "Compensar" Then
                If .ListItems.Item(InitFor).ListSubItems(6) = "Sim" Then
                    MsgBox ("Não é permitido " & NomeCampo & " cheque, pois o mesmo já está compensado."), vbExclamation
                    .ListItems.Item(InitFor).Checked = False
                End If
            Else
                If .ListItems.Item(InitFor).ListSubItems(6) = "Não" Then
                    MsgBox ("Não é permitido cancelar a compensação deste cheque, pois o mesmo ainda não foi compensado."), vbExclamation
                    .ListItems.Item(InitFor).Checked = False
                End If
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lst_cheque_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lst_cheque
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If txtDtValidacao = "" Then GoTo Proximo
                If txtStatus = "Bloqueada" Then GoTo Proximo
                If Cmb_opcao_lista = "Excluir/cancelar" Or Cmb_opcao_lista = "Compensar" Then
                    If .ListItems.Item(InitFor).ListSubItems(6) = "Sim" Then GoTo Proximo
                Else
                    If .ListItems.Item(InitFor).ListSubItems(6) = "Não" Then GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lst_cheque, ColumnHeader
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lst_cheque_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lst_cheque
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If Cmb_opcao_lista = "Excluir/cancelar" Then
                NomeCampo = "cancelar este"
            ElseIf Cmb_opcao_lista = "Compensar" Then
                NomeCampo = "compensar este"
            Else
                NomeCampo = "cancelar compensação deste"
            End If
            
            If txtDtValidacao = "" Then
                MsgBox "Não é possivel " & NomeCampo & " cheque, pois a instituição ainda não foi validada.", vbExclamation
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If txtStatus = "Bloqueada" Then
                MsgBox "Não é possivel " & NomeCampo & " cheque, pois a instituição esta bloqueada.", vbExclamation
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If Cmb_opcao_lista = "Excluir/cancelar" Or Cmb_opcao_lista = "Compensar" Then
                If .ListItems.Item(InitFor).ListSubItems(6) = "Sim" Then
                    MsgBox ("Não é permitido " & NomeCampo & " cheque, pois o mesmo já está compensado."), vbExclamation
                    .ListItems.Item(InitFor).Checked = False
                End If
            Else
                If .ListItems.Item(InitFor).ListSubItems(6) = "Não" Then
                    MsgBox ("Não é permitido cancelar a compensação deste cheque, pois o mesmo ainda não foi compensado."), vbExclamation
                    .ListItems.Item(InitFor).Checked = False
                End If
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lst_cheque_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lst_cheque.ListItems.Count = 0 Then Exit Sub

Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * from tbl_ContasPagar where IdIntConta = " & Lst_cheque.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    Txt_favorecido = IIf(IsNull(TBContas!Favorecido), "", TBContas!Favorecido)
    txtobscheque = IIf(IsNull(TBContas!Obscheque), "", TBContas!Obscheque)
End If
TBContas.Close
Frame7.Enabled = True

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lst_cheque1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lst_cheque1
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lst_cheque1, ColumnHeader
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lst_cheque1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lst_cheque1.ListItems.Count = 0 Then Exit Sub
Set TBContas = CreateObject("adodb.recordset")
TBContas.Open "Select * from tbl_ContasPagar where IdIntConta = " & Lst_cheque1.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBContas.EOF = False Then
    Txt_favorecido = IIf(IsNull(TBContas!Favorecido), "", TBContas!Favorecido)
    txtobscheque = IIf(IsNull(TBContas!Obscheque), "", TBContas!Obscheque)
End If
TBContas.Close
Frame7.Enabled = True

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lst_Contas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lst_Contas, ColumnHeader

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub lst_Duplicata_Click()
On Error GoTo tratar_erro
Titulosselecionados = 0

If ColumnHeader = "" Then
    Contador = 0
    With lst_Duplicata
    For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
            Titulosselecionados = Titulosselecionados + 1
            End If
        Next InitFor
    End With
End If

If Titulosselecionados > 0 Then
    chkRemessa.Enabled = True
    chkEmail.Enabled = True
    chkImprimir.Enabled = True
Else
    chkRemessa.Enabled = False
    chkEmail.Enabled = False
    chkEmailcopia.Enabled = False
    chkImprimir.Enabled = False
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub lst_Duplicata_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro
Titulosselecionados = 0

If ColumnHeader = "" Then
    With lst_Duplicata
    For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Titulosselecionados = Titulosselecionados + 1
                .ListItems.Item(InitFor).Checked = True
            End If
        Next InitFor
    End With
End If

If Titulosselecionados > 0 Then
    chkRemessa.Enabled = True
    chkRemessa.Value = 1
    chkEmail.Enabled = True
    chkImprimir.Enabled = True
Else
    chkRemessa.Value = 0
    chkRemessa.Enabled = False
    chkEmail.Enabled = False
    chkEmailcopia.Enabled = False
    chkImprimir.Enabled = False
End If

ProcOrdenaListView lst_Duplicata, ColumnHeader

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtcarteiraconf_Change()
On Error GoTo tratar_erro

    If txtcarteiraconf.Text <> "" Then
    FramePesquisa.Enabled = True
    Else
    FramePesquisa.Enabled = False
    End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lst_extrato_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lst_extrato, ColumnHeader

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lst_extrato_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With TxtHistoricoExtrato
    .Text = ""
    .Locked = True
    .TabStop = False
    If Lst_extrato.ListItems.Count > 0 And Lst_extrato.SelectedItem <> "" Then
        Set TBFluxo = CreateObject("adodb.recordset")
        TBFluxo.Open "Select IDFluxo, Data, ID_varias, Instituicao, Cheque, Operacao, Obs, Descricao from Tbl_Fluxo_de_Caixa where IDFluxo = " & Lst_extrato.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBFluxo.EOF = False Then
            .Text = IIf(IsNull(TBFluxo!Obs), "", TBFluxo!Obs)
            .Locked = False
            .TabStop = True
        End If
        TBFluxo.Close
    End If
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub lst_Instituicoes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lst_Instituicoes
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If cmb_Opcao_Lista_Instituicao = "Excluir" Then
                    If .ListItems.Item(InitFor).ListSubItems(7) = "Sim" Then GoTo Proximo
                    
                    ProcVerificaRegistroUtilizadoSemMsg "tbl_ContasPagar", "banco = '" & .ListItems(InitFor).ListSubItems(5) & "' and ID_empresa = " & .ListItems(InitFor).ListSubItems(6)
                    If Permitido = False Then GoTo Proximo
                    ProcVerificaRegistroUtilizadoSemMsg "tbl_contas_receber", "banco = '" & .ListItems(InitFor).ListSubItems(5) & "' and ID_empresa = " & .ListItems(InitFor).ListSubItems(6)
                    If Permitido = False Then GoTo Proximo
                    ProcVerificaRegistroUtilizadoSemMsg "tbl_instituicoes_transf", "id_banco_rem = " & .ListItems(InitFor) & " or id_banco_rec = " & .ListItems(InitFor)
                    If Permitido = False Then GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView lst_Instituicoes, ColumnHeader
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub lst_Instituicoes_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With lst_Instituicoes
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If cmb_Opcao_Lista_Instituicao = "Excluir" Then
                If .ListItems.Item(InitFor).ListSubItems(7) = "Sim" Then
                    MsgBox "Não é possivel excluir esta instituição, pois a mesma esta validada.", vbExclamation
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                
                Mensagem = "Não é permitido excluir esta instituição bancária, pois a mesma está sendo utilizado no módulo"
                ProcVerificaRegistroUtilizado "tbl_ContasPagar", "banco = '" & .ListItems(InitFor).ListSubItems(5) & "' and ID_empresa = " & .ListItems(InitFor).ListSubItems(6), "Financeiro/Contas a pagar"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "tbl_contas_receber", "banco = '" & .ListItems(InitFor).ListSubItems(5) & "' and ID_empresa = " & .ListItems(InitFor).ListSubItems(6), "Financeiro/Contas a receber"
                If Permitido = False Then
                    .ListItems.Item(InitFor).Checked = False
                    Exit Sub
                End If
                ProcVerificaRegistroUtilizado "tbl_instituicoes_transf", "id_banco_rem = " & .ListItems(InitFor) & " or id_banco_rec = " & .ListItems(InitFor), "Financeiro/Instituições/Movimentação financeira"
                If Permitido = False Then .ListItems.Item(InitFor).Checked = False
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub lst_Instituicoes_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
        
If lst_Instituicoes.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Instituicoes where Id = " & lst_Instituicoes.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCampos
    ProcCarregaDados
    CodigoLista = lst_Instituicoes.SelectedItem.Index
End If
TBAbrir.Close

ProcCarregaInstituicaoBoleto
DTINI.Value = Date
DTFim.Value = "31/12/" & Year(Date)
ProcCarregacomboCarteira
ProcCarregadadosCedente
ProcCarregaInstrucoesBoleto

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCarregadadosCedente()
On Error GoTo tratar_erro

   Set TBAbrir = CreateObject("adodb.recordset")
    
    TBAbrir.Open "Select Codigo, Razao,email from Empresa where EMPRESA = '" & Cmb_empresa & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        txtCodcedente = TBAbrir!CODIGO
        txtNomecedente = TBAbrir!Razao
        EmailCopia = TBAbrir!Email
    End If
            
    TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcBuscaArquivolicenca()
Agencia = txtAgencia
If Txt_codigo_cedente.Text = "" Then Exit Sub
'Início dos parâmetros obrigatórios da ContaCorrente corrente
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Empresa where Empresa = '" & Cmb_empresa & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Select Case txtNBanco
        Case "001": 'Banco do brasil
            Select Case cmbCarteira
                Case "11 - Simples - Com Registro":
                    txtcarteiraconf = TBAbrir!Registro_boleto & "-001-11.conf"
                    OutrosDadosConfiguracao1 = Carteira1
                Case "11 - Vinculada - Com Registro"
                     ArquivoLicensa = TBAbrir!Registro_boleto & "-001-11Vinculada.conf"
                     txtcarteiraconf = ArquivoLicensa
'                    OutrosDadosConfiguracao1 = txtCodigocedente
'                    OutrosDadosConfiguracao2 = txtCodigocedente
                Case "17 - Direta Especial - Com Registro":
                    txtcarteiraconf = TBAbrir!Registro_boleto & "-001-17.conf"
                    OutrosDadosConfiguracao1 = Carteira1
                Case "17Simples - Direta Especial Simples - Com Registro":
                    txtcarteiraconf = TBAbrir!Registro_boleto & "-001-17SIMPLES.conf"
                    OutrosDadosConfiguracao1 = Carteira1
                Case "17-7 - Direta Especial - Com Registro Convênio 7 dígitos":
                    txtcarteiraconf = TBAbrir!Registro_boleto & "-001-17-7.conf"
                    OutrosDadosConfiguracao1 = Carteira1
                    OutrosDadosConfiguracao2 = "0000000000"
                Case "18 - Simples - Sem Registro":
                    txtcarteiraconf = TBAbrir!Registro_boleto & "-001-18.conf"
                    OutrosDadosConfiguracao1 = Carteira1
                Case "18-7 - Simples - Sem Registro - Convênio 7 dígitos":
                    txtcarteiraconf = TBAbrir!Registro_boleto & "-001-18-7.conf"
                    OutrosDadosConfiguracao1 = Carteira1
            End Select
            Select Case Len(txtAgencia)
                Case 1: AgenciaBol = "0000-" & Agencia
                Case 2: AgenciaBol = "000" & Left(Agencia, 1) & "-" & Right(Agencia, 1)
                Case 3: AgenciaBol = "00" & Left(Agencia, 2) & "-" & Right(Agencia, 1)
                Case 4: AgenciaBol = "0" & Left(Agencia, 3) & "-" & Right(Agencia, 1)
                Case Is >= 5: AgenciaBol = Left(Agencia, 4) & "-" & Right(Agencia, 1)
            End Select
            Select Case Len(ContaCorrente)
                Case 1: ContaCorrenteBol = "00000000-" & ContaCorrente
                Case 2: ContaCorrenteBol = "0000000" & Left(ContaCorrente, 1) & "-" & Right(ContaCorrente, 1)
                Case 3: ContaCorrenteBol = "000000" & Left(ContaCorrente, 2) & "-" & Right(ContaCorrente, 1)
                Case 4: ContaCorrenteBol = "00000" & Left(ContaCorrente, 3) & "-" & Right(ContaCorrente, 1)
                Case 5: ContaCorrenteBol = "0000" & Left(ContaCorrente, 4) & "-" & Right(ContaCorrente, 1)
                Case 6: ContaCorrenteBol = "000" & Left(ContaCorrente, 5) & "-" & Right(ContaCorrente, 1)
                Case 7: ContaCorrenteBol = "00" & Left(ContaCorrente, 6) & "-" & Right(ContaCorrente, 1)
                Case 8: ContaCorrenteBol = "0" & Left(ContaCorrente, 7) & "-" & Right(ContaCorrente, 1)
                Case Is >= 9: ContaCorrenteBol = Left(ContaCorrente, 8) & "-" & Right(ContaCorrente, 1)
            End Select

            Debug.Print AgenciaBol
            
            
            If cmbCarteira = "17-7 - Direta Especial - Com Registro Convênio 7 dígitos" Or cmbCarteira = "18-7 - Simples - Sem Registro - Convênio 7 dígitos" Then
                'Codigocedente = FunTamanhoTextoZeroEsq(Left(Codigocedente, 7), 7)
            Else
                'Codigocedente = FunTamanhoTextoZeroEsq(Left(Codigocedente, 6), 6)
            End If
            
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Banco do brasil"
            Arquivo = "CBR" & Dia & Mes & "." & SeqRemessa
            Layout = "FEBRABAN240"
        Case "033": 'Santander
            If cmbCarteira = "CSR - Cobrança Simples Sem Registro" Or cmbCarteira = "ECR - Cobrança Simples Com Registro" Or cmbCarteira = "COBR-Nova - Cobrança Simples - Rápida Com Registro" Then
                Select Case cmbCarteira
                    Case "CSR - Cobrança Simples Sem Registro": txtcarteiraconf = TBAbrir!Registro_boleto & "-033-CSR.conf"
                    Case "ECR - Cobrança Simples Com Registro":
                        txtcarteiraconf = TBAbrir!Registro_boleto & "-033-ECR.conf"
                        'OutrosDadosConfiguracao1 = Left( txtAgencia, 4) & ProcTamanhoTextoZeroEsq(Codigocedente, 7) & Right(txtContaCorrente, 9) forma antiga com 11 digitos no nosso numero
                        OutrosDadosConfiguracao1 = Left(txtAgencia, 5) & FunTamanhoTextoZeroEsq(Codigocedente, 7) & Left(txtContacorrente, 9)
                    Case "COBR-Nova - Cobrança Simples - Rápida Com Registro":
                        txtcarteiraconf = TBAbrir!Registro_boleto & "-033-COBR-NOVA.conf"
                        OutrosDadosConfiguracao1 = Left(txtAgencia, 5) & FunTamanhoTextoZeroEsq(Codigocedente, 7) & Left(txtContacorrente, 9)
                End Select
                Select Case Len(txtAgencia)
                    Case 1: AgenciaBol = "0000-" & txtAgencia
                    Case 2: AgenciaBol = "000" & Left(txtAgencia, 1) & "-" & Right(txtAgencia, 1)
                    Case 3: AgenciaBol = "00" & Left(txtAgencia, 2) & "-" & Right(txtAgencia, 1)
                    Case 4: AgenciaBol = "0" & Left(txtAgencia, 3) & "-" & Right(txtAgencia, 1)
                    Case Is >= 5: AgenciaBol = Left(txtAgencia, 4) & "-" & Right(txtAgencia, 1)
                End Select
                Select Case Len(txtContacorrente)
                    Case 1: ContaCorrenteBol = "000000000-" & txtContacorrente
                    Case 2: ContaCorrenteBol = "00000000" & Left(txtContacorrente, 1) & "-" & Right(txtContacorrente, 1)
                    Case 3: ContaCorrenteBol = "0000000" & Left(txtContacorrente, 2) & "-" & Right(txtContacorrente, 1)
                    Case 4: ContaCorrenteBol = "000000" & Left(txtContacorrente, 3) & "-" & Right(txtContacorrente, 1)
                    Case 5: ContaCorrenteBol = "00000" & Left(txtContacorrente, 4) & "-" & Right(txtContacorrente, 1)
                    Case 6: ContaCorrenteBol = "0000" & Left(txtContacorrente, 5) & "-" & Right(txtContacorrente, 1)
                    Case 7: ContaCorrenteBol = "000" & Left(txtContacorrente, 6) & "-" & Right(txtContacorrente, 1)
                    Case 8: ContaCorrenteBol = "00" & Left(txtContacorrente, 7) & "-" & Right(txtContacorrente, 1)
                    Case 9: ContaCorrenteBol = "0" & Left(txtContacorrente, 8) & "-" & Right(txtContacorrente, 1)
                    Case Is >= 10: ContaCorrenteBol = Left(txtContacorrente, 9) & "-" & Right(txtContacorrente, 1)
                End Select
                Select Case Len(Codigocedente)
                    Case 1: Codigocedente = "000000-" & Codigocedente
                    Case 2: Codigocedente = "00000" & Left(Codigocedente, 1) & "-" & Right(Codigocedente, 1)
                    Case 3: Codigocedente = "0000" & Left(Codigocedente, 2) & "-" & Right(Codigocedente, 1)
                    Case 4: Codigocedente = "000" & Left(Codigocedente, 3) & "-" & Right(Codigocedente, 1)
                    Case 5: Codigocedente = "00" & Left(Codigocedente, 4) & "-" & Right(Codigocedente, 1)
                    Case 6: Codigocedente = "0" & Left(Codigocedente, 5) & "-" & Right(Codigocedente, 1)
                    Case Is >= 7: Codigocedente = Left(Codigocedente, 6) & "-" & Right(Codigocedente, 1)
                End Select
            Else
                Select Case cmbCarteira
                    Case "COB - Cobrança Simples": txtcarteiraconf = TBAbrir!Registro_boleto & "-033-COB.conf"
                    Case "COBR - Cobrança Simples - Rápida Com Registro": txtcarteiraconf = TBAbrir!Registro_boleto & "-033-COBR.conf"
                End Select
                AgenciaBol = Mid(txtAgencia, 2, 3)
                ContaCorrente = Codigocedente
                Select Case Len(txtContacorrente)
                    Case 1: ContaCorrenteBol = "00" & " " & "00000" & " " & txtContacorrente
                    Case 2: ContaCorrenteBol = "00" & " " & "0000" & Left(txtContacorrente, 1) & " " & Mid(txtContacorrente, 2, 1)
                    Case 3: ContaCorrenteBol = "00" & " " & "000" & Left(txtContacorrente, 2) & " " & Mid(txtContacorrente, 3, 1)
                    Case 4: ContaCorrenteBol = "00" & " " & "00" & Left(txtContacorrente, 3) & " " & Mid(txtContacorrente, 4, 1)
                    Case 5: ContaCorrenteBol = "00" & " " & "0" & Left(txtContacorrente, 4) & " " & Mid(txtContacorrente, 5, 1)
                    Case 6: ContaCorrenteBol = "00" & " " & Left(txtContacorrente, 5) & " " & Mid(txtContacorrente, 6, 1)
                    Case 7: ContaCorrenteBol = "0" & Left(txtContacorrente, 1) & " " & Mid(txtContacorrente, 2, 5) & " " & Mid(txtContacorrente, 7, 1)
                    Case Is >= 8: ContaCorrenteBol = Left(txtContacorrente, 2) & " " & Mid(txtContacorrente, 3, 5) & " " & Mid(txtContacorrente, 8, 1)
                End Select
                Codigocedente = FunTamanhoTextoVazioDir(Left(NomeAgencia, 20), 20)
            End If
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Santander"
            Arquivo = "DB" & Dia & Mes & Right(ano, 2) & "." & SeqRemessa
            Layout = "CNAB400"
        Case "104": 'Caixa
            If cmbCarteira = "SIG14 - SIG Com Registro - Emissão pelo Cedente" Then
                txtcarteiraconf = TBAbrir!Registro_boleto & "-104-SIG14.conf"
                AgenciaBol = FunTamanhoTextoZeroEsq(Left(txtAgencia, 4), 4)
                ContaCorrenteBol = ""
                Codigocedente = FunTamanhoTextoZeroEsq(Left(DS_RetornarNumeros(txtCodigocedente), 6), 6)
                Layout = "SIGCB240"
            Else
                Select Case cmbCarteira
                    Case "CR - Cobrança Rápida": txtcarteiraconf = TBAbrir!Registro_boleto & "-104-CR.conf"
                    Case "SR - Cobrança Sem Registro": txtcarteiraconf = TBAbrir!Registro_boleto & "-104-SR.conf"
                End Select
                AgenciaBol = ""
                ContaCorrenteBol = ""
                Codigocedente = DS_RetornarNumeros(txtCodigocedente)
                Codigocedente = Left(Codigocedente, 4) & "." & Mid(Codigocedente, 5, 3) & "." & Mid(Codigocedente, 8, 8) & "-" & Right(Codigocedente, 1)
                Layout = "CNAB400"
            End If
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Caixa"
            Arquivo = "CB" & Dia & Mes & "." & SeqRemessa
        Case "237": 'Bradesco
            Select Case cmbCarteira
                Case "06 - Sem Registro": txtcarteiraconf = TBAbrir!Registro_boleto & "-237-06.conf"
                Case "09 - Com Registro": txtcarteiraconf = TBAbrir!Registro_boleto & "-237-09.conf"
            End Select
            Select Case Len(txtAgencia)
                Case 1: AgenciaBol = "0000-" & txtAgencia
                Case 2: AgenciaBol = "000" & Left(txtAgencia, 1) & "-" & Right(txtAgencia, 1)
                Case 3: AgenciaBol = "00" & Left(txtAgencia, 2) & "-" & Right(txtAgencia, 1)
                Case 4: AgenciaBol = "0" & Left(txtAgencia, 3) & "-" & Right(txtAgencia, 1)
                Case Is >= 5: AgenciaBol = Left(txtAgencia, 4) & "-" & Right(txtAgencia, 1)
            End Select
            Select Case Len(txtContacorrente)
                Case 1: ContaCorrenteBol = "0000000-" & txtContacorrente
                Case 2: ContaCorrenteBol = "000000" & Left(txtContacorrente, 1) & "-" & Right(txtContacorrente, 1)
                Case 3: ContaCorrenteBol = "00000" & Left(txtContacorrente, 2) & "-" & Right(txtContacorrente, 1)
                Case 4: ContaCorrenteBol = "0000" & Left(txtContacorrente, 3) & "-" & Right(txtContacorrente, 1)
                Case 5: ContaCorrenteBol = "000" & Left(txtContacorrente, 4) & "-" & Right(txtContacorrente, 1)
                Case 6: ContaCorrenteBol = "00" & Left(txtContacorrente, 5) & "-" & Right(txtContacorrente, 1)
                Case 7: ContaCorrenteBol = "0" & Left(txtContacorrente, 6) & "-" & Right(txtContacorrente, 1)
                Case Is >= 8: ContaCorrenteBol = Left(txtContacorrente, 7) & "-" & Right(txtContacorrente, 1)
            End Select
            Codigocedente = FunTamanhoTextoZeroEsq(Left(Codigocedente, 15), 15)
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Bradesco"
            Arquivo = "CB" & Dia & Mes & "." & SeqRemessa
            Layout = "CNAB400"
        Case "341": 'Itaú
            Select Case cmbCarteira
                Case "109 - Direta Eletrônica Sem Emissão - Simples": txtcarteiraconf = TBAbrir!Registro_boleto & "-341-109.conf"
                Case "112 - Escritual Eletrônica - simples / contratual": txtcarteiraconf = TBAbrir!Registro_boleto & "-341-112.conf"
                Case "175 - Sem Registro Sem Emissão": txtcarteiraconf = TBAbrir!Registro_boleto & "-341-175.conf"
            End Select
            AgenciaBol = FunTamanhoTextoZeroEsq(Left(txtAgencia, 4), 4)
            Select Case Len(txtContacorrente)
                Case 1: ContaCorrenteBol = "00000-" & txtContacorrente
                Case 2: ContaCorrenteBol = "0000" & Left(txtContacorrente, 1) & "-" & Right(txtContacorrente, 1)
                Case 3: ContaCorrenteBol = "000" & Left(txtContacorrente, 2) & "-" & Right(txtContacorrente, 1)
                Case 4: ContaCorrenteBol = "00" & Left(txtContacorrente, 3) & "-" & Right(txtContacorrente, 1)
                Case 5: ContaCorrenteBol = "0" & Left(txtContacorrente, 4) & "-" & Right(txtContacorrente, 1)
                Case Is >= 6: ContaCorrenteBol = Left(txtContacorrente, 5) & "-" & Right(txtContacorrente, 1)
            End Select
            Codigocedente = Txt_codigo_cedente
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Itaú"
            'Arquivo = Dia & Mes & Right(Ano, 2) & Seq1
            Arquivo = Dia & Mes & Right(ano, 2) & SeqRemessa
            Layout = "CNAB400"
        Case "356": 'ABN e Real
            Select Case cmbCarteira
                Case "20 - Cobrança Simples": txtcarteiraconf = TBAbrir!Registro_boleto & "-356-20.conf"
            End Select
            AgenciaBol = FunTamanhoTextoZeroEsq(Left(txtAgencia, 4), 4)
            Select Case Len(txtContacorrente)
                Case 1: ContaCorrenteBol = "000000-" & txtContacorrente
                Case 2: ContaCorrenteBol = "00000" & Left(txtContacorrente, 1) & "-" & Right(txtContacorrente, 1)
                Case 3: ContaCorrenteBol = "0000" & Left(txtContacorrente, 2) & "-" & Right(txtContacorrente, 1)
                Case 4: ContaCorrenteBol = "000" & Left(txtContacorrente, 3) & "-" & Right(txtContacorrente, 1)
                Case 5: ContaCorrenteBol = "00" & Left(txtContacorrente, 4) & "-" & Right(txtContacorrente, 1)
                Case 6: ContaCorrenteBol = "0" & Left(txtContacorrente, 5) & "-" & Right(txtContacorrente, 1)
                Case Is > 7: ContaCorrenteBol = Left(txtContacorrente, 6) & "-" & Right(txtContacorrente, 1)
            End Select
            Codigocedente = FunTamanhoTextoZeroEsq(Left(Codigocedente, 9), 9)
            Diretorio = Localrel & "\Boletos\Arquivos remessa\ABN e Real"
            Arquivo = "CB" & Dia & Mes & "." & SeqRemessa
            Layout = "CNAB400"
        Case "399": 'HSBC
            Select Case cmbCarteira
                Case "CNR - Sem Registro": txtcarteiraconf = TBAbrir!Registro_boleto & "-399-CNR.conf"
            End Select
            Codigocedente = FunTamanhoTextoZeroEsq(Left(Codigocedente, 7), 7)
            Diretorio = Localrel & "\Boletos\Arquivos remessa\HSBC"
            Arquivo = "D" & Dia & Mes & ano & "." & SeqRemessa
            Layout = "CNAB400"
        Case "409": 'Unibanco
            Select Case cmbCarteira
                Case "Especial": txtcarteiraconf = TBAbrir!Registro_boleto & "-409-ESPECIAL.conf"
            End Select
            AgenciaBol = FunTamanhoTextoZeroEsq(Left(txtAgencia, 4), 4)
            Select Case Len(txtContacorrente)
                Case 1: ContaCorrenteBol = "000" & "." & "000" & "-" & txtContacorrente
                Case 2: ContaCorrenteBol = "000" & "." & "00" & Left(txtContacorrente, 1) & "-" & Mid(txtContacorrente, 2, 1)
                Case 3: ContaCorrenteBol = "000" & "." & "0" & Left(txtContacorrente, 2) & "-" & Mid(txtContacorrente, 3, 1)
                Case 4: ContaCorrenteBol = "000" & "." & Left(txtContacorrente, 3) & "-" & Mid(txtContacorrente, 4, 1)
                Case 5: ContaCorrenteBol = "00" & Left(txtContacorrente, 1) & "." & Mid(txtContacorrente, 2, 3) & "-" & Mid(txtContacorrente, 5, 1)
                Case 6: ContaCorrenteBol = "0" & Left(txtContacorrente, 2) & "." & Mid(txtContacorrente, 3, 3) & "-" & Mid(txtContacorrente, 6, 1)
                Case Is >= 7: ContaCorrenteBol = Left(txtContacorrente, 3) & "." & Mid(txtContacorrente, 4, 3) & "-" & Mid(txtContacorrente, 7, 1)
            End Select
            Codigocedente = ContaCorrente
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Unibanco"
            Arquivo = "CBR" & Dia & Mes & "." & SeqRemessa
            Layout = "CNAB240"
    End Select
End If

End Sub

Public Sub ProcCarregaDados()
On Error GoTo tratar_erro

If IsNull(TBAbrir!ID_empresa) = False And TBAbrir!ID_empresa <> "" Then ProcPuxaDadosComboEmpresa Cmb_empresa, TBAbrir!ID_empresa
txtCodBanco = TBAbrir!ID
Txt_IDBanco.Text = TBAbrir!ID

txtData1 = IIf(IsNull(TBAbrir!data), "", Format(TBAbrir!data, "dd/mm/yy"))
txtResponsavel = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
txtDtValidacao = IIf(IsNull(TBAbrir!DtValidacao), "", TBAbrir!DtValidacao)
txtRespValidacao = IIf(IsNull(TBAbrir!RespValidacao), "", TBAbrir!RespValidacao)
cmbFamilia = IIf(IsNull(TBAbrir!Txt_familia), "", TBAbrir!Txt_familia)
txtNBanco = IIf(IsNull(TBAbrir!int_NBanco), "", TBAbrir!int_NBanco)
txtAgencia = IIf(IsNull(TBAbrir!txt_Agencia), "", TBAbrir!txt_Agencia)
Txt_codigo_cedente = IIf(IsNull(TBAbrir!codigo_cedente), "", TBAbrir!codigo_cedente)
Txt_codigo_cedente1 = IIf(IsNull(TBAbrir!Codigo_cedente_registrado), "", TBAbrir!Codigo_cedente_registrado)
Txt_nome_agencia = IIf(IsNull(TBAbrir!Nome_agencia), "", TBAbrir!Nome_agencia)
txtDescricao.Text = IIf(IsNull(TBAbrir!Txt_descricao), "", TBAbrir!Txt_descricao)
Caption = "Administrativo - Financeiro - Instituições - (Instituição bancária : " & TBAbrir!Txt_descricao & ")"
txtConta = IIf(IsNull(TBAbrir!txt_conta), "", TBAbrir!txt_conta)
txtgerente.Text = IIf(IsNull(TBAbrir!txt_gerente), "", TBAbrir!txt_gerente)
txtFone = IIf(IsNull(TBAbrir!txt_fone), "", TBAbrir!txt_fone)
txtFAX = IIf(IsNull(TBAbrir!Txt_fax), "", TBAbrir!Txt_fax)
txtStatus = IIf(TBAbrir!Bloqueado = True, "Bloqueada", "Liberada")

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Usuarios_setor.* from Usuarios_setor where ID = " & IIf(IsNull(TBAbrir!ID_CC), 0, TBAbrir!ID_CC), Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    If IsNull(TBFI!CODIGO) = False And TBFI!CODIGO <> "" Then
        If IsNull(TBFI!DtBloq) = False Then
            Cmb_centro.AddItem TBFI!CODIGO & " - " & IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
            Cmb_centro.ItemData(Cmb_centro.NewIndex) = TBFI!ID
        End If
        Cmb_centro = TBFI!CODIGO & " - " & IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
    Else
        If IsNull(TBFI!DtBloq) = False Then
            Cmb_centro.AddItem IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
            Cmb_centro.ItemData(Cmb_centro.NewIndex) = TBFI!ID
        End If
        Cmb_centro = IIf(IsNull(TBFI!Setor), "", TBFI!Setor)
    End If
End If
TBFI.Close

txtobs = IIf(IsNull(TBAbrir!Txt_obs), "", TBAbrir!Txt_obs)
txtsaldo = IIf(IsNull(TBAbrir!Saldo), "0,00", Format(TBAbrir!Saldo, "###,##0.00"))
txtLimite = IIf(IsNull(TBAbrir!Limite_desconto), "0,00", Format(TBAbrir!Limite_desconto, "###,##0.00"))
txtUtilizado = IIf(IsNull(TBAbrir!Limite_utilizado), "0,00", Format(TBAbrir!Limite_utilizado, "###,##0.00"))
Novo_Banco = False
Frame2.Enabled = True
ProcLimparTudo

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lst_saque_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lst_saque
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If txtDtValidacao = "" Then GoTo Proximo
                If txtStatus = "Bloqueada" Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lst_saque, ColumnHeader
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lst_saque_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lst_saque
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If txtDtValidacao = "" Then
                MsgBox "Não é possivel excluir a movimentação, pois a instituição ainda não foi validada.", vbExclamation
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If txtStatus = "Bloqueada" Then
                MsgBox "Não é possivel alterar a movimentação, pois a instituição esta bloqueada.", vbExclamation
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lst_saque_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lst_saque.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_instituicoes_transf where id_transf = " & Lst_saque.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCamposSaque
    Txt_id_saque = TBLISTA!id_transf
    If IsNull(TBLISTA!data_transf) = False And TBLISTA!data_transf <> "" Then txtdata2.Value = TBLISTA!data_transf
    txtResponsavel2 = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
    
    'Fluxo de Caixa
    Set TBFluxo = CreateObject("adodb.recordset")
    TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBLISTA!IDFluxo), 0, TBLISTA!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
    If TBFluxo.EOF = False Then
        txtObsFluxo1 = IIf(IsNull(TBFluxo!Obs), "", TBFluxo!Obs)
    End If
    TBFluxo.Close
    
    Txt_valor = IIf(IsNull(TBLISTA!valor_transf), "", Format(TBLISTA!valor_transf, "###,##0.00"))
End If
TBLISTA.Close
ProcCarregaListaContas
Frame8.Enabled = True
Novo_Banco2 = False
CodigoLista2 = Lst_saque.SelectedItem.Index

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCarregaListaContas()
On Error GoTo tratar_erro

Valor_total = 0
Lst_Contas.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select tbl_contaspagar.* from tbl_contaspagar INNER JOIN tbl_ContasPagar_Saque ON tbl_contaspagar.IDintconta = tbl_ContasPagar_Saque.IDintconta where tbl_ContasPagar_Saque.IDSaque = " & Lst_saque.SelectedItem & " and tbl_contaspagar.logsit = 'S' order by tbl_contaspagar.DataBaixa desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista(2).Min = 0
    PBLista(2).Max = TBLISTA.RecordCount
    PBLista(2).Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lst_Contas.ListItems
            .Add , , TBLISTA!IDintconta
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!DataBaixa), "", Format(TBLISTA!DataBaixa, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Txt_fornecedor), "", TBLISTA!Txt_fornecedor)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!dbl_valorpagto), "", Format(TBLISTA!dbl_valorpagto, "###,##0.00"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!ValorPago), "", Format(TBLISTA!ValorPago, "###,##0.00"))
            
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select sum(Valor) as Qtde from tbl_contas_antecipacao where ID_conta = " & TBLISTA!IDintconta & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                .Item(.Count).SubItems(5) = IIf(IsNull(TBContas!Qtde), "", Format(TBContas!Qtde, "###,##0.00"))
            End If
            TBContas.Close
            
            Valor_total = Valor_total + IIf(IsNull(TBLISTA!ValorPago), 0, TBLISTA!ValorPago)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista(2).Value = Contador
     Loop
End If
TBLISTA.Close
LblValortotal.Caption = "Valor total pago = " & Format(Valor_total, "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lst_tarifa_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lst_tarifa
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If txtDtValidacao = "" Then GoTo Proximo
                If txtStatus = "Bloqueada" Then GoTo Proximo
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lst_tarifa, ColumnHeader
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lst_tarifa_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lst_tarifa
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If txtDtValidacao = "" Then
                MsgBox "Não é possivel excluir a movimentação, pois a instituição ainda não foi validada.", vbExclamation
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If txtStatus = "Bloqueada" Then
                MsgBox "Não é possivel excluir a movimentação, pois a instituição esta bloqueada.", vbExclamation
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lst_tarifa_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lst_tarifa.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select id_transf, IDintconta, Tipo from tbl_instituicoes_transf where id_transf = " & Lst_tarifa.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcLimpaCamposTarifa
    Txt_id_tarifa = TBAbrir!id_transf
    Set TBLISTA = CreateObject("adodb.recordset")
    If TBAbrir!Tipo = "P" Then
        TBLISTA.Open "Select DataBaixa, resppag, ValorPago, Class_conta, FormaBaixa, IDFluxo from tbl_ContasPagar where idintconta = " & TBAbrir!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
        If TBLISTA.EOF = False Then
            txtdata3.Value = TBLISTA!DataBaixa
            txtResponsavel3 = IIf(IsNull(TBLISTA!resppag), "", TBLISTA!resppag)
            Cmb_operacao = "Débito"
            Txt_valor1 = IIf(IsNull(TBLISTA!ValorPago), "", Format(TBLISTA!ValorPago, "###,##0.00"))
            
            NomeCampo = "o tipo do documento"
            If IsNull(TBLISTA!Class_conta) = False And TBLISTA!Class_conta <> "" Then Cmb_tipo.Text = TBLISTA!Class_conta
            NomeCampo = "a forma da baixa"
            If IsNull(TBLISTA!FormaBaixa) = False And TBLISTA!FormaBaixa <> "" Then cmb_forma1 = TBLISTA!FormaBaixa
         End If
    Else
        TBLISTA.Open "Select Data_pagamento, resprec, valortitulorecebido, Tipo_doc, FormaBaixa, IDFluxo from tbl_contas_receber where idintconta = " & TBAbrir!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
        If TBLISTA.EOF = False Then
            txtdata3.Value = TBLISTA!Data_pagamento
            txtResponsavel3 = IIf(IsNull(TBLISTA!resprec), "", TBLISTA!resprec)
            Cmb_operacao = "Crédito"
            Txt_valor1 = IIf(IsNull(TBLISTA!valortitulorecebido), "", Format(TBLISTA!valortitulorecebido, "###,##0.00"))
            
            NomeCampo = "o tipo do documento"
            If IsNull(TBLISTA!Tipo_doc) = False And TBLISTA!Tipo_doc <> "" Then Cmb_tipo.Text = TBLISTA!Tipo_doc
            NomeCampo = "a forma da baixa"
            If IsNull(TBLISTA!FormaBaixa) = False And TBLISTA!FormaBaixa <> "" Then cmb_forma1 = TBLISTA!FormaBaixa
        End If
    End If

1:
    'Conta contábil
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select F.int_codfamilia, F.CODIGO, F.Txt_descricao from (Familia_financeiro FF INNER JOIN tbl_instituicoes_transf IT ON IT.IDintconta = FF.IDConta) INNER JOIN tbl_familia F ON FF.ID_PC = F.int_codfamilia where FF.IDconta = " & TBAbrir!IDintconta & " and FF.TipoConta = '" & TBAbrir!Tipo & "' and FF.Deposito_transf = 'False'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Txt_ID_PC = IIf(IsNull(TBFI!int_codfamilia), "", TBFI!int_codfamilia)
        Txt_codigo_PC = IIf(IsNull(TBFI!CODIGO), "", TBFI!CODIGO)
        Txt_descricao_PC = IIf(IsNull(TBFI!Txt_descricao), "", TBFI!Txt_descricao)
    End If
    TBFI.Close
    
    'Fluxo de Caixa
    Set TBFluxo = CreateObject("adodb.recordset")
    TBFluxo.Open "Select Obs from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBLISTA!IDFluxo), 0, TBLISTA!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
    If TBFluxo.EOF = False Then
        txtObsFluxo2 = IIf(IsNull(TBFluxo!Obs), "", TBFluxo!Obs)
    End If
    TBFluxo.Close
End If
TBAbrir.Close
Frame4.Enabled = True
Novo_Banco3 = False
CodigoLista3 = Lst_tarifa.SelectedItem.Index

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        MsgBox ("Não foi encontrado " & NomeCampo & " desta tarifa."), vbExclamation
        GoTo 1
    End If
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub lst_transferencias_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With lst_transferencias
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If txtDtValidacao = "" Then GoTo Proximo
                If txtStatus = "Bloqueada" Then GoTo Proximo
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select * from tbl_instituicoes_transf where id_transf = " & .ListItems(InitFor) & " and Tipo = 'D' and FormaBaixa = 'CHEQUE'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = False Then
                    Cheque = "Cheque n. " & TBFIltro!NDoctoBaixa
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from tbl_Fluxo_de_caixa where Instituicao = '" & txtDescricao & "' and Descricao = '" & Cheque & "' and Bloqueado = 'False'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then GoTo Proximo
                End If
                TBFIltro.Close
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView lst_transferencias, ColumnHeader
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub lst_transferencias_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With lst_transferencias
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If txtDtValidacao = "" Then
                MsgBox "Não é possivel excluir a movimentação, pois a instituição ainda não foi validada.", vbExclamation
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            If txtStatus = "Bloqueada" Then
                MsgBox "Não é possivel excluir a movimentação, pois a instituição esta bloqueada.", vbExclamation
                .ListItems.Item(InitFor).Checked = False
                Exit Sub
            End If
            Set TBFIltro = CreateObject("adodb.recordset")
            TBFIltro.Open "Select * from tbl_instituicoes_transf where id_transf = " & .ListItems(InitFor) & " and Tipo = 'D' and FormaBaixa = 'CHEQUE'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFIltro.EOF = False Then
                Cheque = "Cheque n. " & TBFIltro!NDoctoBaixa
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from tbl_Fluxo_de_caixa where Instituicao = '" & txtDescricao & "' and Descricao = '" & Cheque & "' and Bloqueado = 'False'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    MsgBox ("Não é permitido excluir este depósito em cheque, pois o mesmo já está compensado."), vbExclamation
                    .ListItems.Item(InitFor).Checked = False
                End If
                TBAbrir.Close
            End If
            TBFIltro.Close
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub lst_transferencias_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lst_transferencias.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from tbl_instituicoes_transf where id_transf = " & lst_transferencias.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCamposTransf
    Tipo = IIf(IsNull(TBLISTA!Tipo), "", TBLISTA!Tipo)
    If Tipo = "T" Then OptTransferencia.Value = True
    If Tipo = "D" Then OptDeposito.Value = True
    txtid = TBLISTA!id_transf
    If IsNull(TBLISTA!FormaBaixa) = False And TBLISTA!FormaBaixa <> "" And Tipo <> "" Then cmb_forma = TBLISTA!FormaBaixa
    NomeCampo = "a instituição bancária recebedora"
    If IsNull(TBLISTA!banco_recebedor) = False And TBLISTA!banco_recebedor <> "" Then cmbrecebedor = TBLISTA!banco_recebedor
1:
    mskvalor.Text = IIf(IsNull(TBLISTA!valor_transf), "", Format(TBLISTA!valor_transf, "###,##0.00"))
    If IsNull(TBLISTA!data_transf) = False And TBLISTA!data_transf <> "" Then txtdata.Value = TBLISTA!data_transf
    txtResponsavel1 = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
    txtCheque = IIf(IsNull(TBLISTA!NDoctoBaixa), "", TBLISTA!NDoctoBaixa)
    
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select Favorecido from tbl_ContasPagar where NDoctoBaixa = '" & txtCheque & "' and Banco = '" & txtDescricao & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        txtfavorecido = IIf(IsNull(TBContas!Favorecido), "", TBContas!Favorecido)
    End If
    TBContas.Close
    
    'Conta contábil
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select F.int_codfamilia, F.CODIGO, F.Txt_descricao from Familia_financeiro FF INNER JOIN tbl_familia F ON FF.ID_PC = F.int_codfamilia where FF.IDconta = " & TBLISTA!id_transf & " and FF.Tipoconta = 'P' and FF.Deposito_transf = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Txt_ID_PC_instituicao = IIf(IsNull(TBFI!int_codfamilia), 0, TBFI!int_codfamilia)
        Txt_codigo_PC_instituicao = IIf(IsNull(TBFI!CODIGO), "", TBFI!CODIGO)
        Txt_descricao_PC_instituicao = IIf(IsNull(TBFI!Txt_descricao), "", TBFI!Txt_descricao)
    End If
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select F.int_codfamilia, F.CODIGO, F.Txt_descricao from Familia_financeiro FF INNER JOIN tbl_familia F ON FF.ID_PC = F.int_codfamilia where FF.IDconta = " & TBLISTA!id_transf & " and FF.Tipoconta = 'R' and FF.Deposito_transf = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Txt_ID_PC_instituicao_rec = IIf(IsNull(TBFI!int_codfamilia), 0, TBFI!int_codfamilia)
        Txt_codigo_PC_instituicao_rec = IIf(IsNull(TBFI!CODIGO), "", TBFI!CODIGO)
        Txt_descricao_PC_instituicao_rec = IIf(IsNull(TBFI!Txt_descricao), "", TBFI!Txt_descricao)
    End If
    TBFI.Close
End If
TBLISTA.Close
frm_filtro.Enabled = True
Novo_Banco1 = False
CodigoLista1 = lst_transferencias.SelectedItem.Index

Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        MsgBox ("Não foi encontrado " & NomeCampo & " desse registro."), vbExclamation
        GoTo 1
    End If
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

Lst_extrato.ListItems.Clear

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

Lst_extrato.ListItems.Clear

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub mskvalor_Change()
On Error GoTo tratar_erro

If mskvalor.Text <> "" Then
    VerifNumero = mskvalor.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        mskvalor.Text = ""
        mskvalor.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcCarregaListaTransf()
On Error GoTo tratar_erro

If Instituicao_Localizar_Transf = "" Then Exit Sub
lst_transferencias.ListItems.Clear
valor = 0
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open Instituicao_Localizar_Transf, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista(1).Min = 0
    PBLista(1).Max = TBLISTA.RecordCount
    PBLista(1).Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With lst_transferencias.ListItems
            .Add , , TBLISTA!id_transf
                .Item(.Count).SubItems(1) = Format(TBLISTA!data_transf, "dd/mm/yy")
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            If IsNull(TBLISTA!Tipo) = False And TBLISTA!Tipo <> "" Then
                Select Case TBLISTA!Tipo
                    Case "D": .Item(.Count).SubItems(3) = "Depósito"
                    Case "T": .Item(.Count).SubItems(3) = "Transferência"
                End Select
            End If
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!banco_remetente), "", TBLISTA!banco_remetente)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!banco_recebedor), "", TBLISTA!banco_recebedor)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!valor_transf), "", Format(TBLISTA!valor_transf, "###,##0.00"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!id_banco_rec), "", TBLISTA!id_banco_rec)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!id_banco_rem), "", TBLISTA!id_banco_rem)
        End With
        valor = valor + TBLISTA!valor_transf
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista(1).Value = Contador
    Loop
End If
TBLISTA.Close
Txt_vlr_total_deptran = Format(valor, "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcCarregaListaSaque()
On Error GoTo tratar_erro

If Instituicao_Localizar_Saque = "" Then Exit Sub
Valor1 = 0
Valor2 = 0
Valor3 = 0
Valor_total = 0
Lst_saque.ListItems.Clear
Lst_Contas.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open Instituicao_Localizar_Saque, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista(2).Min = 0
    PBLista(2).Max = TBLISTA.RecordCount
    PBLista(2).Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        'Verifica se o saldo do saque é maior que zero
        Valor_total = 0
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Sum(Valor_utilizado) as Valor_Total from tbl_ContasPagar_Saque where IDSaque = " & TBLISTA!id_transf, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Valor_total = IIf(IsNull(TBAbrir!Valor_total), 0, TBAbrir!Valor_total)
        End If
        TBAbrir.Close
        
        With Lst_saque.ListItems
            .Add , , TBLISTA!id_transf
            .Item(.Count).SubItems(1) = Format(TBLISTA!data_transf, "dd/mm/yy")
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!valor_transf), 0, Format(TBLISTA!valor_transf, "###,##0.00"))
            .Item(.Count).SubItems(3) = Format(Valor_total, "###,##0.00")
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!valor_transf), "", Format(TBLISTA!valor_transf - Valor_total, "###,##0.00"))
            Valor1 = Valor1 + IIf(IsNull(TBLISTA!valor_transf), 0, TBLISTA!valor_transf) 'Valor Saque
            Valor2 = Valor2 + Valor_total 'Valor utilizado
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista(2).Value = Contador
    Loop
End If
TBLISTA.Close
Valor3 = Valor1 - Valor2 'Valor saque - Valor utilizado(saldo)
TxtDisponivel.Text = Format(Valor1, "###,##0.00")
TxtValorSaqueUtilizado.Text = Format(Valor2, "###,##0.00")
TxtSaldoSaque.Text = Format(Valor3, "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcCarregaListaTarifa()
On Error GoTo tratar_erro

valor = 0
Valor1 = 0
Lst_tarifa.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
If Instituicao_Localizar_Tarifa = "" Then Exit Sub
TBLISTA.Open Instituicao_Localizar_Tarifa, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista(3).Min = 0
    PBLista(3).Max = TBLISTA.RecordCount
    PBLista(3).Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lst_tarifa.ListItems
            .Add , , TBLISTA!id_transf
            If TBLISTA!Tipo = "P" Then
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select * from tbl_ContasPagar where IDintconta = " & TBLISTA!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = False Then
                    .Item(.Count).SubItems(1) = Format(TBFIltro!DataBaixa, "dd/mm/yy")
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBFIltro!resppag), "", TBFIltro!resppag)
                    .Item(.Count).SubItems(3) = "Débito"
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBFIltro!ValorPago), "", Format(TBFIltro!ValorPago, "###,##0.00"))
                    valor = valor + IIf(IsNull(TBFIltro!ValorPago), 0, TBFIltro!ValorPago)
                End If
            Else
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select * from tbl_contas_receber where IDintconta = " & TBLISTA!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = False Then
                    .Item(.Count).SubItems(1) = Format(TBFIltro!Data_pagamento, "dd/mm/yy")
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBFIltro!resprec), "", TBFIltro!resprec)
                    .Item(.Count).SubItems(3) = "Crédito"
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBFIltro!valortitulorecebido), "", Format(TBFIltro!valortitulorecebido, "###,##0.00"))
                    Valor1 = Valor1 + IIf(IsNull(TBFIltro!valortitulorecebido), 0, TBFIltro!valortitulorecebido)
                End If
            End If
            TBFIltro.Close
            
            'Conta contábil
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select F.CODIGO, F.Txt_descricao from (Familia_financeiro FF INNER JOIN tbl_instituicoes_transf IT ON IT.IDintconta = FF.IDConta) INNER JOIN tbl_familia F ON FF.ID_PC = F.int_codfamilia where FF.IDconta = " & TBLISTA!IDintconta & " and FF.TipoConta = '" & TBLISTA!Tipo & "' and FF.Deposito_transf = 'False'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                .Item(.Count).SubItems(4) = IIf(IsNull(TBFI!CODIGO), "", TBFI!CODIGO)
                .Item(.Count).SubItems(5) = IIf(IsNull(TBFI!Txt_descricao), "", TBFI!Txt_descricao)
            End If
            TBFI.Close
            
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista(3).Value = Contador
    Loop
End If
TBLISTA.Close
Txt_valor_total_tarifas = Format(valor, "###,##0.00")
Txt_valor_total_tarifas1 = Format(Valor1, "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCarregaListaExtrato()
On Error GoTo tratar_erro

Lst_extrato.ListItems.Clear
TotalCredito = 0
TotalDebito = 0

With Lst_extrato.ListItems
    .Add , , ""
    .Item(.Count).SubItems(1) = Format(Dataini, "dd/mm/yy")
    .Item(.Count).SubItems(2) = "SALDO"
    .Item(.Count).SubItems(3) = ""
    .Item(.Count).SubItems(4) = Format(Saldo_Anterior, "###,##0.00")
End With

If TBLISTA.EOF = False Then
    data = TBLISTA!data
    PBLista(4).Min = 0
    PBLista(4).Max = TBLISTA.RecordCount
    PBLista(4).Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        TBLISTA!Saldo_Ant = Format(Saldo_Anterior, "###,##0.00")
        With Lst_extrato.ListItems
            
            If data <> TBLISTA!data Then
                TBLISTA.MovePrevious
                .Add , , ""
                .Item(.Count).SubItems(1) = Format(TBLISTA!data, "dd/mm/yy")
                .Item(.Count).SubItems(2) = "SALDO"
                .Item(.Count).SubItems(3) = ""
                .Item(.Count).SubItems(4) = Format(Saldo_Anterior, "###,##0.00")
                TBLISTA.MoveNext
            End If
            
            .Add , , TBLISTA!IDFluxo
            .Item(.Count).SubItems(1) = Format(TBLISTA!data, "dd/mm/yy")
                        
            If TBLISTA!Operacao = "Crédito" Then
                TabelaFiltro = "tbl_Contas_receber"
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
            Else
                TabelaFiltro = "tbl_ContasPagar"
                If TBLISTA!valor >= 0 Then valor = "-" & TBLISTA!valor Else valor = TBLISTA!valor * -1
                .Item(.Count).SubItems(3) = Format(valor, "###,##0.00")
            End If
            
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from " & TabelaFiltro & " where IDFluxo = " & TBLISTA!IDFluxo & " and (Antecipacao = 'True' or Devolucao = 'True')", Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                If TBContas!Antecipacao = True Then Texto = " (ANTECIPAÇÃO)" Else Texto = " (DEVOLUÇÃO)"
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Obs), "", TBLISTA!Obs) & Texto
            Else
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Obs), "", TBLISTA!Obs)
            End If
            TBContas.Close
        
            If TBLISTA!Operacao = "Crédito" Then
                Saldo_Anterior = Format(Saldo_Anterior + IIf(IsNull(TBLISTA!valor), "", TBLISTA!valor))
            Else
                Saldo_Anterior = Format(Saldo_Anterior - IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor))
            End If
            TBLISTA!Saldo_Atual = Format(Saldo_Anterior, "###,##0.00")
            .Item(.Count).SubItems(4) = ""
        End With
        TBLISTA.Update
        data = TBLISTA!data
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista(4).Value = Contador
    Loop
Else
    MsgBox ("Não existem movimentações neste período."), vbInformation
    NomeRel = "Instituicoes_extrato bancario_saldos.rpt"
    FormulaRel_Instituicao1 = "{tbl_Instituicoes.txt_Descricao} = '" & txtDescricao & "' and {tbl_contas_receber.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
End If
TBLISTA.Close

With Lst_extrato.ListItems
    .Add , , ""
    .Item(.Count).SubItems(1) = Format(data, "dd/mm/yy")
    .Item(.Count).SubItems(2) = "SALDO"
    .Item(.Count).SubItems(3) = ""
    .Item(.Count).SubItems(4) = Format(Saldo_Anterior, "###,##0.00")
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcCarregaListaCheque()
On Error GoTo tratar_erro

If Cheques_Emitidos = True Then
    Quant = 0
    valor = 0
    quantidade = 0
    Valor_total = 0
    Cheque = ""
    ChequeC = ""
    Lst_cheque.ListItems.Clear
    Lst_cheque1.ListItems.Clear
    Conexao.Execute "DELETE from Cheques_Relatorios"
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from Cheques_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
    If StrSql_Instituicoes_Localizar_Cheque <> "" Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open StrSql_Instituicoes_Localizar_Cheque, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                If TBAbrir!NDoctoBaixa <> "" Then
                    TBGravar.AddNew
                    ProcEnviaDadosCheque
                    TBGravar!Tipo = 1
                    TBGravar.Update
                End If
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
    End If
    If StrSql_Instituicoes_Localizar_Cheque_Cancelados <> "" Then
        'Grava cheques cancelados
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open StrSql_Instituicoes_Localizar_Cheque_Cancelados, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                If TBAbrir!NDoctoBaixa <> "" Then
                    TBGravar.AddNew
                    ProcEnviaDadosCheque
                    TBGravar!Tipo = 2
                    TBGravar.Update
                End If
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
    End If
       
    'Carrega Lista
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select * from Cheques_Relatorios order by Tipo, Data, Cheque, Fornecedor", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        PBLista(5).Min = 0
        PBLista(5).Max = TBLISTA.RecordCount
        PBLista(5).Value = 1
        Contador = 0
        Do While TBLISTA.EOF = False
            If TBLISTA!Tipo = 1 Then
                With Lst_cheque.ListItems
                    .Add , , TBLISTA!ID_conta
                    .Item(.Count).SubItems(1) = Format(TBLISTA!data, "dd/mm/yy")
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Cheque), "", TBLISTA!Cheque)
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Fornecedor), "", Trim(TBLISTA!Fornecedor))
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Obs), "", TBLISTA!Obs)
                    If TBLISTA!Compensado = True Then .Item(.Count).SubItems(6) = "Sim" Else .Item(.Count).SubItems(6) = "Não"
                End With
                If Cheque <> IIf(IsNull(TBLISTA!Cheque), "", TBLISTA!Cheque) Then Quant = Quant + 1
                valor = valor + IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor)
                Cheque = IIf(IsNull(TBLISTA!Cheque), "", TBLISTA!Cheque)
                Permitido = True
            Else
                With Lst_cheque1.ListItems
                    .Add , , TBLISTA!ID_conta
                    .Item(.Count).SubItems(1) = Format(TBLISTA!data, "dd/mm/yy")
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Cheque), "", TBLISTA!Cheque)
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Fornecedor), "", TBLISTA!Fornecedor)
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Obs), "", TBLISTA!Obs)
                    Set TBFIltro = CreateObject("adodb.recordset")
                    TBFIltro.Open "Select * from Cheques_Cancelados where ID_conta = " & TBLISTA!ID_conta, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFIltro.EOF = False Then
                        .Item(.Count).SubItems(6) = IIf(IsNull(TBFIltro!Data_cancelamento), "", Format(TBFIltro!Data_cancelamento, "dd/mm/yy"))
                        .Item(.Count).SubItems(7) = IIf(IsNull(TBFIltro!Responsavel), "", TBFIltro!Responsavel)
                        .Item(.Count).SubItems(8) = IIf(IsNull(TBFIltro!motivo), "", Trim(TBFIltro!motivo))
                    End If
                    TBFIltro.Close
                End With
                If ChequeC <> IIf(IsNull(TBLISTA!Cheque), "", TBLISTA!Cheque) Then quantidade = quantidade + 1
                Valor_total = Valor_total + IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor)
                ChequeC = IIf(IsNull(TBLISTA!Cheque), "", TBLISTA!Cheque)
                Permitido = False
            End If
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista(5).Value = Contador
        Loop
    End If
    TBLISTA.Close
    'Carrega Totais
    Txt_qtde_ativo = Quant
    Txt_qtde_cancelado = quantidade
    Txt_qtde_total = Quant + quantidade
    Txt_valor_ativo = Format(valor, "###,##0.00")
    Txt_valor_cancelado = Format(Valor_total, "###,##0.00")
    Txt_valor_total = Format(valor + Valor_total, "###,##0.00")
Else
    Lista_cheque.ListItems.Clear
    BomPara = ""
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open StrSql_Instituicoes_Localizar_Cheque_Recebidos, Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        PBLista(5).Min = 0
        PBLista(5).Max = TBLISTA.RecordCount
        PBLista(5).Value = 1
        Contador = 0
        Do While TBLISTA.EOF = False
            With Lista_cheque.ListItems
                .Add , , TBLISTA!IDintconta
                .Item(.Count).SubItems(1) = Format(TBLISTA!Data_pagamento, "dd/mm/yy")
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!NDoctoBaixa), "", TBLISTA!NDoctoBaixa)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Nome_Razao), "", TBLISTA!Nome_Razao)
                If TBLISTA!status = "TÍTULO LIQUIDADO ANTECIPADO" Then
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                Else
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!valortitulorecebido), "", Format(TBLISTA!valortitulorecebido, "###,##0.00"))
                End If
                If TBLISTA!FormaBaixa = "CHEQUE PRÉ-DATADO" Then
                    If IsNull(TBLISTA!Bom_para) = False And TBLISTA!Bom_para <> "" Then BomPara = "- BOM PARA: " & Format(TBLISTA!Bom_para, "dd/mm/yy")
                    If IsNull(TBLISTA!Obs) = False And TBLISTA!Obs <> "" Then .Item(.Count).SubItems(5) = Trim(TBLISTA!Obs) & " (PRÉ-DATADO) " & BomPara Else .Item(.Count).SubItems(5) = "(PRÉ-DATADO) " & BomPara
                Else
                    If IsNull(TBLISTA!Obs) = False And TBLISTA!Obs <> "" Then .Item(.Count).SubItems(5) = Trim(TBLISTA!Obs)
                End If
                'Verifica se o cheque já foi compensado
                Cheque = "Cheque n. " & TBLISTA!NDoctoBaixa
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select * from tbl_Fluxo_de_caixa where Operacao = 'Crédito' and Instituicao = '" & txtDescricao & "' and Descricao = '" & Cheque & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = False Then
                    If TBFIltro!Bloqueado = True Then .Item(.Count).SubItems(6) = "Não" Else .Item(.Count).SubItems(6) = "Sim"
                Else
                    .Item(.Count).SubItems(6) = "N/C"
                End If
                TBFIltro.Close
            End With
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista(5).Value = Contador
        Loop
    End If
    TBLISTA.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcEnviaDadosCheque()
On Error GoTo tratar_erro

BomPara = ""
MotivoTexto = ""
If TBAbrir!status = "CHEQUE CANCELADO" Then
    Set TBContas = CreateObject("adodb.recordset")
    TBContas.Open "Select Motivo from Cheques_Cancelados where ID_conta = " & TBAbrir!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
    If TBContas.EOF = False Then
        MotivoTexto = IIf(IsNull(TBContas!motivo), "", TBContas!motivo)
    End If
End If

TBGravar!ID_conta = TBAbrir!IDintconta
If TBAbrir!FormaBaixa = "CHEQUE PRÉ-DATADO" Then
    If IsNull(TBAbrir!Bom_para) = False And TBAbrir!Bom_para <> "" Then BomPara = "- BOM PARA: " & Format(TBAbrir!Bom_para, "dd/mm/yy")
    If TBAbrir!status = "CHEQUE CANCELADO" Then
        TBGravar!Obs = MotivoTexto & " (PRÉ-DATADO) " & BomPara
    Else
        If IsNull(TBAbrir!Obs) = False And TBAbrir!Obs <> "" Then TBGravar!Obs = Trim(TBAbrir!Obs) & " (PRÉ-DATADO) " & BomPara Else TBGravar!Obs = "(PRÉ-DATADO) " & BomPara
    End If
Else
    If TBAbrir!status = "CHEQUE CANCELADO" Then
        TBGravar!Obs = MotivoTexto
    Else
        If IsNull(TBAbrir!Obs) = False And TBAbrir!Obs <> "" Then TBGravar!Obs = Trim(TBAbrir!Obs)
    End If
End If
TBGravar!data = TBAbrir!DataBaixa
TBGravar!Cheque = TBAbrir!NDoctoBaixa
If TBAbrir!status = "TÍTULO LIQUIDADO ANTECIPADO" Then
    TBGravar!valor = TBAbrir!dbl_valorpagto
Else
    TBGravar!valor = TBAbrir!ValorPago
End If

If TBAbrir!status = "CHEQUE CANCELADO" Then
    TBGravar!Fornecedor = TBAbrir!Txt_fornecedor & " (CANCELADO)"
Else
    TBGravar!Fornecedor = TBAbrir!Txt_fornecedor
    'Verifica se o cheque já foi compensado
    Cheque = "Cheque n. " & TBAbrir!NDoctoBaixa
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select * from tbl_Fluxo_de_caixa where Operacao = 'Débito' and Instituicao = '" & txtDescricao & "' and Descricao = '" & Cheque & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        If TBFIltro!Bloqueado = True Then TBGravar!Compensado = False Else TBGravar!Compensado = True
    End If
    TBFIltro.Close
End If
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVisualizarContas()
On Error GoTo tratar_erro

If Lst_extrato.ListItems.Count = 0 Or Lst_extrato.SelectedItem = "" Then Exit Sub
frm_Instituicoes2_lista_contas.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub mskvalor_LostFocus()
On Error GoTo tratar_erro

mskvalor = Format(mskvalor, "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub OptDeposito_Click()
On Error GoTo tratar_erro

Tipo = "D"
With cmb_forma
    .Clear
    .AddItem "Dinheiro"
    .AddItem "CHEQUE"
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub OptTransferencia_Click()
On Error GoTo tratar_erro

Tipo = "T"
With cmb_forma
    .Clear
    .AddItem "DOC"
    .AddItem "TED"
    .AddItem "TEV"
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

If txtCodBanco = "" Then
    SSTab1.Tab = 0
    Exit Sub
End If
Select Case SSTab1.Tab
    Case 0:
        Cmb_empresa.Visible = True
        If lst_Instituicoes.Visible = True Then lst_Instituicoes.SetFocus
    Case 1:
        Cmb_empresa.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Select Case SSTab3.Tab
            Case 0: lst_transferencias.SetFocus
            Case 1: Lst_saque.SetFocus
            Case 2: Lst_tarifa.SetFocus
        End Select
    Case 2:
        Cmb_empresa.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Lst_extrato.SetFocus
        TxtHistoricoExtrato = ""
    Case 3:
        Cmb_empresa.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        If SSTab2.Tab = 0 Then Lst_cheque.SetFocus Else Lst_cheque1.SetFocus
        Txt_favorecido = ""
        txtobscheque = ""
    Case 4:
        Cmb_empresa.Visible = False
        ProcVerificaProsseguir
        If Permitido = False Then Exit Sub
        Lista_cheque.SetFocus
    Case 5:
        Cmb_empresa.Visible = False
        ProcCarregaComboCliente

End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerificaProsseguir()
On Error GoTo tratar_erro

Permitido = True
If Novo_Banco = True Then
    MsgBox ("Salve a instituição bancária antes de prosseguir."), vbExclamation
    Permitido = False
    SSTab1.Tab = 0
    imgSalvar.SetFocus
    Exit Sub
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

With Cmb_opcao_lista
    .Clear
    If SSTab2.Tab = 0 Then
        .AddItem "Excluir/cancelar"
        .AddItem "Compensar"
        .AddItem "Cancelar compensação"
        .Text = "Cancelar compensação"
    Else
        .AddItem "Excluir/cancelar"
        .Text = "Excluir/cancelar"
    End If
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub SSTab3_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

With USToolBar2
    .ButtonState(5) = 5
    .ButtonState(8) = 0
    Select Case SSTab3.Tab
        Case 0:
            .ButtonState(5) = 0
            .ButtonState(8) = 5
            If lst_transferencias.Visible = True Then lst_transferencias.SetFocus
            ProcCarregaListaTransf
        Case 1:
            Lst_saque.SetFocus
            ProcCarregaListaSaque
        Case 2:
            Lst_tarifa.SetFocus
            ProcCarregaListaTarifa
    End Select
    .Refresh
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_Valor_Change()
On Error GoTo tratar_erro

If Txt_valor <> "" Then
    VerifNumero = Txt_valor
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor = ""
        Txt_valor.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_Valor_LostFocus()
On Error GoTo tratar_erro

Txt_valor = Format(Txt_valor, "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_valor1_Change()
On Error GoTo tratar_erro

If Txt_valor1 <> "" Then
    VerifNumero = Txt_valor1
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor1 = ""
        Txt_valor1.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_valor1_LostFocus()
On Error GoTo tratar_erro

Txt_valor1 = Format(Txt_valor1, "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtCheque_LostFocus()
On Error GoTo tratar_erro

TxtHistDepTranf = ""
If txtCheque <> "" And txtCheque <> "0" Then
    If Novo_Banco1 = True Then
        Set TBContas = CreateObject("adodb.recordset")
        TBContas.Open "Select * from tbl_ContasPagar where NDoctoBaixa = '" & txtCheque & "' and Banco = '" & txtDescricao & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBContas.EOF = False Then
            Select Case cmb_forma
                Case "CHEQUE": NomeCampo = "número de cheque"
                Case "CHEQUE PRÉ-DATADO": NomeCampo = "número de cheque"
                Case "DOC": NomeCampo = "número de DOC"
                Case "TED": NomeCampo = "número de TED"
                Case "TEV": NomeCampo = "número de TEV"
            End Select
            MsgBox ("Não é permitido utilizar este " & NomeCampo & ", pois o mesmo já foi utilizado em outra conta."), vbExclamation
            txtCheque = ""
            txtCheque.SetFocus
            TBContas.Close
            Exit Sub
        End If
        TBContas.Close
    End If
        
    Select Case cmb_forma
        Case "CHEQUE": TxtHistDepTranf = "Cheque n. " & txtCheque
        Case "CHEQUE PRÉ-DATADO": TxtHistDepTranf = "Cheque n. " & txtCheque
        Case "DOC": TxtHistDepTranf = "Doc n. " & txtCheque
        Case "TED": TxtHistDepTranf = "Ted n. " & txtCheque
        Case "TEV": TxtHistDepTranf = "Tev n. " & txtCheque
    End Select
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtConta_Change()
On Error GoTo tratar_erro

If txtConta.Text <> "" Then
    VerifNumero = txtConta.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtConta.Text = ""
        txtConta.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub TxtHistDepTranf_Change()
On Error GoTo tratar_erro

txtObsFluxo = TxtHistDepTranf

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtLimite_Change()
On Error GoTo tratar_erro

If txtLimite <> "" Then
    VerifNumero = txtLimite
    ProcVerificaNumero
    If VerifNumero = False Then
        txtLimite = ""
        txtLimite.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtLimite_LostFocus()
On Error GoTo tratar_erro

txtLimite = Format(txtLimite, "###,##0.00")
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtNBanco_Change()
On Error GoTo tratar_erro

If txtNBanco.Text <> "" Then
    VerifNumero = txtNBanco.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNBanco.Text = ""
        txtNBanco.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtSaldo_Change()
On Error GoTo tratar_erro

If txtsaldo <> "" Then
    VerifNumero = txtsaldo
    ProcVerificaNumero
    If VerifNumero = False Then
        txtsaldo = ""
        txtsaldo.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtsaldo_LostFocus()
On Error GoTo tratar_erro

txtsaldo.Text = Format(txtsaldo.Text, "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtCodBanco = ""
txtData1 = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
txtDtValidacao.Text = ""
txtRespValidacao.Text = ""
txtLimite = "0,00"
txtUtilizado = "0,00"
cmbFamilia = ""
txtNBanco = ""
txtAgencia = ""
Txt_codigo_cedente = ""
Txt_codigo_cedente1 = ""
Txt_nome_agencia = ""
txtDescricao.Text = ""
txtConta = ""
txtgerente.Text = ""
txtFone = ""
txtFAX = ""
ProcCarregaComboSetor Cmb_centro, "Setor IS NOT NULL and DtBloq IS NULL and (Consolidacao = 'False' or Consolidacao is null)", "", False, True, False, "", True, False
txtobs = ""
txtsaldo = "0,00"
txtStatus = "Liberado"
CodigoLista = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaSaldoBancario()
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from tbl_instituicoes where txt_Descricao = '" & txtDescricao.Text & "' and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    txtsaldo.Text = Format(TBFIltro!Saldo, "###,##0.00")
End If
TBFIltro.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal Key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcFiltrar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: ProcStatus
    Case 8: ProcValidarRegistros lst_Instituicoes, "Financeiro/Instituições"
    Case 9: ProcAtualizar
    Case 11: ProcAjuda
    Case 12: ProcSair
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal Key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoMovimentacao
    Case 2: ProcLocalizarMovimentacao
    Case 3: ProcSalvarMovimentacao
    Case 4: ProcExcluirMovimentacao
    Case 5: ProcImprimirMovimentacao
    Case 6: ProcAnterior
    Case 7: ProcProximo
    Case 8: ProcCopiarTarifa
    Case 9: ProcAtualizarMovimentacao
    Case 11: ProcAjuda
    Case 12: ProcSair
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub USToolBar3_ButtonClick(ByVal ButtonIndex As Integer, ByVal Key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrarExtrato
    Case 2: ProcSalvarExtrato
    Case 3: ProcImprimirExtrato
    Case 4: ProcAnterior
    Case 5: ProcProximo
    Case 6: ProcVisualizarContas
    Case 7: ProcAtualizarExtrato
    Case 9: ProcAjuda
    Case 10: ProcSair
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub USToolBar4_ButtonClick(ByVal ButtonIndex As Integer, ByVal Key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrarChequeEmitido
    Case 2: ProcSalvarChequeEmitido
    Case 3: ProcExcluirChequeEmitido
    Case 4: ProcImprimirChequeEmitido
    Case 5: ProcAnterior
    Case 6: ProcProximo
    Case 7: ProcCopiaChequeEmitido
    Case 8: ProcCompensarChequeEmitido
    Case 9: ProcCancelarCompensacaoChequeEmitido
    Case 11: ProcAjuda
    Case 12: ProcSair
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub USToolBar5_ButtonClick(ByVal ButtonIndex As Integer, ByVal Key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrarChequeRecebido
    Case 2: ProcExcluirChequeRecebido
    Case 3: ProcAnterior
    Case 4: ProcProximo
    Case 5: ProcCompensarChequeRecebido
    Case 6: ProcCancelarCompensacaoChequeRecebido
    Case 8: ProcAjuda
    Case 9: ProcSair
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcStatus()
On Error GoTo tratar_erro

If Alterar = False Then
    MsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation
    Exit Sub
End If
Permitido = False
With lst_Instituicoes
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then Permitido = True
    Next InitFor
End With
If Permitido = False Then
    MsgBox ("Informe a(s) instituição(ões) antes de alterar o status."), vbExclamation
    Exit Sub
End If
Compras_Fornecedores = False
Financeiro_Instituicao = True
frmCompras_fornecedores_bloq.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

frm_Instituicoes_localizar.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub
