VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Baixas_Receber 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Financeiro - Contas a receber | Baixar"
   ClientHeight    =   7470
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   8295
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
   Icon            =   "Frm_Baixas_Receber.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   62
      Top             =   7065
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   820
      DibPicture      =   "Frm_Baixas_Receber.frx":030A
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "Frm_Baixas_Receber.frx":8B04
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   150
      TabIndex        =   33
      Top             =   510
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   11456
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
      TabCaption(0)   =   "Dados para baixa"
      TabPicture(0)   =   "Frm_Baixas_Receber.frx":8E1E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "USToolBar1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "USImageList1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Contas antecipadas"
      TabPicture(1)   =   "Frm_Baixas_Receber.frx":8E3A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Lista"
      Tab(1).Control(1)=   "USToolBar2"
      Tab(1).Control(2)=   "PbLista"
      Tab(1).Control(3)=   "Frame9"
      Tab(1).Control(4)=   "Frame5"
      Tab(1).ControlCount=   5
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Height          =   675
         Left            =   -74940
         TabIndex        =   58
         Top             =   1320
         Width           =   7875
         Begin MSComCtl2.DTPicker msk_fltFim 
            Height          =   315
            Left            =   6390
            TabIndex        =   24
            ToolTipText     =   "Data final."
            Top             =   210
            Width           =   1305
            _ExtentX        =   2302
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
            Format          =   155713537
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker msk_fltInicio 
            Height          =   315
            Left            =   4500
            TabIndex        =   23
            ToolTipText     =   "Data inicio."
            Top             =   210
            Width           =   1305
            _ExtentX        =   2302
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
            Format          =   155713537
            CurrentDate     =   39057
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vencimento de :"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3270
            TabIndex        =   60
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Até :"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5985
            TabIndex        =   59
            Top             =   240
            Width           =   360
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74940
         TabIndex        =   53
         Top             =   5460
         Width           =   7875
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
            Left            =   3360
            TabIndex        =   27
            ToolTipText     =   "Número da página."
            Top             =   180
            Width           =   465
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
            Left            =   1830
            TabIndex        =   26
            Text            =   "30"
            ToolTipText     =   "Número de registros por página."
            Top             =   180
            Width           =   465
         End
         Begin DrawSuite2022.USButton cmdPagProx 
            Height          =   315
            Left            =   5430
            TabIndex        =   31
            ToolTipText     =   "Próxima página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "Frm_Baixas_Receber.frx":8E56
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
            Left            =   3840
            TabIndex        =   28
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
         Begin DrawSuite2022.USButton cmdPagUlt 
            Height          =   315
            Left            =   5970
            TabIndex        =   32
            ToolTipText     =   "Última página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "Frm_Baixas_Receber.frx":C5FA
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
            Left            =   4890
            TabIndex        =   30
            ToolTipText     =   "Página anterior."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "Frm_Baixas_Receber.frx":FE86
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
         Begin DrawSuite2022.USButton cmdPagPrim 
            Height          =   315
            Left            =   4350
            TabIndex        =   29
            ToolTipText     =   "Primeira página."
            Top             =   180
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   556
            DibPicture      =   "Frm_Baixas_Receber.frx":13995
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
         Begin VB.Label lblRegistros 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de reg.: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   180
            TabIndex        =   56
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lblPaginas 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pág.: 0 de: 0"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6600
            TabIndex        =   55
            Top             =   240
            Width           =   945
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carr.            reg. p/ pág."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1440
            TabIndex        =   54
            Top             =   240
            Width           =   1785
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados do extrato bancário"
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
         Left            =   55
         TabIndex        =   48
         Top             =   5440
         Width           =   7875
         Begin VB.TextBox Txt_historico 
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
            MaxLength       =   255
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   "Histórico padrão do lançamento."
            Top             =   450
            Width           =   3750
         End
         Begin VB.TextBox txtObsFluxo 
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
            Left            =   3940
            MaxLength       =   255
            TabIndex        =   22
            ToolTipText     =   "Histórico do lançamento."
            Top             =   450
            Width           =   3750
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Histórico padrão do lançamento"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   923
            TabIndex        =   50
            Top             =   240
            Width           =   2265
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Histórico do lançamento"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4960
            TabIndex        =   49
            Top             =   240
            Width           =   1710
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3665
         Left            =   55
         TabIndex        =   36
         Top             =   1780
         Width           =   7875
         Begin VB.TextBox Txt_dias_atraso 
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
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "Dias em atraso."
            Top             =   1650
            Width           =   1125
         End
         Begin VB.CheckBox Chk_multa 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Multa"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5400
            TabIndex        =   16
            Top             =   1440
            Width           =   765
         End
         Begin VB.TextBox Txt_multa 
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
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Valor da multa."
            Top             =   1650
            Width           =   1245
         End
         Begin VB.TextBox txt_ndocumento 
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
            Left            =   3150
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Número do documento."
            Top             =   1050
            Width           =   1515
         End
         Begin VB.TextBox txtSaldo 
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
            Left            =   6180
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Saldo."
            Top             =   420
            Width           =   1485
         End
         Begin VB.TextBox txtSaldoAtual 
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
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Saldo anterior."
            Top             =   420
            Width           =   1485
         End
         Begin VB.TextBox txt_conta 
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
            Left            =   3150
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Número da conta corrente."
            Top             =   420
            Width           =   1515
         End
         Begin VB.TextBox txt_ValorPago 
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
            Left            =   6180
            TabIndex        =   10
            ToolTipText     =   "Valor baixado."
            Top             =   1050
            Width           =   1485
         End
         Begin VB.TextBox txtDesconto 
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
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Valor do desconto."
            Top             =   1650
            Width           =   1245
         End
         Begin VB.TextBox txtjuros 
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
            Left            =   3900
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Valor diário do juros de mora."
            Top             =   1650
            Width           =   1245
         End
         Begin VB.TextBox txt_VlrDocto 
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
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Valor do documento."
            Top             =   1050
            Width           =   1485
         End
         Begin VB.TextBox txtObs 
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
            Height          =   1275
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            ToolTipText     =   "Observações."
            Top             =   2220
            Width           =   7485
         End
         Begin VB.CheckBox chkdesconto 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Desconto"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6547
            TabIndex        =   18
            Top             =   1440
            Width           =   990
         End
         Begin VB.CheckBox chkjuros 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Juros diário"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3960
            TabIndex        =   14
            Top             =   1440
            Width           =   1125
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
            ItemData        =   "Frm_Baixas_Receber.frx":17A86
            Left            =   180
            List            =   "Frm_Baixas_Receber.frx":17A88
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            ToolTipText     =   "Forma da baixa."
            Top             =   1050
            Width           =   2655
         End
         Begin VB.ComboBox cmb_Banco 
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
            Left            =   180
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Instituição bancária."
            Top             =   420
            Width           =   2975
         End
         Begin VB.CommandButton CmdForma 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   315
            Left            =   2820
            Picture         =   "Frm_Baixas_Receber.frx":17A8A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Localizar forma da baixa."
            Top             =   1050
            Width           =   315
         End
         Begin MSComCtl2.DTPicker txt_DtPagto 
            Height          =   315
            Left            =   180
            TabIndex        =   11
            ToolTipText     =   "Data da baixa."
            Top             =   1650
            Width           =   1275
            _ExtentX        =   2249
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
            Format          =   182059009
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker Cmb_data_movimentacao 
            Height          =   315
            Left            =   1470
            TabIndex        =   12
            ToolTipText     =   "Data da movimentação."
            Top             =   1650
            Width           =   1275
            _ExtentX        =   2249
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
            Format          =   182059009
            CurrentDate     =   39057
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. moviment."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1590
            TabIndex        =   57
            Top             =   1440
            Width           =   1020
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Dias em atraso"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2790
            TabIndex        =   47
            Top             =   1440
            Width           =   1065
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6727
            TabIndex        =   46
            Top             =   210
            Width           =   390
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo anterior"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4920
            TabIndex        =   45
            Top             =   210
            Width           =   1005
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Valor baixado*"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6180
            TabIndex        =   44
            Top             =   840
            Width           =   1485
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. baixa"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   487
            TabIndex        =   43
            Top             =   1440
            Width           =   660
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Forma da baixa*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   907
            TabIndex        =   42
            Top             =   840
            Width           =   1200
         End
         Begin VB.Label LblDocumento 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "N° do documento"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3150
            TabIndex        =   41
            Top             =   840
            Width           =   1515
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Instituição bancária*"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   917
            TabIndex        =   40
            Top             =   210
            Width           =   1500
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   " Conta corrente"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3337
            MousePointer    =   4  'Icon
            TabIndex        =   39
            Top             =   210
            Width           =   1140
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Observações"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3450
            TabIndex        =   38
            Top             =   2010
            Width           =   945
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Vlr. documento"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4875
            TabIndex        =   37
            Top             =   840
            Width           =   1095
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
         Height          =   465
         Left            =   55
         TabIndex        =   35
         Top             =   1320
         Width           =   7875
         Begin VB.CheckBox Chk_mov_total 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Criar movimentação total no extrato"
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
            Left            =   2340
            TabIndex        =   1
            Top             =   180
            Width           =   3435
         End
         Begin VB.CheckBox chbparcial 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Baixar parcialmente"
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
            Left            =   180
            TabIndex        =   0
            Top             =   180
            Width           =   2175
         End
      End
      Begin DrawSuite2022.USProgressBar PbLista 
         Height          =   255
         Left            =   -74940
         TabIndex        =   34
         Top             =   6090
         Width           =   7875
         _ExtentX        =   13891
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
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   4635
         Top             =   480
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "Frm_Baixas_Receber.frx":17B8C
         Count           =   1
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   60
         TabIndex        =   51
         Top             =   330
         Width           =   7875
         _ExtentX        =   13891
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
         ButtonCaption1  =   "Baixar"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Baixar (F3)"
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
         ButtonWidth1    =   44
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
         ButtonLeft2     =   48
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   52
         ButtonTop3      =   2
         ButtonWidth3    =   41
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   95
         ButtonTop4      =   2
         ButtonWidth4    =   30
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonKey5      =   "5"
         ButtonAlignment5=   2
         BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState5    =   5
         ButtonLeft5     =   127
         ButtonTop5      =   2
         ButtonWidth5    =   24
         ButtonHeight5   =   24
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74940
         TabIndex        =   52
         Top             =   330
         Width           =   7875
         _ExtentX        =   13891
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
            Name            =   "MS Sans Serif"
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
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   7080
            Top             =   210
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "Frm_Baixas_Receber.frx":19AE6
            Count           =   1
         End
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   3420
         Left            =   -74940
         TabIndex        =   25
         Top             =   2010
         Visible         =   0   'False
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   6033
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
            Text            =   "Dt. emissão"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "D"
            Text            =   "Dt. vencto."
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Valor"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "N"
            Text            =   "Saldo"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Nota fiscal"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Parcela"
            Object.Width           =   1499
         EndProperty
      End
   End
End
Attribute VB_Name = "frm_Baixas_Receber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TBLISTA_Contas_Receber_Baixar As ADODB.Recordset 'OK
Dim permitido_devolucao As Boolean

Private Sub chbparcial_Click()
On Error GoTo tratar_erro

If chbparcial.Value = 1 Then
    txt_ValorPago.Text = ""
    With chkjuros
        .Value = 0
        .Enabled = False
    End With
    With txtjuros
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
    With Chk_multa
        .Value = 0
        .Enabled = False
    End With
    With Txt_multa
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
    With chkdesconto
        .Value = 0
        .Enabled = False
    End With
    With txtDesconto
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
Else
    txt_ValorPago.Text = txt_VlrDocto.Text
    chkjuros.Enabled = True
    Chk_multa.Enabled = True
    chkdesconto.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_multa_Click()
On Error GoTo tratar_erro

If Chk_multa.Value = 1 Then
    'Verifica percentual de multa que está cadastrado no boleto e calcula valor da multa
    Multa = ""
    With frmContas_Receber
        For InitFor = 1 To .Lista.ListItems.Count
            If .Lista.ListItems.Item(InitFor).Checked = True Then
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select Multa from tbl_Detalhes_Recebimento where IDContaReceber = " & .Lista.ListItems.Item(InitFor) & " and Multa IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    valor = txt_VlrDocto
                    Multa = Format((valor * TBContas!Multa) / 100, "###,##0.00")
                End If
            End If
        Next InitFor
    End With
    With Txt_multa
        .Text = Multa
        .Locked = False
        .TabStop = True
        .SetFocus
    End With
    
    chkdesconto.Value = 0
    With txtDesconto
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
    With txt_ValorPago
        .Locked = True
        .TabStop = False
    End With
Else
    With Txt_multa
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
    With txtDesconto
        .Locked = False
        .TabStop = True
    End With
    With txt_ValorPago
        .Locked = False
        .TabStop = True
    End With
    ProcCalculaJurosMulta
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkdesconto_Click()
On Error GoTo tratar_erro

If chkdesconto.Value = 1 Then
    'Verifica percentual de desconto que está cadastrado no boleto e calcula valor do desconto
    DescontoTexto = ""
    With frmContas_Receber
        For InitFor = 1 To .Lista.ListItems.Count
            If .Lista.ListItems.Item(InitFor).Checked = True Then
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select Desconto from tbl_Detalhes_Recebimento where IDContaReceber = " & .Lista.ListItems.Item(InitFor) & " and Desconto IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    valor = txt_VlrDocto
                    DescontoTexto = Format((valor * TBContas!Desconto) / 100, "###,##0.00")
                End If
            End If
        Next InitFor
    End With
    With txtDesconto
        .Text = DescontoTexto
        .Locked = False
        .TabStop = True
        .SetFocus
    End With
    
    chkjuros.Value = 0
    With txtjuros
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
    Chk_multa.Value = 0
    With Txt_multa
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
    With txt_ValorPago
        .Locked = True
        .TabStop = False
    End With
Else
    With txtDesconto
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
    With txt_ValorPago
        .Locked = False
        .TabStop = True
    End With
    ProcCalculaDesconto
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkjuros_Click()
On Error GoTo tratar_erro

If chkjuros.Value = 1 Then
    'Verifica valor do juros diário que está cadastrado no boleto
    Juros = ""
    With frmContas_Receber
        For InitFor = 1 To .Lista.ListItems.Count
            If .Lista.ListItems.Item(InitFor).Checked = True Then
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select Juros from tbl_Detalhes_Recebimento where IDContaReceber = " & .Lista.ListItems.Item(InitFor) & " and Juros IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    valor = txt_VlrDocto
                    Juros = Format((valor * TBContas!Juros) / 100, "###,##0.00")
                End If
            End If
        Next InitFor
    End With
    With txtjuros
        .Text = Juros
        .Locked = False
        .TabStop = True
        .SetFocus
    End With
    
    chkdesconto.Value = 0
    With txtDesconto
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
    With txt_ValorPago
        .Locked = True
        .TabStop = False
    End With
Else
    With txtjuros
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
    With txt_ValorPago
        .Locked = False
        .TabStop = True
    End With
    ProcCalculaJurosMulta
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmb_forma_Click()
On Error GoTo tratar_erro

LblDocumento.Caption = "N° do documento"
With txt_ndocumento
    .Text = ""
    .Locked = True
    .TabStop = False
End With
With Txt_historico
    .Text = ""
    .Locked = False
    .TabStop = True
End With
If chbparcial.Value = 0 Then
    txt_ValorPago = txt_VlrDocto
    If chkjuros.Value = 1 Or Chk_multa.Value = 1 Then ProcCalculaJurosMulta
    If chkdesconto.Value = 1 Then ProcCalculaDesconto
End If

Contador = 0
With frmContas_Receber
    For InitFor = 1 To .Lista.ListItems.Count
        If .Lista.ListItems.Item(InitFor).Checked = True Then
            'Verifica se existe(m) conta(s) descontada selecionada e bloqueia a opção de criar valor total no extrato
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from tbl_Contas_receber where IdIntConta = " & .Lista.ListItems.Item(InitFor) & " and Status = 'DUPLICATA DESCONTADA EM ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Permitido = False
            End If
            TBFI.Close
            
            Contador = Contador + 1
            If Contador > 1 Then GoTo Prosseguir
        End If
    Next InitFor
End With

Prosseguir:
    If Contador > 1 Then Chk_mov_total.Enabled = True
    With txt_ValorPago
        .Locked = False
        .TabStop = True
    End With
    ProcAtualizaSaldo
    
    If cmb_forma = "DOC" Or cmb_forma = "TED" Or cmb_forma = "MALOTE" Then
        If Contador > 1 Then
            With Chk_mov_total
                .Value = 1
                .Enabled = False
            End With
        End If
        If cmb_forma = "DOC" Then
            LblDocumento.Caption = "N° do DOC*"
        ElseIf cmb_forma = "TED" Then
                LblDocumento.Caption = "N° do TED*"
            Else
                LblDocumento.Caption = "N° do malote*"
        End If
        With txt_ndocumento
           .Locked = False
           .TabStop = True
        End With
        With Txt_historico
            .Locked = True
            .TabStop = False
        End With
    Else
        Select Case cmb_forma.Text
            Case "CHEQUE":
                If Contador > 1 Then
                    With Chk_mov_total
                        .Value = 1
                        .Enabled = False
                    End With
                End If
                LblDocumento.Caption = "N° do cheque*"
                With txt_ndocumento
                   .Locked = False
                   .TabStop = True
                End With
                With Txt_historico
                    .Locked = True
                    .TabStop = False
                End With
            Case "CHEQUE PRÉ-DATADO":
                If Contador > 1 Then
                    With Chk_mov_total
                        .Value = 1
                        .Enabled = False
                    End With
                End If
                LblDocumento.Caption = "N° do cheque*"
                With txt_ndocumento
                   .Locked = False
                   .TabStop = True
                End With
                With Txt_historico
                    .Locked = True
                    .TabStop = False
                End With
        End Select
    End If
    
    ProcCriaHistPadrao
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub CmdForma_Click()
On Error GoTo tratar_erro

Financeiro_Contas_Pagar = False
Financeiro_Forma_Pgto_Pagar = False
Financeiro_Contas_Receber = False
Financeiro_Forma_Pgto_Receber = True
frmContas_Forma_Pagamento.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Contas_Receber_Baixar.AbsolutePage <> 2 Then
    If TBLISTA_Contas_Receber_Baixar.AbsolutePage = -3 Then
        ProcExibePagina (TBLISTA_Contas_Receber_Baixar.PageCount - 1), Devolucao
    Else
        TBLISTA_Contas_Receber_Baixar.AbsolutePage = TBLISTA_Contas_Receber_Baixar.AbsolutePage - 2
        ProcExibePagina (TBLISTA_Contas_Receber_Baixar.AbsolutePage), Devolucao
    End If
Else
    ProcExibePagina (1), Devolucao
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
    TBLISTA_Contas_Receber_Baixar.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBLISTA_Contas_Receber_Baixar.AbsolutePage), Devolucao
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Contas_Receber_Baixar.AbsolutePage = 1
ProcExibePagina (TBLISTA_Contas_Receber_Baixar.AbsolutePage), Devolucao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBLISTA_Contas_Receber_Baixar.AbsolutePage <> -3 Then
    If TBLISTA_Contas_Receber_Baixar.AbsolutePage = 1 Then
        ProcExibePagina (2), Devolucao
    Else
        ProcExibePagina (TBLISTA_Contas_Receber_Baixar.AbsolutePage), Devolucao
    End If
Else
    ProcExibePagina (TBLISTA_Contas_Receber_Baixar.PageCount), Devolucao
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBLISTA_Contas_Receber_Baixar.AbsolutePage = TBLISTA_Contas_Receber_Baixar.PageCount
ProcExibePagina (TBLISTA_Contas_Receber_Baixar.AbsolutePage), Devolucao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
    Case vbKeyF3: ProcBaixar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
   
Private Sub Cmb_banco_Click()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_instituicoes where ID = " & cmb_Banco.ItemData(cmb_Banco.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txt_Conta.Text = IIf(IsNull(TBAbrir!txt_Conta), "", TBAbrir!txt_Conta)
    txtSaldoAtual = IIf(IsNull(TBAbrir!Saldo), "", Format(TBAbrir!Saldo, "##,##0.00"))
    ProcAtualizaSaldo
End If
TBAbrir.Close

ProcCriaHistPadrao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Permitido1 = False
SSTab1.Tab = 0
ProcCarregaToolBar1 Me, 7875, 5, True
ProcCarregaToolBar2 Me, 7875, 5, True

txt_DtPagto.Value = Date
Cmb_data_movimentacao.Value = Date
msk_fltInicio.Value = Date
msk_fltFim.Value = Date

With frmContas_Receber
    ProcCarregaComboBancoFinanceiro cmb_Banco, "txt_descricao IS NOT NULL and ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex) & " and Bloqueado = 'false' and DtValidacao IS NOT NULL", False
    ProcCarregaComboForma
        
    'Verifica valor total das contas selecionadas
    valor = 0
    Contador = 0
    For InitFor = 1 To .Lista.ListItems.Count
        If .Lista.ListItems.Item(InitFor).Checked = True Then
            valor = valor + .Lista.ListItems.Item(InitFor).ListSubItems(3)
            
            Antecipacao = False
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_Contas_receber where IdIntConta = " & .Lista.ListItems.Item(InitFor) & " and Antecipacao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                Antecipacao = True
            End If
            
            Devolucao = False
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_Contas_receber where IdIntConta = " & .Lista.ListItems.Item(InitFor) & " and Devolucao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                Devolucao = True
            End If
            TBContas.Close
            
            'Verifica se existe(m) conta(s) descontada selecionada e bloqueia a opção de criar valor total no extrato
            Permitido = True
            Set TBContas = CreateObject("adodb.recordset")
            TBContas.Open "Select * from tbl_Contas_receber where IdIntConta = " & .Lista.ListItems.Item(InitFor) & " and Status = 'DUPLICATA DESCONTADA EM ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
            If TBContas.EOF = False Then
                Permitido = False
            End If
            TBContas.Close
            
            Contador = Contador + 1
        End If
    Next InitFor
    
     With SSTab1
        If Contador = 1 And (Antecipacao = False Or Devolucao = True) Then
            .TabVisible(1) = True
            .TabsPerRow = 2
            If Devolucao = True Then .TabCaption(1) = "Contas a receber/baixadas" Else .TabCaption(1) = "Contas antecipadas"
        Else
            .TabVisible(1) = False
            .TabsPerRow = 1
        End If
    End With
        
    txt_VlrDocto = Format(valor, "###,##0.00")
    With txt_ValorPago
        .Text = Format(valor, "###,##0.00")
        If Contador > 1 Then
            .Locked = True
            .TabStop = False
        Else
            .Locked = False
            .TabStop = True
        End If
    End With
        
    If Contador > 1 Or Antecipacao = True Or Devolucao = True Then
        chbparcial.Enabled = False
        
        With Chk_mov_total
            If Contador > 1 And Permitido = True Then
                .Enabled = True
                .Value = 1
            Else
                .Enabled = False
            End If
        End With
        
        chkjuros.Enabled = False
        Chk_multa.Enabled = False
        chkdesconto.Enabled = False
    Else
        chbparcial.Enabled = True
        Chk_mov_total.Enabled = False
        chkjuros.Enabled = False
        Chk_multa.Enabled = False
        chkdesconto.Enabled = False
    End If
    
    If Contador = 1 Then
        For InitFor = 1 To .Lista.ListItems.Count
            If .Lista.ListItems.Item(InitFor).Checked = True Then
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select * from tbl_contas_receber where IdIntConta = " & .Lista.ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    NomeCampo = "a instituição bancária"
                    If IsNull(TBContas!Banco) = False And TBContas!Banco <> "" Then cmb_Banco = TBContas!Banco
                    NomeCampo = "a forma da baixa"
                    If IsNull(TBContas!FormaBaixa) = False And TBContas!FormaBaixa <> "" Then cmb_forma = TBContas!FormaBaixa
1:
                    If chbparcial.Enabled = True Then
                        If TBContas!status = "DUPLICATA DESCONTADA EM ABERTO" Then chbparcial.Enabled = False Else chbparcial.Enabled = True
                    End If
                    txtObs = IIf(IsNull(TBContas!Observacoes), "", Trim(TBContas!Observacoes))
                End If
                TBContas.Close
                ProcVerifDiasAtraso
            End If
        Next InitFor
    End If
End With
       
Exit Sub
tratar_erro:
    If Err.Number = "383" Then
        If NomeCampo = "a instituição bancária" Then
            USMsgBox ("Não foi encontrado a instituição bancária ou a mesma está bloqueada."), vbExclamation, "CAPRIND v5.0"
        Else
            USMsgBox ("Não foi encontrado " & NomeCampo & " desta conta."), vbExclamation, "CAPRIND v5.0"
        End If
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaComboForma()
On Error GoTo tratar_erro

ProcCarregaComboFormaPgtoRcbto cmb_forma, "Tipo = 'R'"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBaixar()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente baixar esta(s) conta(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Acao = "baixar"
    If cmb_Banco.Text = "" Then
        NomeCampo = "a instituição bancária"
        ProcVerificaAcao
        cmb_Banco.SetFocus
        Exit Sub
    End If
    If cmb_forma.Text = "" Then
        NomeCampo = "a forma da baixa"
        ProcVerificaAcao
        cmb_forma.SetFocus
        Exit Sub
    End If
    If cmb_forma = "CHEQUE" Or cmb_forma = "CHEQUE PRÉ-DATADO" Or cmb_forma = "DOC" Or cmb_forma = "TED" Or cmb_forma = "MALOTE" Then
        If txt_ndocumento = "" Or txt_ndocumento = "0" Then
            Select Case cmb_forma
                Case "CHEQUE": NomeCampo = "o número do cheque"
                Case "CHEQUE PRÉ-DATADO": NomeCampo = "o número do cheque"
                Case "DOC": NomeCampo = "o número do DOC"
                Case "TED": NomeCampo = "o número do TED"
                Case "MALOTE": NomeCampo = "o número do malote"
            End Select
            ProcVerificaAcao
            txt_ndocumento.SetFocus
            Exit Sub
        End If
    End If
    If chkjuros.Value = 1 Then
        valor = IIf(txtjuros = "", 0, txtjuros)
        If valor <= 0 Then
            NomeCampo = "o valor de juros mora diário"
            ProcVerificaAcao
            txtjuros.SetFocus
            Exit Sub
        End If
    End If
    If Chk_multa.Value = 1 Then
        valor = IIf(Txt_multa = "", 0, Txt_multa)
        If valor <= 0 Then
            NomeCampo = "o valor da multa"
            ProcVerificaAcao
            Txt_multa.SetFocus
            Exit Sub
        End If
    End If
    If chkdesconto.Value = 1 Then
        valor = IIf(txtDesconto = "", 0, txtDesconto)
        If valor <= 0 Then
            NomeCampo = "o valor do desconto"
            ProcVerificaAcao
            txtDesconto.SetFocus
            Exit Sub
        End If
    End If
    
    'Verifica se tem antecipação e não deixa criar totalização no extrato
    If Antecipacao = False And Devolucao = False Then
        For InitFor1 = 1 To Lista.ListItems.Count
            If Lista.ListItems.Item(InitFor1).Checked = True Then
                If USMsgBox("Não será criado a movimentação total no extrato, pois existe(m) antecipação(ões) selecionada(s). Deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Chk_mov_total.Value = 0
                GoTo Prosseguir
            End If
        Next InitFor1
    End If

Prosseguir:
    valor = IIf(txt_ValorPago = "", 0, txt_ValorPago)
    If valor <= 0 And Devolucao = False Or valor >= 0 And Devolucao = True Then
        NomeCampo = "o valor baixado"
        ProcVerificaAcao
        txt_ValorPago.SetFocus
        Exit Sub
    End If
    
    'Verifica se o total de devoluções é maior ou iqual o valor da conta
    If Devolucao = True Then
        Valor_Produto = 0
        For InitFor1 = 1 To Lista.ListItems.Count
            If Lista.ListItems.Item(InitFor1).Checked = True Then Valor_Produto = Valor_Produto + Lista.ListItems(InitFor1).SubItems(3)
        Next InitFor1
        If valor > Valor_Produto Then
            USMsgBox ("Não foi possível baixar, pois o valor da(s) conta(s) selecionada(s) é menor que o valor da devolução."), vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
    End If
    
    ProcCriaHistPadrao
    Permitido1 = True
    If chbparcial.Value = 1 Then ProcReceberParcial Else ProcReceberIntegral
    If Permitido1 = False Then Exit Sub
    
    Unload Me
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcReceberParcial()
On Error GoTo tratar_erro
    
If IsNumeric(txt_ValorPago.Text) = True Then VP = txt_ValorPago.Text
If IsNumeric(txt_VlrDocto.Text) = True Then VD = txt_VlrDocto.Text
If VP = VD Then
    USMsgBox ("Não é permitido baixar parcial, pois o valor recebido é o mesmo que o valor total da conta."), vbInformation, "CAPRIND v5.0"
    Permitido1 = False
    Exit Sub
End If
If VP > VD Then
    USMsgBox ("Não é permitido baixar parcial, pois o valor recebido é maior que o valor total da conta."), vbInformation, "CAPRIND v5.0"
    Permitido1 = False
    Exit Sub
End If
ID_variasTexto = 0
With frmContas_Receber
    For InitFor = 1 To .Lista.ListItems.Count
        If .Lista.ListItems.Item(InitFor).Checked = True Then
            'Edita conta principal
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_contas_receber where idintconta = " & .Lista.ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                TBAbrir!ValorPendente = Format(txt_VlrDocto.Text - txt_ValorPago.Text, "###,##0.00")
                TBAbrir!valor = TBAbrir!ValorPendente
                TBAbrir!ValorExtenso = FunValorExtenso(TBAbrir!ValorPendente)
                If TBAbrir!valorprincipal = 0 Then TBAbrir!valorprincipal = txt_VlrDocto.Text
                TBAbrir!tituloref = TBAbrir!IDintconta
                TBAbrir!status = "TÍTULO RECEBIDO PARCIAL"
                TBAbrir!Banco = cmb_Banco
                TBAbrir.Update
                
                'Fluxo de Caixa
                Set TBFluxo = CreateObject("adodb.recordset")
                TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBAbrir!IDFluxo), 0, TBAbrir!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                If TBFluxo.EOF = False Then
                    TBFluxo!valor = TBAbrir!ValorPendente
                    TBFluxo!Instituicao = cmb_Banco
                    TBFluxo.Update
                End If
                TBFluxo.Close
                
                'Cria e paga nova conta parcial
                Set TBCorretiva = CreateObject("adodb.recordset")
                TBCorretiva.Open "Select * from tbl_contas_receber", Conexao, adOpenKeyset, adLockOptimistic
                TBCorretiva.AddNew
                TBCorretiva!tituloref = TBAbrir!IDintconta
                TBCorretiva!RecebidoParcial = txt_ValorPago.Text
                TBCorretiva!Parcial = True
                TBCorretiva!status = "TÍTULO RECEBIDO PARCIAL"
                If pubUsuario <> "" Then TBCorretiva!resprec = pubUsuario
                TBCorretiva!Antecipacao = False
                TBCorretiva!Devolucao = False
                TBCorretiva!Data_transacao = TBAbrir!Data_transacao
                TBCorretiva!Tipo_doc = TBAbrir!Tipo_doc
                TBCorretiva!txt_ndocumento = IIf(IsNull(TBAbrir!txt_ndocumento), "", TBAbrir!txt_ndocumento)
                TBCorretiva!NFiscal = TBAbrir!NFiscal
                TBCorretiva!Proposta = TBAbrir!Proposta
                TBCorretiva!IDCliente = TBAbrir!IDCliente
                TBCorretiva!Tipo = TBAbrir!Tipo
                TBCorretiva!Nome_Razao = TBAbrir!Nome_Razao
                TBCorretiva!Cidade = TBAbrir!Cidade
                TBCorretiva!Estado = TBAbrir!Estado
                TBCorretiva!emissao = TBAbrir!emissao
                TBCorretiva!Vencimento = TBAbrir!Vencimento
                TBCorretiva!valor = txt_ValorPago
                TBCorretiva!ValorExtenso = FunValorExtenso(txt_ValorPago)
                TBCorretiva!Parcela = TBAbrir!Parcela
                TBCorretiva!Observacoes = TBAbrir!Observacoes
                TBCorretiva!Logsit = "S"
                TBCorretiva!valorprincipal = TBAbrir!valorprincipal
                TBCorretiva!ID_nota = TBAbrir!ID_nota
                'Dados de recebimento
                TBCorretiva!FormaBaixa = cmb_forma.Text
                TBCorretiva!Data_pagamento = txt_DtPagto.Value
                TBCorretiva!Data_movimentacao = Cmb_data_movimentacao.Value
                TBCorretiva!valortitulorecebido = txt_ValorPago.Text
                TBCorretiva!valorprincipal = TBAbrir!valorprincipal
                TBCorretiva!NDoctoBaixa = txt_ndocumento.Text
                TBCorretiva!Banco = cmb_Banco.Text
                TBCorretiva!Obs = txtObs.Text
                TBCorretiva!Dias_atraso = IIf(Txt_dias_atraso = "", 0, Txt_dias_atraso)
                TBCorretiva!ValorPendente = Format(TBAbrir!ValorPendente, "###,##0.00")
                TBCorretiva!ID_empresa = TBAbrir!ID_empresa
                TBCorretiva!Bloqueado = False
                TBCorretiva.Update
                
                Valor1 = txt_ValorPago
                
                'Verifica se o valor baixado é igual ao valor da conta de antecipação selecionada na lista
                Valor2 = 0
                If TBCorretiva!Antecipacao = False And TBCorretiva!Devolucao = False Then
                    For InitFor1 = 1 To Lista.ListItems.Count
                        If Lista.ListItems.Item(InitFor1).Checked = True Then Valor2 = Valor2 + Lista.ListItems.Item(InitFor1).ListSubItems(3)
                    Next InitFor1
                End If
                
                'Fluxo de Caixa
                Set TBFluxo = CreateObject("adodb.recordset")
                TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                If TBFluxo.EOF = True Then TBFluxo.AddNew
                TBFluxo!IDintconta = TBCorretiva!IDintconta
                TBFluxo!Operacao = "Crédito"
                TBFluxo!Data = Cmb_data_movimentacao
                If cmb_forma <> "CHEQUE" And cmb_forma <> "CHEQUE PRÉ-DATADO" And cmb_forma <> "DOC" And cmb_forma <> "TED" And Txt_historico <> "" Then
                    TBFluxo!Descricao = Txt_historico
                Else
                    TBFluxo!Descricao = TBAbrir!Nome_Razao
                End If
                TBFluxo!status = "S"
                TBFluxo!int_NotaFiscal = TBAbrir!NFiscal
                TBFluxo!Obs = IIf(txtObsFluxo = "", TBFluxo!Descricao, txtObsFluxo)
                
                If Valor2 = 0 Then 'Valor antecipado igual a 0
                    Valor3 = Valor1
                    If TBAbrir!titulodesc = True Then TBFluxo!Bloqueado = True Else TBFluxo!Bloqueado = False
                ElseIf Valor1 > Valor2 Then 'Valor pago maior que o valor antecipado
                        Valor3 = Valor1 - Valor2
                        If TBAbrir!titulodesc = True Then TBFluxo!Bloqueado = True Else TBFluxo!Bloqueado = False
                    Else
                        Valor3 = Valor1 'Valor pago menor ou igual ao valor antecipado
                        TBFluxo!Bloqueado = True
                End If
                TBFluxo!valor = Format(Valor3, "###,##0.00")
                                
                TBFluxo!ID_empresa = TBAbrir!ID_empresa
                TBFluxo!Documento = IIf(IsNull(TBAbrir!txt_ndocumento), "", TBAbrir!txt_ndocumento)
                TBFluxo!Instituicao = cmb_Banco
                TBFluxo!Hora = Format(Now, "hh:mm:ss")
                If txt_ndocumento <> "" Then TBFluxo!Cheque = txt_ndocumento
                TBFluxo!tituloref = TBAbrir!IDintconta
                If (cmb_forma = "CHEQUE" Or cmb_forma = "CHEQUE PRÉ-DATADO" Or cmb_forma = "DOC" Or cmb_forma = "TED" Or cmb_forma = "MALOTE") And TBAbrir!status <> "DUPLICATA DESCONTADA EM ABERTO" Then
                    TBFluxo!Bloqueado = True
                    TextoFiltroData = "Data = '" & Format(Cmb_data_movimentacao.Value, "Short Date") & "' and"
                    Select Case cmb_forma
                        Case "CHEQUE":
                            Descricao = "Cheque n. " & txt_ndocumento
                            TextoFiltroData = ""
                        Case "CHEQUE PRÉ-DATADO":
                            Descricao = "Cheque n. " & txt_ndocumento
                            TextoFiltroData = ""
                        Case "DOC": Descricao = "Doc n. " & txt_ndocumento
                        Case "TED": Descricao = "Ted n. " & txt_ndocumento
                        Case "MALOTE": Descricao = "Malote n. " & txt_ndocumento
                    End Select
                    
                    'Cria registro com o valor total da operação
                    valor = 0
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select Sum(Valor) as ValorTotal from tbl_Fluxo_de_caixa where " & TextoFiltroData & " Operacao = 'Crédito' and Descricao = '" & Descricao & "' and Cheque = '" & txt_ndocumento & "' and Instituicao = '" & cmb_Banco & "' and IdIntConta <> " & TBCorretiva!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        valor = IIf(IsNull(TBFI!ValorTotal), 0, TBFI!ValorTotal)
                    End If
                    TBFI.Close
                    valor = valor + Valor1
                    
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from tbl_Fluxo_de_caixa where Data = '" & Format(Cmb_data_movimentacao.Value, "Short Date") & "' and Operacao = 'Crédito' and Descricao = '" & Descricao & "' and Cheque = '" & txt_ndocumento & "' and Instituicao = '" & cmb_Banco & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBGravar.EOF = True Then
                        TBGravar.AddNew
                        TBGravar!Operacao = "Crédito"
                        TBGravar!Data = Cmb_data_movimentacao
                        TBGravar!Bloqueado = False
                        If cmb_forma = "CHEQUE" Or cmb_forma = "CHEQUE PRÉ-DATADO" Or cmb_forma = "DOC" Or cmb_forma = "TED" Or cmb_forma = "MALOTE" Then
                            If cmb_forma = "CHEQUE" Or cmb_forma = "CHEQUE PRÉ-DATADO" Then TBGravar!Bloqueado = True
                            TBGravar!Descricao = Descricao
                        End If
                        TBGravar!status = "S"
                        TBGravar!Instituicao = cmb_Banco
                        TBGravar!Hora = TBFluxo!Hora
                        TBGravar!Cheque = txt_ndocumento
                        TBGravar!Obs = IIf(txtObsFluxo = "", TBGravar!Descricao, txtObsFluxo)
                        TBGravar!ID_empresa = TBAbrir!ID_empresa
                    End If
                    TBGravar!valor = Format(valor, "###,##0.00")
                    TBGravar.Update
                    ID_variasTexto = IIf(IsNull(TBGravar!ID_varias), 0, TBGravar!ID_varias)
                    TBGravar.Close
                End If
                TBFluxo.Update
                Conexao.Execute "Update tbl_contas_receber set IDFluxo = " & TBFluxo!IDFluxo & ", ID_varias = " & ID_variasTexto & " where IDIntconta = " & TBCorretiva!IDintconta
                TBFluxo.Close
                
                Qtd = Valor1 'Valor recebido
                If Devolucao = True Then
                    ProcDevolucao TBCorretiva!IDintconta
                ElseIf Valor2 > 0 Then
                        ProcAntecipacao TBCorretiva!IDintconta
                End If
                
                If TBAbrir!status <> "DUPLICATA DESCONTADA EM ABERTO" And cmb_forma <> "CHEQUE" And cmb_forma <> "CHEQUE PRÉ-DATADO" Then
                    Set TBItem = CreateObject("adodb.recordset")
                    TBItem.Open "Select * from tbl_instituicoes where ID = " & cmb_Banco.ItemData(cmb_Banco.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
                    If TBItem.EOF = False Then
                        If Devolucao = True Then TBItem!Saldo = Format(TBItem!Saldo + Valor1, "###,##0.00") Else TBItem!Saldo = Format(TBItem!Saldo + Qtd, "###,##0.00")
                        TBItem.Update
                    End If
                    TBItem.Close
                End If
                
                'Família de contas
                qt = 0
                ValorTotal = txt_VlrDocto
                Set TBFamilia = CreateObject("adodb.recordset")
                TBFamilia.Open "Select * from familia_financeiro where IDConta = " & .Lista.ListItems.Item(InitFor) & " and tipoconta = 'R' and Pago_recebido = 'False' order by ID_PC", Conexao, adOpenKeyset, adLockOptimistic
                If TBFamilia.EOF = False Then
                    Contador = TBFamilia.RecordCount
                    Do While TBFamilia.EOF = False
                        'Verifica a porcentagem representada pelo valor da família
                        Valor2 = TBFamilia!valor
                        Valor1 = Format((Valor2 * 100) / ValorTotal, "###,##0.0000000000")
                        
                        Qtde = txt_ValorPago
                        valor = Format((Qtde * Valor1) / 100, "###,##0.00")
                        
                        TBFamilia!Pago_recebido = True
                        TBFamilia!IDConta = TBCorretiva!IDintconta
                        qt = TBFamilia!valor - valor 'Valor a receber
                        TBFamilia!valor = valor
                        If qt > 0 Then
                            Set TBCiclo = CreateObject("adodb.recordset")
                            TBCiclo.Open "select * FROM familia_financeiro", Conexao, adOpenKeyset, adLockOptimistic
                            TBCiclo.AddNew
                            TBCiclo!ID_PC = TBFamilia!ID_PC
                            TBCiclo!IDConta = TBAbrir!IDintconta
                            TBCiclo!IDnota = TBFamilia!IDnota
                            TBCiclo!TipoConta = TBFamilia!TipoConta
                            TBCiclo!Pago_recebido = False
                            TBCiclo!valor = qt
                            TBCiclo.Update
                            TBCiclo.Close
                        End If
                        TBFamilia.Update
                        TBFamilia.MoveNext
                    Loop
                End If
                TBFamilia.Close
                TBCorretiva.Close
                
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select local_troca from Troca_titulo where ID = " & TBAbrir!IDtrocatitulo, Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    Set TBReceber = CreateObject("adodb.recordset")
                    TBReceber.Open "Select Sum(tbl_contas_receber.Valor) as Valor from tbl_contas_receber INNER JOIN troca_titulo on tbl_contas_receber.Idtrocatitulo = troca_titulo.ID where troca_titulo.Local_troca = '" & TBContas!local_troca & "' and tbl_contas_receber.ID_empresa = " & TBAbrir!ID_empresa & " and tbl_contas_receber.status = 'DUPLICATA DESCONTADA EM ABERTO' and tbl_contas_receber.Logsit = 'N'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBReceber.EOF = False Then
                        valor = IIf(IsNull(TBReceber!valor), 0, TBReceber!valor)
                        NovoValor = Replace(valor, ",", ".")
                        Conexao.Execute "Update tbl_Instituicoes Set Limite_utilizado = " & NovoValor & " where txt_Descricao = '" & TBContas!local_troca & "' and ID_empresa = " & TBAbrir!ID_empresa
                    End If
                    TBReceber.Close
                End If
                TBContas.Close
                        
                '==================================
                Modulo = "Financeiro/Contas a receber"
                Evento = "Baixar conta parcial"
                ID_documento = .Lista.ListItems.Item(InitFor)
                Documento = "Documento: " & TBAbrir!txt_ndocumento
                Documento1 = ""
                ProcGravaEvento
                '==================================
            End If
        End If
    Next InitFor
End With
USMsgBox ("Conta baixada parcialmente com sucesso."), vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcReceberIntegral()
On Error GoTo tratar_erro

With frmContas_Receber
    ID_varias = 0
    If Chk_mov_total.Value = 1 Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from tbl_Contas_Varias", Conexao, adOpenKeyset, adLockOptimistic
        TBAbrir.AddNew
        TBAbrir.Update
        ID_varias = TBAbrir!ID
        TBAbrir.Close
    End If
    ID_variasTexto = ID_varias
    
    Contador = 0
    For InitFor = 1 To .Lista.ListItems.Count
        If .Lista.ListItems.Item(InitFor).Checked = True Then
            Contador = Contador + 1
            If Contador > 1 Then GoTo Prosseguir
        End If
    Next InitFor

Prosseguir:
    For InitFor = 1 To .Lista.ListItems.Count
        If .Lista.ListItems.Item(InitFor).Checked = True Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_contas_receber  where idintconta = " & .Lista.ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                If TBAbrir!Antecipacao = False Then
                    TBAbrir!Logsit = "S"
                    If Contador > 1 Then TBAbrir!valortitulorecebido = TBAbrir!valor Else TBAbrir!valortitulorecebido = txt_ValorPago.Text
                    Valor1 = TBAbrir!valortitulorecebido
                Else
                    TBAbrir!valortitulorecebido = 0
                    Valor1 = TBAbrir!valor
                End If
            
                TBAbrir!FormaBaixa = cmb_forma.Text
                TBAbrir!Data_pagamento = txt_DtPagto.Value
                TBAbrir!Data_movimentacao = Cmb_data_movimentacao.Value
                If TBAbrir!status = "TÍTULO RECEBIDO PARCIAL" Then
                    TBAbrir!status = "TÍTULO RECEBIDO PARCIAL LIQUIDADO"
                    TBAbrir!ValorPendente = 0
                    TBAbrir!tituloref = TBAbrir!IDintconta
                Else
                    If TBAbrir!status = "DUPLICATA DESCONTADA EM ABERTO" Then
                            TBAbrir!status = "DUPLICATA DESCONTADA LIQUIDADA"
                        ElseIf TBAbrir!Antecipacao = True Then
                                TBAbrir!status = "TÍTULO LIQUIDADO ANTECIPADO"
                            ElseIf TBAbrir!Devolucao = True Then
                                    TBAbrir!status = "TÍTULO DEVOLVIDO LIQUIDADO"
                                Else
                                    TBAbrir!status = "TÍTULO LIQUIDADO"
                    End If
                End If
                If pubUsuario <> "" Then TBAbrir!resprec = pubUsuario
                TBAbrir!NDoctoBaixa = txt_ndocumento.Text
                TBAbrir!Banco = cmb_Banco.Text
                If Contador > 1 Then TBAbrir!Obs = TBAbrir!Observacoes Else TBAbrir!Obs = txtObs.Text
                TBAbrir!Dias_atraso = IIf(Txt_dias_atraso = "", 0, Txt_dias_atraso)
                
                'Família de contas
                Conexao.Execute "Update familia_financeiro Set Pago_recebido = 'True' where IDConta = " & TBAbrir!IDintconta & " and tipoconta = 'R' and Pago_recebido = 'False'"
                
                valor = txt_VlrDocto
                If chkjuros.Value = 1 Then
                    TBAbrir!Juros_valor = Format(txtjuros, "###,##0.0000000")
                    Valor_IPI = txtjuros
                    Valor_IPI = Valor_IPI * 100
                    TBAbrir!Juros = Valor_IPI / valor
                End If
                If Chk_multa.Value = 1 Then
                    TBAbrir!Multa_valor = Format(Txt_multa, "###,##0.0000000")
                    Valor_IPI = Txt_multa
                    Valor_IPI = Valor_IPI * 100
                    TBAbrir!Multa = Valor_IPI / valor
                End If
                If chkdesconto.Value = 1 Then
                    TBAbrir!Desconto_valor = Format(txtDesconto, "###,##0.0000000")
                    Valor_IPI = txtDesconto
                    Valor_IPI = Valor_IPI * 100
                    TBAbrir!Desconto = Valor_IPI / valor
                End If
                TBAbrir!ID_varias = ID_varias
                TBAbrir.Update
                
                ProcGavarPCJurosMulta TBAbrir!IDintconta, IIf(IsNull(TBAbrir!ID_nota), 0, TBAbrir!ID_nota), IIf(IsNull(TBAbrir!Juros_valor), 0, TBAbrir!Juros_valor) * TBAbrir!Dias_atraso, IIf(IsNull(TBAbrir!Multa_valor), 0, TBAbrir!Multa_valor), "R", True
                
                'Verifica se o valor baixado é igual ao valor da conta de antecipação selecionada na lista
                Valor2 = 0
                If TBAbrir!Antecipacao = False And TBAbrir!Devolucao = False Then
                    For InitFor1 = 1 To Lista.ListItems.Count
                        If Lista.ListItems.Item(InitFor1).Checked = True Then Valor2 = Valor2 + Lista.ListItems.Item(InitFor1).ListSubItems(3)
                    Next InitFor1
                End If
                                
                'Fluxo de Caixa
                Set TBFluxo = CreateObject("adodb.recordset")
                TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBAbrir!IDFluxo), 0, TBAbrir!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                If TBFluxo.EOF = True Then TBFluxo.AddNew
                TBFluxo!IDintconta = TBAbrir!IDintconta
                TBFluxo!Operacao = "Crédito"
                TBFluxo!Data = Cmb_data_movimentacao
                If cmb_forma <> "CHEQUE" And cmb_forma <> "CHEQUE PRÉ-DATADO" And cmb_forma <> "DOC" And cmb_forma <> "TED" And Txt_historico <> "" Then
                    TBFluxo!Descricao = Txt_historico
                Else
                    TBFluxo!Descricao = TBAbrir!Nome_Razao
                End If
                TBFluxo!status = "S"
                TBFluxo!int_NotaFiscal = TBAbrir!NFiscal
                TBFluxo!Documento = TBAbrir!txt_ndocumento
                TBFluxo!Instituicao = cmb_Banco
                TBFluxo!Hora = Format(Now, "hh:mm:ss")
                TBFluxo!Obs = IIf(txtObsFluxo = "", TBFluxo!Descricao, txtObsFluxo)
                
                If Valor2 = 0 Then 'Valor antecipado igual a 0
                    Valor3 = Valor1
                    If TBAbrir!titulodesc = True Then TBFluxo!Bloqueado = True Else TBFluxo!Bloqueado = False
                ElseIf Valor1 > Valor2 Then 'Valor pago maior que o valor antecipado
                        Valor3 = Valor1 - Valor2
                        If TBAbrir!titulodesc = True Then TBFluxo!Bloqueado = True Else TBFluxo!Bloqueado = False
                    Else
                        Valor3 = Valor1 'Valor pago menor ou igual ao valor antecipado
                        TBFluxo!Bloqueado = True
                End If
                TBFluxo!valor = Format(Valor3, "###,##0.00")
                                
                TBFluxo!ID_empresa = TBAbrir!ID_empresa
                TBFluxo!ID_varias = 0
                If txt_ndocumento <> "" Then TBFluxo!Cheque = txt_ndocumento
                TBFluxo!tituloref = TBAbrir!IDintconta
                If (cmb_forma = "CHEQUE" Or cmb_forma = "CHEQUE PRÉ-DATADO" Or cmb_forma = "DOC" Or cmb_forma = "TED" Or cmb_forma = "MALOTE" Or ID_varias <> 0) And TBAbrir!status <> "DUPLICATA DESCONTADA EM ABERTO" Then
                    TBFluxo!Bloqueado = True
                    TextoFiltroData = "Data = '" & Format(Cmb_data_movimentacao.Value, "Short Date") & "' and"
                    Select Case cmb_forma
                        Case "CHEQUE":
                            Descricao = "Cheque n. " & txt_ndocumento
                            TextoFiltroData = ""
                        Case "CHEQUE PRÉ-DATADO":
                            Descricao = "Cheque n. " & txt_ndocumento
                            TextoFiltroData = ""
                        Case "DOC": Descricao = "Doc n. " & txt_ndocumento
                        Case "TED": Descricao = "Ted n. " & txt_ndocumento
                        Case "MALOTE": Descricao = "Malote n. " & txt_ndocumento
                    End Select
                    
                    'Cria registro com o valor total da operação
                    valor = 0
                    Set TBFI = CreateObject("adodb.recordset")
                    If ID_varias = 0 Or cmb_forma = "CHEQUE" Or cmb_forma = "CHEQUE PRÉ-DATADO" Or cmb_forma = "DOC" Or cmb_forma = "TED" Or cmb_forma = "MALOTE" Then
                        TextoFiltro1 = "Sum(Valor) as Valortotal from tbl_Fluxo_de_caixa"
                        TextoFiltro = TextoFiltroData & " Operacao = 'Crédito' and Descricao = '" & Descricao & "' and Cheque = '" & txt_ndocumento & "' and Instituicao = '" & cmb_Banco & "'"
                    Else
                        TextoFiltro1 = "Sum(" & IIf(TBAbrir!Antecipacao = False, "valortitulorecebido", "Valor") & ") as Valortotal from tbl_contas_receber"
                        TextoFiltro = "ID_varias = " & ID_varias
                    End If
                    TBFI.Open "Select " & TextoFiltro1 & " where " & TextoFiltro & " and IdIntConta <> " & TBAbrir!IDintconta, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        valor = IIf(IsNull(TBFI!ValorTotal), 0, TBFI!ValorTotal)
                    End If
                    TBFI.Close
                    valor = Format(valor + Valor3, "###,##0.00")
                    
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "Select * from tbl_Fluxo_de_caixa where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
                    If TBGravar.EOF = True Then
                        TBGravar.AddNew
                        TBGravar!Operacao = "Crédito"
                        TBGravar!Data = Cmb_data_movimentacao
                        TBGravar!Bloqueado = False
                        If cmb_forma = "CHEQUE" Or cmb_forma = "CHEQUE PRÉ-DATADO" Or cmb_forma = "DOC" Or cmb_forma = "TED" Or cmb_forma = "MALOTE" Then
                            If cmb_forma = "CHEQUE" Or cmb_forma = "CHEQUE PRÉ-DATADO" Then TBGravar!Bloqueado = True
                            TBGravar!Descricao = Descricao
                        ElseIf Contador > 1 Then
                                TBGravar!Descricao = "RCBTO. VARIAS CONTAS"
                        End If
                        TBGravar!status = "S"
                        TBGravar!Instituicao = cmb_Banco
                        TBGravar!Hora = TBFluxo!Hora
                        TBGravar!Cheque = txt_ndocumento
                        TBGravar!Obs = IIf(txtObsFluxo = "", TBGravar!Descricao, txtObsFluxo)
                        TBGravar!ID_empresa = TBAbrir!ID_empresa
                        TBGravar!ID_varias = ID_varias
                    End If
                    TBGravar!valor = valor
                    TBGravar.Update
                    ID_variasTexto = IIf(IsNull(TBGravar!ID_varias), 0, TBGravar!ID_varias)
                    TBGravar.Close
                End If
                TBFluxo.Update
                Conexao.Execute "Update tbl_contas_receber set IDFluxo = " & TBFluxo!IDFluxo & ", ID_varias = " & ID_variasTexto & " where IDIntconta = " & TBAbrir!IDintconta
                TBFluxo.Close
                
                Qtd = Valor1 'Valor recebido
                If Devolucao = True Then
                    ProcDevolucao TBAbrir!IDintconta
                ElseIf Valor2 > 0 Then
                        ProcAntecipacao TBAbrir!IDintconta
                End If
                
                If TBAbrir!status <> "DUPLICATA DESCONTADA LIQUIDADA" And cmb_forma <> "CHEQUE" And cmb_forma <> "CHEQUE PRÉ-DATADO" Then
                    Set TBItem = CreateObject("adodb.recordset")
                    TBItem.Open "Select * from tbl_instituicoes where ID = " & cmb_Banco.ItemData(cmb_Banco.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
                    If TBItem.EOF = False Then
                        If Devolucao = True Then TBItem!Saldo = Format(TBItem!Saldo - Valor_Produto, "###,##0.00") Else TBItem!Saldo = Format(TBItem!Saldo + Qtd, "###,##0.00")
                        TBItem.Update
                    End If
                    TBItem.Close
                End If
                
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select local_troca from Troca_titulo where ID = " & IIf(IsNull(TBAbrir!IDtrocatitulo), 0, TBAbrir!IDtrocatitulo), Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    Set TBReceber = CreateObject("adodb.recordset")
                    TBReceber.Open "Select Sum(tbl_contas_receber.Valor) as Valor from tbl_contas_receber INNER JOIN troca_titulo on tbl_contas_receber.Idtrocatitulo = troca_titulo.ID where troca_titulo.Local_troca = '" & TBContas!local_troca & "' and tbl_contas_receber.ID_empresa = " & TBAbrir!ID_empresa & " and tbl_contas_receber.status = 'DUPLICATA DESCONTADA EM ABERTO' and tbl_contas_receber.Logsit = 'N'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBReceber.EOF = False Then
                        valor = IIf(IsNull(TBReceber!valor), 0, TBReceber!valor)
                        NovoValor = Replace(valor, ",", ".")
                        Conexao.Execute "Update tbl_Instituicoes Set Limite_utilizado = " & NovoValor & " where txt_Descricao = '" & TBContas!local_troca & "' and ID_empresa = " & TBAbrir!ID_empresa
                    End If
                    TBReceber.Close
                End If
                TBContas.Close
                
                '==================================
                Modulo = "Financeiro/Contas a receber"
                Evento = "Baixar conta"
                ID_documento = TBAbrir!IDintconta
                Documento = "Documento: " & TBAbrir!txt_ndocumento
                Documento1 = ""
                ProcGravaEvento
                '==================================
            End If
            TBAbrir.Close
        End If
    Next InitFor
End With
USMsgBox ("Conta(s) baixada(s) com sucesso."), vbInformation, "CAPRIND v5.0"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0:
        If cmb_Banco.Visible = True Then cmb_Banco.SetFocus
        Lista.Visible = False
    Case 1:
        If Lista.Visible = True Then Lista.SetFocus
        Lista.Visible = True
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro
    
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_dias_atraso_Change()
On Error GoTo tratar_erro

If Txt_dias_atraso.Text <> "" Then
    VerifNumero = Txt_dias_atraso.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_dias_atraso.Text = ""
        Exit Sub
    End If
End If
ProcCalculaJurosMulta

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txt_DtPagto_Change()
On Error GoTo tratar_erro

ProcVerifDiasAtraso
Cmb_data_movimentacao = txt_DtPagto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerifDiasAtraso()
On Error GoTo tratar_erro

Data = txt_DtPagto

Contador = 0
With frmContas_Receber
    For InitFor = 1 To .Lista.ListItems.Count
        If .Lista.ListItems.Item(InitFor).Checked = True Then
            DataFim = .Lista.ListItems.Item(InitFor).ListSubItems(2)
            Contador = Contador + 1
            If Contador > 1 Then GoTo Prosseguir
        End If
    Next InitFor
End With

Prosseguir:
    If Contador = 1 Then
        With chkdesconto
            .Value = 0
            .Enabled = True
        End With
        If Data > DataFim Then
            With Txt_dias_atraso
                .Text = Data - DataFim
                .Locked = False
                .TabStop = True
            End With
            
            If chbparcial.Value = 0 Then
                chkjuros.Enabled = True
                Chk_multa.Enabled = True
            End If
        Else
            With Txt_dias_atraso
                .Text = 0
                .Locked = True
                .TabStop = Fase
            End With
            
            With chkjuros
                .Value = 0
                .Enabled = False
            End With
            With txtjuros
                .Text = ""
                .Locked = True
                .TabStop = False
            End With
            With Chk_multa
                .Value = 0
                .Enabled = False
            End With
            With Txt_multa
                .Text = ""
                .Locked = True
                .TabStop = False
            End With
        End If
        If chkjuros.Value = 1 Or Chk_multa.Value = 1 Then ProcCalculaJurosMulta
        If chkdesconto.Value = 1 Then ProcCalculaDesconto
    End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_historico_Change()
On Error GoTo tratar_erro

txtObsFluxo = Txt_historico

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_multa_Change()
On Error GoTo tratar_erro

If Txt_multa.Text <> "" Then
    VerifNumero = Txt_multa.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_multa.Text = ""
        Exit Sub
    End If
End If
ProcCalculaJurosMulta

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_multa_LostFocus()
On Error GoTo tratar_erro

If Txt_multa <> "" Then Txt_multa.Text = Format(Txt_multa.Text, "###,##0.0000000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txt_ndocumento_LostFocus()
On Error GoTo tratar_erro

Txt_historico = ""
ProcCriaHistPadrao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriaHistPadrao()
On Error GoTo tratar_erro

Nbanco = ""
If cmb_Banco <> "" Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select int_NBanco from tbl_instituicoes where ID = " & cmb_Banco.ItemData(cmb_Banco.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Nbanco = TBAbrir!int_NBanco
    End If
    TBAbrir.Close
End If

Select Case cmb_forma
    Case "CHEQUE": Txt_historico = "Cheque n. " & txt_ndocumento
    Case "CHEQUE PRÉ-DATADO": Txt_historico = "Cheque n. " & txt_ndocumento
    Case "DOC":
        Txt_historico = "Doc n. " & txt_ndocumento
        If Nbanco = 104 Then txtObsFluxo = "DOC ELET"
    Case "TED":
        Txt_historico = "Ted n. " & txt_ndocumento
        If Nbanco = 104 Then txtObsFluxo = "CRED TED"
    Case "MALOTE": Txt_historico = "Malote n. " & txt_ndocumento
    Case "TRANSFERÊNCIA ENTRE CONTAS": If Nbanco = 104 Then Txt_historico = "CRED TEV"
    Case "TEV": If Nbanco = 104 Then Txt_historico = "CRED TEV"
    Case "BOLETO": If Nbanco = 104 Then Txt_historico = "COB COMPE"
    Case "BOLETO BANCÁRIO": If Nbanco = 104 Then Txt_historico = "COB COMPE"
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txt_ValorPago_Change()
On Error GoTo tratar_erro

If txt_ValorPago.Text <> "" Then
    VerifNumero = txt_ValorPago.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txt_ValorPago.Text = ""
        txt_ValorPago.SetFocus
        Exit Sub
    End If
    ProcAtualizaSaldo
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaSaldo()
On Error GoTo tratar_erro

If cmb_Banco <> "" And cmb_forma <> "CHEQUE" And cmb_forma <> "CHEQUE PRÉ-DATADO" Then
    SaldoAtual = IIf(txtSaldoAtual = "", 0, txtSaldoAtual)
    valor = IIf(txt_ValorPago = "", 0, txt_ValorPago)
    txtSaldo = Format(SaldoAtual + valor, "###,##0.00")
Else
    txtSaldo = txtSaldoAtual
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txt_ValorPago_LostFocus()
On Error GoTo tratar_erro

txt_ValorPago.Text = Format(txt_ValorPago.Text, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesconto_Change()
On Error GoTo tratar_erro

If txtDesconto.Text <> "" Then
    VerifNumero = txtDesconto.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtDesconto.Text = ""
        Exit Sub
    End If
End If
ProcCalculaDesconto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaDesconto()
On Error GoTo tratar_erro

valor = txt_VlrDocto
Valor_IPI = IIf(txtDesconto = "", 0, txtDesconto)
valor = valor - Valor_IPI
txt_ValorPago.Text = Format(valor, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtdesconto_LostFocus()
On Error GoTo tratar_erro

If txtDesconto <> "" Then txtDesconto.Text = Format(txtDesconto.Text, "###,##0.0000000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtJuros_Change()
On Error GoTo tratar_erro

If txtjuros.Text <> "" Then
    VerifNumero = txtjuros.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtjuros.Text = ""
        Exit Sub
    End If
End If
ProcCalculaJurosMulta

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaJurosMulta()
On Error GoTo tratar_erro

If chbparcial.Value = 0 Then
    valor = txt_VlrDocto
    Valor_IPI = IIf(txtjuros = "", 0, txtjuros)
    Valor_IPI = Valor_IPI * IIf(Txt_dias_atraso = "", 0, Txt_dias_atraso)
    valor = valor + Valor_IPI
    ValorTotal = IIf(Txt_multa = "", 0, Txt_multa)
    txt_ValorPago.Text = Format(valor + ValorTotal, "###,##0.00")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtjuros_LostFocus()
On Error GoTo tratar_erro

If txtjuros <> "" Then txtjuros.Text = Format(txtjuros.Text, "###,##0.0000000")
    
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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcBaixar
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
    Case 1: ProcFiltrar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaAntecipacao()
On Error GoTo tratar_erro

With Lista
    .ListItems.Clear
    .ColumnHeaders(5).Width = 1100
    .ColumnHeaders(6).Width = 2000
End With

TextoFiltro = ""
With frmContas_Receber
    For InitFor = 1 To .Lista.ListItems.Count
        If .Lista.ListItems.Item(InitFor).Checked = True Then
            If TextoFiltro = "" Then
                If Contador > 1 Then TextoFiltro = "("
                TextoFiltro = TextoFiltro & "CR.IdIntConta = " & .Lista.ListItems.Item(InitFor)
            Else
                TextoFiltro = TextoFiltro & " or CR.IdIntConta = " & .Lista.ListItems.Item(InitFor)
            End If
        End If
    Next InitFor
    If Contador > 1 Then TextoFiltro = TextoFiltro & ")"
End With

lblRegistros.Caption = "Nº de reg.: 0"
lblPaginas.Caption = "Pág.: 0 de: 0"
Set TBLISTA_Contas_Receber_Baixar = CreateObject("adodb.recordset")
TBLISTA_Contas_Receber_Baixar.Open "Select CR1.IdIntConta, CR1.emissao, CR1.Vencimento, CR1.Valor, CR1.Saldo_antecipacao, CR1.nfiscal, CR1.Parcela from tbl_Contas_receber AS CR LEFT OUTER JOIN tbl_Contas_receber AS CR1 ON CR.ID_empresa = CR1.ID_empresa AND CR.IDcliente = CR1.IDcliente where " & TextoFiltro & " and CR1.Antecipacao = 1 and CR1.status = 'TÍTULO LIQUIDADO ANTECIPADO' and CR1.Saldo_antecipacao > 0 and CR1.Bloqueado = 'False' and CR1.vencimento Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' group by CR1.IdIntConta, CR1.emissao, CR1.Vencimento, CR1.Valor, CR1.Saldo_antecipacao, CR1.nfiscal, CR1.Parcela order by CR1.Vencimento, CR1.IdIntConta", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Contas_Receber_Baixar.EOF = False Then ProcExibePagina (1), Devolucao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina, Devolucao As Boolean)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBLISTA_Contas_Receber_Baixar.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBLISTA_Contas_Receber_Baixar.AbsolutePage = Pagina
TamanhoPagina = TBLISTA_Contas_Receber_Baixar.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBLISTA_Contas_Receber_Baixar.RecordCount - IIf(Pagina > 1, (TBLISTA_Contas_Receber_Baixar.PageSize * (Pagina - 1)), 0), TBLISTA_Contas_Receber_Baixar.PageSize)
PBLista.Value = 1
Contador = 0
Do While TBLISTA_Contas_Receber_Baixar.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        If Devolucao = True Then
            .Add , , TBLISTA_Contas_Receber_Baixar!IDintconta
            .Item(.Count).SubItems(1) = Format(TBLISTA_Contas_Receber_Baixar!emissao, "dd/mm/yy")
            .Item(.Count).SubItems(2) = Format(TBLISTA_Contas_Receber_Baixar!Vencimento, "dd/mm/yy")
            .Item(.Count).SubItems(3) = Format(TBLISTA_Contas_Receber_Baixar!valor, "###,##0.00")
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Contas_Receber_Baixar!NFiscal), "", TBLISTA_Contas_Receber_Baixar!NFiscal)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Contas_Receber_Baixar!Parcela), "", TBLISTA_Contas_Receber_Baixar!Parcela)
        Else
            .Add , , TBLISTA_Contas_Receber_Baixar!IDintconta
            .Item(.Count).SubItems(1) = Format(TBLISTA_Contas_Receber_Baixar!emissao, "dd/mm/yy")
            .Item(.Count).SubItems(2) = Format(TBLISTA_Contas_Receber_Baixar!Vencimento, "dd/mm/yy")
            .Item(.Count).SubItems(3) = Format(TBLISTA_Contas_Receber_Baixar!valor, "###,##0.00")
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA_Contas_Receber_Baixar!Saldo_antecipacao), "", Format(TBLISTA_Contas_Receber_Baixar!Saldo_antecipacao, "###,##0.00"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA_Contas_Receber_Baixar!NFiscal), "", TBLISTA_Contas_Receber_Baixar!NFiscal)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA_Contas_Receber_Baixar!Parcela), "", TBLISTA_Contas_Receber_Baixar!Parcela)
        End If
    End With
    TBLISTA_Contas_Receber_Baixar.MoveNext
    ContadorReg = ContadorReg + 1
    Contador = Contador + 1
    PBLista.Value = Contador
Loop
lblRegistros.Caption = "Nº de reg.: " & TBLISTA_Contas_Receber_Baixar.RecordCount
If TBLISTA_Contas_Receber_Baixar.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Pág.: 1 de: " & TBLISTA_Contas_Receber_Baixar.PageCount
ElseIf TBLISTA_Contas_Receber_Baixar.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Pág.: " & TBLISTA_Contas_Receber_Baixar.PageCount & " de: " & TBLISTA_Contas_Receber_Baixar.PageCount
    Else
        lblPaginas.Caption = "Pág.: " & TBLISTA_Contas_Receber_Baixar.AbsolutePage - 1 & " de: " & TBLISTA_Contas_Receber_Baixar.PageCount
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAntecipacao(IDConta As Long)
On Error GoTo tratar_erro

With Lista
    For InitFor1 = 1 To .ListItems.Count
        If .ListItems.Item(InitFor1).Checked = True Then
            qt = 0
            If Qtd > 0 Then
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "SELECT * from tbl_Contas_antecipacao", Conexao, adOpenKeyset, adLockOptimistic
                TBContas.AddNew
                TBContas!ID_conta = IDConta
                TBContas!ID_antecipacao = .ListItems(InitFor1)
                Qtde = .ListItems.Item(InitFor1).SubItems(4) 'Valor antecipado
                If Qtde >= Qtd Then qt = Qtd Else qt = Qtde
                TBContas!valor = Format(qt, "###,##0.00")
                TBContas!Tipo = "R"
                TBContas.Update
                
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select * from tbl_contas_receber where idintconta = " & .ListItems.Item(InitFor1), Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    TBContas!Saldo_antecipacao = Format(TBContas!Saldo_antecipacao - qt, "###,##0.00")
                    If TBContas!Saldo_antecipacao = 0 Then TBContas!Logsit = "S" Else TBContas!Logsit = "N"
                    TBContas.Update
                End If
                TBContas.Close
            End If
            Qtd = Qtd - qt
            Valor2 = Valor2 - qt
        End If
    Next InitFor1
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
If Devolucao = True Then
    ProcCarregaListaDevolucao
ElseIf Antecipacao = False Then
        ProcCarregaListaAntecipacao
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaDevolucao()
On Error GoTo tratar_erro

With Lista
    .ListItems.Clear
    .ColumnHeaders(5).Width = 0
    .ColumnHeaders(6).Width = 3100
End With

TextoFiltro = ""
With frmContas_Receber
    For InitFor = 1 To .Lista.ListItems.Count
        If .Lista.ListItems.Item(InitFor).Checked = True Then
            If TextoFiltro = "" Then
                If Contador > 1 Then TextoFiltro = "("
                TextoFiltro = TextoFiltro & "CR.IdIntConta = " & .Lista.ListItems.Item(InitFor)
            Else
                TextoFiltro = TextoFiltro & " or CR.IdIntConta = " & .Lista.ListItems.Item(InitFor)
            End If
        End If
    Next InitFor
    If Contador > 1 Then TextoFiltro = TextoFiltro & ")"
End With

lblRegistros.Caption = "Nº de reg.: 0"
lblPaginas.Caption = "Pág.: 0 de: 0"
Set TBLISTA_Contas_Receber_Baixar = CreateObject("adodb.recordset")
TBLISTA_Contas_Receber_Baixar.Open "Select CR1.IdIntConta, CR1.emissao, CR1.Vencimento, CR1.Valor - SUM(ISNULL(CD.Valor, 0)) AS Valor, CR1.Nfiscal, CR1.Parcela from tbl_contas_devolucao AS CD RIGHT OUTER JOIN tbl_Contas_receber AS CR1 ON CD.ID_conta = CR1.IdIntConta and CD.Tipo = 'R' RIGHT OUTER JOIN tbl_Contas_receber AS CR ON CR1.ID_empresa = CR.ID_empresa AND CR1.IDCliente = CR.IDCliente where " & TextoFiltro & " and CR.Devolucao = 1 and CR1.Antecipacao = 0 and CR1.Bloqueado = 'False' and CR1.Devolucao = 0 and CR1.vencimento Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' group by CR1.IdIntConta, CR1.emissao, CR1.Vencimento, CR1.valor, CR1.nfiscal, CR1.Parcela Having (CR1.Valor - Sum(IsNull(CD.Valor, 0)) > 0) order by CR1.Vencimento, CR1.IdIntConta", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA_Contas_Receber_Baixar.EOF = False Then ProcExibePagina (1), Devolucao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcDevolucao(IDConta As Long)
On Error GoTo tratar_erro

Qtd = Valor1 'Valor pago (devolvido)
Valor_Produto = 0 'Soma  do valor das contas recebidas

If USMsgBox("Deseja que mostre a baixa no extrato bancário/fluxo de caixa?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    permitido_devolucao = True
Else
    permitido_devolucao = False
End If

With Lista
    For InitFor1 = 1 To .ListItems.Count
        If .ListItems.Item(InitFor1).Checked = True Then
            If Qtd < 0 Then
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "SELECT * from tbl_Contas_devolucao", Conexao, adOpenKeyset, adLockOptimistic
                TBContas.AddNew
                TBContas!ID_Devolucao = IDConta
                Qtde = .ListItems.Item(InitFor1).SubItems(3) 'Valor da conta
                SaqueUtilizado = Format((Qtd * -1), "0.00")
                
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "SELECT * from tbl_Contas_receber where IDIntconta = " & .ListItems.Item(InitFor1), Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    TBContas!Logsit = TBFI!Logsit
                    If TBFI!Logsit = "S" Then
                        Valor_Produto = Valor_Produto + Qtde
                        TBContas!ID_conta = .ListItems(InitFor1)
                    Else
                        'Precisa estar com o format na linha abaixo devido a um erro do proprio vb6
                        If Format(Qtde, "0.00") > SaqueUtilizado Then
                            ProcCriarContaReceberParcialDev
                        Else
                            TBContas!ID_conta = .ListItems(InitFor1)
                            ProcReceberIntegralDev
                        End If
                        If Permitido2 = False Then Conexao.Execute "Update FC set FC.Bloqueado = 'True' from tbl_Fluxo_de_caixa FC INNER JOIN tbl_contas_receber CR ON CR.IDFluxo = FC.IDFluxo where CR.IDIntconta = " & IDConta
                    End If
                End If
                TBFI.Close
                TBContas!valor = Format(IIf(Qtde > SaqueUtilizado, SaqueUtilizado, Qtde), "###,##0.00")
                TBContas!Tipo = "R"
                TBContas.Update
                TBContas.Close
            End If
            Qtd = Qtd + Qtde
        End If
    Next InitFor1
End With
SaqueUtilizado = Format((Valor1 * -1), "0.00")
If Valor_Produto > SaqueUtilizado Then Valor_Produto = Valor1 * -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarContaReceberParcialDev()
On Error GoTo tratar_erro

TBFI!ValorPendente = Format(Qtde + Qtd, "###,##0.00")
TBFI!valor = TBFI!ValorPendente
If TBFI!valorprincipal = 0 Then TBFI!valorprincipal = Qtde
TBFI!tituloref = TBFI!IDintconta
TBFI!status = "TÍTULO RECEBIDO PARCIAL"
TBFI!Banco = cmb_Banco
TBFI.Update

'Fluxo de Caixa
Set TBFluxo = CreateObject("adodb.recordset")
TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBFI!IDFluxo), 0, TBFI!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
If TBFluxo.EOF = False Then
    TBFluxo!valor = TBFI!ValorPendente
    TBFluxo!Instituicao = cmb_Banco
    TBFluxo.Update
End If
TBFluxo.Close

'Cria e recebe nova conta parcial
Set TBCorretiva = CreateObject("adodb.recordset")
TBCorretiva.Open "Select * from tbl_contas_receber", Conexao, adOpenKeyset, adLockOptimistic
TBCorretiva.AddNew
TBCorretiva!tituloref = TBFI!IDintconta
TBCorretiva!valor = Qtd * -1
TBCorretiva!RecebidoParcial = Qtd * -1
TBCorretiva!Parcial = True
TBCorretiva!status = "TÍTULO RECEBIDO PARCIAL"
If pubUsuario <> "" Then TBCorretiva!resprec = pubUsuario
TBCorretiva!Antecipacao = False
TBCorretiva!Devolucao = False
TBCorretiva!Data_transacao = TBFI!Data_transacao
TBCorretiva!Tipo_doc = TBFI!Tipo_doc
TBCorretiva!txt_ndocumento = IIf(IsNull(TBFI!txt_ndocumento), "", TBFI!txt_ndocumento)
TBCorretiva!NFiscal = TBFI!NFiscal
TBCorretiva!Proposta = TBFI!Proposta
TBCorretiva!emissao = TBFI!emissao
TBCorretiva!valortitulorecebido = Qtd * -1
TBCorretiva!ID_nota = TBFI!ID_nota
TBCorretiva!Parcela = TBFI!Parcela
TBCorretiva!IDCliente = TBFI!IDCliente
TBCorretiva!Logsit = "S"
TBCorretiva!Tipo = TBFI!Tipo
TBCorretiva!Nome_Razao = TBFI!Nome_Razao
TBCorretiva!Cidade = TBFI!Cidade
TBCorretiva!Estado = TBFI!Estado
TBCorretiva!Tipo = TBFI!Tipo
TBCorretiva!valorprincipal = TBFI!valorprincipal
TBCorretiva!ID_nota = TBFI!ID_nota
TBCorretiva!Vencimento = TBFI!Vencimento
'Dados do recebimento
TBCorretiva!FormaBaixa = cmb_forma.Text
TBCorretiva!Data_pagamento = txt_DtPagto.Value
TBCorretiva!Data_movimentacao = Cmb_data_movimentacao.Value
TBCorretiva!valortitulorecebido = Qtd * -1
TBCorretiva!valorprincipal = TBFI!valorprincipal
TBCorretiva!NDoctoBaixa = txt_ndocumento.Text
TBCorretiva!Banco = cmb_Banco.Text
TBCorretiva!Obs = txtObs.Text
TBCorretiva!ValorPendente = TBFI!ValorPendente
TBCorretiva!Dias_atraso = IIf(Txt_dias_atraso = "", 0, Txt_dias_atraso)
TBCorretiva!ID_empresa = TBFI!ID_empresa
TBCorretiva.Update

TBContas!ID_conta = TBCorretiva!IDintconta

'Fluxo de Caixa
Set TBFluxo = CreateObject("adodb.recordset")
TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
If TBFluxo.EOF = True Then TBFluxo.AddNew
TBFluxo!IDintconta = TBCorretiva!IDintconta
TBFluxo!Operacao = "Crédito"
TBFluxo!Data = Cmb_data_movimentacao
TBFluxo!Descricao = TBFI!Nome_Razao
TBFluxo!status = "S"
TBFluxo!int_NotaFiscal = IIf(IsNull(TBFI!txt_ndocumento), "", TBFI!txt_ndocumento)
TBFluxo!Obs = IIf(txtObsFluxo = "", TBFluxo!Descricao, txtObsFluxo)
TBFluxo!valor = Qtd * -1

If TBFI!titulodesc = True Then
    Permitido2 = False
    TBFluxo!Bloqueado = True
Else
    If permitido_devolucao = True Then
        Permitido2 = True
        TBFluxo!Bloqueado = False
    Else
        Permitido2 = False
        TBFluxo!Bloqueado = True
    End If
End If

TBFluxo!ID_empresa = TBFI!ID_empresa
TBFluxo!Instituicao = cmb_Banco
TBFluxo!Hora = Format(Now, "hh:mm:ss")
If txt_ndocumento <> "" Then TBFluxo!Cheque = txt_ndocumento
TBFluxo!tituloref = TBFI!IDintconta
TBFluxo.Update
Conexao.Execute "Update tbl_contas_receber set IDFluxo = " & TBFluxo!IDFluxo & " where IdIntConta = " & TBCorretiva!IDintconta
TBFluxo.Close

'Família de contas
qt = 0
ValorTotal = Qtde
Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "Select * from familia_financeiro where IDConta = " & TBFI!IDintconta & " and tipoconta = 'R' and Pago_recebido = 'False' order by ID_PC", Conexao, adOpenKeyset, adLockOptimistic
If TBFamilia.EOF = False Then
    Contador = TBFamilia.RecordCount
    Do While TBFamilia.EOF = False
        'Verifica a porcentagem representada pelo valor da família
        Valor2 = TBFamilia!valor
        VlrIPI = Format((Valor2 * 100) / ValorTotal, "###,##0.0000000000")
        
        valor = Format(((Qtd * -1) * VlrIPI) / 100, "###,##0.00")
    
        TBFamilia!Pago_recebido = True
        TBFamilia!IDConta = TBCorretiva!IDintconta
        qt = TBFamilia!valor - valor 'Valor a pagar
        TBFamilia!valor = valor
        If qt > 0 Then
            Set TBCiclo = CreateObject("adodb.recordset")
            TBCiclo.Open "select * FROM familia_financeiro", Conexao, adOpenKeyset, adLockOptimistic
            TBCiclo.AddNew
            TBCiclo!ID_PC = TBFamilia!ID_PC
            TBCiclo!IDConta = TBFI!IDintconta
            TBCiclo!IDnota = TBFamilia!IDnota
            TBCiclo!TipoConta = TBFamilia!TipoConta
            TBCiclo!Pago_recebido = False
            TBCiclo!valor = qt
            TBCiclo.Update
            TBCiclo.Close
        End If
        TBFamilia.Update
        TBFamilia.MoveNext
    Loop
End If
TBFamilia.Close
TBCorretiva.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcReceberIntegralDev()
On Error GoTo tratar_erro

TBFI!Logsit = "S"
TBFI!valortitulorecebido = Qtde

TBFI!FormaBaixa = cmb_forma.Text
TBFI!Data_pagamento = txt_DtPagto.Value
TBFI!Data_movimentacao = Cmb_data_movimentacao.Value
If TBFI!status = "TÍTULO RECEBIDO PARCIAL" Then
    TBFI!status = "TÍTULO RECEBIDO PARCIAL LIQUIDADO"
    TBFI!ValorPendente = 0
    TBFI!tituloref = TBFI!IDintconta
Else
    If TBFI!titulodesc = True Then TBFI!status = "DUPLICATA DESCONTADA LIQUIDADA" Else TBFI!status = "TÍTULO LIQUIDADO"
End If
If pubUsuario <> "" Then TBFI!resprec = pubUsuario
TBFI!NDoctoBaixa = txt_ndocumento.Text
TBFI!Banco = cmb_Banco.Text
TBFI!Obs = txtObs.Text
TBFI.Update

'Família de contas
Conexao.Execute "Update familia_financeiro Set Pago_recebido = 'True' where IDConta = " & TBFI!IDintconta & " and tipoconta = 'R' and Pago_recebido = 'False'"

'Fluxo de Caixa
Set TBFluxo = CreateObject("adodb.recordset")
TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBFI!IDFluxo), 0, TBFI!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
If TBFluxo.EOF = True Then TBFluxo.AddNew
TBFluxo!IDintconta = TBFI!IDintconta
TBFluxo!Operacao = "Crédito"
TBFluxo!Data = Cmb_data_movimentacao
TBFluxo!Descricao = TBFI!Nome_Razao
TBFluxo!status = "S"
TBFluxo!int_NotaFiscal = TBFI!txt_ndocumento
TBFluxo!Instituicao = cmb_Banco
TBFluxo!Hora = Format(Now, "hh:mm:ss")
TBFluxo!Obs = IIf(txtObsFluxo = "", TBFluxo!Descricao, txtObsFluxo)

If TBFI!titulodesc = True Then
    Permitido2 = False
    TBFluxo!Bloqueado = True
Else
    If permitido_devolucao = True Then
        Permitido2 = True
        TBFluxo!Bloqueado = False
    Else
        Permitido2 = False
        TBFluxo!Bloqueado = True
    End If
End If

TBFluxo!valor = Qtde
TBFluxo!ID_empresa = TBFI!ID_empresa
TBFluxo!ID_varias = 0
If txt_ndocumento <> "" Then TBFluxo!Cheque = txt_ndocumento
TBFluxo.Update
Conexao.Execute "Update tbl_contas_receber set IDFluxo = " & TBFluxo!IDFluxo & " where IdIntConta = " & TBFI!IDintconta
TBFluxo.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
