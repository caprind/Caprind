VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCFI_Saida 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estoque - Almoxarifado - Retirar"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCFI_Saida.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   6225
      Left            =   0
      TabIndex        =   8
      Top             =   30
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   10980
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
      TabCaption(0)   =   "Dados da retirada"
      TabPicture(0)   =   "frmCFI_Saida.frx":21F49
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
      TabCaption(1)   =   "Destino/aplicação"
      TabPicture(1)   =   "frmCFI_Saida.frx":21F65
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Lista"
      Tab(1).Control(1)=   "PBLista"
      Tab(1).Control(2)=   "USToolBar2"
      Tab(1).Control(3)=   "USImageList2"
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados requisitante"
         Height          =   1365
         Left            =   60
         TabIndex        =   31
         Top             =   2790
         Width           =   9315
         Begin VB.ComboBox cmbCentrodecusto 
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
            Left            =   4500
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   38
            ToolTipText     =   "Funcionário."
            Top             =   375
            Width           =   4695
         End
         Begin VB.ComboBox cmbFuncionario 
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
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   34
            ToolTipText     =   "Funcionário."
            Top             =   375
            Width           =   4395
         End
         Begin VB.ComboBox cmbMaquina 
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
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   33
            ToolTipText     =   "Posto de trabalho."
            Top             =   900
            Width           =   1755
         End
         Begin VB.TextBox txtDescricao_maquina 
            Alignment       =   2  'Center
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
            Left            =   1890
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Descrição do posto de trabalho."
            Top             =   915
            Width           =   7275
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Index           =   15
            Left            =   6270
            TabIndex        =   39
            Top             =   180
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Funcionário"
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
            Left            =   1905
            TabIndex        =   37
            Top             =   180
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Posto de trabalho"
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
            Left            =   390
            TabIndex        =   36
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição posto de trabalho"
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
            Left            =   4860
            TabIndex        =   35
            Top             =   720
            Width           =   2010
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados do item saida almoxarifado"
         Height          =   1485
         Left            =   60
         TabIndex        =   23
         Top             =   1290
         Width           =   9315
         Begin VB.TextBox Txt_cod_ref 
            Alignment       =   1  'Right Justify
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
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   40
            TabStop         =   0   'False
            ToolTipText     =   "Código de referência."
            Top             =   465
            Width           =   1665
         End
         Begin VB.TextBox txtfamilia 
            Alignment       =   2  'Center
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
            Left            =   3210
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Família."
            Top             =   465
            Width           =   5985
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   2  'Center
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
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   465
            Width           =   1395
         End
         Begin VB.TextBox txtdescricao 
            Alignment       =   2  'Center
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
            Left            =   150
            Locked          =   -1  'True
            MaxLength       =   255
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Descrição."
            Top             =   1020
            Width           =   9015
         End
         Begin VB.TextBox txtid 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   270
            TabIndex        =   30
            Text            =   "0"
            Top             =   1560
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. de referência"
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
            Left            =   1740
            TabIndex        =   41
            Top             =   270
            Width           =   1350
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Index           =   4
            Left            =   5955
            TabIndex        =   29
            Top             =   270
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            Index           =   3
            Left            =   4335
            TabIndex        =   27
            Top             =   810
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. interno"
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
            Left            =   450
            TabIndex        =   26
            Top             =   270
            Width           =   900
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Informações retirada almoxarifado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2025
         Left            =   75
         TabIndex        =   9
         Top             =   4140
         Width           =   9315
         Begin DrawSuite2022.USButton btnRetirar 
            Height          =   855
            Left            =   7980
            TabIndex        =   42
            Top             =   990
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   1508
            DibPicture      =   "frmCFI_Saida.frx":21F81
            Caption         =   "Retirar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            PicSize         =   3
            PicSizeH        =   32
            PicSizeW        =   32
            Theme           =   4
         End
         Begin VB.TextBox txtOrdem 
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
            MaxLength       =   50
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   "Código interno."
            Top             =   495
            Width           =   1065
         End
         Begin VB.TextBox txtEst_At 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
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
            Left            =   8010
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Estoque atualizado."
            Top             =   495
            Width           =   1065
         End
         Begin VB.TextBox txtquantretirada 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
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
            Left            =   6360
            TabIndex        =   4
            Text            =   "0,000"
            ToolTipText     =   "Quantidade retirada."
            Top             =   495
            Width           =   1635
         End
         Begin VB.TextBox txtEstoque 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            Text            =   "0,000"
            ToolTipText     =   "Quantidade em estoque."
            Top             =   495
            Width           =   1065
         End
         Begin VB.ComboBox Cmb_RE 
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
            ItemData        =   "frmCFI_Saida.frx":237D5
            Left            =   3960
            List            =   "frmCFI_Saida.frx":237D7
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Número da rastreabilidade de estoque."
            Top             =   495
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker mskprev 
            Height          =   315
            Left            =   2610
            TabIndex        =   1
            ToolTipText     =   "Previsão de devolução."
            Top             =   495
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   180092929
            CurrentDate     =   39057
         End
         Begin MSComCtl2.DTPicker txtDataretirada 
            Height          =   315
            Left            =   1260
            TabIndex        =   0
            ToolTipText     =   "Data de retirada."
            Top             =   495
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   180092929
            CurrentDate     =   39057
         End
         Begin VB.TextBox txtObservacao 
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
            TabIndex        =   6
            ToolTipText     =   "Observações."
            Top             =   1155
            Width           =   7755
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° Ordem"
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
            Left            =   360
            TabIndex        =   22
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Est. atualiz."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   8055
            TabIndex        =   17
            Top             =   300
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde. retirada"
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
            Left            =   6660
            TabIndex        =   16
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Est. real"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   5475
            TabIndex        =   15
            Top             =   300
            Width           =   675
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº RE"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4335
            TabIndex        =   14
            Top             =   300
            Width           =   450
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Index           =   5
            Left            =   3585
            TabIndex        =   13
            Top             =   930
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prev. devolução"
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
            Left            =   2685
            TabIndex        =   12
            Top             =   300
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   150
            TabIndex        =   11
            Top             =   3030
            Width           =   45
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dt. retirada"
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
            Left            =   1500
            TabIndex        =   10
            Top             =   300
            Width           =   840
         End
      End
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   -73530
         Top             =   2310
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCFI_Saida.frx":237D9
         Count           =   1
      End
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   7440
         Top             =   570
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCFI_Saida.frx":2649D
         Count           =   1
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   75
         TabIndex        =   18
         Top             =   330
         Width           =   9270
         _ExtentX        =   16351
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
         ButtonLeft2     =   45
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
         ButtonLeft3     =   49
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
         ButtonLeft4     =   87
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
         ButtonLeft5     =   115
         ButtonTop5      =   2
         ButtonWidth5    =   24
         ButtonHeight5   =   24
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   -74925
         TabIndex        =   19
         Top             =   330
         Width           =   9270
         _ExtentX        =   16351
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
         ButtonCaption1  =   "Novo"
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
         ButtonWidth1    =   33
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
         ButtonLeft2     =   37
         ButtonTop2      =   2
         ButtonWidth2    =   39
         ButtonHeight2   =   21
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
         ButtonLeft3     =   78
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
         ButtonLeft4     =   82
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
         ButtonLeft5     =   120
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
         ButtonLeft6     =   148
         ButtonTop6      =   2
         ButtonWidth6    =   24
         ButtonHeight6   =   24
      End
      Begin DrawSuite2022.USProgressBar PBLista 
         Height          =   255
         Left            =   -74925
         TabIndex        =   20
         Top             =   5100
         Width           =   9270
         _ExtentX        =   16351
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
      Begin MSComctlLib.ListView Lista 
         Height          =   3750
         Left            =   -74925
         TabIndex        =   7
         Top             =   1335
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   6615
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
         MousePointer    =   99
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
            Object.Width           =   512
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "Cod produto"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   8431
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Família"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Tag             =   "T"
            Text            =   "Un."
            Object.Width           =   1058
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCFI_Saida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Qtd_Real        As Double 'OK
Dim Qtd_Retirada    As Double 'OK

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
cmbFuncionario.ListIndex = -1
cmbMaquina.ListIndex = -1
txtDescricao_maquina = ""
Txt_cod_ref = ""
Cmb_RE.ListIndex = -1
txtDataretirada.Value = Date
txtCodigo.Text = ""
txtfamilia.Text = ""
txtdescricao.Text = ""
txtEstoque.Text = "0,0000"
txtQuantRetirada.Text = "0,0000"
mskprev.Value = Date
txtObservacao.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnRetirar_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente fazer essa retirada do almoxarifado?", vbYesNo, "CAPRIND v5.0") = vbYes Then
ProcRetirada
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_RE_Click()
On Error GoTo tratar_erro

If Cmb_RE = "" Then Exit Sub
quantidade = 0
Txt_cod_ref = ""
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select ref, Saldo from estoque_controle_Saldo_RE where IDestoque = " & Cmb_RE, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    Txt_cod_ref = IIf(IsNull(TBEstoque!Ref), "", TBEstoque!Ref)
    quantidade = IIf(IsNull(TBEstoque!Saldo), 0, TBEstoque!Saldo)
End If
TBEstoque.Close
txtEstoque = Format(quantidade, "###,##0.0000")
txtEst_At = Format(quantidade, "###,##0.0000")

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

Private Sub ProcNovo()
On Error GoTo tratar_erro

frmCFI_locprod.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbMaquina_Click()
On Error GoTo tratar_erro

If cmbMaquina = "" Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Descricao from Cadmaquinas where idmaquina = " & cmbMaquina.ItemData(cmbMaquina.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtDescricao_maquina = IIf(IsNull(TBAbrir!Descricao), "", TBAbrir!Descricao)
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case SSTab.Tab
    Case 0:
        Select Case KeyCode
            Case vbKeyEscape: Unload Me
            Case vbKeyF3: ProcRetirada
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyEscape: Unload Me
            Case vbKeyF4: ProcExcluir
            Case vbKeyInsert: ProcNovo
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 9270, 5, True
ProcCarregaToolBar2 Me, 9270, 6, True
ProcLimpaVariaveisPrincipais

If Qualidade_Almox = True Then Caption = "Qualidade - Almoxarifado - Retirar"
SSTab.Tab = 0
With frmCFI
    txtCodigo.Text = .txtCodinterno
    txtdescricao.Text = .txtdescricao
    txtfamilia.Text = .txtfamilia
    txtDataretirada.Value = Date
    mskprev.Value = Date
    txtEstoque.Text = .txtquantestoque.Text
    txtEst_At = txtEstoque
End With

Conexao.Execute "DELETE from CFI_itens WHERE ID_CFI = 0"
ProcCarregaComboFuncionario cmbFuncionario, "Situacao <> 'Afastado' and Situacao <> 'Demitido'", False
ProcCarregaComboPostoTrab cmbMaquina, "Bloqueado = 'False'", False, False

Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select IdEstoque from estoque_controle_Saldo_RE where Codigo = '" & txtCodigo.Text & "' and Saldo <> 0 group by  IdEstoque order by IdEstoque", Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    Do While TBEstoque.EOF = False
        Cmb_RE.AddItem TBEstoque!IDEstoque
        TBEstoque.MoveNext
    Loop
End If
TBEstoque.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView lst_NatOp, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Select Case SSTab.Tab
    Case 0: If txtDataretirada.Visible = True Then txtDataretirada.SetFocus
    Case 1:
        If Lista.Visible = True Then Lista.SetFocus
        ProcCarregaLista
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtquantretirada_change()
On Error GoTo tratar_erro

If txtQuantRetirada.Text <> "" Then
    VerifNumero = txtQuantRetirada.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQuantRetirada.Text = ""
        txtQuantRetirada.SetFocus
        Exit Sub
    End If
End If
Qtd_Real = txtEstoque
Qtd_Retirada = IIf(txtQuantRetirada = "", "0", txtQuantRetirada)
ValorTotal = Qtd_Real - Qtd_Retirada
txtEst_At = Format(ValorTotal, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQuantRetirada_GotFocus()
On Error GoTo tratar_erro

If txtQuantRetirada = "0,0000" Then txtQuantRetirada = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtquantretirada_LostFocus()
On Error GoTo tratar_erro

txtQuantRetirada.Text = Format(txtQuantRetirada.Text, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcRetirada()
On Error GoTo tratar_erro

Acao = "retirar"
If Cmb_RE.Text = "" Then
    NomeCampo = "o número de rastreabilidade do estoque"
    ProcVerificaAcao
    Cmb_RE.SetFocus
    Exit Sub
End If
QuantEmpenho = IIf(txtQuantRetirada = "", 0, txtQuantRetirada)
If QuantEmpenho <= 0 Then
    NomeCampo = "a quantidade retirada"
    ProcVerificaAcao
    txtQuantRetirada.SetFocus
    Exit Sub
End If
If cmbFuncionario.Text = "" Then
    NomeCampo = "o funcionário"
    ProcVerificaAcao
    cmbFuncionario.SetFocus
    Exit Sub
End If

Qtd_Real = IIf(txtEstoque = "", 0, txtEstoque)
If QuantEmpenho > Qtd_Real Then
    USMsgBox ("A quantidade retirada não pode ser maior que a quantidade em estoque."), vbExclamation, "CAPRIND v5.0"
    txtQuantRetirada.SetFocus
    Exit Sub
End If

Cont = 0
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_controle where IDestoque = " & Cmb_RE, Conexao, adOpenKeyset, adLockOptimistic
If TBEstoque.EOF = False Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "SELECT * from CFI where idcfi = " & txtId.Text, Conexao, adOpenKeyset, adLockOptimistic
    TBProduto.AddNew
    TBProduto!Responsavel = pubUsuario
    TBProduto!status = "EM ABERTO"
    TBProduto!Codigo_produto = txtCodigo
    TBProduto!Descricao = txtdescricao
    TBProduto!Familia = txtfamilia
    TBProduto!Funcionario = cmbFuncionario
    If cmbMaquina <> "" Then TBProduto!ID_Maquina = cmbMaquina.ItemData(cmbMaquina.ListIndex)
    TBProduto!IDEstoque = TBEstoque!IDEstoque
    TBProduto!LOTE = TBEstoque!LOTE
    TBProduto!Dataretirada = txtDataretirada.Value
    TBProduto!dataprevisao = mskprev.Value
    TBProduto!Observacao = txtObservacao
    TBProduto!Quantretirada = QuantEmpenho
    
    TBProduto.Update
    Cont = TBProduto!IDCFI
    Conexao.Execute "Update CFI_itens Set ID_CFI = " & Cont & " where ID_CFI = 0"
    TBProduto.Close
    TBEstoque!estoque_real = Format(TBEstoque!estoque_real - QuantEmpenho, "###,##0.0000")
    TBEstoque!estoque_real_PC = TBEstoque!estoque_real
    TBEstoque!estoque_venda = Format(TBEstoque!estoque_venda - QuantEmpenho, "###,##0.0000")
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "select * from Estoque_movimentacao where id_cfi = " & Cont, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = True Then
        TBGravar.AddNew
        TBGravar!Destino = "Interno"
        TBGravar!Terceiros = False
        TBGravar!Id_cfi = Cont
        TBGravar!IDEstoque = TBEstoque!IDEstoque
        TBGravar!Operacao = "SAIDA_ALMOXARIFADO"
        TBGravar!Desenho = TBEstoque!Desenho
        TBGravar!Descricao = TBEstoque!Descricao
        TBGravar!Data = txtDataretirada
        TBGravar!Saida = QuantEmpenho
        TBGravar!Saida_PC = QuantEmpenho
        TBGravar!Responsavel = pubUsuario
        TBGravar!LOTE = TBEstoque!LOTE
        TBGravar!Requisitante = cmbFuncionario
        TBGravar!VlrUnit = IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario)
        TBGravar!vlrTotal = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * QuantEmpenho, "###,##0.00")
        TBGravar!Familia = IIf(IsNull(TBEstoque!Classe), "", TBEstoque!Classe)
           Set TBFamilia = CreateObject("adodb.recordset")
           TBFamilia.Open "select * from ProjFamilia where Familia = '" & TBEstoque!Classe & "'", Conexao, adOpenKeyset, adLockOptimistic
           If TBFamilia.EOF = False Then
           TBGravar!Grupo = TBFamilia!Grupo
           End If
           TBFamilia.Close
    
        TBGravar!Destino = "Interno"
        TBGravar.Update
    End If
    TBGravar.Close
        
    'Atualiza valor do material no estoque
    'Estoque_controle
    quantestoque = IIf(IsNull(TBEstoque!estoque_real), 0, TBEstoque!estoque_real)
    TBEstoque!Valor_total = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * quantestoque, "###,##0.00")
    TBEstoque.Update
'==================================================================================================
' :Aqui começa a criar credito no centro de custo
'==================================================================================================

    'Centro de custo
'    Set TBItem = CreateObject("adodb.recordset")
'    TBItem.Open "Select * from projproduto where desenho = '" & txtcodigo & "'", Conexao, adOpenKeyset, adLockOptimistic
'    If TBItem.EOF = False Then
'        Codproduto = TBItem!Codproduto
'        IDAntigo = IIf(IsNull(TBItem!ID_PC), 0, TBItem!ID_PC)
'    End If
'    TBItem.Close
'
'    valor = TBProduto!VlrTotal
'    If Requisicao_materiais = True Then
'        ProcCriaCreditoCCProdutoItem
'        Set TBMateriaprima = CreateObject("adodb.recordset")
'        TBMateriaprima.Open "Select * from Requisicao_materiais_lista where idlista = " & Listamaterial.SelectedItem & " and ID_CC is not null", Conexao, adOpenKeyset, adLockOptimistic
'        If TBMateriaprima.EOF = False Then
'            If TBMateriaprima!ID_CC <> "" Then
'                Set TBFI = CreateObject("adodb.recordset")
'                TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
'                TBFI.AddNew
'                ProcEnviaDadosCCRealizado TBMateriaprima!ID_CC
'                TBFI.Update
'
'                'Grava movimentação no centro consolidado
'                Set TBAfericao = CreateObject("adodb.recordset")
'                TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBMateriaprima!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
'                If TBAfericao.EOF = False Then
'                    Do While TBAfericao.EOF = False
'                        Set TBFI = CreateObject("adodb.recordset")
'                        TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
'                        TBFI.AddNew
'                        ProcEnviaDadosCCRealizado TBAfericao!ID_CC
'                        TBFI.Update
'
'                        Set TBCiclo = CreateObject("adodb.recordset")
'                        TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
'                        If TBCiclo.EOF = False Then
'                            Do While TBCiclo.EOF = False
'                                Set TBFI = CreateObject("adodb.recordset")
'                                TBFI.Open "Select * from CC_realizado", Conexao, adOpenKeyset, adLockOptimistic
'                                TBFI.AddNew
'                                ProcEnviaDadosCCRealizado TBCiclo!ID_CC
'                                TBFI.Update
'                                TBCiclo.MoveNext
'                            Loop
'                        End If
'                        TBCiclo.Close
'
'                        TBAfericao.MoveNext
'                    Loop
'                End If
'                TBAfericao.Close
'            End If
'        End If
'        TBMateriaprima.Close
'    Else
'        ProcCriaCreditoCCProdutoItem
'    End If
'    TBProduto.Close
'End If
'
'If Chk_expedir.Value = 0 And IsNumeric(Txt_RM) = True Then ProcAtualizaCTMaterialOrdem Cmb_empresa.ItemData(Cmb_empresa.ListIndex), Txt_RM
'==========================================================================================================================================
    
    
    USMsgBox ("Produto retirado do estoque com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = Formulario
    Evento = "Retirar"
    ID_documento = Cont
    Documento = "Cód. interno: " & txtCodigo.Text & " - RE: " & Cmb_RE.Text & " - Lote: " & TBEstoque!LOTE
    Documento1 = ""
    ProcGravaEvento
    '==================================
    
    TBEstoque.Close
    
    ProcLimpaCampos
    With frmCFI
        .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
        .ProcLimpaCampos
        .Frame2.Enabled = False
    End With
    Unload Me
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from CFI_Itens where ID_CFI = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Codproduto), "", TBLISTA!Codproduto)
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "select * from Projproduto where codproduto = " & TBLISTA!Codproduto, Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                .Item(.Count).SubItems(2) = IIf(IsNull(TBItem!Desenho), "", TBItem!Desenho)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBItem!Descricao), "", Trim(TBItem!Descricao))
                .Item(.Count).SubItems(4) = IIf(IsNull(TBItem!Classe), "", Trim(TBItem!Classe))
                .Item(.Count).SubItems(5) = IIf(IsNull(TBItem!Unidade), "", TBItem!Unidade)
            End If
            TBItem.Close
            TBLISTA.MoveNext
            Contador = Contador + 1
            PBLista.Value = Contador
        End With
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcRetirada
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
    Case 1: ProcNovo
    Case 2: ProcExcluir
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcExcluir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) produto(s)(ns)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from CFI_itens WHERE id = " & .ListItems.Item(InitFor)
            '==================================
            Modulo = Formulario
            Evento = "Excluir produto de destino/aplicação"
            ID_documento = Lista.SelectedItem
            Documento = "Cód. interno: " & .ListItems.Item(InitFor).SubItems(2)
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) produto(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Produto(s) excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
