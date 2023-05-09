VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCertificado_qualidade 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Qualidade - Ensaios - Certificado da qualidade"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   15360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin TabDlg.SSTab SSTab1 
      Height          =   4785
      Left            =   0
      TabIndex        =   67
      Top             =   5295
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   8440
      _Version        =   393216
      Tab             =   2
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
      TabCaption(0)   =   "Ultra-som"
      TabPicture(0)   =   "frmCertificado_qualidade.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lista_ultra"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Líquido penetrante"
      TabPicture(1)   =   "frmCertificado_qualidade.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lista_liquido"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Análise química"
      TabPicture(2)   =   "frmCertificado_qualidade.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "USToolBar2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lista_carcaca"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin MSComctlLib.ListView lista_carcaca 
         Height          =   2025
         Left            =   180
         TabIndex        =   53
         Top             =   2520
         Width           =   14910
         _ExtentX        =   26300
         _ExtentY        =   3572
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
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
            Object.Tag             =   "T"
            Text            =   "Material"
            Object.Width           =   15090
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "Certificado"
            Object.Width           =   10523
         EndProperty
      End
      Begin MSComctlLib.ListView lista_liquido 
         Height          =   4005
         Left            =   -74760
         TabIndex        =   27
         Top             =   525
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   7064
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
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
            Text            =   "Número"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   3607
         EndProperty
      End
      Begin MSComctlLib.ListView lista_ultra 
         Height          =   2685
         Left            =   -74760
         TabIndex        =   26
         Top             =   1860
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   4736
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
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
            Text            =   "Número"
            Object.Width           =   12792
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "D"
            Text            =   "Data"
            Object.Width           =   12792
         EndProperty
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   4365
         Left            =   -74925
         TabIndex        =   154
         Top             =   330
         Width           =   15200
         Begin VB.TextBox txtData_liquido 
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
            Left            =   3675
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Metal base."
            Top             =   390
            Width           =   1275
         End
         Begin VB.TextBox txtQTDE_liquido 
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
            Left            =   4965
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Metal base."
            Top             =   390
            Width           =   1095
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Conclusão"
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
            Height          =   735
            Left            =   3660
            TabIndex        =   155
            Top             =   810
            Width           =   2415
            Begin VB.OptionButton optReprovado_liquido 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Reprovado"
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
               Height          =   255
               Left            =   1260
               TabIndex        =   31
               Top             =   330
               Width           =   1095
            End
            Begin VB.OptionButton optAprovado_liquido 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Aprovado"
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
               Height          =   255
               Left            =   180
               TabIndex        =   30
               Top             =   330
               Width           =   1035
            End
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Index           =   20
            Left            =   4140
            TabIndex        =   157
            Top             =   180
            Width           =   345
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde"
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
            Left            =   5332
            TabIndex        =   156
            Top             =   180
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   4365
         Left            =   -74925
         TabIndex        =   148
         Top             =   330
         Width           =   15200
         Begin VB.TextBox txtEspess_ultra 
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
            MaxLength       =   50
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Espessura do revest."
            Top             =   490
            Width           =   5415
         End
         Begin VB.TextBox txtTransdutor 
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
            Left            =   10145
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Metal base."
            Top             =   1090
            Width           =   4905
         End
         Begin VB.Frame frameConclusao 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Conclusão"
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
            Height          =   645
            Left            =   10140
            TabIndex        =   153
            Top             =   160
            Width           =   4905
            Begin VB.OptionButton optAprovado 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Aprovado"
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
               Height          =   255
               Left            =   480
               TabIndex        =   21
               Top             =   300
               Width           =   1035
            End
            Begin VB.OptionButton optReprovado 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Reprovado"
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
               Height          =   255
               Left            =   1620
               TabIndex        =   22
               Top             =   300
               Width           =   1095
            End
         End
         Begin VB.TextBox txtAparelho 
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
            MaxLength       =   100
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Metal base."
            Top             =   1090
            Width           =   4990
         End
         Begin VB.TextBox txtSuperf_ultra 
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
            MaxLength       =   50
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "Superfície do revest."
            Top             =   1090
            Width           =   4990
         End
         Begin VB.TextBox txtQtde_ultra 
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
            Left            =   5595
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Quantidade"
            Top             =   490
            Width           =   3075
         End
         Begin VB.TextBox txtData_ultra 
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
            Left            =   8685
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Data."
            Top             =   490
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Espessura do revest."
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
            Left            =   2122
            TabIndex        =   175
            Top             =   280
            Width           =   1530
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Transdutor"
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
            Left            =   12200
            TabIndex        =   170
            Top             =   880
            Width           =   795
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Aparelho"
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
            Left            =   7333
            TabIndex        =   152
            Top             =   880
            Width           =   645
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Superfície do revest."
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
            Left            =   1925
            TabIndex        =   151
            Top             =   880
            Width           =   1500
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde"
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
            Left            =   6952
            TabIndex        =   150
            Top             =   285
            Width           =   360
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Left            =   9180
            TabIndex        =   149
            Top             =   280
            Width           =   345
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   885
         Left            =   -74925
         TabIndex        =   141
         Top             =   330
         Width           =   11745
         Begin VB.CommandButton cmdContasPorcentagem 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   1440
            MouseIcon       =   "frmCertificado_qualidade.frx":0054
            MousePointer    =   99  'Custom
            Picture         =   "frmCertificado_qualidade.frx":01A6
            Style           =   1  'Graphical
            TabIndex        =   147
            ToolTipText     =   "Enviar para o financeiro em porcentagem (F8)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdContas 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   810
            MouseIcon       =   "frmCertificado_qualidade.frx":03A5
            MousePointer    =   99  'Custom
            Picture         =   "frmCertificado_qualidade.frx":04F7
            Style           =   1  'Graphical
            TabIndex        =   146
            ToolTipText     =   "Enviar para o financeiro (F7)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdProximo_comercial 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2715
            MouseIcon       =   "frmCertificado_qualidade.frx":070F
            MousePointer    =   99  'Custom
            Picture         =   "frmCertificado_qualidade.frx":0861
            Style           =   1  'Graphical
            TabIndex        =   145
            ToolTipText     =   "Próximo pedido."
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdAnt_comercial 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2070
            MouseIcon       =   "frmCertificado_qualidade.frx":0BAA
            MousePointer    =   99  'Custom
            Picture         =   "frmCertificado_qualidade.frx":0CFC
            Style           =   1  'Graphical
            TabIndex        =   144
            ToolTipText     =   "Pedido anterior."
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdImpProposta_Comercial 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3345
            MouseIcon       =   "frmCertificado_qualidade.frx":1045
            MousePointer    =   99  'Custom
            Picture         =   "frmCertificado_qualidade.frx":1197
            Style           =   1  'Graphical
            TabIndex        =   143
            ToolTipText     =   "Visualizar impressão (F5)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton imgGravar 
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
            Height          =   615
            Left            =   180
            MouseIcon       =   "frmCertificado_qualidade.frx":1986
            MousePointer    =   99  'Custom
            Picture         =   "frmCertificado_qualidade.frx":1AD8
            Style           =   1  'Graphical
            TabIndex        =   142
            ToolTipText     =   "Salvar (F3)"
            Top             =   180
            Width           =   630
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         Height          =   885
         Left            =   -74925
         TabIndex        =   132
         Top             =   330
         Width           =   11745
         Begin VB.CommandButton cmdCancelarFaturamento 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   2730
            MouseIcon       =   "frmCertificado_qualidade.frx":22B1
            MousePointer    =   99  'Custom
            Picture         =   "frmCertificado_qualidade.frx":2403
            Style           =   1  'Graphical
            TabIndex        =   140
            ToolTipText     =   "Cancelar liberação para faturamento do produto/item (F8)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdProximo_Detalhe 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3990
            MouseIcon       =   "frmCertificado_qualidade.frx":2996
            MousePointer    =   99  'Custom
            Picture         =   "frmCertificado_qualidade.frx":2AE8
            Style           =   1  'Graphical
            TabIndex        =   139
            ToolTipText     =   "Próximo pedido."
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmd_Anterior_Detalhe 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3360
            MouseIcon       =   "frmCertificado_qualidade.frx":2E31
            MousePointer    =   99  'Custom
            Picture         =   "frmCertificado_qualidade.frx":2F83
            Style           =   1  'Graphical
            TabIndex        =   138
            ToolTipText     =   "Pedido anterior."
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdImp_Proposta_Detalhe 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4635
            MouseIcon       =   "frmCertificado_qualidade.frx":32CC
            MousePointer    =   99  'Custom
            Picture         =   "frmCertificado_qualidade.frx":341E
            Style           =   1  'Graphical
            TabIndex        =   137
            ToolTipText     =   "Visualizar impressão (F5)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdgravarlista 
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
            Height          =   615
            Left            =   810
            MouseIcon       =   "frmCertificado_qualidade.frx":3C0D
            MousePointer    =   99  'Custom
            Picture         =   "frmCertificado_qualidade.frx":3D5F
            Style           =   1  'Graphical
            TabIndex        =   136
            ToolTipText     =   "Salvar (F3)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdexcluirlista 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   1455
            MouseIcon       =   "frmCertificado_qualidade.frx":4538
            MousePointer    =   99  'Custom
            Picture         =   "frmCertificado_qualidade.frx":468A
            Style           =   1  'Graphical
            TabIndex        =   135
            ToolTipText     =   "Excluir (F4)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdAgregar 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   180
            MouseIcon       =   "frmCertificado_qualidade.frx":4ED9
            MousePointer    =   99  'Custom
            Picture         =   "frmCertificado_qualidade.frx":502B
            Style           =   1  'Graphical
            TabIndex        =   134
            ToolTipText     =   "Novo (Insert)"
            Top             =   180
            Width           =   630
         End
         Begin VB.CommandButton cmdfaturamento 
            BackColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   2085
            MouseIcon       =   "frmCertificado_qualidade.frx":5551
            MousePointer    =   99  'Custom
            Picture         =   "frmCertificado_qualidade.frx":56A3
            Style           =   1  'Graphical
            TabIndex        =   133
            ToolTipText     =   "Liberar produto/item para faturamento (F7)"
            Top             =   180
            Width           =   630
         End
      End
      Begin VB.Frame Frame9 
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
         ForeColor       =   &H8000000E&
         Height          =   5445
         Left            =   -74925
         TabIndex        =   91
         Top             =   1200
         Width           =   11745
         Begin VB.TextBox txtID_produto 
            Height          =   285
            Left            =   1980
            TabIndex        =   128
            Top             =   4140
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.TextBox txt_VlrICMS 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   615
            MaxLength       =   50
            TabIndex        =   127
            ToolTipText     =   "Valor do ICMS"
            Top             =   6480
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txt_BaseICMS 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   -1560
            MaxLength       =   50
            TabIndex        =   126
            ToolTipText     =   "Base de cálculo do ICMS"
            Top             =   6480
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   3465
            Left            =   30
            TabIndex        =   92
            Top             =   150
            Width           =   11685
            Begin VB.ComboBox cmbreferencia 
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
               Left            =   3270
               MouseIcon       =   "frmCertificado_qualidade.frx":5F1D
               MousePointer    =   99  'Custom
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   116
               ToolTipText     =   "Código de referencia."
               Top             =   285
               Width           =   2475
            End
            Begin VB.CommandButton cmdfiltrar 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2490
               MouseIcon       =   "frmCertificado_qualidade.frx":6227
               MousePointer    =   99  'Custom
               Picture         =   "frmCertificado_qualidade.frx":6379
               Style           =   1  'Graphical
               TabIndex        =   115
               ToolTipText     =   "Filtrar por código interno."
               Top             =   285
               Width           =   315
            End
            Begin VB.TextBox txtRev_cod 
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
               Left            =   1950
               Locked          =   -1  'True
               MaxLength       =   50
               MouseIcon       =   "frmCertificado_qualidade.frx":6794
               MousePointer    =   99  'Custom
               TabIndex        =   114
               TabStop         =   0   'False
               Text            =   "0"
               ToolTipText     =   "Revisão do produto/item."
               Top             =   285
               Width           =   525
            End
            Begin VB.CheckBox OPTnovoman 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Criar novo produto (cód. manual) ?"
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   6000
               TabIndex        =   113
               Top             =   330
               Width           =   2835
            End
            Begin VB.CheckBox OPTnovo 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Criar novo produto (cód. automático) ?"
               ForeColor       =   &H00000000&
               Height          =   405
               Left            =   6000
               TabIndex        =   112
               Top             =   0
               Width           =   3015
            End
            Begin VB.TextBox txtvalor_total 
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
               Height          =   315
               Left            =   9990
               Locked          =   -1  'True
               MaxLength       =   50
               MouseIcon       =   "frmCertificado_qualidade.frx":6A9E
               MousePointer    =   99  'Custom
               TabIndex        =   111
               TabStop         =   0   'False
               ToolTipText     =   "Valor total."
               Top             =   3030
               Width           =   1545
            End
            Begin VB.TextBox txtdbl_valoripi 
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
               Height          =   315
               Left            =   6540
               Locked          =   -1  'True
               MaxLength       =   50
               MouseIcon       =   "frmCertificado_qualidade.frx":6BF0
               MousePointer    =   99  'Custom
               TabIndex        =   110
               TabStop         =   0   'False
               ToolTipText     =   "Valor do IPI."
               Top             =   3030
               Width           =   1455
            End
            Begin VB.TextBox txtvalorunitariodesc 
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
               Height          =   315
               Left            =   4620
               Locked          =   -1  'True
               MaxLength       =   50
               MouseIcon       =   "frmCertificado_qualidade.frx":6EFA
               MousePointer    =   99  'Custom
               TabIndex        =   109
               TabStop         =   0   'False
               ToolTipText     =   "Valor unitário com desconto."
               Top             =   3030
               Width           =   1425
            End
            Begin VB.TextBox txtvalordesconto 
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
               Height          =   315
               Left            =   3360
               Locked          =   -1  'True
               MaxLength       =   50
               MouseIcon       =   "frmCertificado_qualidade.frx":7204
               MousePointer    =   99  'Custom
               TabIndex        =   108
               TabStop         =   0   'False
               ToolTipText     =   "Valor do desconto."
               Top             =   3030
               Width           =   1245
            End
            Begin VB.TextBox txtdesconto 
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
               Height          =   315
               Left            =   2445
               MaxLength       =   12
               MousePointer    =   1  'Arrow
               TabIndex        =   107
               Text            =   "0"
               ToolTipText     =   "Valor do desconto (%)."
               Top             =   3030
               Width           =   915
            End
            Begin VB.CommandButton cmdlistaproduto 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2820
               MouseIcon       =   "frmCertificado_qualidade.frx":750E
               MousePointer    =   99  'Custom
               Picture         =   "frmCertificado_qualidade.frx":7660
               Style           =   1  'Graphical
               TabIndex        =   106
               ToolTipText     =   "Localizar produto/item."
               Top             =   285
               Width           =   315
            End
            Begin VB.ComboBox cmbfamilia 
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
               Left            =   2460
               MouseIcon       =   "frmCertificado_qualidade.frx":7762
               MousePointer    =   99  'Custom
               Style           =   2  'Dropdown List
               TabIndex        =   105
               ToolTipText     =   "Familia."
               Top             =   2400
               Width           =   9075
            End
            Begin VB.ComboBox cmbun 
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
               Left            =   1560
               MouseIcon       =   "frmCertificado_qualidade.frx":7A6C
               MousePointer    =   99  'Custom
               Style           =   2  'Dropdown List
               TabIndex        =   104
               ToolTipText     =   "Unidade."
               Top             =   2400
               Width           =   885
            End
            Begin VB.TextBox txtvalorunitario 
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
               Height          =   315
               Left            =   1200
               MaxLength       =   12
               MouseIcon       =   "frmCertificado_qualidade.frx":7D76
               MousePointer    =   99  'Custom
               TabIndex        =   103
               ToolTipText     =   "Valor unitário."
               Top             =   3030
               Width           =   1245
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
               Height          =   315
               Left            =   180
               MaxLength       =   50
               MouseIcon       =   "frmCertificado_qualidade.frx":8080
               MousePointer    =   99  'Custom
               TabIndex        =   102
               ToolTipText     =   "Código interno do produto/item."
               Top             =   285
               Width           =   1785
            End
            Begin VB.TextBox txtQuantidade 
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
               Height          =   315
               Left            =   180
               MaxLength       =   12
               MouseIcon       =   "frmCertificado_qualidade.frx":838A
               MousePointer    =   99  'Custom
               TabIndex        =   101
               ToolTipText     =   "Quantidade."
               Top             =   3030
               Width           =   1005
            End
            Begin VB.ComboBox cmbfiscal 
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
               ItemData        =   "frmCertificado_qualidade.frx":8694
               Left            =   180
               List            =   "frmCertificado_qualidade.frx":8696
               MouseIcon       =   "frmCertificado_qualidade.frx":8698
               MousePointer    =   99  'Custom
               Style           =   2  'Dropdown List
               TabIndex        =   100
               ToolTipText     =   "Classificação fiscal."
               Top             =   2400
               Width           =   1065
            End
            Begin VB.TextBox txtint_icms 
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
               Height          =   315
               Left            =   8010
               Locked          =   -1  'True
               MaxLength       =   50
               MouseIcon       =   "frmCertificado_qualidade.frx":89A2
               MousePointer    =   99  'Custom
               TabIndex        =   99
               TabStop         =   0   'False
               ToolTipText     =   "Porcentagem do ICMS."
               Top             =   3030
               Width           =   465
            End
            Begin VB.TextBox txtInt_ipi 
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
               Height          =   315
               Left            =   6060
               Locked          =   -1  'True
               MaxLength       =   50
               MouseIcon       =   "frmCertificado_qualidade.frx":8CAC
               MousePointer    =   99  'Custom
               TabIndex        =   98
               TabStop         =   0   'False
               ToolTipText     =   "Porcentagem do IPI."
               Top             =   3030
               Width           =   465
            End
            Begin VB.TextBox txtvalor_icms 
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
               Height          =   315
               Left            =   8490
               Locked          =   -1  'True
               MaxLength       =   50
               MouseIcon       =   "frmCertificado_qualidade.frx":8FB6
               MousePointer    =   99  'Custom
               TabIndex        =   97
               TabStop         =   0   'False
               ToolTipText     =   "Valor do ICMS."
               Top             =   3030
               Width           =   1485
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
               Height          =   585
               Left            =   180
               MaxLength       =   5000
               MouseIcon       =   "frmCertificado_qualidade.frx":92C0
               MousePointer    =   99  'Custom
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   96
               ToolTipText     =   "Descrição comercial do produto/item."
               Top             =   1515
               Width           =   11355
            End
            Begin VB.TextBox txtdesctecnica 
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
               Left            =   180
               MaxLength       =   255
               MouseIcon       =   "frmCertificado_qualidade.frx":95CA
               MousePointer    =   99  'Custom
               TabIndex        =   95
               ToolTipText     =   "Descricao técnica do item."
               Top             =   900
               Width           =   7725
            End
            Begin VB.CommandButton cmdCF 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   1230
               MouseIcon       =   "frmCertificado_qualidade.frx":98D4
               MousePointer    =   99  'Custom
               Picture         =   "frmCertificado_qualidade.frx":9A26
               Style           =   1  'Graphical
               TabIndex        =   94
               ToolTipText     =   "Abrir módulo para consulta de classificação fiscal."
               Top             =   2400
               Width           =   315
            End
            Begin VB.TextBox txtpccliente 
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
               Left            =   7920
               MaxLength       =   255
               MouseIcon       =   "frmCertificado_qualidade.frx":9B28
               MousePointer    =   99  'Custom
               TabIndex        =   93
               ToolTipText     =   "Pedido do cliente."
               Top             =   900
               Width           =   2145
            End
            Begin MSMask.MaskEdBox mskprazo 
               Height          =   315
               Left            =   10080
               TabIndex        =   117
               ToolTipText     =   "Prazo de entrega."
               Top             =   900
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rev.                                Código referencia"
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
               Left            =   2040
               TabIndex        =   125
               Top             =   75
               Width           =   3060
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000A&
               BackStyle       =   0  'Transparent
               Caption         =   $"frmCertificado_qualidade.frx":9E32
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
               Index           =   2
               Left            =   300
               TabIndex        =   124
               Top             =   2850
               Width           =   10800
            End
            Begin VB.Image imgCalendario 
               Height          =   360
               Left            =   11175
               MouseIcon       =   "frmCertificado_qualidade.frx":9EF6
               MousePointer    =   99  'Custom
               Picture         =   "frmCertificado_qualidade.frx":A048
               Stretch         =   -1  'True
               ToolTipText     =   "Abrir calendário."
               Top             =   870
               Width           =   330
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   $"frmCertificado_qualidade.frx":A4CB
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
               Left            =   480
               TabIndex        =   123
               Top             =   2220
               Width           =   6705
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição técnica do produto/item"
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
               Left            =   420
               TabIndex        =   122
               Top             =   690
               Width           =   2445
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Descrição comercial do produto/item"
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
               Left            =   420
               TabIndex        =   121
               Top             =   1290
               Width           =   2595
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Prazo"
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
               Left            =   10387
               TabIndex        =   120
               Top             =   690
               Width           =   480
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Pedido do cliente"
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
               Left            =   8385
               TabIndex        =   119
               Top             =   690
               Width           =   1215
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Caption         =   "Código interno"
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
               Left            =   420
               TabIndex        =   118
               Top             =   75
               Width           =   1230
            End
         End
         Begin MSComctlLib.ListView Listprod 
            Height          =   1815
            Left            =   180
            TabIndex        =   129
            Top             =   3600
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   3201
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
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
            NumItems        =   12
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Seq."
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Cód. interno"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Descrição"
               Object.Width           =   7938
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Qtde."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Vlr. unitário"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Desc. (%)"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Vlr. desconto"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Vlr. unit. c/ Desc."
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Text            =   "Vlr. total"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Status"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   10
               Text            =   "Prazo final"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Pedido do cliente"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label69 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Valor do ICMS"
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
            Left            =   795
            TabIndex        =   131
            Top             =   6240
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label71 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Base de Calculo do ICMS"
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
            Left            =   -1335
            TabIndex        =   130
            Top             =   6240
            Visible         =   0   'False
            Width           =   1770
         End
      End
      Begin VB.Frame Frame12 
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
         ForeColor       =   &H8000000E&
         Height          =   5445
         Left            =   -74925
         TabIndex        =   68
         Top             =   1200
         Width           =   11745
         Begin VB.TextBox txtID_servico 
            Height          =   285
            Left            =   2250
            TabIndex        =   90
            Top             =   4500
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   3465
            Left            =   30
            TabIndex        =   69
            Top             =   150
            Width           =   11685
            Begin VB.ComboBox cmbreferencia_serv 
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
               Left            =   3270
               MouseIcon       =   "frmCertificado_qualidade.frx":A55C
               MousePointer    =   99  'Custom
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   89
               ToolTipText     =   "Código de referencia."
               Top             =   285
               Width           =   2475
            End
            Begin VB.CommandButton cmdfiltrar_serv 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2490
               MouseIcon       =   "frmCertificado_qualidade.frx":A866
               MousePointer    =   99  'Custom
               Picture         =   "frmCertificado_qualidade.frx":A9B8
               Style           =   1  'Graphical
               TabIndex        =   88
               ToolTipText     =   "Filtrar por código interno."
               Top             =   285
               Width           =   315
            End
            Begin VB.ComboBox txtunservico 
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
               Left            =   180
               MouseIcon       =   "frmCertificado_qualidade.frx":ADD3
               MousePointer    =   99  'Custom
               Style           =   2  'Dropdown List
               TabIndex        =   87
               ToolTipText     =   "Unidade."
               Top             =   3030
               Width           =   735
            End
            Begin VB.TextBox txtqtservico 
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
               Height          =   315
               Left            =   900
               MaxLength       =   50
               MouseIcon       =   "frmCertificado_qualidade.frx":B0DD
               MousePointer    =   99  'Custom
               TabIndex        =   86
               ToolTipText     =   "Quantidade."
               Top             =   3030
               Width           =   885
            End
            Begin VB.TextBox txtRev_serv 
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
               Left            =   1950
               Locked          =   -1  'True
               MaxLength       =   50
               MouseIcon       =   "frmCertificado_qualidade.frx":B3E7
               MousePointer    =   99  'Custom
               TabIndex        =   85
               TabStop         =   0   'False
               Text            =   "0"
               ToolTipText     =   "Revisão do serviço."
               Top             =   285
               Width           =   525
            End
            Begin VB.CheckBox optnovoservico 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Criar novo serviço (cód. automático) ?"
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   6000
               TabIndex        =   84
               Top             =   30
               Width           =   3045
            End
            Begin VB.CheckBox OPTnovoservicoman 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Criar novo serviço (cód. manual) ?"
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   6000
               TabIndex        =   83
               Top             =   330
               Width           =   2835
            End
            Begin VB.ComboBox cmbfamiliaservico 
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
               Left            =   180
               MouseIcon       =   "frmCertificado_qualidade.frx":B6F1
               MousePointer    =   99  'Custom
               Style           =   2  'Dropdown List
               TabIndex        =   82
               ToolTipText     =   "Familia."
               Top             =   2430
               Width           =   11355
            End
            Begin VB.TextBox txtdesccomservico 
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
               Height          =   585
               Left            =   180
               MaxLength       =   5000
               MouseIcon       =   "frmCertificado_qualidade.frx":B9FB
               MousePointer    =   99  'Custom
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   81
               ToolTipText     =   "Descrição comercial do serviço."
               Top             =   1515
               Width           =   11355
            End
            Begin VB.TextBox txtiss 
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
               Height          =   315
               Left            =   7320
               TabIndex        =   80
               ToolTipText     =   "Porcentagem do ISS."
               Top             =   3030
               Width           =   945
            End
            Begin VB.TextBox txtvlrISS 
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
               Height          =   315
               Left            =   8280
               Locked          =   -1  'True
               TabIndex        =   79
               TabStop         =   0   'False
               ToolTipText     =   "Valor do ISS."
               Top             =   3030
               Width           =   1680
            End
            Begin VB.TextBox txtvalorunitariodesc2 
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
               Height          =   315
               Left            =   5685
               Locked          =   -1  'True
               MaxLength       =   50
               MouseIcon       =   "frmCertificado_qualidade.frx":BD05
               MousePointer    =   99  'Custom
               TabIndex        =   78
               TabStop         =   0   'False
               ToolTipText     =   "Valor unitário com desconto."
               Top             =   3030
               Width           =   1620
            End
            Begin VB.TextBox txtvalordesconto2 
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
               Height          =   315
               Left            =   4170
               Locked          =   -1  'True
               MaxLength       =   50
               MouseIcon       =   "frmCertificado_qualidade.frx":C00F
               MousePointer    =   99  'Custom
               TabIndex        =   77
               TabStop         =   0   'False
               ToolTipText     =   "Valor do desconto."
               Top             =   3030
               Width           =   1500
            End
            Begin VB.TextBox txtdesconto2 
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
               Height          =   315
               Left            =   3180
               MaxLength       =   50
               MousePointer    =   1  'Arrow
               TabIndex        =   76
               Text            =   "0"
               ToolTipText     =   "Valor do desconto (%)."
               Top             =   3030
               Width           =   975
            End
            Begin VB.TextBox txtcodservico 
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
               Height          =   315
               Left            =   180
               MaxLength       =   50
               MouseIcon       =   "frmCertificado_qualidade.frx":C319
               MousePointer    =   99  'Custom
               TabIndex        =   75
               ToolTipText     =   "Código interno do serviço."
               Top             =   285
               Width           =   1785
            End
            Begin VB.CommandButton cmdlistaservicos 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Height          =   315
               Left            =   2820
               MouseIcon       =   "frmCertificado_qualidade.frx":C623
               MousePointer    =   99  'Custom
               Picture         =   "frmCertificado_qualidade.frx":C775
               Style           =   1  'Graphical
               TabIndex        =   74
               ToolTipText     =   "Localizar seviços."
               Top             =   285
               Width           =   315
            End
            Begin VB.TextBox txtdescservico 
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
               Left            =   180
               MaxLength       =   255
               MouseIcon       =   "frmCertificado_qualidade.frx":C877
               MousePointer    =   99  'Custom
               TabIndex        =   73
               ToolTipText     =   "Descrição técnica do serviço."
               Top             =   900
               Width           =   7725
            End
            Begin VB.TextBox txtvlrunitservico 
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
               Height          =   315
               Left            =   1800
               MaxLength       =   50
               MouseIcon       =   "frmCertificado_qualidade.frx":CB81
               MousePointer    =   99  'Custom
               TabIndex        =   72
               ToolTipText     =   "Valor unitário."
               Top             =   3030
               Width           =   1365
            End
            Begin VB.TextBox txtvlrtotalservico 
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
               Height          =   315
               Left            =   9960
               Locked          =   -1  'True
               MaxLength       =   50
               MouseIcon       =   "frmCertificado_qualidade.frx":CE8B
               MousePointer    =   99  'Custom
               TabIndex        =   71
               TabStop         =   0   'False
               ToolTipText     =   "Valor total."
               Top             =   3030
               Width           =   1575
            End
            Begin VB.TextBox txtpcclienteserv 
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
               Left            =   7920
               MaxLength       =   255
               MouseIcon       =   "frmCertificado_qualidade.frx":CFDD
               MousePointer    =   99  'Custom
               TabIndex        =   70
               ToolTipText     =   "Pedido do cliente."
               Top             =   900
               Width           =   2145
            End
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   180
         TabIndex        =   177
         Top             =   510
         Width           =   15000
         _ExtentX        =   26458
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
         ButtonLeft2     =   37
         ButtonTop2      =   2
         ButtonWidth2    =   38
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Excluir"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Excluir (F4)"
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
         ButtonLeft3     =   77
         ButtonTop3      =   2
         ButtonWidth3    =   39
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonKey4      =   "4"
         ButtonAlignment4=   2
         ButtonStyle4    =   -1
         BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState4    =   5
         ButtonLeft4     =   118
         ButtonTop4      =   2
         ButtonWidth4    =   24
         ButtonHeight4   =   24
         ButtonUseMaskColor4=   0   'False
         Begin DrawSuite2022.USImageList USImageList2 
            Left            =   3060
            Top             =   60
            _ExtentX        =   900
            _ExtentY        =   767
            Img1            =   "frmCertificado_qualidade.frx":D2E7
            Count           =   1
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   4365
         Left            =   75
         TabIndex        =   158
         Top             =   330
         Width           =   15200
         Begin VB.Frame Frame9_carcaca 
            BackColor       =   &H00E0E0E0&
            Height          =   915
            Left            =   14010
            TabIndex        =   167
            Top             =   1170
            Width           =   1070
            Begin VB.TextBox txtTexto_carcaca9 
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
               Left            =   195
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   180
               Width           =   645
            End
            Begin VB.TextBox txtNumero_carcaca9 
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
               Left            =   90
               MaxLength       =   50
               TabIndex        =   52
               Top             =   510
               Width           =   855
            End
         End
         Begin VB.Frame Frame8_carcaca 
            BackColor       =   &H00E0E0E0&
            Height          =   915
            Left            =   12915
            TabIndex        =   166
            Top             =   1170
            Width           =   1070
            Begin VB.TextBox txtNumero_carcaca8 
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
               Left            =   90
               MaxLength       =   50
               TabIndex        =   50
               Top             =   510
               Width           =   855
            End
            Begin VB.TextBox txtTexto_carcaca8 
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
               Left            =   190
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   180
               Width           =   645
            End
         End
         Begin VB.Frame Frame2_carcaca 
            BackColor       =   &H00E0E0E0&
            Height          =   915
            Left            =   6330
            TabIndex        =   160
            Top             =   1170
            Width           =   1070
            Begin VB.TextBox txtTexto_carcaca2 
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
               Left            =   190
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   180
               Width           =   645
            End
            Begin VB.TextBox txtNumero_carcaca2 
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
               Left            =   90
               MaxLength       =   50
               TabIndex        =   39
               Top             =   510
               Width           =   855
            End
         End
         Begin VB.TextBox txtID_Elemento 
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
            Left            =   960
            MaxLength       =   50
            MouseIcon       =   "frmCertificado_qualidade.frx":F0B8
            MousePointer    =   99  'Custom
            TabIndex        =   173
            Text            =   "0"
            ToolTipText     =   "Metal base."
            Top             =   1770
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtid_carcaca 
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
            Left            =   240
            MaxLength       =   50
            MouseIcon       =   "frmCertificado_qualidade.frx":F3C2
            MousePointer    =   99  'Custom
            TabIndex        =   171
            Text            =   "0"
            ToolTipText     =   "Metal base."
            Top             =   1770
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Frame Frame5_carcaca 
            BackColor       =   &H00E0E0E0&
            Height          =   915
            Left            =   9630
            TabIndex        =   163
            Top             =   1170
            Width           =   1070
            Begin VB.TextBox txtNumero_carcaca5 
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
               Left            =   90
               MaxLength       =   50
               TabIndex        =   45
               Top             =   510
               Width           =   855
            End
            Begin VB.TextBox txtTexto_carcaca5 
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
               Left            =   190
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   180
               Width           =   645
            End
         End
         Begin VB.TextBox txtProduto 
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
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Produto."
            Top             =   1770
            Width           =   1875
         End
         Begin VB.CommandButton cmdProduto 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2010
            Picture         =   "frmCertificado_qualidade.frx":F6CC
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Cadastrar elementos químicos por produto."
            Top             =   1770
            Width           =   315
         End
         Begin VB.Frame Frame1_carcaca 
            BackColor       =   &H00E0E0E0&
            Height          =   915
            Left            =   5220
            TabIndex        =   159
            Top             =   1170
            Width           =   1070
            Begin VB.TextBox txtNumero_carcaca1 
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
               MaxLength       =   50
               TabIndex        =   37
               Top             =   480
               Width           =   855
            End
            Begin VB.TextBox txtTexto_carcaca1 
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
               Left            =   220
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   180
               Width           =   645
            End
         End
         Begin VB.Frame Frame3_carcaca 
            BackColor       =   &H00E0E0E0&
            Height          =   915
            Left            =   7440
            TabIndex        =   161
            Top             =   1170
            Width           =   1070
            Begin VB.TextBox txtTexto_carcaca3 
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
               Left            =   190
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   180
               Width           =   645
            End
            Begin VB.TextBox txtNumero_carcaca3 
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
               Left            =   90
               MaxLength       =   50
               TabIndex        =   41
               Top             =   510
               Width           =   855
            End
         End
         Begin VB.Frame Frame4_carcaca 
            BackColor       =   &H00E0E0E0&
            Height          =   915
            Left            =   8535
            TabIndex        =   162
            Top             =   1170
            Width           =   1070
            Begin VB.TextBox txtTexto_carcaca4 
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
               Left            =   190
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   180
               Width           =   645
            End
            Begin VB.TextBox txtNumero_carcaca4 
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
               Left            =   90
               MaxLength       =   50
               TabIndex        =   43
               Top             =   510
               Width           =   855
            End
         End
         Begin VB.Frame Frame6_carcaca 
            BackColor       =   &H00E0E0E0&
            Height          =   915
            Left            =   10725
            TabIndex        =   164
            Top             =   1170
            Width           =   1070
            Begin VB.TextBox txtTexto_carcaca6 
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
               Left            =   190
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   180
               Width           =   645
            End
            Begin VB.TextBox txtNumero_carcaca6 
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
               Left            =   90
               MaxLength       =   50
               TabIndex        =   47
               Top             =   510
               Width           =   855
            End
         End
         Begin VB.Frame Frame7_carcaca 
            BackColor       =   &H00E0E0E0&
            Height          =   915
            Left            =   11820
            TabIndex        =   165
            Top             =   1170
            Width           =   1070
            Begin VB.TextBox txtNumero_carcaca7 
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
               Left            =   120
               MaxLength       =   50
               TabIndex        =   18
               Top             =   510
               Width           =   855
            End
            Begin VB.TextBox txtTexto_carcaca7 
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
               Left            =   220
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   180
               Width           =   645
            End
         End
         Begin VB.Frame frameTipo_carcaca 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tipo de amostra"
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
            Height          =   615
            Left            =   2460
            TabIndex        =   54
            Top             =   1470
            Width           =   2715
            Begin VB.OptionButton optGranulado 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Granulado"
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
               Height          =   255
               Left            =   1260
               TabIndex        =   35
               Top             =   300
               Width           =   1065
            End
            Begin VB.OptionButton optCavaco 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cavaco"
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
               Height          =   255
               Left            =   150
               TabIndex        =   34
               Top             =   300
               Width           =   1035
            End
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Produto"
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
            Left            =   960
            TabIndex        =   172
            Top             =   1530
            Width           =   570
         End
      End
   End
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
      Height          =   4305
      Left            =   55
      TabIndex        =   55
      Top             =   900
      Width           =   15225
      Begin VB.TextBox txtPedido_interno 
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
         Left            =   2670
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Pedido interno."
         Top             =   390
         Width           =   1635
      End
      Begin VB.CommandButton cmdPedido 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         Picture         =   "frmCertificado_qualidade.frx":F7CE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Localizar pedido interno."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtQtde 
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
         Left            =   13560
         MaxLength       =   50
         TabIndex        =   15
         ToolTipText     =   "Quantidade."
         Top             =   2190
         Width           =   1485
      End
      Begin VB.TextBox txtObs 
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
         Height          =   1365
         Left            =   180
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         ToolTipText     =   "Observação."
         Top             =   2790
         Width           =   14865
      End
      Begin VB.TextBox txtMetal_adicao 
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
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Metal adição."
         Top             =   1590
         Width           =   2385
      End
      Begin VB.TextBox txtDescricao_metal 
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
         Left            =   3030
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Descrição do metal adição"
         Top             =   1590
         Width           =   12015
      End
      Begin VB.CommandButton cmdMetal 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2580
         Picture         =   "frmCertificado_qualidade.frx":F8D0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Localizar metal adição."
         Top             =   1590
         Width           =   315
      End
      Begin VB.TextBox txtID_cliente 
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
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   390
         Width           =   1065
      End
      Begin VB.CommandButton cmdDesenho 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2580
         Picture         =   "frmCertificado_qualidade.frx":F9D2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Localizar produto."
         Top             =   990
         Width           =   315
      End
      Begin VB.TextBox txtLote 
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
         Left            =   9930
         MaxLength       =   50
         TabIndex        =   14
         ToolTipText     =   "Lote."
         Top             =   2190
         Width           =   3615
      End
      Begin VB.TextBox txtCorpo 
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
         TabIndex        =   13
         ToolTipText     =   "Mat. do corpo."
         Top             =   2190
         Width           =   9705
      End
      Begin VB.TextBox txtPedido_cliente 
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
         Left            =   4770
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Pedido do cliente."
         Top             =   390
         Width           =   2415
      End
      Begin VB.TextBox txtDescricao 
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
         Left            =   3030
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Descrição do item."
         Top             =   990
         Width           =   12015
      End
      Begin VB.TextBox txtdesenho 
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
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   990
         Width           =   2385
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   8280
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Nome do cliente."
         Top             =   390
         Width           =   6765
      End
      Begin VB.TextBox txtData 
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
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Data."
         Top             =   390
         Width           =   1125
      End
      Begin VB.TextBox txtID 
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
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Número do certificado"
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   7350
         TabIndex        =   174
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde"
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
         Left            =   14122
         TabIndex        =   169
         Top             =   1980
         Width           =   360
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Obs"
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
         Left            =   7470
         TabIndex        =   168
         Top             =   2580
         Width           =   285
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Metal adição"
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
         Index           =   22
         Left            =   922
         TabIndex        =   66
         Top             =   1380
         Width           =   900
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Index           =   21
         Left            =   8692
         TabIndex        =   65
         Top             =   1380
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Lote"
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
         Left            =   11580
         TabIndex        =   64
         Top             =   1980
         Width           =   315
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mat. do corpo"
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
         Left            =   4530
         TabIndex        =   63
         Top             =   1980
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido do cliente"
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
         Left            =   5310
         TabIndex        =   62
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido interno"
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
         Left            =   2970
         TabIndex        =   61
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Index           =   7
         Left            =   8692
         TabIndex        =   60
         Top             =   780
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Código interno"
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
         Left            =   847
         TabIndex        =   59
         Top             =   780
         Width           =   1050
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Index           =   0
         Left            =   11550
         TabIndex        =   58
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "N° certificado"
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
         Left            =   420
         TabIndex        =   57
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   1905
         TabIndex        =   56
         Top             =   180
         Width           =   345
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   30
      TabIndex        =   176
      Top             =   0
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   1720
      ButtonCount     =   9
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
      ButtonStyle5    =   -1
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
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonAlignment6=   2
      ButtonType6     =   1
      ButtonStyle6    =   -1
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState6    =   -1
      ButtonLeft6     =   209
      ButtonTop6      =   4
      ButtonWidth6    =   2
      ButtonHeight6   =   54
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Ajuda"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Ajuda (F1)"
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
      ButtonLeft7     =   213
      ButtonTop7      =   2
      ButtonWidth7    =   36
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonCaption8  =   "Sair"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Sair (Esc)"
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
      ButtonLeft8     =   251
      ButtonTop8      =   2
      ButtonWidth8    =   26
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonEnabled9  =   0   'False
      BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState9    =   5
      ButtonLeft9     =   279
      ButtonTop9      =   2
      ButtonWidth9    =   24
      ButtonHeight9   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   8790
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCertificado_qualidade.frx":FAD4
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmCertificado_qualidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Certificado_Qualidade As Boolean 'OK
Dim Novo_Carcaca               As Boolean 'OK

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

frmCertificado_qualidade_Abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdDesenho_Click()
On Error GoTo tratar_erro

Ultrasom = False
Liquido = False
frmLiquido_item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirAnaliseQuimica()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If lista_carcaca.ListItems.Count = 0 And txtid_carcaca = "0" Then
    USMsgBox ("Informe o material antes de excluir."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
If Sair = True Then GoTo Pula
    If USMsgBox("Deseja realmente excuir os registros do material?", vbYesNo, "CAPRIND v5.0") = vbYes Then
Pula:
    Conexao.Execute "DELETE from certificado_Quimica where id = " & txtid_carcaca
    '==================================
    Modulo = "Qualidade/Ensaios/Controle de certificados"
    Evento = "Excluir analíse química"
    Documento = txtId & "-" & txtid_carcaca
    ProcGravaEvento
    '==================================
    ProcLimpacampos_carcaca
    Novo_Carcaca = False
    Frame7.Enabled = False
    USMsgBox ("Registros do material excluídos com sucesso."), vbInformation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdMetal_Click()
On Error GoTo tratar_erro

Ultrasom = False
frmUltraSom_item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoAnaliseQuimica()
On Error GoTo tratar_erro

If Novo_Certificado_Qualidade = True Then
    USMsgBox ("Salve o certificado antes de Prosseguir."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
If lista_carcaca.ListItems.Count = 0 Then
    USMsgBox ("Informe o material antes de criar novo registro."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from certificado_Quimica where id = " & txtid_carcaca, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    ProcLimpacampos_carcaca
    TBGravar.AddNew
    TBGravar!Desenho = lista_carcaca.SelectedItem
    TBGravar!Certificado = lista_carcaca.SelectedItem.ListSubItems(1)
    TBGravar!id_certificado = txtId
    TBGravar.Update
    txtid_carcaca = TBGravar!ID
    Frame7.Enabled = True
    Novo_Carcaca = True
    txtProduto.SetFocus
Else
    USMsgBox ("Não é permitido criar novo registro, pois já existe cadastro para o material " & lista_carcaca.SelectedItem & "."), vbInformation, "CAPRIND v5.0"
End If
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Novo_Certificado_Qualidade = True Then Exit Sub
ProcLimpaCampos
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Certificado_qualidade", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
TBGravar!Data = Format(Date, "dd/mm/yy")
TBGravar.Update
txtId = TBGravar!ID
TBGravar.Close
txtData = Format(Date, "DD/mm/yy")
Novo_Certificado_Qualidade = True
Frame2.Enabled = True
Ultrasom = False
Liquido = False
frmLiquido_pedido.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdpedido_Click()
On Error GoTo tratar_erro

Ultrasom = False
Liquido = False
frmLiquido_pedido.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdProduto_Click()
On Error GoTo tratar_erro

frmCertificado_qualidade_analisequimica.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Certificado_Qualidade = True Then
    If USMsgBox("O certificado ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Certificado_Qualidade = True Then
            Exit Sub
        Else
            Unload Me
        End If
    Else
        If txtId.Text <> "" Then
            Sair = True
            ProcExcluir
        End If
    End If
End If
Select Case SSTab1.Tab
    Case 2
        If Novo_Carcaca = True Then
            If USMsgBox("Os dados da análise química ainda não foram salvos, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
                ProcSalvarAnaliseQuimica
                If Novo_Carcaca = True Then Exit Sub
            Else
                Sair = True
                ProcExcluirAnaliseQuimica
            End If
        End If
End Select
Conexao.Execute "DELETE from Certificado_qualidade WHERE Responsavel = 'null'"
Conexao.Execute "DELETE from Certificado_Quimica WHERE Responsavel = 'null'"
Novo_Carcaca = False
Novo_Certificado_Qualidade = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = ""
txtData = ""
txtPedido_interno = ""
txtPedido_cliente = ""
txtid_cliente = ""
txtCliente = ""
txtdesenho = ""
txtdescricao = ""
txtQtde = ""
txtCorpo = ""
txtMetal_adicao = ""
txtDescricao_metal = ""
txtLote = ""
txtObs = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarAnaliseQuimica()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame7.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If lista_carcaca.ListItems.Count = 0 And txtId = "0" Then
    USMsgBox ("Informe o material antes de salvar."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from certificado_Quimica where id = " & txtid_carcaca, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    ProcEnviadados_carcaca
    TBGravar.Update
    If Novo_Carcaca = True Then
        USMsgBox ("Novo registro cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Novo analíse química"
    Else
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar analíse química"
    End If
End If
'==================================
Modulo = "Qualidade/Ensaios/Controle de certificados"
Documento = txtId & "-" & txtid_carcaca
ProcGravaEvento
'==================================
Novo_Carcaca = False
TBGravar.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar2 Me, 14940, 4, True
ProcCarregaToolBar1 Me, 15200, 9, True
Formulario = "Qualidade/Ensaios/Controle de certificados"
Direitos
ProcLimpaVariaveisPrincipais
SSTab1.Tab = 0
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF2: ProcLocalizar
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyEscape: ProcSair
End Select

    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Qualidade/Ensaios/Controle de certificados"
Direitos
ProcLimpaVariaveisPrincipais
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If txtId = "" Then
    USMsgBox ("Informe o certificado antes de excluir."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
If Sair = True Then GoTo Pula
    If USMsgBox("Deseja realmente excuir o certificado?", vbYesNo, "CAPRIND v5.0") = vbYes Then
Pula:
    Conexao.Execute "DELETE from Certificado_qualidade where id = " & txtId
    Conexao.Execute "DELETE from certificado_Quimica where id_certificado = " & txtId
    USMsgBox ("Certificado excluído com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Qualidade/Ensaios/Controle de certificados"
    Evento = "Excluir"
    Documento = txtId
    ProcGravaEvento
    '==================================
    ProcLimpaCampos
    SSTab1.Tab = 0
    ProcLimpacampos_ultra
    Novo_Certificado_Qualidade = False
    Frame2.Enabled = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If txtId = "" Then
    USMsgBox ("Informe o certificado antes de visualizar impressão."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
NomeRel = "CQ_certificado_qualidade.rpt"
ProcImprimirRel "{Certificado_qualidade.id}= " & txtId, ""

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
If txtId = "" Then
    ProcVerificaSalvar
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Certificado_qualidade where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
ProcEnviaDados
TBGravar.Update
If Novo_Certificado_Qualidade = True Then
    USMsgBox ("Novo certificado cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
End If
'==================================
Modulo = "Qualidade/Ensaios/Controle de certificados"
Documento = txtId
ProcGravaEvento
'==================================
Novo_Certificado_Qualidade = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub lista_carcaca_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView lista_carcaca, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_carcaca_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

ProcLimpacampos_carcaca
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from certificado_Quimica where id_certificado = " & txtId & " and desenho = '" & lista_carcaca.SelectedItem & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    ProcPuxadados_carcaca
    Frame7.Enabled = True
Else
    Frame7.Enabled = False
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_liquido_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView lista_liquido, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub lista_liquido_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lista_liquido.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Liquido_penetrante where id = " & lista_liquido.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtData_liquido = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
    txtQTDE_liquido = IIf(IsNull(TBAbrir!Qtde), "", Format(TBAbrir!Qtde, "###,##0.0000"))
    If TBAbrir!Conclusao = "Aprovado" Then optAprovado_liquido.Value = True
    If TBAbrir!Conclusao = "Reprovado" Then optReprovado_liquido.Value = True
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub lista_ultra_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView lista_ultra, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub lista_ultra_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lista_ultra.ListItems.Count = 0 Then Exit Sub
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from UltraSom where id = " & lista_ultra.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    txtData_ultra = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
    txtQtde_ultra = IIf(IsNull(TBAbrir!Qtde), "", Format(TBAbrir!Qtde, "###,##0.0000"))
    txtEspess_ultra = IIf(IsNull(TBAbrir!Espess), "", Format(TBAbrir!Espess, "###,##0.00"))
    txtSuperf_ultra = IIf(IsNull(TBAbrir!Superficie), "", TBAbrir!Superficie)
    If TBAbrir!Conclusao = "Aprovado" Then optAprovado.Value = True
    If TBAbrir!Conclusao = "Reprovado" Then optReprovado.Value = True
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select * from UltraSom_inspetores where id = " & TBAbrir!idInspetores, Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        txtAparelho = IIf(IsNull(TBLISTA!Aparelho), "", TBLISTA!Aparelho)
        txtTransdutor = IIf(IsNull(TBLISTA!Transdutor), "", TBLISTA!Transdutor)
    End If
    TBLISTA.Close
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab
    Case 0
        ProcCarregalista_ultra
        ProcLimpacampos_ultra
    Case 1
        ProcLimpacampos_liquido
        ProcCarregalista_liquido
    Case 2
        ProcCarregalista_carcaca
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesenho_Change()
On Error GoTo tratar_erro

If txtdesenho = "" Then Exit Sub
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    txtImagem = IIf(IsNull(TBItem!imagem), "", TBItem!imagem)
    txtdescricao = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
End If
TBItem.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados()
On Error GoTo tratar_erro

TBGravar!Responsavel = pubUsuario
TBGravar!Data = txtData
TBGravar!Pedido = txtPedido_interno
TBGravar!Pedido_cliente = txtPedido_cliente
TBGravar!IDCliente = IIf(txtid_cliente = "", 0, txtid_cliente)
TBGravar!Qtde = IIf(txtQtde = "", 0, txtQtde)
TBGravar!Cliente = txtCliente
TBGravar!Desenho = txtdesenho
TBGravar!Metal_adicao = txtMetal_adicao
TBGravar!corpo = txtCorpo
TBGravar!LOTE = txtLote
TBGravar!Obs = Trim(txtObs)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

txtId = TBAbrir!ID
txtData = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
txtPedido_interno = IIf(IsNull(TBAbrir!Pedido), "", TBAbrir!Pedido)
txtPedido_cliente = IIf(IsNull(TBAbrir!Pedido_cliente), "", TBAbrir!Pedido_cliente)
txtid_cliente = IIf(IsNull(TBAbrir!IDCliente), "", TBAbrir!IDCliente)
txtCliente = IIf(IsNull(TBAbrir!Cliente), "", TBAbrir!Cliente)
txtdesenho = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
txtMetal_adicao = IIf(IsNull(TBAbrir!Metal_adicao), "", TBAbrir!Metal_adicao)
txtQtde = IIf(IsNull(TBAbrir!Qtde), "", Format(TBAbrir!Qtde, "###,##0.0000"))
txtCorpo = IIf(IsNull(TBAbrir!corpo), "", TBAbrir!corpo)
txtLote = IIf(IsNull(TBAbrir!LOTE), "", TBAbrir!LOTE)
txtObs = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
Novo_Certificado_Qualidade = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub txtMetal_adicao_Change()
On Error GoTo tratar_erro

If txtMetal_adicao = "" Then Exit Sub

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & txtMetal_adicao & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    txtDescricao_metal = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
End If
TBItem.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Sub ProcLimpacampos_ultra()
On Error GoTo tratar_erro

txtData_ultra = ""
txtQtde_ultra = ""
txtEspess_ultra = ""
txtSuperf_ultra = ""
txtAparelho = ""
txtTransdutor = ""
optAprovado.Value = False
optReprovado.Value = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregalista_ultra()
On Error GoTo tratar_erro

lista_ultra.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from UltraSom where pedido_interno = '" & txtPedido_interno & "' and desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With lista_ultra.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "DD/mm/yy"))
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

Sub ProcCarregalista_liquido()
On Error GoTo tratar_erro

lista_liquido.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Liquido_penetrante where pedido_interno = '" & txtPedido_interno & "' and desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With lista_liquido.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "DD/mm/yy"))
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

Sub ProcCarregalista_carcaca()
On Error GoTo tratar_erro

lista_carcaca.ListItems.Clear
If txtLote = "" Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Estoque_Controle where lote = '" & txtLote & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With lista_carcaca.ListItems
            .Add , , TBLISTA!Desenho
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Certificado), "", TBLISTA!Certificado)
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

Sub ProcLimpacampos_liquido()
On Error GoTo tratar_erro

txtData_liquido = ""
txtQTDE_liquido = ""
optAprovado_liquido.Value = False
optReprovado_liquido.Value = False
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpacampos_carcaca()
On Error GoTo tratar_erro

txtid_carcaca = 0
txtID_Elemento = 0
optCavaco.Value = False
optGranulado.Value = False
txtProduto = ""
txtTexto_carcaca1 = ""
txtTexto_carcaca2 = ""
txtTexto_carcaca3 = ""
txtTexto_carcaca4 = ""
txtTexto_carcaca5 = ""
txtTexto_carcaca6 = ""
txtTexto_carcaca7 = ""
txtTexto_carcaca8 = ""
txtTexto_carcaca9 = ""
txtNumero_carcaca1 = ""
txtNumero_carcaca2 = ""
txtNumero_carcaca3 = ""
txtNumero_carcaca4 = ""
txtNumero_carcaca5 = ""
txtNumero_carcaca6 = ""
txtNumero_carcaca7 = ""
txtNumero_carcaca8 = ""
txtNumero_carcaca9 = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviadados_carcaca()
On Error GoTo tratar_erro

TBGravar!Responsavel = pubUsuario
If optCavaco = True Then TBGravar!AMOSTRA = "Cavaco"
If optGranulado = True Then TBGravar!AMOSTRA = "Granulado"
TBGravar!Numero1 = IIf(txtNumero_carcaca1 = "", 0, txtNumero_carcaca1)
TBGravar!Numero2 = IIf(txtNumero_carcaca2 = "", 0, txtNumero_carcaca2)
TBGravar!Numero3 = IIf(txtNumero_carcaca3 = "", 0, txtNumero_carcaca3)
TBGravar!Numero4 = IIf(txtNumero_carcaca4 = "", 0, txtNumero_carcaca4)
TBGravar!Numero5 = IIf(txtNumero_carcaca5 = "", 0, txtNumero_carcaca5)
TBGravar!Numero6 = IIf(txtNumero_carcaca6 = "", 0, txtNumero_carcaca6)
TBGravar!Numero7 = IIf(txtNumero_carcaca7 = "", 0, txtNumero_carcaca7)
TBGravar!Numero8 = IIf(txtNumero_carcaca8 = "", 0, txtNumero_carcaca8)
TBGravar!Numero9 = IIf(txtNumero_carcaca9 = "", 0, txtNumero_carcaca9)
TBGravar!ID_Elemento = txtID_Elemento

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxadados_carcaca()
On Error GoTo tratar_erro

txtid_carcaca = TBAbrir!ID
If TBAbrir!AMOSTRA = "Cavaco" Then optCavaco.Value = True
If TBAbrir!AMOSTRA = "Granulado" Then optGranulado.Value = True
txtID_Elemento = IIf(IsNull(TBAbrir!ID_Elemento), 0, TBAbrir!ID_Elemento)
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Certificado_Analise where ID = " & txtID_Elemento, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    txtProduto = IIf(IsNull(TBFIltro!Produto), "", TBFIltro!Produto)
    txtTexto_carcaca1 = IIf(IsNull(TBFIltro!Texto), "", TBFIltro!Texto)
    txtTexto_carcaca2 = IIf(IsNull(TBFIltro!Texto2), "", TBFIltro!Texto2)
    txtTexto_carcaca3 = IIf(IsNull(TBFIltro!Texto3), "", TBFIltro!Texto3)
    txtTexto_carcaca4 = IIf(IsNull(TBFIltro!texto4), "", TBFIltro!texto4)
    txtTexto_carcaca5 = IIf(IsNull(TBFIltro!Texto5), "", TBFIltro!Texto5)
    txtTexto_carcaca6 = IIf(IsNull(TBFIltro!Texto6), "", TBFIltro!Texto6)
    txtTexto_carcaca7 = IIf(IsNull(TBFIltro!Texto7), "", TBFIltro!Texto7)
    txtTexto_carcaca8 = IIf(IsNull(TBFIltro!Texto8), "", TBFIltro!Texto8)
    txtTexto_carcaca9 = IIf(IsNull(TBFIltro!Texto9), "", TBFIltro!Texto9)
End If
TBFIltro.Close
txtNumero_carcaca1 = IIf(IsNull(TBAbrir!Numero1), "", Format(TBAbrir!Numero1, "###,##0.00"))
txtNumero_carcaca2 = IIf(IsNull(TBAbrir!Numero2), "", Format(TBAbrir!Numero2, "###,##0.00"))
txtNumero_carcaca3 = IIf(IsNull(TBAbrir!Numero3), "", Format(TBAbrir!Numero3, "###,##0.00"))
txtNumero_carcaca4 = IIf(IsNull(TBAbrir!Numero4), "", Format(TBAbrir!Numero4, "###,##0.00"))
txtNumero_carcaca5 = IIf(IsNull(TBAbrir!Numero5), "", Format(TBAbrir!Numero5, "###,##0.00"))
txtNumero_carcaca6 = IIf(IsNull(TBAbrir!Numero6), "", Format(TBAbrir!Numero6, "###,##0.00"))
txtNumero_carcaca7 = IIf(IsNull(TBAbrir!Numero7), "", Format(TBAbrir!Numero7, "###,##0.00"))
txtNumero_carcaca8 = IIf(IsNull(TBAbrir!Numero8), "", Format(TBAbrir!Numero8, "###,##0.00"))
txtNumero_carcaca9 = IIf(IsNull(TBAbrir!Numero9), "", Format(TBAbrir!Numero9, "###,##0.00"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNumero_carcaca1_LostFocus()
On Error GoTo tratar_erro

If txtNumero_carcaca1.Text <> "" Then
    VerifNumero = txtNumero_carcaca1.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNumero_carcaca1.Text = ""
        txtNumero_carcaca1.SetFocus
        Exit Sub
    End If
    txtNumero_carcaca1.Text = Format(txtNumero_carcaca1.Text, "###,##0.00")
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNumero_carcaca2_LostFocus()
On Error GoTo tratar_erro

If txtNumero_carcaca2.Text <> "" Then
    VerifNumero = txtNumero_carcaca2.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNumero_carcaca2.Text = ""
        txtNumero_carcaca2.SetFocus
        Exit Sub
    End If
    txtNumero_carcaca2.Text = Format(txtNumero_carcaca2.Text, "###,##0.00")
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNumero_carcaca3_LostFocus()
On Error GoTo tratar_erro

If txtNumero_carcaca3.Text <> "" Then
    VerifNumero = txtNumero_carcaca3.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNumero_carcaca3.Text = ""
        txtNumero_carcaca3.SetFocus
        Exit Sub
    End If
    txtNumero_carcaca3.Text = Format(txtNumero_carcaca3.Text, "###,##0.00")
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNumero_carcaca4_LostFocus()
On Error GoTo tratar_erro

If txtNumero_carcaca4.Text <> "" Then
    VerifNumero = txtNumero_carcaca4.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNumero_carcaca4.Text = ""
        txtNumero_carcaca4.SetFocus
        Exit Sub
    End If
    txtNumero_carcaca4.Text = Format(txtNumero_carcaca4.Text, "###,##0.00")
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNumero_carcaca5_LostFocus()
On Error GoTo tratar_erro

If txtNumero_carcaca5.Text <> "" Then
    VerifNumero = txtNumero_carcaca5.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNumero_carcaca5.Text = ""
        txtNumero_carcaca5.SetFocus
        Exit Sub
    End If
    txtNumero_carcaca5.Text = Format(txtNumero_carcaca5.Text, "###,##0.00")
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNumero_carcaca6_LostFocus()
On Error GoTo tratar_erro

If txtNumero_carcaca6.Text <> "" Then
    VerifNumero = txtNumero_carcaca6.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNumero_carcaca6.Text = ""
        txtNumero_carcaca6.SetFocus
        Exit Sub
    End If
    txtNumero_carcaca6.Text = Format(txtNumero_carcaca6.Text, "###,##0.00")
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNumero_carcaca7_LostFocus()
On Error GoTo tratar_erro

If txtNumero_carcaca7.Text <> "" Then
    VerifNumero = txtNumero_carcaca7.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNumero_carcaca7.Text = ""
        txtNumero_carcaca7.SetFocus
        Exit Sub
    End If
    txtNumero_carcaca7.Text = Format(txtNumero_carcaca7.Text, "###,##0.00")
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNumero_carcaca8_LostFocus()
On Error GoTo tratar_erro

If txtNumero_carcaca8.Text <> "" Then
    VerifNumero = txtNumero_carcaca8.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNumero_carcaca8.Text = ""
        txtNumero_carcaca8.SetFocus
        Exit Sub
    End If
    txtNumero_carcaca8.Text = Format(txtNumero_carcaca8.Text, "###,##0.00")
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNumero_carcaca9_LostFocus()
On Error GoTo tratar_erro

If txtNumero_carcaca9.Text <> "" Then
    VerifNumero = txtNumero_carcaca9.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNumero_carcaca9.Text = ""
        txtNumero_carcaca9.SetFocus
        Exit Sub
    End If
    txtNumero_carcaca9.Text = Format(txtNumero_carcaca9.Text, "###,##0.00")
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtqtde_LostFocus()
On Error GoTo tratar_erro

If txtQtde.Text <> "" Then
    VerifNumero = txtQtde.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde.Text = ""
        txtQtde.SetFocus
        Exit Sub
    End If
    txtQtde.Text = Format(txtQtde.Text, "###,##0.0000")
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcLocalizar
    Case 3: ProcSalvar
    Case 4: ProcExcluir
    Case 5: ProcImprimir
    'Case 7: ProcAjuda
    Case 8: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoAnaliseQuimica
    Case 2: ProcSalvarAnaliseQuimica
    Case 3: ProcExcluirAnaliseQuimica
  End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

