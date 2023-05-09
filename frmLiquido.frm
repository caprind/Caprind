VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmLiquido 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Qualidade - Ensaios - Líquido penetrante"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
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
      Height          =   9135
      Left            =   55
      TabIndex        =   32
      Top             =   900
      Width           =   15225
      Begin VB.TextBox txtImagem 
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
         Left            =   4680
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Imagem."
         Top             =   3810
         Width           =   4500
      End
      Begin VB.CommandButton cmdImagem 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   9180
         Picture         =   "frmLiquido.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Abrir arquivo."
         Top             =   3810
         Width           =   255
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
         ToolTipText     =   "Número do líquido penetrante."
         Top             =   390
         Width           =   1335
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
         Picture         =   "frmLiquido.frx":00B5
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Localizar pedido interno."
         Top             =   390
         Width           =   315
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
         Left            =   2250
         Picture         =   "frmLiquido.frx":01B7
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Localizar produto."
         Top             =   990
         Width           =   315
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
         Left            =   9540
         TabIndex        =   52
         Top             =   3570
         Width           =   2415
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
            Left            =   150
            TabIndex        =   24
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
            Left            =   1200
            TabIndex        =   25
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.Frame frameAceitacao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Critério de aceitação"
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
         Left            =   11970
         TabIndex        =   51
         Top             =   3570
         Width           =   3075
         Begin VB.OptionButton OptD 
            BackColor       =   &H00E0E0E0&
            Caption         =   "D"
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
            Left            =   2400
            TabIndex        =   29
            Top             =   300
            Width           =   465
         End
         Begin VB.OptionButton optC 
            BackColor       =   &H00E0E0E0&
            Caption         =   "C"
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
            Left            =   1770
            TabIndex        =   28
            Top             =   300
            Width           =   465
         End
         Begin VB.OptionButton optB 
            BackColor       =   &H00E0E0E0&
            Caption         =   "B1/B2"
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
            Left            =   810
            TabIndex        =   27
            Top             =   300
            Width           =   795
         End
         Begin VB.OptionButton optA 
            BackColor       =   &H00E0E0E0&
            Caption         =   "A"
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
            TabIndex        =   26
            Top             =   300
            Width           =   465
         End
      End
      Begin VB.ComboBox cmbLocal 
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
         ItemData        =   "frmLiquido.frx":02B9
         Left            =   180
         List            =   "frmLiquido.frx":02C3
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         ToolTipText     =   "Local ensaiado."
         Top             =   3810
         Width           =   4500
      End
      Begin VB.Frame frameRevetador 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Revetador"
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
         Height          =   1545
         Left            =   7613
         TabIndex        =   46
         Top             =   1950
         Width           =   7432
         Begin VB.TextBox txtLote_revetador 
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
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   20
            ToolTipText     =   "Lote."
            Top             =   1080
            Width           =   6165
         End
         Begin VB.TextBox txtFabricante_revetador 
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
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   19
            ToolTipText     =   "Fabricante."
            Top             =   720
            Width           =   6165
         End
         Begin VB.TextBox txtMarca_revetador 
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
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   18
            ToolTipText     =   "Marca."
            Top             =   360
            Width           =   6165
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Marca :"
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
            Left            =   510
            TabIndex        =   49
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fabricante :"
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
            Left            =   180
            TabIndex        =   48
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Lote :"
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
            Left            =   630
            TabIndex        =   47
            Top             =   1080
            Width           =   420
         End
      End
      Begin VB.Frame framePenetrante 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Penetrante"
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
         Height          =   1545
         Left            =   180
         TabIndex        =   43
         Top             =   1950
         Width           =   7425
         Begin VB.TextBox txtMarca_penetrante 
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
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   15
            ToolTipText     =   "Marca."
            Top             =   360
            Width           =   6165
         End
         Begin VB.TextBox txtLote_penetrante 
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
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   17
            ToolTipText     =   "Lote."
            Top             =   1080
            Width           =   6165
         End
         Begin VB.TextBox txtFabricante_penetrante 
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
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   16
            ToolTipText     =   "Fabricante."
            Top             =   720
            Width           =   6165
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Marca :"
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
            Left            =   510
            TabIndex        =   57
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Lote :"
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
            Left            =   630
            TabIndex        =   45
            Top             =   1080
            Width           =   420
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fabricante :"
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
            Left            =   180
            TabIndex        =   44
            Top             =   720
            Width           =   870
         End
      End
      Begin VB.TextBox txtLimpeza 
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
         Left            =   6840
         MaxLength       =   50
         TabIndex        =   13
         ToolTipText     =   "Metódo de limpeza."
         Top             =   1590
         Width           =   3885
      End
      Begin VB.TextBox txtRemovedor 
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
         Left            =   3330
         MaxLength       =   50
         TabIndex        =   12
         ToolTipText     =   "Removedor."
         Top             =   1590
         Width           =   3495
      End
      Begin VB.TextBox txtNorma 
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
         Left            =   10740
         MaxLength       =   50
         TabIndex        =   14
         ToolTipText     =   "Norma."
         Top             =   1590
         Width           =   4305
      End
      Begin VB.TextBox txtEstado 
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
         MaxLength       =   30
         TabIndex        =   10
         ToolTipText     =   "Estado da superfície."
         Top             =   1590
         Width           =   2175
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
         Left            =   2370
         MaxLength       =   50
         TabIndex        =   11
         ToolTipText     =   "Quantidade."
         Top             =   1590
         Width           =   945
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
         ToolTipText     =   "pedido do cliente."
         Top             =   390
         Width           =   2415
      End
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
         Left            =   2700
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Descrição do item."
         Top             =   990
         Width           =   12345
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
         Width           =   2055
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
         ToolTipText     =   "Nomo do cliente."
         Top             =   390
         Width           =   6765
      End
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
         ForeColor       =   &H00000000&
         Height          =   4575
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Observações."
         Top             =   4410
         Width           =   14850
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
         Index           =   21
         Left            =   7485
         TabIndex        =   56
         Top             =   180
         Width           =   495
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
         Left            =   1920
         TabIndex        =   55
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Número líquido"
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
         Left            =   322
         TabIndex        =   54
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Imagem"
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
         Left            =   6645
         TabIndex        =   53
         Top             =   3600
         Width           =   570
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Local ensaiado"
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
         Left            =   1905
         TabIndex        =   50
         Top             =   3600
         Width           =   1050
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Norma"
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
         Left            =   12690
         TabIndex        =   42
         Top             =   1380
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Removedor"
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
         Left            =   4665
         TabIndex        =   41
         Top             =   1380
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Metódo limpeza"
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
         Left            =   8242
         TabIndex        =   40
         Top             =   1380
         Width           =   1110
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado superf."
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
         Left            =   735
         TabIndex        =   39
         Top             =   1380
         Width           =   1065
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
         Left            =   2662
         TabIndex        =   38
         Top             =   1380
         Width           =   360
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
         Left            =   5370
         TabIndex        =   37
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
         TabIndex        =   36
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
         Left            =   8527
         TabIndex        =   35
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
         Left            =   682
         TabIndex        =   34
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
         Left            =   11415
         TabIndex        =   33
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Index           =   1
         Left            =   7133
         TabIndex        =   31
         Top             =   4200
         Width           =   945
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   58
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
         Left            =   11610
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmLiquido.frx":0302
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmLiquido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Liquido As Boolean 'OK

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

frmLiquido_Abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdDesenho_Click()
On Error GoTo tratar_erro

Ultrasom = False
Liquido = True
frmLiquido_item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdImagem_Click()
On Error GoTo tratar_erro

If txtImagem = "" Then Exit Sub
fotos = txtImagem
If Right(fotos, 3) = "pdf" Or Right(fotos, 3) = "PDF" Then
    frmproj_produto_abrir_arquivos.Show 1
Else
    frmproj_produto_abrir_imagem.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

ProcLimpaCampos
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from liquido_penetrante", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
TBGravar!Data = Format(Date, "dd/mm/yy")
TBGravar.Update
txtId = TBGravar!ID
TBGravar.Close
txtData = Format(Date, "DD/mm/yy")
Novo_Liquido = True
Frame2.Enabled = True
Ultrasom = False
Liquido = True
frmLiquido_pedido.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdpedido_Click()
On Error GoTo tratar_erro

Ultrasom = False
Liquido = True
frmLiquido_pedido.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Liquido = True Then
    If USMsgBox("o líquida penetrante ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Liquido = True Then
            Exit Sub
        Else
            Unload Me
        End If
    Else
        If txtId.Text <> "" Then Conexao.Execute "DELETE from Liquido_penetrante WHERE id = " & txtId.Text
    End If
Else
    Conexao.Execute "DELETE from Liquido_penetrante WHERE responsavel = 'null'"
End If
Novo_Liquido = False
ProcLimpaVariaveisPrincipais
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
txtEstado = ""
txtQtde = ""
txtRemovedor = ""
txtLimpeza = ""
txtNorma = ""
txtMarca_penetrante = ""
txtMarca_revetador = ""
txtFabricante_penetrante = ""
txtFabricante_revetador = ""
txtLote_penetrante = ""
txtLote_revetador = ""
cmbLocal.ListIndex = -1
txtImagem = ""
optA.Value = False
optB.Value = False
optC.Value = False
OptD.Value = False
optAprovado.Value = False
optReprovado.Value = False
txtObs = ""

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

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15200, 9, True
ProcLimpaVariaveisPrincipais
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

ProcLimpaVariaveisPrincipais
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If txtId = "" Then
    USMsgBox ("Informe o líquido penetrante antes de excluir."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente excuir o líquido penetrante?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Conexao.Execute "DELETE from liquido_penetrante where id = " & txtId
    '==================================
    Modulo = "Qualidade/Ensaios/Líquido penetrante"
    Evento = "Excluir"
    Documento = txtId
    ProcGravaEvento
    '==================================
    ProcLimpaCampos
    Novo_Liquido = False
    Frame2.Enabled = False
    USMsgBox ("Registro excluído com sucesso."), vbInformation, "CAPRIND v5.0"
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If txtId = "" Then
    USMsgBox ("Informe o líquido penetrante antes de visualizar impressão."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
NomeRel = "CQ_liquido_penetrante.rpt"
ProcImprimirRel "{Liquido_penetrante.id}= " & txtId, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If txtId = "" Then
    ProcVerificaSalvar
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Liquido_penetrante where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    ProcEnviaDados
    TBGravar.Update
    If Novo_Liquido = True Then
        USMsgBox ("Novo líquido penetrante cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Qualidade/Ensaios/Líquido penetrante"
        Evento = "Novo"
        Documento = txtId
        ProcGravaEvento
        '==================================
    Else
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Qualidade/Ensaios/Líquido penetrante"
        Evento = "Alterar"
        Documento = txtId
        ProcGravaEvento
        '==================================
    End If
    Novo_Liquido = False
End If

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
TBGravar!pedido_interno = txtPedido_interno
TBGravar!Pedido_cliente = txtPedido_cliente
TBGravar!IDCliente = IIf(txtid_cliente = "", 0, txtid_cliente)
TBGravar!Cliente = txtCliente
TBGravar!Desenho = txtdesenho
TBGravar!Estado_superfice = txtEstado
TBGravar!Qtde = IIf(txtQtde = "", 0, txtQtde)
TBGravar!Removedor = txtRemovedor
TBGravar!Limpeza = txtLimpeza
TBGravar!Norma = txtNorma
TBGravar!Marca_penet = txtMarca_penetrante
TBGravar!Fabricante_penet = txtFabricante_penetrante
TBGravar!Lote_penet = txtLote_penetrante
TBGravar!Marca_rev = txtMarca_revetador
TBGravar!Fabricante_rev = txtFabricante_revetador
TBGravar!Lote_rev = txtLote_revetador
TBGravar!Local = IIf(cmbLocal = "", Null, cmbLocal)
If optA = True Then TBGravar!Criterio_aceitacao = "A"
If optB = True Then TBGravar!Criterio_aceitacao = "B"
If optC = True Then TBGravar!Criterio_aceitacao = "C"
If OptD = True Then TBGravar!Criterio_aceitacao = "D"
If optAprovado = True Then TBGravar!Conclusao = "Aprovado"
If optReprovado = True Then TBGravar!Conclusao = "Reprovado"
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
txtPedido_interno = IIf(IsNull(TBAbrir!pedido_interno), "", TBAbrir!pedido_interno)
txtPedido_cliente = IIf(IsNull(TBAbrir!Pedido_cliente), "", TBAbrir!Pedido_cliente)
txtid_cliente = IIf(IsNull(TBAbrir!IDCliente), "", TBAbrir!IDCliente)
txtCliente = IIf(IsNull(TBAbrir!Cliente), "", TBAbrir!Cliente)
txtdesenho = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
txtEstado = IIf(IsNull(TBAbrir!Estado_superfice), "", TBAbrir!Estado_superfice)
txtQtde = IIf(IsNull(TBAbrir!Qtde), "", Format(TBAbrir!Qtde, "###,##0.0000"))
txtRemovedor = IIf(IsNull(TBAbrir!Removedor), "", TBAbrir!Removedor)
txtLimpeza = IIf(IsNull(TBAbrir!Limpeza), "", TBAbrir!Limpeza)
txtNorma = IIf(IsNull(TBAbrir!Norma), "", TBAbrir!Norma)
txtMarca_penetrante = IIf(IsNull(TBAbrir!Marca_penet), "", TBAbrir!Marca_penet)
txtFabricante_penetrante = IIf(IsNull(TBAbrir!Fabricante_penet), "", TBAbrir!Fabricante_penet)
txtLote_penetrante = IIf(IsNull(TBAbrir!Lote_penet), "", TBAbrir!Lote_penet)
txtMarca_revetador = IIf(IsNull(TBAbrir!Marca_rev), "", TBAbrir!Marca_rev)
txtFabricante_revetador = IIf(IsNull(TBAbrir!Fabricante_rev), "", TBAbrir!Fabricante_rev)
txtLote_revetador = IIf(IsNull(TBAbrir!Lote_rev), "", TBAbrir!Lote_rev)
If IsNull(TBAbrir!Local) = False Then
    cmbLocal = TBAbrir!Local
End If
If TBAbrir!Criterio_aceitacao = "A" Then optA.Value = True
If TBAbrir!Criterio_aceitacao = "B" Then optB.Value = True
If TBAbrir!Criterio_aceitacao = "C" Then optC.Value = True
If TBAbrir!Criterio_aceitacao = "D" Then OptD.Value = True
If TBAbrir!Conclusao = "Aprovado" Then optAprovado.Value = True
If TBAbrir!Conclusao = "Reprovado" Then optReprovado.Value = True
txtObs = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
Novo_Liquido = False

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

