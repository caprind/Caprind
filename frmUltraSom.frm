VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmUltraSom 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Qualidade - Ensaios - Ultra-som"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   15360
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin VB.Frame frameAceitacao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Critério de aceitação"
      Enabled         =   0   'False
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
      Height          =   1305
      Left            =   13500
      TabIndex        =   60
      Top             =   8700
      Width           =   1785
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
         Height          =   210
         Left            =   540
         TabIndex        =   30
         Top             =   240
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
         Height          =   210
         Left            =   540
         TabIndex        =   31
         Top             =   500
         Width           =   795
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
         Height          =   210
         Left            =   540
         TabIndex        =   32
         Top             =   760
         Width           =   465
      End
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
         Height          =   210
         Left            =   540
         TabIndex        =   33
         Top             =   1020
         Width           =   465
      End
   End
   Begin VB.Frame frameConclusao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Conclusão"
      Enabled         =   0   'False
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
      Left            =   13495
      TabIndex        =   59
      Top             =   7950
      Width           =   1785
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
         Left            =   360
         TabIndex        =   29
         Top             =   420
         Width           =   1095
      End
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
         Left            =   360
         TabIndex        =   28
         Top             =   210
         Width           =   1035
      End
   End
   Begin VB.Frame framePenetrante 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Inspetor"
      Enabled         =   0   'False
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
      Height          =   2055
      Left            =   55
      TabIndex        =   55
      Top             =   7950
      Width           =   13395
      Begin VB.TextBox txtcodigo 
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
         TabIndex        =   23
         Top             =   450
         Width           =   1545
      End
      Begin VB.TextBox txtNome 
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
         Left            =   2210
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Nome."
         Top             =   450
         Width           =   11005
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
         Left            =   210
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Aparelho."
         Top             =   1050
         Width           =   13005
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
         Left            =   210
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Transdutor."
         Top             =   1620
         Width           =   13005
      End
      Begin VB.CommandButton cmdInspetor 
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
         Left            =   1770
         Picture         =   "frmUltraSom.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Localizar inspetor."
         Top             =   450
         Width           =   315
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
         Index           =   12
         Left            =   705
         TabIndex        =   61
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
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
         Left            =   7510
         TabIndex        =   58
         Top             =   240
         Width           =   405
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
         Index           =   15
         Left            =   6390
         TabIndex        =   57
         Top             =   840
         Width           =   645
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
         Index           =   16
         Left            =   6315
         TabIndex        =   56
         Top             =   1410
         Width           =   795
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
      Height          =   6855
      Left            =   55
      TabIndex        =   34
      Top             =   1005
      Width           =   15225
      Begin VB.TextBox Txt_responsavel 
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
         Left            =   3990
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   390
         Width           =   5775
      End
      Begin VB.TextBox txtEspess 
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
         Left            =   6630
         MaxLength       =   50
         TabIndex        =   15
         ToolTipText     =   "Espessura do revestimento."
         Top             =   2190
         Width           =   1695
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
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Metal adição."
         Top             =   1590
         Width           =   3705
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
         Left            =   4320
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   1590
         Width           =   10695
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
         Left            =   3900
         Picture         =   "frmUltraSom.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Localizar metal adição."
         Top             =   1590
         Width           =   315
      End
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
         Left            =   4470
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Imagem."
         Top             =   2820
         Width           =   5865
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
         Left            =   11490
         Picture         =   "frmUltraSom.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Localizar pedido interno."
         Top             =   390
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
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   990
         Width           =   1005
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
         Left            =   9870
         Picture         =   "frmUltraSom.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Localizar produtos/itens."
         Top             =   990
         Width           =   315
      End
      Begin VB.ComboBox cmbSuperficie 
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
         ItemData        =   "frmUltraSom.frx":0408
         Left            =   180
         List            =   "frmUltraSom.frx":0415
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Superficie revestida."
         Top             =   2820
         Width           =   4275
      End
      Begin VB.TextBox txtMetal_base 
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
         TabIndex        =   14
         ToolTipText     =   "Metal base."
         Top             =   2190
         Width           =   6435
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
         Left            =   9990
         MaxLength       =   50
         TabIndex        =   17
         ToolTipText     =   "Norma."
         Top             =   2190
         Width           =   5025
      End
      Begin VB.TextBox txtCalibracao 
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
         Left            =   10770
         MaxLength       =   50
         TabIndex        =   21
         ToolTipText     =   "Ref. calibração."
         Top             =   2820
         Width           =   4245
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
         Left            =   8340
         MaxLength       =   50
         TabIndex        =   16
         ToolTipText     =   "Quantidade."
         Top             =   2190
         Width           =   1635
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
         Left            =   11910
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Pedido do cliente."
         Top             =   390
         Width           =   3105
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
         Left            =   9780
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Pedido interno."
         Top             =   390
         Width           =   1665
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
         Left            =   10320
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   990
         Width           =   4695
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
         Left            =   8100
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   990
         Width           =   1725
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
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Nome do cliente."
         Top             =   990
         Width           =   6885
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
         Left            =   2250
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "cadastro."
         Top             =   390
         Width           =   1695
      End
      Begin VB.TextBox txtID 
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
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Número do ultra-som."
         Top             =   390
         Width           =   2025
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
         Height          =   3285
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         ToolTipText     =   "Observações."
         Top             =   3450
         Width           =   14835
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Index           =   18
         Left            =   6420
         TabIndex        =   54
         Top             =   180
         Width           =   915
      End
      Begin VB.OLE OLE1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         DisplayType     =   1  'Icon
         Height          =   315
         HostName        =   "Abrir"
         Left            =   10350
         TabIndex        =   20
         Top             =   2820
         Width           =   315
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         Left            =   600
         TabIndex        =   52
         Top             =   780
         Width           =   165
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Espess. revest."
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
         Left            =   8595
         TabIndex        =   51
         Top             =   1980
         Width           =   1125
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
         Left            =   1582
         TabIndex        =   50
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
         Left            =   9292
         TabIndex        =   49
         Top             =   1380
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Caminho da imagem"
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
         Left            =   6690
         TabIndex        =   48
         Top             =   2610
         Width           =   1425
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Superficie revest."
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
         Left            =   1680
         TabIndex        =   47
         Top             =   2610
         Width           =   1275
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
         Left            =   12270
         TabIndex        =   46
         Top             =   1980
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Metal base"
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
         Left            =   3007
         TabIndex        =   45
         Top             =   1980
         Width           =   780
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. calibração"
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
         Left            =   12322
         TabIndex        =   44
         Top             =   2610
         Width           =   1080
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde."
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
         Left            =   7267
         TabIndex        =   43
         Top             =   1980
         Width           =   420
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
         Left            =   12855
         TabIndex        =   42
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
         Left            =   10095
         TabIndex        =   41
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
         Left            =   12322
         TabIndex        =   40
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
         Left            =   8437
         TabIndex        =   39
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
         Left            =   4395
         TabIndex        =   38
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "N° ultra-som"
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
         Index           =   3
         Left            =   645
         TabIndex        =   37
         Top             =   180
         Width           =   1080
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
         Left            =   2925
         TabIndex        =   36
         Top             =   180
         Width           =   345
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
         Left            =   7125
         TabIndex        =   35
         Top             =   3240
         Width           =   945
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I"
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
         Height          =   315
         Left            =   10350
         TabIndex        =   53
         Top             =   2880
         Width           =   315
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   62
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
         Left            =   14220
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmUltraSom.frx":043E
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmUltraSom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Ultrasom As Boolean 'OK

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

frmUltraSom_Abrir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdDesenho_Click()
On Error GoTo tratar_erro

Ultrasom = True
Liquido = False
frmLiquido_item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdInspetor_Click()
On Error GoTo tratar_erro

frmUltraSom_inspetores.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdMetal_Click()
On Error GoTo tratar_erro

Ultrasom = True
frmUltraSom_item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

ProcLimpaCampos
Novo_Ultrasom = True
Frame2.Enabled = True
framePenetrante.Enabled = True
frameConclusao.Enabled = True
frameAceitacao.Enabled = True
cmdpedido_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdpedido_Click()
On Error GoTo tratar_erro

Ultrasom = True
Liquido = False
frmLiquido_pedido.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Ultrasom = True Then
    If USMsgBox("O ultra-som ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Ultrasom = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Ultrasom = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = ""
txtData = Format(Date, "DD/mm/yy")
Txt_responsavel = pubUsuario
txtPedido_interno = ""
txtPedido_cliente = ""
txtid_cliente = ""
txtCliente = ""
txtdesenho = ""
txtdescricao = ""
txtMetal_adicao = ""
txtDescricao_metal = ""
txtMetal_base = ""
txtEspess = ""
txtQtde = ""
txtNorma = ""
cmbSuperficie.ListIndex = -1
txtImagem = Localrel & "\imagens\caprind.bmp"
txtCalibracao = ""
txtObs = ""
txtCodigo = ""
txtNome = ""
txtAparelho = ""
txtTransdutor = ""
optAprovado.Value = False
optReprovado.Value = False
optA.Value = False
optB.Value = False
optC.Value = False
OptD.Value = False
fotos = Localrel & "\imagens\caprind.bmp"
OLE1.CreateLink (fotos)
1:

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15200, 9, True
Formulario = "Qualidade/Ensaios/Ultra-som"
Direitos
ProcLimpaVariaveisPrincipais
    
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

Formulario = "Qualidade/Ensaios/Ultra-som"
Direitos
ProcLimpaVariaveisPrincipais
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If txtId = "" Then
    USMsgBox ("Informe o ultra-som antes de excluir."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If USMsgBox("Deseja realmente excuir este ultra-som?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Conexao.Execute "DELETE from ultrasom where id = " & txtId
    USMsgBox ("Ultra-som excluído com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Qualidade/Ensaios/Ultra-som"
    Evento = "Excluir"
    ID_documento = txtId
    Documento = "Número do ultra-som: " & txtId
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcLimpaCampos
    Novo_Ultrasom = False
    Frame2.Enabled = False
    framePenetrante.Enabled = False
    frameConclusao.Enabled = False
    frameAceitacao.Enabled = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If txtId = "" Then
    USMsgBox ("Informe o ultra-som antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
NomeRel = "CQ_ultra_som.rpt"
ProcImprimirRel "{UltraSom.id}= " & txtId, ""
        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Frame2.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from ultrasom where id = " & IIf(txtId = "", 0, txtId), Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
ProcEnviaDados
TBGravar.Update
txtId = TBGravar!ID
TBGravar.Close
If Novo_Ultrasom = True Then
    USMsgBox ("Novo ultra-som cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
End If
'==================================
Modulo = "Qualidade/Ensaios/Ultra-som"
ID_documento = txtId
Documento = "Número do ultra-som: " & txtId
Documento1 = ""
ProcGravaEvento
'==================================
Novo_Ultrasom = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCodigo_Change()
On Error GoTo tratar_erro

txtNome = ""
txtAparelho = ""
txtTransdutor = ""

If txtCodigo = "" Then Exit Sub
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from UltraSom_inspetores where id = " & txtCodigo, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    txtNome = IIf(IsNull(TBItem!Nome), "", TBItem!Nome)
    txtAparelho = IIf(IsNull(TBItem!Aparelho), "", TBItem!Aparelho)
    txtTransdutor = IIf(IsNull(TBItem!Transdutor), "", TBItem!Transdutor)
End If
TBItem.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtDesenho_Change()
On Error GoTo tratar_erro

fotos = Localrel & "\imagens\caprind.bmp"
txtImagem = fotos
txtdescricao = ""
OLE1.CreateLink (fotos)

If txtdesenho = "" Then Exit Sub
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    If TBItem!imagem <> "" Then
        fotos = IIf(IsNull(TBItem!imagem), "", TBItem!imagem)
        txtImagem = fotos
        OLE1.CreateLink (fotos)
    End If
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

If txtData = "" Then TBGravar!Data = Date Else TBGravar!Data = txtData
If Txt_responsavel = "" Then TBGravar!Responsavel = pubUsuario Else TBGravar!Responsavel = Txt_responsavel
TBGravar!pedido_interno = txtPedido_interno
TBGravar!Pedido_cliente = txtPedido_cliente
TBGravar!IDCliente = IIf(txtid_cliente = "", 0, txtid_cliente)
TBGravar!Cliente = txtCliente
TBGravar!Desenho = txtdesenho
TBGravar!Metal_adicao = txtMetal_adicao
TBGravar!Metal_base = txtMetal_base
TBGravar!Espess = IIf(txtEspess = "", 0, txtEspess)
TBGravar!Qtde = IIf(txtQtde = "", 0, txtQtde)
TBGravar!Norma = txtNorma
TBGravar!Superficie = IIf(cmbSuperficie = "", Null, cmbSuperficie)
TBGravar!Calibracao = txtCalibracao
TBGravar!Obs = Trim(txtObs)
TBGravar!idInspetores = IIf(txtCodigo = "", Null, txtCodigo)
If optAprovado = True Then
    TBGravar!Conclusao = "Aprovado"
ElseIf optReprovado = True Then
        TBGravar!Conclusao = "Reprovado"
    Else
        TBGravar!Conclusao = ""
End If
If optA = True Then
    TBGravar!Criterio_aceitacao = "A"
ElseIf optB = True Then
        TBGravar!Criterio_aceitacao = "B"
    ElseIf optC = True Then
            TBGravar!Criterio_aceitacao = "C"
        ElseIf OptD = True Then
                TBGravar!Criterio_aceitacao = "D"
            Else
                TBGravar!Criterio_aceitacao = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

txtId = TBAbrir!ID
txtData = IIf(IsNull(TBAbrir!Data), "", Format(TBAbrir!Data, "dd/mm/yy"))
Txt_responsavel = IIf(IsNull(TBAbrir!Responsavel), "", TBAbrir!Responsavel)
txtPedido_interno = IIf(IsNull(TBAbrir!pedido_interno), "", TBAbrir!pedido_interno)
txtPedido_cliente = IIf(IsNull(TBAbrir!Pedido_cliente), "", TBAbrir!Pedido_cliente)
txtid_cliente = IIf(IsNull(TBAbrir!IDCliente), "", TBAbrir!IDCliente)
txtCliente = IIf(IsNull(TBAbrir!Cliente), "", TBAbrir!Cliente)
txtdesenho = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
txtMetal_adicao = IIf(IsNull(TBAbrir!Metal_adicao), "", TBAbrir!Metal_adicao)
txtMetal_base = IIf(IsNull(TBAbrir!Metal_base), "", TBAbrir!Metal_base)
txtEspess = IIf(IsNull(TBAbrir!Espess), "", Format(TBAbrir!Espess, "###,##0.00"))
txtQtde = IIf(IsNull(TBAbrir!Qtde), "", Format(TBAbrir!Qtde, "###,##0.0000"))
txtNorma = IIf(IsNull(TBAbrir!Norma), "", TBAbrir!Norma)
If IsNull(TBAbrir!Superficie) = False And TBAbrir!Superficie <> "" Then cmbSuperficie = TBAbrir!Superficie
txtCalibracao = IIf(IsNull(TBAbrir!Calibracao), "", TBAbrir!Calibracao)
txtObs = IIf(IsNull(TBAbrir!Obs), "", TBAbrir!Obs)
txtCodigo = IIf(IsNull(TBAbrir!idInspetores), "", TBAbrir!idInspetores)
Select Case TBAbrir!Conclusao
    Case "Aprovado": optAprovado.Value = True
    Case "Reprovado": optReprovado.Value = True
End Select
Select Case TBAbrir!Criterio_aceitacao
    Case "A": optA.Value = True
    Case "B": optB.Value = True
    Case "C": optC.Value = True
    Case "D": OptD.Value = True
End Select
Novo_Ultrasom = False
Frame2.Enabled = True
framePenetrante.Enabled = True
frameConclusao.Enabled = True
frameAceitacao.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEspess_Change()
On Error GoTo tratar_erro

If txtEspess.Text <> "" Then
    VerifNumero = txtEspess.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtEspess.Text = ""
        txtEspess.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEspess_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtEspess

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEspess_LostFocus()
On Error GoTo tratar_erro

txtEspess.Text = Format(txtEspess.Text, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtMetal_adicao_Change()
On Error GoTo tratar_erro

txtDescricao_metal = ""

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

Private Sub txtQtde_Change()
On Error GoTo tratar_erro

If txtQtde.Text <> "" Then
    VerifNumero = txtQtde.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde.Text = ""
        txtQtde.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtQtde

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtqtde_LostFocus()
On Error GoTo tratar_erro

txtQtde.Text = Format(txtQtde.Text, "###,##0.0000")
    
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

