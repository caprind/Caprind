VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRelatorios_Custos_Prev_Real_Just 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Custos - Centro de custo - Previsto x realizado - Justificativa"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15270
   ClipControls    =   0   'False
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtID 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2130
      TabIndex        =   37
      Text            =   "0"
      Top             =   6780
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtIDCCusto 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3210
      TabIndex        =   36
      Text            =   "0"
      Top             =   6780
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtIDCContabil 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4290
      TabIndex        =   35
      Text            =   "0"
      Top             =   6780
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtIDEmpresa 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5370
      TabIndex        =   34
      Text            =   "0"
      Top             =   6780
      Visible         =   0   'False
      Width           =   1065
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   11490
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmRelatorios_Custos_Prev_Real_Just.frx":0000
      Count           =   1
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   840
      Left            =   60
      TabIndex        =   16
      Top             =   990
      Width           =   15150
      Begin VB.TextBox txtEmpresa 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Empresa."
         Top             =   390
         Width           =   4125
      End
      Begin VB.TextBox txtContaContabil 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Conta contábil."
         Top             =   390
         Width           =   5715
      End
      Begin VB.TextBox txtCentroCusto 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Centro de custo."
         Top             =   390
         Width           =   4905
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
         Left            =   1875
         TabIndex        =   33
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conta contabil"
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
         TabIndex        =   18
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Centro de custo"
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
         Left            =   6105
         TabIndex        =   17
         Top             =   180
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4065
      Left            =   60
      TabIndex        =   15
      Top             =   5670
      Width           =   15150
      _ExtentX        =   26723
      _ExtentY        =   7170
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Empresa"
         Object.Width           =   3826
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Centro de custo"
         Object.Width           =   3826
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Conta contábil"
         Object.Width           =   4621
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "De"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Até"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Rev."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Vlr. previsto"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Vlr. realizado"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Vlr. variação"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "% variação"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Object.Tag             =   "S"
         Text            =   "Enviado"
         Object.Width           =   1587
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   840
      Left            =   60
      TabIndex        =   19
      Top             =   1830
      Width           =   15150
      Begin VB.TextBox txtRev 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Revisão prevista."
         Top             =   390
         Width           =   555
      End
      Begin VB.TextBox txtAteAno 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8340
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Ano até."
         Top             =   390
         Width           =   885
      End
      Begin VB.TextBox txtDeAno 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6540
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Ano de."
         Top             =   390
         Width           =   885
      End
      Begin VB.TextBox txtPorcentagemVariacao 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   13590
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Percentual da variação."
         Top             =   390
         Width           =   1365
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Data do cadastro."
         Top             =   390
         Width           =   1125
      End
      Begin VB.TextBox txtResponsavel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   390
         Width           =   4305
      End
      Begin VB.TextBox txtValorVariacao 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   12330
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Valor da variação."
         Top             =   390
         Width           =   1245
      End
      Begin VB.TextBox txtValorRealizado 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11070
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Valor realizado."
         Top             =   390
         Width           =   1245
      End
      Begin VB.TextBox txtValorPrevisto 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9810
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Valor previsto."
         Top             =   390
         Width           =   1245
      End
      Begin VB.TextBox txtDe 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Mês de."
         Top             =   390
         Width           =   885
      End
      Begin VB.TextBox txtAte 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Mês até."
         Top             =   390
         Width           =   885
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ano"
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
         Left            =   8640
         TabIndex        =   39
         Top             =   180
         Width           =   285
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ano"
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
         Left            =   6840
         TabIndex        =   38
         Top             =   180
         Width           =   285
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rev."
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
         Left            =   9345
         TabIndex        =   31
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% variação"
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
         Left            =   13860
         TabIndex        =   29
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label9 
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
         Left            =   570
         TabIndex        =   26
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label8 
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
         Left            =   3015
         TabIndex        =   25
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor variação"
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
         Left            =   12435
         TabIndex        =   24
         Top             =   180
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor realizado"
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
         Left            =   11160
         TabIndex        =   23
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor previsto"
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
         Left            =   9930
         TabIndex        =   22
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De"
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
         Left            =   5985
         TabIndex        =   21
         Top             =   180
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Até"
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
         Left            =   7755
         TabIndex        =   20
         Top             =   180
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   2990
      Left            =   60
      TabIndex        =   27
      Top             =   2670
      Width           =   15150
      Begin VB.TextBox txtJustificativa 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2475
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         ToolTipText     =   "Justificativa."
         Top             =   390
         Width           =   14805
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Justificativa"
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
         Left            =   6975
         TabIndex        =   28
         Top             =   180
         Width           =   1155
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   30
      Top             =   0
      Width           =   15150
      _ExtentX        =   26723
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
      ButtonCaption4  =   "Enviar"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Enviar justificativa por e-mail (F7]"
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
      ButtonWidth4    =   38
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonAlignment5=   2
      ButtonType5     =   1
      ButtonStyle5    =   -1
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   -1
      ButtonLeft5     =   176
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   180
      ButtonTop6      =   2
      ButtonWidth6    =   36
      ButtonHeight6   =   22
      ButtonUseMaskColor6=   0   'False
      ButtonCaption7  =   "Sair"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Sair (Esc)"
      ButtonKey7      =   "7"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft7     =   218
      ButtonTop7      =   2
      ButtonWidth7    =   27
      ButtonHeight7   =   22
      ButtonUseMaskColor7=   0   'False
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonAlignment8=   2
      BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState8    =   5
      ButtonLeft8     =   247
      ButtonTop8      =   2
      ButtonWidth8    =   24
      ButtonHeight8   =   24
      ButtonUseMaskColor8=   0   'False
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   32
      Top             =   9750
      Width           =   15150
      _ExtentX        =   26723
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
Attribute VB_Name = "frmRelatorios_Custos_Prev_Real_Just"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF7: ProcEnviarEmail
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 8, True
With frmRelatorios_Custos_Prev_Real
    txtIDEmpresa = .Cmb_empresa.ItemData(0)
    txtIDCCusto = .Lista_resumido.SelectedItem
    txtIDCContabil = .Lista_res_PC.SelectedItem
    txtCentroCusto = .cmbTexto.Text
    txtContaContabil = .Lista_res_PC.SelectedItem.SubItems(1)
    txtEmpresa = .Cmb_empresa
    txtDe = .Cmb_mes_de
    txtDeAno = .Cmb_ano_de
    txtAte = .Cmb_mes_ate
    txtAteAno = .Cmb_ano_ate
    txtRev = .Txt_rev_prev
    txtResponsavel = pubUsuario
    txtData = Format(Date, "dd/mm/yy")
    
    Texto1 = ""
    Texto = ""
    Numero = 0
    Numero1 = Len(.Lista_res_PC.SelectedItem.ListSubItems(3))
    If Numero1 <> 1 Then
        Do While Numero1 <> 0
            If Texto = "|" Then
                If txtValorPrevisto = "" Then
                    txtValorPrevisto = Texto1
                ElseIf txtValorRealizado = "" Then
                        txtValorRealizado = Texto1
                    ElseIf txtValorVariacao = "" Then
                            txtValorVariacao = Texto1
                        ElseIf txtPorcentagemVariacao = "" Then
                                txtPorcentagemVariacao = Texto1
                End If
                Texto1 = ""
            Else
                Texto1 = Trim(Texto1 & Texto)
            End If
            Texto = Left(.Lista_res_PC.SelectedItem.ListSubItems(3), (Numero + 1))
            Texto = Right(Texto, Len(Texto) - Numero)
            Numero = Numero + 1
            Numero1 = Numero1 - 1
        Loop
        If Numero1 = 0 Then txtPorcentagemVariacao = Texto1 & "%"
    End If
End With

ProcLimpaCampos
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select ID, Justificativa from CustosJustificativa where IDEmpresa = " & txtIDEmpresa & " and IDCCusto = " & txtIDCCusto & " and IDCContabil = " & txtIDCContabil & " and PInicio = '" & txtDe & "' and AnoInicio = " & txtDeAno & " and PFim = '" & txtAte & "' and AnoFim = " & txtAteAno & " and Rev = '" & txtRev & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtId = TBLISTA!ID
    txtJustificativa = TBLISTA!Justificativa
End If
TBLISTA.Close
ProcCarregaLista

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
If txtJustificativa = "" Then
    NomeCampo = "a justificativa"
    ProcVerificaAcao
    If Frame4.Enabled = True Then txtJustificativa.SetFocus
    Exit Sub
End If
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from CustosJustificativa where ID = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = True Then
    TBLISTA.AddNew
    USMsgBox ("Nova justificativa cadastrada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
End If
TBLISTA!IDempresa = txtIDEmpresa
TBLISTA!IDCCusto = txtIDCCusto
TBLISTA!IDCContabil = txtIDCContabil
TBLISTA!Data = txtData
TBLISTA!Responsavel = txtResponsavel
TBLISTA!PInicio = txtDe
TBLISTA!AnoInicio = txtDeAno
TBLISTA!PFim = txtAte
TBLISTA!AnoFim = txtAteAno
TBLISTA!Rev = txtRev
TBLISTA!Previsto = txtValorPrevisto
TBLISTA!Realizado = txtValorRealizado
TBLISTA!Variacao = txtValorVariacao
TBLISTA!PercVariacao = txtPorcentagemVariacao
TBLISTA!Justificativa = txtJustificativa
TBLISTA.Update
txtId = TBLISTA!ID
TBLISTA.Close
'==================================
Modulo = "Custos/Centro de custo/Justificativa"
ID_documento = txtId
Documento = "Centro de custo: " & txtCentroCusto & " - Conta contábil: " & txtContaContabil
Documento1 = ""
ProcGravaEvento
'==================================
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " voce não tem acesso a este recurso.")
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir esta(s) justificativa(s)?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            '==================================
            Modulo = "Custos/Centro de custo/Justificativa"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            Documento = "Centro de custo: " & .ListItems(InitFor).ListSubItems(2) & " - Conta contábil: " & .ListItems(InitFor).ListSubItems(3)
            Documento1 = ""
            ProcGravaEvento
            '==================================
            Conexao.Execute "DELETE from CustosJustificativa where id = " & .ListItems(InitFor)
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe a(s) justificativa(s) antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Justificativa(s) excluída(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtData = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
txtJustificativa.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviarEmail()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then
    USMsgBox ("Não é possivel enviar o email pois não existe justificativa cadastrada."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
Compras_Pedido = False
Custos_justificativa = True
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Caminho from Empresa_armazenamento_PDF where ID_empresa = " & txtIDEmpresa & " and Relatorio = 'Justificativa centro de custo'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    NomeRel = "Custos_previsto_realizado_justificativa.rpt"
    If Len(TBAbrir!caminho) = 3 Then caminho = TBAbrir!caminho Else caminho = TBAbrir!caminho & "\"
    Nome_anexo = Replace(txtCentroCusto, "/", "_") & " - " & txtDe & "-" & txtDeAno & " Até " & txtAte & "-" & txtAteAno & ".pdf"
    ProcGerarPDF caminho & Nome_anexo, "{CustosJustificativa.IDCCusto} = " & txtIDCCusto & " and {CustosJustificativa.PInicio} = '" & txtDe & "' and {CustosJustificativa.AnoInicio} = " & txtDeAno & " and {CustosJustificativa.PFim} = '" & txtAte & "' and {CustosJustificativa.AnoFim} = " & txtAteAno, ""
    
    FrmEnviarEmail.Txt_anexo = caminho & Nome_anexo
End If
TBAbrir.Close
FrmEnviarEmail.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CJ.ID, EMP.Empresa, US.Codigo as CCUSTO, US.Setor, FAM.Codigo as CCONTABIL, FAM.txt_descricao, CJ.PInicio, CJ.AnoInicio, CJ.PFim, CJ.AnoFim, CJ.Rev, CJ.Previsto, CJ.Realizado, CJ.Variacao, CJ.PercVariacao, CJ.Email_enviado from (((CustosJustificativa CJ INNER JOIN Usuarios_Setor US ON CJ.IDCCusto = US.ID) INNER JOIN Empresa EMP ON CJ.IDEmpresa = EMP.Codigo) INNER JOIN tbl_Familia FAM ON CJ.IDCContabil = FAM.int_codfamilia) where CJ.IDCCusto = " & txtIDCCusto & " and PInicio = '" & txtDe & "' and AnoInicio = " & txtDeAno & " and PFim = '" & txtAte & "' and AnoFim = '" & txtAteAno & "' order by CJ.ID", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = TBLISTA!Empresa
            .Item(.Count).SubItems(2) = TBLISTA!CCUSTO & " - " & TBLISTA!Setor
            .Item(.Count).SubItems(3) = TBLISTA!CCONTABIL & " - " & TBLISTA!Txt_descricao
            .Item(.Count).SubItems(4) = TBLISTA!PInicio & "/" & TBLISTA!AnoInicio
            .Item(.Count).SubItems(5) = TBLISTA!PFim & "/" & TBLISTA!AnoFim
            .Item(.Count).SubItems(6) = TBLISTA!Rev
            .Item(.Count).SubItems(7) = Format(TBLISTA!Previsto, "###,##0.00")
            .Item(.Count).SubItems(8) = Format(TBLISTA!Realizado, "###,##0.00")
            .Item(.Count).SubItems(9) = Format(TBLISTA!Variacao, "###,##0.00")
            .Item(.Count).SubItems(10) = TBLISTA!PercVariacao
            .Item(.Count).SubItems(11) = IIf(TBLISTA!Email_Enviado = True, "Sim", "Não")
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPuxaDados()
On Error GoTo tratar_erro

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from CustosJustificativa where ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtId = TBLISTA!ID
    txtIDEmpresa = TBLISTA!IDempresa
    txtIDCCusto = TBLISTA!IDCCusto
    txtIDCContabil = TBLISTA!IDCContabil
    txtDe = TBLISTA!PInicio
    txtDeAno = TBLISTA!AnoInicio
    txtAte = TBLISTA!PFim
    txtAteAno = TBLISTA!AnoFim
    txtRev = TBLISTA!Rev
    txtResponsavel = TBLISTA!Responsavel
    txtData = Format(TBLISTA!Data, "dd/mm/yy")
    txtValorPrevisto = Format(TBLISTA!Previsto, "###,##0.00")
    txtValorRealizado = Format(TBLISTA!Realizado, "###,##0.00")
    txtValorVariacao = Format(TBLISTA!Variacao, "###,##0.00")
    txtPorcentagemVariacao = TBLISTA!PercVariacao
    txtJustificativa = TBLISTA!Justificativa
End If
TBLISTA.Close

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

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

ProcPuxaDados

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    Case 2: ProcExcluir
    Case 3: ProcImprimir
    Case 4: ProcEnviarEmail
    'Case 6: ProcAjuda
    Case 7: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
NomeRel = "Custos_previsto_realizado_justificativa.rpt"
ProcImprimirRel "{CustosJustificativa.IDCCusto} = " & txtIDCCusto & " and {CustosJustificativa.PInicio} = '" & txtDe & "' and {CustosJustificativa.AnoInicio} = " & txtDeAno & " and {CustosJustificativa.PFim} = '" & txtAte & "' and {CustosJustificativa.AnoFim} = " & txtAteAno, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
