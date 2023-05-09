VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmPCP_relatorios_resultados 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PCP - Relatórios - Resultados da ordem"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
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
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
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
   Begin VB.CheckBox chkSemEscopo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fora do escopo"
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
      Height          =   210
      Left            =   3915
      TabIndex        =   13
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox chkEscopo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Escopo"
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
      Height          =   210
      Left            =   2910
      TabIndex        =   12
      Top             =   1080
      Value           =   1  'Checked
      Width           =   915
   End
   Begin VB.CheckBox chkOS 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Carregar lista por OS"
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
      Height          =   210
      Left            =   5580
      TabIndex        =   14
      Top             =   1080
      Width           =   2025
   End
   Begin VB.CheckBox Chk_nao_concluida 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Não concluída"
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
      Height          =   210
      Left            =   180
      TabIndex        =   10
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1425
   End
   Begin VB.CheckBox Chk_concluida 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Concluída"
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
      Height          =   210
      Left            =   1695
      TabIndex        =   11
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1125
   End
   Begin VB.OptionButton optEmissao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Emissão"
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
      Left            =   12390
      TabIndex        =   16
      Top             =   1080
      Width           =   885
   End
   Begin VB.OptionButton optConclusao 
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
      Height          =   195
      Left            =   13320
      TabIndex        =   17
      Top             =   1080
      Width           =   1035
   End
   Begin VB.OptionButton optPrazo 
      BackColor       =   &H00E0E0E0&
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
      Height          =   195
      Left            =   11250
      TabIndex        =   15
      Top             =   1080
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Frame Frame1 
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
      Height          =   1485
      Left            =   60
      TabIndex        =   45
      Top             =   8250
      Width           =   15195
      Begin VB.TextBox Txt_outras 
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
         Left            =   6450
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Custo total de outras despesas."
         Top             =   1020
         Width           =   1520
      End
      Begin VB.TextBox Txt_eficiencia_exec 
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
         Left            =   12375
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Eficiência de execução."
         Top             =   1020
         Width           =   1290
      End
      Begin VB.TextBox Txt_eficiencia_prep 
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
         Left            =   11070
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Eficiência de preparação."
         Top             =   1020
         Width           =   1290
      End
      Begin VB.TextBox txtEficiencia 
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
         Left            =   13680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Eficiência média."
         Top             =   1020
         Width           =   1320
      End
      Begin VB.TextBox txtTotal_peca 
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
         Left            =   9590
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Custo total por peça."
         Top             =   1020
         Width           =   1470
      End
      Begin VB.TextBox txtReal_lote 
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
         Left            =   11070
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Tempo total real por lote."
         Top             =   390
         Width           =   1290
      End
      Begin VB.TextBox txtTotal 
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
         Left            =   7980
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Custo total."
         Top             =   1020
         Width           =   1590
      End
      Begin VB.TextBox txtPrevisto_lote 
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
         Left            =   9590
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Tempo total previsto por lote."
         Top             =   390
         Width           =   1470
      End
      Begin VB.TextBox txtTerceiros 
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
         Left            =   4965
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Custo total de terceiros."
         Top             =   1020
         Width           =   1470
      End
      Begin VB.TextBox txtMaterial 
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
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Custo total de material."
         Top             =   1020
         Width           =   1470
      End
      Begin VB.TextBox txtReal_peca 
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
         Left            =   7980
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Tempo total real por peça."
         Top             =   390
         Width           =   1590
      End
      Begin VB.TextBox txtPrev_peca 
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
         Left            =   6450
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Tempo total previsto por peça."
         Top             =   390
         Width           =   1520
      End
      Begin VB.TextBox txtValor_ref 
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
         Left            =   4965
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Valor total de refugo."
         Top             =   390
         Width           =   1470
      End
      Begin VB.TextBox txtQtde_ref 
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
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total refugada."
         Top             =   390
         Width           =   1470
      End
      Begin VB.TextBox txtQtde_prod 
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
         Left            =   1830
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total produzida."
         Top             =   390
         Width           =   1630
      End
      Begin VB.TextBox txtQtde_prev 
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
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total prevista."
         Top             =   390
         Width           =   1630
      End
      Begin VB.TextBox txtMO_realLote 
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
         Left            =   1830
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Custo total de mão de obra real por lote."
         Top             =   1020
         Width           =   1630
      End
      Begin VB.TextBox txtMO_prevLote 
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
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Custo total de mão de obra previsto por lote."
         Top             =   1020
         Width           =   1630
      End
      Begin VB.TextBox txtMO_realPeca 
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
         Left            =   13680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Custo total de mão de obra real por peça."
         Top             =   390
         Width           =   1320
      End
      Begin VB.TextBox txtMO_prevPeca 
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
         Left            =   12375
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Custo total de mão de obra previsto por peça."
         Top             =   390
         Width           =   1290
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CT. outras desp."
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
         Left            =   6540
         TabIndex        =   70
         Top             =   810
         Width           =   1350
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Efic. média"
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
         Left            =   13890
         TabIndex        =   69
         Top             =   810
         Width           =   900
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Efic. exec."
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
         Left            =   12608
         TabIndex        =   68
         Top             =   810
         Width           =   825
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Efic. prep."
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
         Left            =   11310
         TabIndex        =   63
         Top             =   810
         Width           =   810
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CT. total peça"
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
         Left            =   9748
         TabIndex        =   62
         Top             =   810
         Width           =   1155
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custo total"
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
         Left            =   8310
         TabIndex        =   61
         Top             =   810
         Width           =   930
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CT. terceiros"
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
         Left            =   5190
         TabIndex        =   60
         Top             =   810
         Width           =   1065
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CT. material"
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
         Left            =   3705
         TabIndex        =   59
         Top             =   810
         Width           =   1020
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CT. MO real lote"
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
         Left            =   1980
         TabIndex        =   58
         Top             =   810
         Width           =   1320
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CT. MO prev. lote"
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
         Left            =   283
         TabIndex        =   57
         Top             =   810
         Width           =   1425
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CT. MO real pç"
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
         Left            =   13748
         TabIndex        =   56
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CT. MO pre. pç"
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
         Left            =   12428
         TabIndex        =   55
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TT. real lote"
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
         Left            =   11213
         TabIndex        =   54
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TT. prev. lote"
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
         Left            =   9770
         TabIndex        =   53
         Top             =   180
         Width           =   1110
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TT. real peça"
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
         Left            =   8235
         TabIndex        =   52
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TT. prev. peça"
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
         Left            =   6623
         TabIndex        =   51
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor refugo"
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
         Left            =   5183
         TabIndex        =   50
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. refugada"
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
         Left            =   3585
         TabIndex        =   49
         Top             =   180
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. produzida"
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
         Left            =   1973
         TabIndex        =   48
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. prevista"
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
         Left            =   390
         TabIndex        =   46
         Top             =   180
         Width           =   1200
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   1035
      Left            =   13285
      TabIndex        =   38
      Top             =   1290
      Width           =   1965
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   570
         TabIndex        =   8
         ToolTipText     =   "Data final."
         Top             =   570
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
         Format          =   489422849
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   570
         TabIndex        =   7
         ToolTipText     =   "Data inicio."
         Top             =   210
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
         Format          =   489422849
         CurrentDate     =   39057
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "De :"
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
         Left            =   240
         TabIndex        =   40
         Top             =   270
         Width           =   300
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Até :"
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
         TabIndex        =   39
         Top             =   630
         Width           =   360
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5865
      Left            =   60
      TabIndex        =   9
      Top             =   2340
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   10345
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
      NumItems        =   28
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Ordem"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "OS"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Cód. de ref."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Qtde. prevista"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Qtde. produzida"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Qtde. refugada"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Valor refugo"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Object.Tag             =   "D"
         Text            =   "TT. prev. peça"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Object.Tag             =   "D"
         Text            =   "TT. real peça"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Object.Tag             =   "D"
         Text            =   "TT. prev. lote"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Object.Tag             =   "D"
         Text            =   "TT. real lote"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Object.Tag             =   "N"
         Text            =   "CT. MO prev. peça"
         Object.Width           =   2734
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Object.Tag             =   "N"
         Text            =   "CT. MO real peça"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Object.Tag             =   "N"
         Text            =   "CT. MO prev. lote"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Object.Tag             =   "N"
         Text            =   "CT. MO real lote"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Object.Tag             =   "N"
         Text            =   "CT. material"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Object.Tag             =   "N"
         Text            =   "CT. terceiros"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Object.Tag             =   "N"
         Text            =   "CT. outras desp."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   22
         Object.Tag             =   "N"
         Text            =   "CT. total"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Object.Tag             =   "N"
         Text            =   "CT. total peça"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   24
         Object.Tag             =   "N"
         Text            =   "Efic. prep."
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   25
         Object.Tag             =   "N"
         Text            =   "Efic. exec."
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   26
         Object.Tag             =   "N"
         Text            =   "Efic. média"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   27
         Object.Tag             =   "T"
         Text            =   "Res. validado"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.Frame Frame4 
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
      Height          =   1035
      Left            =   3180
      TabIndex        =   41
      Top             =   1290
      Width           =   10095
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
         ItemData        =   "frmPCP_relatorios_resultados.frx":0000
         Left            =   180
         List            =   "frmPCP_relatorios_resultados.frx":0019
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Opções para filtro."
         Top             =   480
         Width           =   2085
      End
      Begin VB.TextBox txtTexto 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         ToolTipText     =   "Texto para pesquisa."
         Top             =   480
         Width           =   7605
      End
      Begin VB.ComboBox cmbTexto 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         ItemData        =   "frmPCP_relatorios_resultados.frx":0076
         Left            =   2280
         List            =   "frmPCP_relatorios_resultados.frx":0078
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Texto para pesquisa."
         Top             =   480
         Visible         =   0   'False
         Width           =   7605
      End
      Begin VB.Label Label8 
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
         Left            =   802
         TabIndex        =   43
         Top             =   270
         Width           =   840
      End
      Begin VB.Label Label9 
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
         Left            =   5347
         TabIndex        =   42
         Top             =   270
         Width           =   1470
      End
   End
   Begin VB.Frame Frame7 
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
      Height          =   1035
      Left            =   55
      TabIndex        =   47
      Top             =   1290
      Width           =   1635
      Begin VB.OptionButton Opt_individual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Individual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   0
         Top             =   330
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton Opt_comparativo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comparativo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   600
         Width           =   1425
      End
   End
   Begin VB.Frame Frame5 
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
      Height          =   1035
      Left            =   1710
      TabIndex        =   44
      Top             =   1290
      Width           =   1455
      Begin VB.OptionButton optDetalhado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detalhado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   330
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optResumido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Resumido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   600
         Width           =   1155
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   65
      Top             =   9750
      Width           =   11715
      _ExtentX        =   20664
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   67
      Top             =   0
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
      ButtonCaption2  =   "Relatório"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Relatório (F5)"
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
      ButtonWidth2    =   51
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
      ButtonLeft3     =   93
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
      ButtonLeft4     =   97
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
      ButtonLeft5     =   135
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
      ButtonLeft6     =   163
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   13620
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmPCP_relatorios_resultados.frx":007A
         Count           =   1
      End
   End
   Begin VB.Label Lbl_relatorio 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registros encontrados: 0000 - 00:00:00"
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
      Left            =   11880
      TabIndex        =   66
      Top             =   9780
      Width           =   3315
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Período por :"
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
      Left            =   10050
      TabIndex        =   64
      Top             =   1050
      Width           =   1065
   End
End
Attribute VB_Name = "frmPCP_relatorios_resultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Chk_concluida_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_nao_concluida_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkOS_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
ProcLimpaCamposTotais
If chkOS.Value = 1 Then Lista.ColumnHeaders(4).Width = 1200 Else Lista.ColumnHeaders(4).Width = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
frmPCP_relatorios_resultados_menuimpressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcAbrir
    Case vbKeyF5: ProcImprimir
    'Case vbKeyF1: ProcAjudar
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Posicao = 0
Lista.ListItems.Clear
If TBLISTA.EOF = False Then
    Posicao = TBLISTA.RecordCount
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Ordem), "", TBLISTA!Ordem)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!OS), "", TBLISTA!OS)
            
            If cmbfiltrarpor = "Ordem" Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Desenho, N_Referencia, Produto from producao where Ordem = " & TBLISTA!maquina, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!Desenho), "", TBAbrir!Desenho)
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBAbrir!N_referencia), "", TBAbrir!N_referencia)
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBAbrir!Produto), "", TBAbrir!Produto)
                End If
                TBAbrir.Close
            End If
            
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!QtdePrev), "0,00", Format(TBLISTA!QtdePrev, "###,##0.00"))
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!qtdeOK), "0,00", Format(TBLISTA!qtdeOK, "###,##0.00"))
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!qtdeNC), "0,00", Format(TBLISTA!qtdeNC, "###,##0.00"))
            .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!Refugo), "0,00", Format(TBLISTA!Refugo, "###,##0.00"))
            .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!Data1), "00:00:00", TBLISTA!Data1)
            .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA!Data2), "00:00:00", TBLISTA!Data2)
            .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA!Data3), "00:00:00", TBLISTA!Data3)
            .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA!Data4), "00:00:00", TBLISTA!Data4)
            .Item(.Count).SubItems(15) = IIf(IsNull(TBLISTA!Qtdetotalprod), "0,00000", Format(TBLISTA!Qtdetotalprod, "###,##0.0000000000"))
            .Item(.Count).SubItems(16) = IIf(IsNull(TBLISTA!Terceiros), "0,00000", Format(TBLISTA!Terceiros, "###,##0.0000000000"))
            .Item(.Count).SubItems(17) = IIf(IsNull(TBLISTA!impostos), "0,00000", Format(TBLISTA!impostos, "###,##0.0000000000"))
            .Item(.Count).SubItems(18) = IIf(IsNull(TBLISTA!Lucro), "0,00", Format(TBLISTA!Lucro, "###,##0.00"))
            .Item(.Count).SubItems(19) = IIf(IsNull(TBLISTA!material), "0,00", Format(TBLISTA!material, "###,##0.00"))
            .Item(.Count).SubItems(20) = IIf(IsNull(TBLISTA!Servicos), "0,00", Format(TBLISTA!Servicos, "###,##0.00"))
            .Item(.Count).SubItems(21) = IIf(IsNull(TBLISTA!Numero4), "0,00", Format(TBLISTA!Numero4, "###,##0.00"))
            .Item(.Count).SubItems(22) = IIf(IsNull(TBLISTA!Total), "0,00", Format(TBLISTA!Total, "###,##0.00"))
            .Item(.Count).SubItems(23) = IIf(IsNull(TBLISTA!Total_peca), "0,00000", Format(TBLISTA!Total_peca, "###,##0.0000000000"))
            .Item(.Count).SubItems(24) = IIf(IsNull(TBLISTA!Numero1), "0,00%", Format(TBLISTA!Numero1, "###,##0.00") & "%")
            .Item(.Count).SubItems(25) = IIf(IsNull(TBLISTA!Numero2), "0,00%", Format(TBLISTA!Numero2, "###,##0.00") & "%")
            .Item(.Count).SubItems(26) = IIf(IsNull(TBLISTA!Eficiencia), "0,00%", Format(TBLISTA!Eficiencia, "###,##0.00") & "%")
            .Item(.Count).SubItems(27) = IIf(TBLISTA!Numero3 = 0, "Não", "Sim")
        End With
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
    
End If
TBLISTA.Close

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtQtde_prev = IIf(IsNull(TBLISTA!QtdePrevista), "0,00", Format(TBLISTA!QtdePrevista, "###,##0.00"))
    txtQtde_prod = IIf(IsNull(TBLISTA!QtdeProduzida), "0,00", Format(TBLISTA!QtdeProduzida, "###,##0.00"))
    txtQtde_ref = IIf(IsNull(TBLISTA!qtdeNC), "0,00", Format(TBLISTA!qtdeNC, "###,##0.00"))
    txtValor_ref = IIf(IsNull(TBLISTA!Lucro), "0,00", Format(TBLISTA!Lucro, "###,##0.00"))
    txtPrev_peca = IIf(IsNull(TBLISTA!Data1), "00:00:00", TBLISTA!Data1)
    txtReal_peca = IIf(IsNull(TBLISTA!Data2), "00:00:00", TBLISTA!Data2)
    txtPrevisto_lote = IIf(IsNull(TBLISTA!Data3), "00:00:00", TBLISTA!Data3)
    txtReal_lote = IIf(IsNull(TBLISTA!Data4), "00:00:00", TBLISTA!Data4)
    txtMO_prevPeca = IIf(IsNull(TBLISTA!CustoMat), "0,00000", Format(TBLISTA!CustoMat, "###,##0.0000000000"))
    txtMO_realPeca = IIf(IsNull(TBLISTA!Terceros), "0,00000", Format(TBLISTA!Terceros, "###,##0.0000000000"))
    txtMO_prevLote = IIf(IsNull(TBLISTA!CustoObra), "0,00", Format(TBLISTA!CustoObra, "###,##0.00"))
    txtMO_realLote = IIf(IsNull(TBLISTA!Valor1), "0,00", Format(TBLISTA!Valor1, "###,##0.00"))
    txtMaterial = IIf(IsNull(TBLISTA!Valor2), "0,00", Format(TBLISTA!Valor2, "###,##0.00"))
    txtTerceiros = IIf(IsNull(TBLISTA!Valor3), "0,00", Format(TBLISTA!Valor3, "###,##0.00"))
    Txt_outras = IIf(IsNull(TBLISTA!Numero4), "0,00", Format(TBLISTA!Numero4, "###,##0.00"))
    txtTotal = IIf(IsNull(TBLISTA!Total1), "0,00", Format(TBLISTA!Total1, "###,##0.00"))
    txtTotal_peca = IIf(IsNull(TBLISTA!Total2), "0,00000", Format(TBLISTA!Total2, "###,##0.0000000000"))
    Txt_eficiencia_prep = IIf(IsNull(TBLISTA!Numero1), "0,00%", Format(TBLISTA!Numero1, "###,##0.00") & "%")
    Txt_eficiencia_exec = IIf(IsNull(TBLISTA!Numero2), "0,00%", Format(TBLISTA!Numero2, "###,##0.00") & "%")
    txtEficiencia = IIf(IsNull(TBLISTA!QtdeOrdem), "0,00%", Format(TBLISTA!QtdeOrdem, "###,##0.00") & "%")
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposTotais()
On Error GoTo tratar_erro

Lbl_relatorio.Caption = "Registros encontrados: 0000 - 00:00:00"
txtQtde_prev = ""
txtQtde_prod = ""
txtQtde_ref = ""
txtValor_ref = ""
txtPrev_peca = ""
txtReal_peca = ""
txtPrevisto_lote = ""
txtReal_lote = ""
Txt_eficiencia_prep = ""
Txt_eficiencia_exec = ""
txtEficiencia = ""
txtMO_prevPeca = ""
txtMO_realPeca = ""
txtMO_prevLote = ""
txtMO_realLote = ""
txtMaterial = ""
txtTerceiros = ""
Txt_outras = ""
txtTotal = ""
txtTotal_peca = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaVariaveis()
On Error GoTo tratar_erro

TotalGeral = 0
Valor_Cofins_Prod = 0
Valor_Cofins_Serv = 0
Valor_CSLL_Prod = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 5, True
Formulario = "PCP/Relatórios/Resultados da ordem"
Direitos
ProcLimpaVariaveisPrincipais
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
cmbfiltrarpor.Text = "Ordem"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "PCP/Relatórios/Índice de atraso"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboTexto()
On Error GoTo tratar_erro

Lista.ListItems.Clear
ProcLimpaCamposTotais
If cmbfiltrarpor = "Ordem" Then ProcListaOrdem Else ProcListaPadrao

With cmbTexto
    If Opt_individual.Value = True Then
        .Clear
        .Visible = True
        txtTexto = ""
        txtTexto.Visible = False
        
        If cmbfiltrarpor = "Código interno" Or cmbfiltrarpor = "Código de referência" Or cmbfiltrarpor = "Descrição" Then
            Select Case cmbfiltrarpor
                Case "Código interno": NomeCampo = "Desenho"
                Case "Código de referência": NomeCampo = "n_referencia"
                Case "Descrição": NomeCampo = "Produto"
            End Select
            
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select " & NomeCampo & " as NomeCampo1 from producao where " & NomeCampo & " is not null Group by " & NomeCampo, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    If TBAbrir!NomeCampo1 <> "" Then .AddItem TBAbrir!NomeCampo1
                    TBAbrir.MoveNext
                Loop
            End If
            TBAbrir.Close
        End If
        
        Select Case cmbfiltrarpor
            Case "Tipo da ordem":
                .AddItem "Produto final"
                .AddItem "Subconjunto"
                .AddItem "Componente"
            Case "Família":
                Set TBProduto = CreateObject("adodb.recordset")
                TBProduto.Open "select projproduto.Classe from projproduto INNER JOIN producao on projproduto.desenho = producao.desenho Group by projproduto.classe", Conexao, adOpenKeyset, adLockOptimistic
                Do While TBProduto.EOF = False
                    If TBProduto!Classe <> "" Then .AddItem TBProduto!Classe
                    TBProduto.MoveNext
                Loop
                TBProduto.Close
            Case "Cliente":
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Cliente from producao where Cliente <> 'Null' Group by cliente", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Do While TBAbrir.EOF = False
                        If TBAbrir!Cliente <> "" Then .AddItem TBAbrir!Cliente
                        TBAbrir.MoveNext
                    Loop
                End If
                TBAbrir.Close
            Case "Ordem":
                .Visible = False
                txtTexto.Visible = True
        End Select
    End If
End With
Lista.ColumnHeaders(2).Text = cmbfiltrarpor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcListaPadrao()
On Error GoTo tratar_erro

With Lista.ColumnHeaders
    .Item(2).Width = 3500
    If optDetalhado.Value = True Then
        .Item(3).Width = 1200
        .Item(27).Width = 1400
    Else
        .Item(3).Width = 0
        .Item(27).Width = 0
    End If
    .Item(5).Width = 0
    .Item(6).Width = 0
    .Item(7).Width = 0
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcListaOrdem()
On Error GoTo tratar_erro

With Lista.ColumnHeaders
    .Item(2).Width = 1200
    .Item(3).Width = 0
    .Item(5).Width = 1200
    .Item(6).Width = 1400
    .Item(7).Width = 2946
    .Item(27).Width = 1400
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrir()
On Error GoTo tratar_erro

Acao = "filtrar"
If Chk_nao_concluida.Value = 0 And Chk_concluida.Value = 0 Then
    NomeCampo = "o status"
    ProcVerificaAcao
    Exit Sub
End If
If Opt_individual.Value = True Then
'    If cmbTexto.Visible = True And cmbTexto = "" Then
'        NomeCampo = "o texto para pesquisa"
'        ProcVerificaAcao
'        cmbTexto.SetFocus
'        Exit Sub
'    End If
'    If txtTexto.Visible = True And txtTexto = "" Then
'        NomeCampo = "o texto para pesquisa"
'        ProcVerificaAcao
'        txtTexto.SetFocus
'        Exit Sub
'    End If
End If
With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
'If chkEscopo.Value = 0 And chkSemEscopo.Value = 0 Then
'    NomeCampo = "uma das opções de escopo"
'    ProcVerificaAcao
'    Exit Sub
'End If

Inicio = Time
ProcLimpaCamposTotais
ProcLimpaVariaveis
ProcAbrirTabelas
If Permitido = True Then ProcGravarTotalizacoes
Set TBLISTA = CreateObject("adodb.recordset")
If Opt_individual.Value = True And optDetalhado.Value = True Then
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by Data, Maquina", Conexao, adOpenKeyset, adLockOptimistic
Else
    If cmbfiltrarpor = "Ordem" Then Ordenar = "Ordem" Else Ordenar = "Maquina"
    TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by " & Ordenar, Conexao, adOpenKeyset, adLockOptimistic
End If
ProcCarregaLista

intervalo = Time
ElapsedTime (intervalo - Inicio)
Lbl_relatorio.Caption = "Registros encontrados: " & FunTamanhoTextoZeroEsq(Posicao, 4) & " - " & HoraTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirTabelas()
On Error GoTo tratar_erro

'Deleta registros e adiciona novos
ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal

Select Case cmbTexto
    Case "Produto final": Tipo = "E"
    Case "Subconjunto": Tipo = "M"
    Case "Componente": Tipo = "F"
End Select
If Chk_nao_concluida.Value = 0 And Chk_concluida.Value = 1 Then
    Concluida = "PR.Concluida = 'True'"
ElseIf Chk_nao_concluida.Value = 1 And Chk_concluida.Value = 0 Then
        Concluida = "PR.Concluida = 'False'"
    Else
        Concluida = "(PR.Concluida = 'True' or PR.Concluida = 'False')"
End If
If optPrazo.Value = True Then DataFiltro = "PR.prazoentrega"
If optEmissao.Value = True Then DataFiltro = "PR.data"
If optConclusao.Value = True Then DataFiltro = "PR.dataentrega"

'If chkEscopo.Value = 0 And chkSemEscopo.Value = 1 Then
'    FiltroEscopo = " and (PR.Escopo = 'False' or PR.Escopo IS NULL)"
'ElseIf chkEscopo.Value = 1 And chkSemEscopo.Value = 0 Then
'        FiltroEscopo = " and PR.Escopo = 'True'"
'    Else
'        FiltroEscopo = ""
'End If

If chkOS.Value = 1 Then
    If cmbfiltrarpor = "Família" Then
        INNERJOINTEXTO = "PR.CTMaterial, PR.CTOutras, PR.consignacao, PR.DtValidacao_custo, P.classe, PR.Tipo, OS.* from (producao PR INNER JOIN projproduto P ON PR.desenho = P.Desenho) INNER JOIN Ordemservico OS on PR.Ordem = OS.Ordem"
    ElseIf cmbfiltrarpor = "Pedido interno" Then
            INNERJOINTEXTO = "PR.CTMaterial, PR.CTOutras, PR.consignacao, PR.DtValidacao_custo, VP.Ncotacao, PR.Tipo, OS.* from (((producao PR INNER JOIN Producao_pedidos PP ON PR.ordem = PP.ordem) INNER JOIN vendas_carteira VC ON VC.Codigo = PP.IDcarteira) INNER JOIN vendas_proposta VP ON VP.Cotacao = VC.Cotacao) INNER JOIN Ordemservico OS on PR.Ordem = OS.Ordem"
        Else
            INNERJOINTEXTO = "PR.CTMaterial, PR.CTOutras, PR.consignacao, PR.DtValidacao_custo, PR.Tipo, OS.* from producao PR INNER JOIN Ordemservico OS on PR.Ordem = OS.Ordem"
    End If
Else
    If cmbfiltrarpor = "Família" Then
        INNERJOINTEXTO = "PR.*, P.classe from producao PR INNER JOIN projproduto P ON PR.desenho = P.Desenho"
    ElseIf cmbfiltrarpor = "Pedido interno" Then
            INNERJOINTEXTO = "PR.*, VP.Ncotacao from ((producao PR INNER JOIN Producao_pedidos PP ON PR.ordem = PP.ordem) INNER JOIN vendas_carteira VC ON VC.Codigo = PP.IDcarteira) INNER JOIN vendas_proposta VP ON VP.Cotacao = VC.Cotacao"
        Else
            INNERJOINTEXTO = "* from producao PR"
    End If
End If

Select Case cmbfiltrarpor
    Case "Código interno": TextoFiltro = "PR.desenho"
    Case "Código de referência": TextoFiltro = "PR.N_Referencia"
    Case "Descrição": TextoFiltro = "PR.Produto"
    Case "Família": TextoFiltro = "P.Classe"
    Case "Cliente": TextoFiltro = "PR.Cliente"
    Case "Pedido interno": TextoFiltro = "VP.Ncotacao"
    Case "Tipo da ordem": TextoFiltro = "PR.tipo"
    Case "Ordem": TextoFiltro = "PR.Ordem"
End Select
If Opt_individual.Value = True Then
    If txtTexto.Visible = True And txtTexto <> "" Or cmbTexto.Visible = True And cmbTexto <> "" Then
        If cmbfiltrarpor = "Ordem" Then
            TextoFiltro1 = TextoFiltro & " = " & txtTexto & " and"
        Else
            TextoFiltro1 = TextoFiltro & " = '" & IIf(cmbfiltrarpor = "Tipo da ordem", Tipo, cmbTexto) & "' and"
        End If
    Else
        TextoFiltro1 = ""
    End If
    If chkOS.Value = 1 Then Ordenar = "PR.Ordem, OS.Fase" Else Ordenar = "PR.Ordem"
Else
    TextoFiltro1 = ""
    Ordenar = TextoFiltro
End If
Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select " & INNERJOINTEXTO & " where " & TextoFiltro1 & " " & DataFiltro & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' and " & Concluida & FiltroEscopo & " and PR.DtValidacao IS NOT NULL order by " & Ordenar, Conexao, adOpenKeyset, adLockOptimistic
ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

maquina = ""
IDlista = 0
ProcLimpaVariaveis
If TBOrdem.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBOrdem.RecordCount
    PBLista.Value = 1
    contador = 0
    Do While TBOrdem.EOF = False
        Set TBProdutividade = CreateObject("adodb.recordset")
        If Opt_individual.Value = True And optDetalhado.Value = True Then
            TBProdutividade.Open "Select * from Producao_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
            ProcEnviaDadosDetalhado
        Else
            Select Case cmbfiltrarpor
                Case "Código interno": TextoFiltro = TBOrdem!Desenho
                Case "Código de referência": TextoFiltro = TBOrdem!N_referencia & "'"
                Case "Descrição": TextoFiltro = TBOrdem!Produto
                Case "Família": TextoFiltro = TBOrdem!Classe
                Case "Cliente": TextoFiltro = TBOrdem!Cliente
                Case "Pedido interno": TextoFiltro = TBOrdem!Pedido
                Case "Tipo da ordem":
                    Select Case TBOrdem!Tipo
                        Case "E": TextoFiltro = "Produto final"
                        Case "M": TextoFiltro = "Subconjunto"
                        Case "F": TextoFiltro = "Componente"
                    End Select
                Case "Ordem": TextoFiltro = TBOrdem!Ordem
            End Select
            TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TextoFiltro & "' and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
            ProcEnviaDadosResumido
        End If
        TBOrdem.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
    Permitido = True
End If
TBOrdem.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosDetalhado()
On Error GoTo tratar_erro

TBProdutividade.AddNew
TBProdutividade!Ordem = TBOrdem!Ordem
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario

Select Case cmbfiltrarpor
    Case "Código interno": Texto = TBOrdem!Desenho
    Case "Código de referência": Texto = TBOrdem!N_referencia
    Case "Descrição": Texto = TBOrdem!Produto
    Case "Família": Texto = TBOrdem!Classe
    Case "Cliente": Texto = TBOrdem!Cliente
    Case "Pedido interno": Texto = TBOrdem!Ncotacao
    Case "Tipo da ordem":
        Select Case TBOrdem!Tipo
            Case "E": Texto = "Produto final"
            Case "M": Texto = "Subconjunto"
            Case "F": Texto = "Componente"
        End Select
    Case "Ordem": Texto = TBOrdem!Ordem
End Select
TBProdutividade!maquina = Texto

If chkOS.Value = 1 Then
    TBProdutividade!OS = TBOrdem!IDProducao
    TBProdutividade!QtdePrev = IIf(IsNull(TBOrdem!QTOK), 0, TBOrdem!QTOK) 'Quantidade prevista
    TBProdutividade!qtdeOK = IIf(IsNull(TBOrdem!Totalprod), 0, TBOrdem!Totalprod) 'Quantidade produzida
    TBProdutividade!qtdeNC = IIf(IsNull(TBOrdem!QTNC), 0, TBOrdem!QTNC) 'Quantidade refugada
    
    'Tempos
    TBProdutividade!Data1 = IIf(IsNull(TBOrdem!TempoExecucao), 0, TBOrdem!TempoExecucao) 'Previsto por peça
    TBProdutividade!Data2 = IIf(IsNull(TBOrdem!TEUTIL), 0, TBOrdem!TEUTIL) 'Real por peça
    TBProdutividade!Data3 = IIf(IsNull(TBOrdem!TTLPREVS), Null, FormataTempo(TBOrdem!TTLPREVS)) 'Previsto por lote
    TBProdutividade!Data4 = IIf(IsNull(TBOrdem!TETTUTILSEG), Null, FormataTempo(TBOrdem!TETTUTILSEG)) 'Real por lote
    
    'Custos
    TBProdutividade!Qtdetotalprod = IIf(IsNull(TBOrdem!CPPECA), 0, TBOrdem!CPPECA) 'MO prev. peça
    TBProdutividade!Terceiros = IIf(IsNull(TBOrdem!CRPECA), 0, TBOrdem!CRPECA) 'MO real peça
    TBProdutividade!impostos = IIf(IsNull(TBOrdem!CPLOTE), 0, TBOrdem!CPLOTE) 'MO prev. lote
    TBProdutividade!Lucro = IIf(IsNull(TBOrdem!CRLOTE), 0, TBOrdem!CRLOTE) 'MO real lote
    
    'Calcula valor do refugo
    ValorNC = 0
    Set TBAbrir_NFe = CreateObject("adodb.recordset")
    TBAbrir_NFe.Open "Select Sum(TTNC) as QtdeNC from CQ_NC_FABRICA where OS = " & TBOrdem!IDProducao & " and PARECERCQ = 'Rejeitar'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir_NFe.EOF = False Then
        If IIf(IsNull(TBAbrir_NFe!qtdeNC), 0, TBAbrir_NFe!qtdeNC) > 0 Then procCalculaRefugoOS
    End If
    TBAbrir_NFe.Close
    TBProdutividade!Refugo = Format(ValorNC, "###,##0.00") 'Valor refugo
    
    If IDlista <> TBOrdem!Ordem Then
        TBProdutividade!material = IIf(IsNull(TBOrdem!CTMaterial), 0, TBOrdem!CTMaterial) 'Material
        TBProdutividade!Numero4 = IIf(IsNull(TBOrdem!CTOutras), 0, TBOrdem!CTOutras) 'Outras
    End If
    IDlista = TBOrdem!Ordem
Else
    TBProdutividade!QtdePrev = IIf(IsNull(TBOrdem!Quant), 0, TBOrdem!Quant) 'Quantidade prevista
    TBProdutividade!qtdeOK = IIf(IsNull(TBOrdem!QuantProd), 0, TBOrdem!QuantProd) 'Quantidade produzida
    TBProdutividade!qtdeNC = IIf(IsNull(TBOrdem!QuantNC), 0, TBOrdem!QuantNC) 'Quantidade refugada
    
    'Tempos
    TBProdutividade!Data1 = IIf(IsNull(TBOrdem!TPP), Null, TBOrdem!TPP) 'Previsto por peça
    TBProdutividade!Data2 = IIf(IsNull(TBOrdem!tpr), Null, TBOrdem!tpr) 'Real por peça
    TBProdutividade!Data3 = IIf(IsNull(TBOrdem!TTTPrev), Null, TBOrdem!TTTPrev) 'Previsto por lote
    TBProdutividade!Data4 = IIf(IsNull(TBOrdem!TTTReal), Null, TBOrdem!TTTReal) 'Real por lote
    'Custos
    TBProdutividade!Qtdetotalprod = IIf(IsNull(TBOrdem!cpp), 0, TBOrdem!cpp) 'MO prev. peça
    TBProdutividade!Terceiros = IIf(IsNull(TBOrdem!CPR), 0, TBOrdem!CPR) 'MO real peça
    TBProdutividade!impostos = IIf(IsNull(TBOrdem!CTTPrev), 0, TBOrdem!CTTPrev) 'MO prev. lote
    TBProdutividade!Lucro = IIf(IsNull(TBOrdem!CTTReal), 0, TBOrdem!CTTReal) 'MO real lote
    
    TBProdutividade!material = IIf(IsNull(TBOrdem!CTMaterial), 0, TBOrdem!CTMaterial) 'Material
    TBProdutividade!Numero4 = IIf(IsNull(TBOrdem!CTOutras), 0, TBOrdem!CTOutras) 'Outras
End If
TBProdutividade!Numero1 = IIf(IsNull(TBOrdem!Eficiencia_prep), 0, TBOrdem!Eficiencia_prep) 'Eficiencia preparação
TBProdutividade!Numero2 = IIf(IsNull(TBOrdem!Eficiencia_exec), 0, TBOrdem!Eficiencia_exec) 'Eficiencia execução
TBProdutividade!Eficiencia = IIf(IsNull(TBOrdem!Eficiencia), 0, TBOrdem!Eficiencia) 'Eficiencia média
'Custos
TBProdutividade!Servicos = IIf(IsNull(TBOrdem!CTServico), 0, TBOrdem!CTServico) 'Terceiros
Valor1 = IIf(IsNull(TBProdutividade!Lucro), 0, TBProdutividade!Lucro)
Valor2 = IIf(IsNull(TBProdutividade!material), 0, TBProdutividade!material)
Valor3 = IIf(IsNull(TBProdutividade!Servicos), 0, TBProdutividade!Servicos)
Valor_DAS = IIf(IsNull(TBProdutividade!Numero4), 0, TBProdutividade!Numero4)
TBProdutividade!Total = Format(Valor1 + Valor2 + Valor3 + Valor_DAS, "###,##0.00") ' Total
If chkOS.Value = 1 Then
    If IsNull(TBOrdem!Totalprod) = False And TBOrdem!Totalprod > 0 Then TBProdutividade!Total_peca = Format(TBProdutividade!Total / TBOrdem!Totalprod, "###,##0.0000000000") ' Total
Else
                                                       'ORDEM         QTDE. PREVISTA                                QTDE. OK                                              QT. PROD.(OK+NC)                                                                                         CUSTO LOTE                                        CUSTO PEÇA                                CUSTO TERCEIROS                                       CUSTO MATERIAL                                          CUSTO OUTRAS                                        ORDEM CONSIGNADA
    TBProdutividade!Total_peca = FunCalculaValorUnitOrdem(TBOrdem!Ordem, IIf(IsNull(TBOrdem!Quant), 0, TBOrdem!Quant), IIf(IsNull(TBOrdem!QuantProd), 0, TBOrdem!QuantProd), IIf(IsNull(TBOrdem!QuantProd), 0, TBOrdem!QuantProd) + IIf(IsNull(TBOrdem!QuantNC), 0, TBOrdem!QuantNC), IIf(IsNull(TBOrdem!CTTReal), 0, TBOrdem!CTTReal), IIf(IsNull(TBOrdem!CPR), 0, TBOrdem!CPR), IIf(IsNull(TBOrdem!CTServico), 0, TBOrdem!CTServico), IIf(IsNull(TBOrdem!CTMaterial), 0, TBOrdem!CTMaterial), IIf(IsNull(TBOrdem!CTOutras), 0, TBOrdem!CTOutras), TBOrdem!consignacao)
    TBProdutividade!Refugo = Format(ValorNC, "###,##0.00") 'Valor refugo
End If
TBProdutividade!Numero3 = IIf(IsNull(TBOrdem!DtValidacao_custo), 0, 1) 'Validação do custo
OF = TBOrdem!Ordem

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosResumido()
On Error GoTo tratar_erro

If TBProdutividade.EOF = True Then TBProdutividade.AddNew
Select Case cmbfiltrarpor
    Case "Código interno": Texto = TBOrdem!Desenho
    Case "Código de referência": Texto = TBOrdem!N_referencia
    Case "Descrição": Texto = TBOrdem!Produto
    Case "Família": Texto = TBOrdem!Classe
    Case "Cliente": Texto = TBOrdem!Cliente
    Case "Pedido interno": Texto = TBOrdem!Ncotacao
    Case "Tipo da ordem":
        Select Case TBOrdem!Tipo
            Case "E": Texto = "Produto final"
            Case "M": Texto = "Subconjunto"
            Case "F": Texto = "Componente"
        End Select
    Case "Ordem":
        Texto = TBOrdem!Ordem
        TBProdutividade!Ordem = Texto
End Select
TBProdutividade!maquina = Texto
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario

If cmbfiltrarpor = "Ordem" Then
    TBProdutividade!Ordem = TBOrdem!Ordem
    TBProdutividade!Numero3 = IIf(IsNull(TBOrdem!DtValidacao_custo), 0, 1) 'Validação do custo
End If
TBProdutividade!QtdePrev = TBProdutividade!QtdePrev + IIf(IsNull(TBOrdem!Quant), 0, TBOrdem!Quant) 'Quantidade prevista
TBProdutividade!qtdeOK = TBProdutividade!qtdeOK + IIf(IsNull(TBOrdem!QuantProd), "0", TBOrdem!QuantProd) 'Quantidade produzida
TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + IIf(IsNull(TBOrdem!QuantNC), "0", TBOrdem!QuantNC) 'Quantidade refugada

Select Case cmbfiltrarpor
    Case "Código interno": If maquina <> TBOrdem!Desenho Then ProcLimpaVariaveis
    Case "Código de referência": If maquina <> TBOrdem!N_referencia Then ProcLimpaVariaveis
    Case "Descrição": If maquina <> TBOrdem!Produto Then ProcLimpaVariaveis
    Case "Família": If maquina <> TBOrdem!Classe Then ProcLimpaVariaveis
    Case "Cliente": If maquina <> TBOrdem!Cliente Then ProcLimpaVariaveis
    Case "Tipo": If maquina <> TBOrdem!Tipo Then ProcLimpaVariaveis
    Case "Ordem": If maquina <> TBOrdem!Ordem Then ProcLimpaVariaveis
End Select

'Tempos
'Previsto por peça
If IsNull(TBProdutividade!Data1) = False And TBProdutividade!Data1 <> "" Then
    ProcFormataHora (TBProdutividade!Data1)
Else
    ProcFormataHora ("00:00:00")
End If
TotalGeral = s + DecimoSegundos

ProcFormataHora (IIf(IsNull(TBOrdem!TPP), 0, TBOrdem!TPP))
TotalGeral = TotalGeral + s + DecimoSegundos
TBProdutividade!Data1 = FormataTempo(TotalGeral)

'Real por peça
If IsNull(TBProdutividade!Data2) = False And TBProdutividade!Data2 <> "" Then
    ProcFormataHora (TBProdutividade!Data2)
Else
    ProcFormataHora ("00:00:00")
End If
Valor_Cofins_Prod = s + DecimoSegundos

ProcFormataHora (IIf(IsNull(TBOrdem!tpr), 0, TBOrdem!tpr))
Valor_Cofins_Prod = Valor_Cofins_Prod + s + DecimoSegundos
TBProdutividade!Data2 = FormataTempo(Valor_Cofins_Prod)

'Previsto por lote
If IsNull(TBProdutividade!Data3) = False And TBProdutividade!Data3 <> "" Then
    ProcFormataHora (TBProdutividade!Data3)
Else
    ProcFormataHora ("00:00:00")
End If
Valor_Cofins_Serv = s + DecimoSegundos

ProcFormataHora (IIf(IsNull(TBOrdem!TTTPrev), 0, TBOrdem!TTTPrev))
Valor_Cofins_Serv = Valor_Cofins_Serv + s + DecimoSegundos
TBProdutividade!Data3 = FormataTempo(Valor_Cofins_Serv)

'Real por lote
If IsNull(TBProdutividade!Data4) = False And TBProdutividade!Data4 <> "" Then
    ProcFormataHora (TBProdutividade!Data4)
Else
    ProcFormataHora ("00:00:00")
End If
Valor_CSLL_Prod = s + DecimoSegundos

ProcFormataHora (IIf(IsNull(TBOrdem!TTTReal), 0, TBOrdem!TTTReal))
Valor_CSLL_Prod = Valor_CSLL_Prod + s + DecimoSegundos
TBProdutividade!Data4 = FormataTempo(Valor_CSLL_Prod)

'Eficiencia
TBProdutividade!Numero1 = TBProdutividade!Numero1 + IIf(IsNull(TBOrdem!Eficiencia_prep), 0, TBOrdem!Eficiencia_prep)
TBProdutividade!Numero2 = TBProdutividade!Numero2 + IIf(IsNull(TBOrdem!Eficiencia_exec), 0, TBOrdem!Eficiencia_exec)
TBProdutividade!Eficiencia = TBProdutividade!Eficiencia + IIf(IsNull(TBOrdem!Eficiencia), 0, TBOrdem!Eficiencia)
If (contador + 1) = TBOrdem.RecordCount Then
    TBProdutividade!Numero1 = TBProdutividade!Numero1 / TBOrdem.RecordCount
    TBProdutividade!Numero2 = TBProdutividade!Numero2 / TBOrdem.RecordCount
    TBProdutividade!Eficiencia = TBProdutividade!Eficiencia / TBOrdem.RecordCount
End If

'Custos
TBProdutividade!Qtdetotalprod = TBProdutividade!Qtdetotalprod + IIf(IsNull(TBOrdem!cpp), "0", TBOrdem!cpp) 'MO prev. peça
TBProdutividade!Terceiros = TBProdutividade!Terceiros + IIf(IsNull(TBOrdem!CPR), "0", TBOrdem!CPR) 'MO real peça
TBProdutividade!impostos = TBProdutividade!impostos + IIf(IsNull(TBOrdem!CTTPrev), "0", TBOrdem!CTTPrev) 'MO prev. lote
TBProdutividade!Lucro = TBProdutividade!Lucro + IIf(IsNull(TBOrdem!CTTReal), "0", TBOrdem!CTTReal) 'MO real lote
TBProdutividade!material = TBProdutividade!material + IIf(IsNull(TBOrdem!CTMaterial), "0", TBOrdem!CTMaterial) 'Material
TBProdutividade!Servicos = TBProdutividade!Servicos + IIf(IsNull(TBOrdem!CTServico), "0", TBOrdem!CTServico)  'Terceiros
TBProdutividade!Numero4 = TBProdutividade!Numero4 + IIf(IsNull(TBOrdem!CTOutras), 0, TBOrdem!CTOutras) 'Outras
Valor1 = IIf(IsNull(TBProdutividade!Lucro), 0, TBProdutividade!Lucro)
Valor2 = IIf(IsNull(TBProdutividade!material), 0, TBProdutividade!material)
Valor3 = IIf(IsNull(TBProdutividade!Servicos), 0, TBProdutividade!Servicos)
Valor_DAS = IIf(IsNull(TBProdutividade!Numero4), 0, TBProdutividade!Numero4)
TBProdutividade!Total = Format(Valor1 + Valor2 + Valor3 + Valor_DAS, "###,##0.00") ' Total
                                                                                'ORDEM         QTDE. PREVISTA                                QTDE. OK                                              QT. PROD.(OK+NC)                                                                                         CUSTO LOTE                                        CUSTO PEÇA                                CUSTO TERCEIROS                                       CUSTO MATERIAL                                          CUSTO OUTRAS                                        ORDEM CONSIGNADA
TBProdutividade!Total_peca = TBProdutividade!Total_peca + FunCalculaValorUnitOrdem(TBOrdem!Ordem, IIf(IsNull(TBOrdem!Quant), 0, TBOrdem!Quant), IIf(IsNull(TBOrdem!QuantProd), 0, TBOrdem!QuantProd), IIf(IsNull(TBOrdem!QuantProd), 0, TBOrdem!QuantProd) + IIf(IsNull(TBOrdem!QuantNC), 0, TBOrdem!QuantNC), IIf(IsNull(TBOrdem!CTTReal), 0, TBOrdem!CTTReal), IIf(IsNull(TBOrdem!CPR), 0, TBOrdem!CPR), IIf(IsNull(TBOrdem!CTServico), 0, TBOrdem!CTServico), IIf(IsNull(TBOrdem!CTMaterial), 0, TBOrdem!CTMaterial), IIf(IsNull(TBOrdem!CTOutras), 0, TBOrdem!CTOutras), TBOrdem!consignacao)
TBProdutividade!Refugo = Format(TBProdutividade!Refugo + ValorNC, "###,##0.00")
OF = TBOrdem!Ordem

Select Case cmbfiltrarpor
    Case "Código interno": maquina = TBOrdem!Desenho
    Case "Código de referência": maquina = TBOrdem!N_referencia
    Case "Descrição": maquina = TBOrdem!Produto
    Case "Família": maquina = TBOrdem!Classe
    Case "Cliente": maquina = IIf(IsNull(TBOrdem!Cliente), "", TBOrdem!Cliente)
    Case "Tipo": maquina = TBOrdem!Tipo
    Case "Ordem": maquina = TBOrdem!Ordem
End Select

TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarTotalizacoes()
On Error GoTo tratar_erro

Qtd = 0
Qtde = 0
qt = 0
TotalGeral = 0
VltUnit = 0
VlttTotal = 0
quantidade = 0
QTLOTE = 0
VlrSubTotal = 0
vlrTotalProd = 0
Valor1 = 0
Valor2 = 0
Valor3 = 0
Valor_DAS = 0
Valor_total = 0
Valor_Cofins_Prod = 0
Valor_Cofins_Serv = 0
Valor_CSLL_Prod = 0
ValorTotal = 0

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew
TBAbrir!Data_inicial = msk_fltInicio.Value
TBAbrir!Data_final = msk_fltFim.Value
If Opt_individual.Value = True Then
    If cmbfiltrarpor = "Ordem" Then TBAbrir!Texto = cmbfiltrarpor & ") : " & txtTexto Else TBAbrir!Texto = cmbfiltrarpor
Else
    TBAbrir!Texto = cmbfiltrarpor
End If
TBAbrir!Responsavel = pubUsuario
TBAbrir!Modulo = Formulario

contador = 0
Contador2 = 0
CamposFiltro = "QtdeNC, QtdePrev, Ordem, QtdeOK, Numero1, Numero2, Eficiencia, Data1, Data2, Data3, Data4, Qtdetotalprod, Terceiros, Impostos, Lucro, material, Servicos, Numero4, Refugo, Total_peca"
GrupoTexto = ""
If chkOS.Value = 1 Then GrupoTexto = "group by " & CamposFiltro
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select " & CamposFiltro & " from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' " & GrupoTexto & " order by Ordem", Conexao, adOpenKeyset, adLockReadOnly
If TBproducao.EOF = False Then
    Do While TBproducao.EOF = False
        qt = qt + IIf(IsNull(TBproducao!qtdeNC), 0, TBproducao!qtdeNC)
        If chkOS.Value = 1 Then
            Qtd = Qtd + IIf(IsNull(TBproducao!QtdePrev), 0, TBproducao!QtdePrev)
            Set TBCFOP = CreateObject("adodb.recordset")
            TBCFOP.Open "Select Totalprod from Ordemservico where ordem = " & TBproducao!Ordem & " order by fase DESC", Conexao, adOpenKeyset, adLockOptimistic
            If TBCFOP.EOF = False Then Qtde = Qtde + IIf(IsNull(TBCFOP!Totalprod), 0, TBCFOP!Totalprod)
            
            Set TBCFOP = CreateObject("adodb.recordset")
            TBCFOP.Open "Select Sum(QTNC) as QtdeNC from Ordemservico where ordem = " & TBproducao!Ordem, Conexao, adOpenKeyset, adLockOptimistic
            If TBCFOP.EOF = False Then Qtde = Qtde + TBCFOP!qtdeNC

            Set TBCFOP = CreateObject("adodb.recordset")
            TBCFOP.Open "Select Eficiencia_prep, Eficiencia_exec, Eficiencia from producao where ordem = " & TBproducao!Ordem, Conexao, adOpenKeyset, adLockOptimistic
            If TBCFOP.EOF = False Then
                Contador2 = Contador2 + 1
                Valor_Cofins_Serv = Valor_Cofins_Serv + IIf(IsNull(TBCFOP!Eficiencia_prep), 0, TBCFOP!Eficiencia_prep)
                Valor_CSLL_Prod = Valor_CSLL_Prod + IIf(IsNull(TBCFOP!Eficiencia_exec), 0, TBCFOP!Eficiencia_exec)
                ValorTotal = ValorTotal + IIf(IsNull(TBCFOP!Eficiencia), 0, TBCFOP!Eficiencia)
            End If
            TBCFOP.Close
        Else
            Qtd = Qtd + IIf(IsNull(TBproducao!QtdePrev), 0, TBproducao!QtdePrev)
            Qtde = Qtde + IIf(IsNull(TBproducao!qtdeOK), 0, TBproducao!qtdeOK)
                        
            Valor_Cofins_Serv = Valor_Cofins_Serv + IIf(IsNull(TBproducao!Numero1), 0, TBproducao!Numero1)
            Valor_CSLL_Prod = Valor_CSLL_Prod + IIf(IsNull(TBproducao!Numero2), 0, TBproducao!Numero2)
            ValorTotal = ValorTotal + IIf(IsNull(TBproducao!Eficiencia), 0, TBproducao!Eficiencia)
        End If
        
        If TBproducao!Data1 <> "00:00:00" Then
            ProcFormataHora (TBproducao!Data1)
            TotalGeral = TotalGeral + s + DecimoSegundos
        End If
        If TBproducao!Data2 <> "00:00:00" Then
            ProcFormataHora (TBproducao!Data2)
            VltUnit = VltUnit + s + DecimoSegundos
        End If
        If TBproducao!Data3 <> "00:00:00" Then
            ProcFormataHora (TBproducao!Data3)
            VlttTotal = VlttTotal + s + DecimoSegundos
        End If
        If TBproducao!Data4 <> "00:00:00" Then
            ProcFormataHora (TBproducao!Data4)
            quantidade = quantidade + s + DecimoSegundos
        End If
        QTLOTE = QTLOTE + IIf(IsNull(TBproducao!Qtdetotalprod), 0, TBproducao!Qtdetotalprod)
        VlrSubTotal = VlrSubTotal + IIf(IsNull(TBproducao!Terceiros), 0, TBproducao!Terceiros)
        vlrTotalProd = vlrTotalProd + IIf(IsNull(TBproducao!impostos), 0, TBproducao!impostos)
        Valor1 = Valor1 + IIf(IsNull(TBproducao!Lucro), 0, TBproducao!Lucro)
        Valor2 = Valor2 + IIf(IsNull(TBproducao!material), 0, TBproducao!material)
        Valor3 = Valor3 + IIf(IsNull(TBproducao!Servicos), 0, TBproducao!Servicos)
        Valor_DAS = Valor_DAS + IIf(IsNull(TBproducao!Numero4), 0, TBproducao!Numero4)
        Valor_total = Valor_total + IIf(IsNull(TBproducao!Refugo), 0, TBproducao!Refugo)
        Valor_Cofins_Prod = Valor_Cofins_Prod + IIf(IsNull(TBproducao!Total_peca), 0, TBproducao!Total_peca)
        
        contador = contador + 1
        If contador = TBproducao.RecordCount Then
            Valor_Cofins_Serv = Valor_Cofins_Serv / IIf(chkOS.Value = 1, Contador2, contador)
            Valor_CSLL_Prod = Valor_CSLL_Prod / IIf(chkOS.Value = 1, Contador2, contador)
            ValorTotal = ValorTotal / IIf(chkOS.Value = 1, Contador2, contador)
        End If
        TBproducao.MoveNext
    Loop
End If

TBAbrir!QtdePrevista = Qtd 'Quantidade prevista
TBAbrir!QtdeProduzida = Qtde 'Quantidade produzida
TBAbrir!qtdeNC = qt 'Quantidade refugada

TBAbrir!Data1 = FormataTempo(TotalGeral) 'Previsto por peça
TBAbrir!Data2 = FormataTempo(VltUnit) 'Real por peça
TBAbrir!Data3 = FormataTempo(VlttTotal) 'Previsto por lote
TBAbrir!Data4 = FormataTempo(quantidade) 'Real por lote

TBAbrir!Numero1 = Valor_Cofins_Serv 'Eficiencia prep.
TBAbrir!Numero2 = Valor_CSLL_Prod 'Eficiencia exec.
TBAbrir!QtdeOrdem = ValorTotal 'Eficiencia média

TBAbrir!CustoMat = QTLOTE 'MO prev. peça
TBAbrir!Terceros = VlrSubTotal 'MO real peça
TBAbrir!CustoObra = vlrTotalProd 'MO prev. lote
TBAbrir!Valor1 = Valor1 'MO real lote
TBAbrir!Valor2 = Valor2 'Material
TBAbrir!Valor3 = Valor3 'Terceiros
TBAbrir!Numero4 = Valor_DAS 'Outras
TBAbrir!Lucro = Valor_total 'Refugo
TBAbrir!Total1 = Valor1 + Valor2 + Valor3 + Valor_DAS 'Total
TBAbrir!Total2 = Valor_Cofins_Prod 'Total peça

TBAbrir.Update
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_comparativo_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
ProcLimpaCamposTotais
If Opt_comparativo.Value = True Then
    optDetalhado.Enabled = False
    optResumido.Value = True
    cmbTexto.ListIndex = -1
    cmbTexto.Enabled = False
    txtTexto = ""
    txtTexto.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_individual_Click()
On Error GoTo tratar_erro

If Opt_individual.Value = True Then
    optDetalhado.Value = True
    optDetalhado.Enabled = True
    cmbTexto.Enabled = True
    txtTexto.Enabled = True
    ProcCarregaComboTexto
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optConclusao_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optDetalhado_Click()
On Error GoTo tratar_erro

If optDetalhado.Value = True Then
    Lista.ListItems.Clear
    ProcLimpaCamposTotais
    chkOS.Enabled = True
    If cmbfiltrarpor = "Ordem" Then ProcListaOrdem Else ProcListaPadrao
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optEmissao_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optPrazo_Click()
On Error GoTo tratar_erro

Lista.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optResumido_Click()
On Error GoTo tratar_erro

If optResumido.Value = True Then
    Lista.ListItems.Clear
    ProcLimpaCamposTotais
    chkOS.Value = 0
    chkOS.Enabled = False
    If cmbfiltrarpor = "Ordem" Then ProcListaOrdem Else ProcListaPadrao
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

Lista.ListItems.Clear
ProcLimpaCamposTotais
If txtTextot <> "" Then
    VerifNumero = txtTexto
    ProcVerificaNumero
    If VerifNumero = False Then
        txtTexto = ""
        txtTexto.SetFocus
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
    Case 1: ProcAbrir
    Case 2: ProcImprimir
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procCalculaRefugoOS()
On Error GoTo tratar_erro

Dim TBMaterialVlrUnitOrdem As ADODB.Recordset
Dim TBproducaoVlrUnitOrdem As ADODB.Recordset
Dim TBAbrirVlrUnitOrdem As ADODB.Recordset

'Valor NC
Valor_Cofins_Prod = 0
Valor1 = 0 'Serviço
Valor2 = 0 'Material
Valor3 = 0 'Mão de obra
Valor_CSLL_Serv = 0
Valor_IPI = 0

'Custo de material
If TBOrdem!consignacao = False Then
    Set TBMaterialVlrUnitOrdem = CreateObject("adodb.recordset")
    TBMaterialVlrUnitOrdem.Open "Select Valor_saida_estoque, Saida from Producaomaterial where Ordem = " & TBOrdem!Ordem & " order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
    If TBMaterialVlrUnitOrdem.EOF = False Then
        Do While TBMaterialVlrUnitOrdem.EOF = False
            
            'Verifica valor total do material
            Valor_CSLL_Prod = IIf(IsNull(TBMaterialVlrUnitOrdem!Valor_saida_estoque), 0, TBMaterialVlrUnitOrdem!Valor_saida_estoque)
            If TBMaterialVlrUnitOrdem!Saida <> "NÃO" Then
                Set TBproducaoVlrUnitOrdem = CreateObject("adodb.recordset")
                TBproducaoVlrUnitOrdem.Open "Select Totalprod from ordemservico where Ordem = " & TBOrdem!Ordem & " ORDER BY fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
                If TBproducaoVlrUnitOrdem.EOF = False Then
                    Qtd_Prog = IIf(IsNull(TBproducaoVlrUnitOrdem!Totalprod), 0, TBproducaoVlrUnitOrdem!Totalprod) 'Qtde. produzida
                    If Qtd_Prog <> 0 Then
                        Valor_CSLL_Serv = Format(Valor_CSLL_Serv + (Valor_CSLL_Prod / Qtd_Prog), "###,##0.0000000000")
                    ElseIf Qtde <> 0 Then
                            Valor_CSLL_Serv = Format(Valor_CSLL_Serv + (Valor_CSLL_Prod / Qtde), "###,##0.0000000000")
                    End If
                End If
                TBproducaoVlrUnitOrdem.Close
            End If
            TBMaterialVlrUnitOrdem.MoveNext
        Loop
    End If
    TBMaterialVlrUnitOrdem.Close
End If

'Verifica qtde NC da ordem
QuantComprado = 0
Set TBAbrirVlrUnitOrdem = CreateObject("adodb.recordset")
TBAbrirVlrUnitOrdem.Open "Select Sum(TTNC) as QtdeNC from CQ_NC_FABRICA where OS = " & TBOrdem!IDProducao & " and PARECERCQ = 'Rejeitar'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrirVlrUnitOrdem.EOF = False Then
    QuantComprado = IIf(IsNull(TBAbrirVlrUnitOrdem!qtdeNC), 0, TBAbrirVlrUnitOrdem!qtdeNC)
End If

Set TBproducaoVlrUnitOrdem = CreateObject("adodb.recordset")
TBproducaoVlrUnitOrdem.Open "Select Totalprod, CTServico, CRPECA, IDProducao from ordemservico where Ordem = " & TBOrdem!Ordem & " and Fase <= " & TBOrdem!Fase & " ORDER BY fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
If TBproducaoVlrUnitOrdem.EOF = False Then
    Do While TBproducaoVlrUnitOrdem.EOF = False
        'Soma valor unitário do SERVIÇO na OS
        If IsNull(TBproducaoVlrUnitOrdem!Totalprod) = False And TBproducaoVlrUnitOrdem!Totalprod <> "" And TBproducaoVlrUnitOrdem!Totalprod <> "0" Then Valor_IPI = Format(Valor_IPI + (TBproducaoVlrUnitOrdem!CTServico / TBproducaoVlrUnitOrdem!Totalprod), "###,##0.0000000000")
        
        'Soma valor unitário da MÃO DE OBRA na OS
        Valor_Cofins_Prod = Format(Valor_Cofins_Prod + TBproducaoVlrUnitOrdem!CRPECA, "###,##0.0000000000")
        
        Qtd_Prog = 0
        Set TBAbrirVlrUnitOrdem = CreateObject("adodb.recordset")
        TBAbrirVlrUnitOrdem.Open "Select Sum(TTNC) as QtdeNC from CQ_NC_FABRICA where OS = " & TBproducaoVlrUnitOrdem!IDProducao & " and PARECERCQ = 'Rejeitar'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrirVlrUnitOrdem.EOF = False Then
            Qtd_Prog = IIf(IsNull(TBAbrirVlrUnitOrdem!qtdeNC), 0, TBAbrirVlrUnitOrdem!qtdeNC)
            Valor1 = Format(Valor_IPI * Qtd_Prog, "###,##0.00") 'Valor total unitário serviço x qtde. refugada da OS
            Valor3 = Format(Valor_Cofins_Prod * Qtd_Prog, "###,##0.00") 'Valor total unitário mão de obra x qtde. refugada da OS
        End If
        
        TBAbrirVlrUnitOrdem.Close
        TBproducaoVlrUnitOrdem.MoveNext
    Loop
End If

'Valor do material por peça x qtde. refugada
If TBOrdem!QTOK <> 0 Then Valor2 = Format(Valor_CSLL_Serv * QuantComprado, "###,##0.00")
                   'SE  +   MT   +   MO
ValorNC = Format(Valor1 + Valor2 + Valor3, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

