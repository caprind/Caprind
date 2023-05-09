VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frm_Orcamento_Processo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "CAPRIND v5.0 | Orçamento | Processo"
   ClientHeight    =   9120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações da fase do processo de fabricação para formação de preço de venda"
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
      Height          =   2565
      Left            =   300
      TabIndex        =   8
      Top             =   750
      Width           =   9915
      Begin DrawSuite2022.USLabel USLabel3 
         Height          =   195
         Left            =   240
         Top             =   1470
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   344
         Caption         =   "Observações:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         NoHTMLCaption   =   "Observações:"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   165
         Left            =   390
         Top             =   1710
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   291
         Caption         =   "Setup = HHH:MM:SS | Ex: 001:14:23"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         NoHTMLCaption   =   "Setup = HHH:MM:SS | Ex: 001:14:23"
      End
      Begin VB.CheckBox chkPchora 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "O tempo de execução será calculado baseado em quantidade de itens produzidos por tempo de execução"
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
         Height          =   210
         Left            =   210
         TabIndex        =   27
         Top             =   2220
         Width           =   6915
      End
      Begin VB.TextBox txtPecaHora_processos 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   5700
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "1"
         ToolTipText     =   "Peça por hora."
         Top             =   540
         Width           =   975
      End
      Begin VB.ComboBox cmbPosto 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   540
         Width           =   4095
      End
      Begin VB.TextBox txtcodigoposto 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   690
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   540
         Width           =   885
      End
      Begin VB.TextBox txtTotalHora_processos 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   210
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Tempo de execução por peça."
         Top             =   1110
         Width           =   1005
      End
      Begin VB.TextBox txtValorHora_processos 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   2145
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Custo por hora de execução."
         Top             =   1110
         Width           =   900
      End
      Begin VB.TextBox txtValorTotal_processos 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3060
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Valor total."
         Top             =   1110
         Width           =   1035
      End
      Begin VB.TextBox txtFase 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Fase."
         Top             =   540
         Width           =   465
      End
      Begin VB.TextBox txtErro 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   8160
         MaxLength       =   20
         TabIndex        =   4
         ToolTipText     =   "Porcentagem de erro."
         Top             =   540
         Width           =   585
      End
      Begin VB.TextBox txtValorHoraPrep_Processos 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1230
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Valor por hora de preparação."
         Top             =   1110
         Width           =   900
      End
      Begin MSMask.MaskEdBox txtPreparacao_processos 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "H:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   4
         EndProperty
         Height          =   315
         Left            =   6690
         TabIndex        =   2
         ToolTipText     =   "Tempo de preparação previsto."
         Top             =   540
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12640511
         ForeColor       =   0
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin RichTextLib.RichTextBox txtTrabalho 
         Height          =   1065
         Left            =   4110
         TabIndex        =   5
         ToolTipText     =   "Instruções de trabalho."
         Top             =   1110
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   1879
         _Version        =   393217
         BackColor       =   12640511
         BorderStyle     =   0
         ScrollBars      =   2
         TextRTF         =   $"frm_Orcamento_Processo.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DrawSuite2022.USLabel USLabel2 
         Height          =   165
         Left            =   390
         Top             =   1890
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   291
         Caption         =   "Execução = HHH:MM:SS | Ex: 001:14:23"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         NoHTMLCaption   =   "Execução = HHH:MM:SS | Ex: 001:14:23"
      End
      Begin MSMask.MaskEdBox txtExecucao_processos 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "H:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   4
         EndProperty
         Height          =   315
         Left            =   7440
         TabIndex        =   3
         ToolTipText     =   "Tempo de preparação previsto."
         Top             =   540
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12640511
         ForeColor       =   0
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin DrawSuite2022.USButton btnNovo 
         Height          =   525
         Left            =   8850
         TabIndex        =   29
         Top             =   570
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   926
         DibPicture      =   "frm_Orcamento_Processo.frx":007E
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Novo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   8
         PicSize         =   1
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
      End
      Begin DrawSuite2022.USButton btnGravar 
         Height          =   495
         Left            =   8850
         TabIndex        =   30
         Top             =   1125
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         DibPicture      =   "frm_Orcamento_Processo.frx":6362
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Gravar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   8
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
      End
      Begin DrawSuite2022.USButton btnExcluir 
         Height          =   495
         Left            =   8850
         TabIndex        =   31
         Top             =   1665
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         DibPicture      =   "frm_Orcamento_Processo.frx":ED67
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Excluir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   8
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Itens|Tempo"
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
         Left            =   5730
         TabIndex        =   28
         Top             =   330
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Index           =   25
         Left            =   885
         TabIndex        =   25
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Instruções de trabalho (Informações para o colaborador)"
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
         Left            =   4425
         TabIndex        =   24
         Top             =   900
         Width           =   4110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fase"
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
         Index           =   76
         Left            =   270
         TabIndex        =   23
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label1 
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
         Index           =   78
         Left            =   2670
         TabIndex        =   22
         Top             =   330
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Setup"
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
         Index           =   79
         Left            =   6840
         TabIndex        =   21
         Top             =   330
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Execução"
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
         Index           =   80
         Left            =   7440
         TabIndex        =   20
         Top             =   330
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "% Erro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   81
         Left            =   8175
         TabIndex        =   19
         Top             =   330
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tempo|Item"
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
         Index           =   83
         Left            =   270
         TabIndex        =   18
         Top             =   900
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "R$ (Setup)"
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
         Index           =   84
         Left            =   1290
         TabIndex        =   17
         Top             =   900
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "R$ (Exec.)"
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
         Index           =   85
         Left            =   2235
         TabIndex        =   16
         Top             =   900
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "R$ (Total)"
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
         Index           =   86
         Left            =   3217
         TabIndex        =   15
         Top             =   900
         Width           =   720
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   7
      Top             =   8715
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   741
      DibPicture      =   "frm_Orcamento_Processo.frx":185B3
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frm_Orcamento_Processo.frx":22760
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin MSComctlLib.ListView lista_Processos 
      Height          =   5025
      Left            =   300
      TabIndex        =   26
      Top             =   3480
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   8864
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Fase"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Posto de trabalho"
         Object.Width           =   8820
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Código"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "D"
         Text            =   "Exec. x peça"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Valor hora prep."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Valor hora exec."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Valor total"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "frm_Orcamento_Processo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcCarregacomboGrupo()
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
StrSql = "Select Grupo from CadMaquinas GROUP BY GRUPO"
TBFIltro.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then

Do While TBFIltro.EOF = False
cmbGrupo.AddItem TBFIltro!Grupo
TBFIltro.MoveNext
Loop

End If
TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregacomboPosto()
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
StrSql = "Select grupo, Maquina, Descricao from CadMaquinas order by Descricao"
TBFIltro.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then

Do While TBFIltro.EOF = False
cmbPosto.AddItem TBFIltro!Descricao
TBFIltro.MoveNext
Loop

End If
TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnGravar_Click()
On Error GoTo tratar_erro

ProcgravarFase

Set TBFases = CreateObject("adodb.recordset")
StrSql = "Select SUM(vlrtotal) as Total from Vendas_Orcamento_Fases where ID_Orcamento = '" & frm_orcamento.txtId.Text & "'"
TBFases.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBFases.EOF = False Then
frm_orcamento.txtv1.Text = Format(TBFases!Total, "###,##0.00")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcgravarFase()
On Error GoTo tratar_erro
Dim Setup As Date
Dim Execucao As Date


Set TBFases = CreateObject("adodb.recordset")
StrSql = "Select * from Vendas_Orcamento_Fases where ID_Orcamento = '" & frm_orcamento.txtId.Text & "' and fase = " & txtFase.Text & ""
TBFases.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBFases.EOF = True Then
TBFases.AddNew
End If
TBFases!ID_orcamento = frm_orcamento.txtId
TBFases!Fase = txtFase.Text
TBFases!Posto = cmbPosto.Text
TBFases!codigoposto = txtcodigoposto.Text
TBFases!itenstempo = txtPecaHora_processos.Text
TBFases!tempoitens = IIf(txtTotalHora_processos.Text = "", "00:00:00", txtTotalHora_processos.Text)
TBFases!vrlsetup = txtValorHoraPrep_Processos.Text
TBFases!vlrexec = txtValorHora_processos.Text
TBFases!Setup = txtPreparacao_processos.Text
TBFases!Execucao = txtExecucao_processos.Text
TBFases!ERRO = txtErro.Text
TBFases!VlrTotal = txtValorTotal_processos
TBFases!Descricao = txtTrabalho.Text
TBFases!ItemHora = chkPchora.Value
TBFases.Update
TBFases.Close

USMsgBox "Dados gravados com sucesso!", vbInformation, "CAPRIND  V5.0"
ProcCarregaLista_processos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnExcluir_Click()
On Error GoTo tratar_erro

If lista_Processos.ListItems.Count > 0 Then
If USMsgBox("Deseja realmente excluir a fase " & lista_Processos.SelectedItem.ListSubItems.Item(1).Text & "?", vbYesNo, "CAPRIND  v5.0") = vbYes Then
Conexao.Execute ("Delete from Vendas_Orcamento_Fases where id_fases = '" & lista_Processos.SelectedItem & "'")
USMsgBox "Fase excluida com sucesso!", vbInformation, "CAPRIND v5.0"
ProcCarregaLista_processos
End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnNovo_Click()
On Error GoTo tratar_erro

procNovaFase
ProcLimparCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbPosto_Change()
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
StrSql = "Select grupo, Maquina, Descricao from CadMaquinas where descricao = '" & cmbPosto.Text & "'order by Descricao"
TBFIltro.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
txtcodigoposto.Text = TBFIltro!maquina
txtValorHora_processos = IIf(IsNull(TBFIltro!PrecoHora), "", Format(TBFIltro!PrecoHora, "###,##0.00"))
txtValorHoraPrep_Processos = IIf(IsNull(TBFIltro!PrecoHora_Setup), "", Format(TBFIltro!PrecoHora_Setup, "###,##0.00"))
End If
TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procNovaFase()
On Error GoTo tratar_erro

Set TBFases = CreateObject("adodb.recordset")
StrSql = "Select * from Vendas_Orcamento_Fases where ID_Orcamento = '" & frm_orcamento.txtId.Text & "' order by fase"
TBFases.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBFases.EOF = False Then
TBFases.MoveLast
Fase = Int(TBFases!Fase) + 10
Else
Fase = 10
End If

txtFase.Text = Fase

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculamaquina()
On Error GoTo tratar_erro

Qtde = 0
valor = 0
qt = 0
Qtd = 0
ValorTotal = 0
quantidade = 0
quantnovo = 0

'Calcula valor de execução
NovoValor = Replace(txtTotalHora_processos, ",", ".")
If txtTotalHora_processos <> "" Then
    ProcFormataHora (txtTotalHora_processos)
    HoraResultado = DataResultado
    ElapsedTime (HoraResultado)
    Qtde = (s + DecimoSegundos) / 3600
End If
valor = IIf(txtValorHora_processos = "", 0, txtValorHora_processos)
ValorTotal = Qtde * valor

'Calcula valor de preparação
Qtd = IIf(frm_orcamento.txtLote.Text = "", 0, frm_orcamento.txtLote.Text)
Valor1 = IIf(txtValorHoraPrep_Processos = "", 0, txtValorHoraPrep_Processos) / Qtd
txtPreparacao_processos.PromptInclude = False
If Len(txtPreparacao_processos.Text) = 6 Then
    txtPreparacao_processos.PromptInclude = True
    ProcFormataHora (txtPreparacao_processos)
    HoraResultado = DataResultado
    ElapsedTime (HoraResultado)
    qt = s / 3600
End If
quantnovo = IIf(txtErro = "", 0, txtErro)

If qt > 0 Then
    quantidade = qt * Valor1
    ValorTotal = ValorTotal + quantidade
End If
txtValorTotal_processos = Format(ValorTotal, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

If frm_orcamento.txtcodproduto <> "" Then

ProcCarregaLista_processos
ProcCarregacomboPosto

End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimparCampos()
On Error GoTo tratar_erro

'      cmbPosto.Text = ""
      txtcodigoposto.Text = ""
      txtPreparacao_processos.Text = "00:00:00"
      txtExecucao_processos = "00:00:00"
      txtPecaHora_processos = "1"
      txtTotalHora_processos.Text = ""
      txtErro.Text = "0"
      txtValorHoraPrep_Processos.Text = "0,00"
      txtValorHora_processos.Text = "0,00"
      txtValorTotal_processos = "0,00"
      txtdescricao = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPreparacao_processos_Change()
On Error GoTo tratar_erro

ProcCalculaExecucao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub lista_Processos_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If lista_Processos.ListItems.Count > 0 Then
  If lista_Processos.SelectedItem <> "" Then
  
  Set TBLISTA = CreateObject("adodb.recordset")
  
  StrSql = "Select * from Vendas_Orcamento_Fases where ID_Fases = '" & lista_Processos.SelectedItem & "'"
  TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
  If TBLISTA.EOF = False Then
      txtFase.Text = IIf(IsNull(TBLISTA!Fase), "", TBLISTA!Fase)
      cmbPosto.Text = IIf(IsNull(TBLISTA!Posto), "", TBLISTA!Posto)
      txtcodigoposto.Text = IIf(IsNull(TBLISTA!codigoposto), "", TBLISTA!codigoposto)
      txtPreparacao_processos.Text = IIf(IsNull(TBLISTA!Setup), "", TBLISTA!Setup)
      txtExecucao_processos = IIf(IsNull(TBLISTA!Execucao), "", TBLISTA!Execucao)
      txtPecaHora_processos = IIf(IsNull(TBLISTA!itenstempo), "", TBLISTA!itenstempo)
      txtTotalHora_processos.Text = IIf(IsNull(TBLISTA!tempoitens), "", Format(TBLISTA!tempoitens, "HH:MM:SS"))
      txtErro.Text = IIf(IsNull(TBLISTA!ERRO), "", TBLISTA!ERRO)
      txtValorHoraPrep_Processos.Text = IIf(IsNull(TBLISTA!vrlsetup), "", Format(TBLISTA!vrlsetup, "###,##0.00"))
      txtValorHora_processos.Text = IIf(IsNull(TBLISTA!vlrexec), "", Format(TBLISTA!vlrexec, "###,##0.00"))
      txtValorTotal_processos = IIf(IsNull(TBLISTA!VlrTotal), "", Format(TBLISTA!VlrTotal, "###,##0.00"))
      txtdescricao = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
      chkPchora.Value = IIf(TBLISTA!ItemHora = True, "1", "0")
      
  End If
  TBLISTA.Close
  
  End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista_processos()
On Error GoTo tratar_erro

valor = 0
lista_Processos.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")

StrSql = "Select * from Vendas_Orcamento_Fases where ID_Orcamento = '" & frm_orcamento.txtId.Text & "'"
TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Contador = 0
    Do While TBLISTA.EOF = False
        With lista_Processos.ListItems
            .Add , , TBLISTA!ID_fases
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Fase), "", TBLISTA!Fase)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Posto), "", TBLISTA!Posto)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!codigoposto), "", TBLISTA!codigoposto)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!tempoitens), "", Format(TBLISTA!tempoitens, "HH:MM:SS"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!vrlsetup), "", "R$ " & Format(TBLISTA!vrlsetup, "###,##0.00"))
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!vlrexec), "", "R$ " & Format(TBLISTA!vlrexec, "###,##0.00"))
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!VlrTotal), "", "R$ " & Format(TBLISTA!VlrTotal, "###,##0.00"))
            valor = valor + IIf(IsNull(TBLISTA!VlrTotal), 0, TBLISTA!VlrTotal)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
    Loop
    lista_Processos.ListItems.Add , , 1 'TBLISTA!ID_fases
    lista_Processos.ListItems.Item(Contador + 1).SubItems(4) = "TOTAL :"
    lista_Processos.ListItems.Item(Contador + 1).ListSubItems.Item(4).ForeColor = vbRed
    lista_Processos.ListItems.Item(Contador + 1).SubItems(7) = "R$ " & Format(valor, "###,##0.00")
    lista_Processos.ListItems.Item(Contador + 1).ListSubItems.Item(7).ForeColor = vbRed
    

frm_orcamento.txtv1.Text = Format(valor, "###,##0.00")

End If
TBLISTA.Close

        
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkPchora_Click()
On Error GoTo tratar_erro

With txtPecaHora_processos
    If chkPchora.Value = 1 Then
        .Locked = False
        .TabStop = True
    Else
        .Locked = True
        .TabStop = False
        .Text = 1
    End If
End With
ProcCalculaExecucao
ProcCalculamaquina

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub txtPecaHora_processos_Change()
On Error GoTo tratar_erro

If txtPecaHora_processos.Text <> "" Then
    VerifNumero = txtPecaHora_processos.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPecaHora_processos.Text = ""
        txtPecaHora_processos.SetFocus
        Exit Sub
    End If
End If
ProcCalculaExecucao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub txtExecucao_processos_Change()
On Error GoTo tratar_erro

ProcCalculaExecucao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCalculaExecucao()
On Error GoTo tratar_erro

txtPreparacao_processos.PromptInclude = True
txtExecucao_processos.PromptInclude = False
If Len(txtExecucao_processos.Text) < 6 Then
    txtExecucao_processos.PromptInclude = True
    txtTotalHora_processos = ""
    Exit Sub
End If

txtExecucao_processos.PromptInclude = True
If txtExecucao_processos > "23:59:59" Then
    ProcFormataHora (txtExecucao_processos)
    Familiatext = DataResultado
    TotalGeral = FunCalculaSegPC(Familiatext, txtPecaHora_processos)
Else
    If txtPecaHora_processos <> "" Then
        TotalGeral = FunCalculaSegPC(txtExecucao_processos, txtPecaHora_processos)
    End If
End If
If txtErro <> "" And txtErro <> "0" Then
    quantnovo = (TotalGeral * txtErro) / 100
    TotalGeral = TotalGeral + quantnovo
End If
Texto = FormataTempo(TotalGeral)
txtTotalHora_processos = Texto
ProcCalculamaquina
'txtExecucao_processos.PromptInclude = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbPosto_Click()
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
StrSql = "Select * from CadMaquinas where descricao = '" & cmbPosto.Text & "'order by Descricao"
TBFIltro.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
txtcodigoposto.Text = TBFIltro!maquina
txtValorHora_processos = IIf(IsNull(TBFIltro!PrecoHora), "", Format(TBFIltro!PrecoHora, "###,##0.00"))
txtValorHoraPrep_Processos = IIf(IsNull(TBFIltro!PrecoHora_Setup), "", Format(TBFIltro!PrecoHora_Setup, "###,##0.00"))
End If
TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

