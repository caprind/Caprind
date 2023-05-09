VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form FrmProd_Resultados 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "PCP | Gerenciamento de ordem - Resultados da ordem detalhado"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   12225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmProd_Resultados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   767
      DibPicture      =   "FrmProd_Resultados.frx":000C
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
      Icon            =   "FrmProd_Resultados.frx":A12F
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   52
      Top             =   3570
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   714
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Total mão de obra (Tempo / valor)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1785
      Left            =   8160
      TabIndex        =   44
      Top             =   570
      Width           =   3930
      Begin VB.TextBox Txt_custo_real_lote_total 
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
         Height          =   315
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor real por lote."
         Top             =   1320
         Width           =   1140
      End
      Begin VB.TextBox Txt_custo_previsto_lote_total 
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
         Height          =   315
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor previsto por lote."
         Top             =   990
         Width           =   1140
      End
      Begin VB.TextBox Txt_custo_real_peca_total 
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
         Height          =   315
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor real por peça."
         Top             =   660
         Width           =   1140
      End
      Begin VB.TextBox Txt_custo_previsto_peca_total 
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
         Height          =   315
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor previsto por peça."
         Top             =   330
         Width           =   1140
      End
      Begin VB.TextBox Txt_tempo_real_lote_total 
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
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Tempo real por lote."
         Top             =   1320
         Width           =   1050
      End
      Begin VB.TextBox Txt_tempo_previsto_lote_total 
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
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Tempo previsto por lote."
         Top             =   990
         Width           =   1050
      End
      Begin VB.TextBox Txt_tempo_real_peca_total 
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
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Tempo real por peça."
         Top             =   660
         Width           =   1050
      End
      Begin VB.TextBox Txt_tempo_previsto_peca_total 
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
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Tempo previsto por peça."
         Top             =   330
         Width           =   1050
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Alinhar à Direita
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Real do lote :"
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
         Left            =   495
         TabIndex        =   48
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Alinhar à Direita
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Previsto do lote :"
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
         Left            =   225
         TabIndex        =   47
         Top             =   990
         Width           =   1230
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Alinhar à Direita
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Real p/ item :"
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
         Left            =   495
         TabIndex        =   46
         Top             =   660
         Width           =   960
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Alinhar à Direita
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Previsto p/ item :"
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
         Left            =   225
         TabIndex        =   45
         Top             =   330
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Execução (Tempo / valor)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1785
      Left            =   4200
      TabIndex        =   39
      Top             =   570
      Width           =   3930
      Begin VB.TextBox Txt_custo_real_lote_exec 
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
         Height          =   315
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor real por lote."
         Top             =   1320
         Width           =   1140
      End
      Begin VB.TextBox Txt_custo_previsto_lote_exec 
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
         Height          =   315
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor previsto por lote."
         Top             =   990
         Width           =   1140
      End
      Begin VB.TextBox Txt_custo_real_peca_exec 
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
         Height          =   315
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor real por peça."
         Top             =   660
         Width           =   1140
      End
      Begin VB.TextBox Txt_custo_previsto_peca_exec 
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
         Height          =   315
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor previsto por peça."
         Top             =   330
         Width           =   1140
      End
      Begin VB.TextBox Txt_tempo_real_lote_exec 
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
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Tempo real por lote."
         Top             =   1320
         Width           =   1050
      End
      Begin VB.TextBox Txt_tempo_previsto_lote_exec 
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
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Tempo previsto por lote."
         Top             =   990
         Width           =   1050
      End
      Begin VB.TextBox Txt_tempo_real_peca_exec 
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
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Tempo real por peça."
         Top             =   660
         Width           =   1050
      End
      Begin VB.TextBox Txt_tempo_previsto_peca_exec 
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
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Tempo previsto por peça."
         Top             =   330
         Width           =   1050
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Alinhar à Direita
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Real do lote :"
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
         Left            =   495
         TabIndex        =   43
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Alinhar à Direita
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Previsto do lote :"
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
         Left            =   225
         TabIndex        =   42
         Top             =   990
         Width           =   1230
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Alinhar à Direita
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Real p/ item :"
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
         Left            =   495
         TabIndex        =   41
         Top             =   660
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Alinhar à Direita
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Previsto p/ item :"
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
         Left            =   225
         TabIndex        =   40
         Top             =   330
         Width           =   1230
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Preparação (Tempo / valor)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1785
      Left            =   240
      TabIndex        =   34
      Top             =   570
      Width           =   3930
      Begin VB.TextBox Txt_tempo_previsto_peca_prep 
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
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Tempo previsto por peça."
         Top             =   330
         Width           =   1050
      End
      Begin VB.TextBox Txt_tempo_real_peca_prep 
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
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Tempo real por peça."
         Top             =   660
         Width           =   1050
      End
      Begin VB.TextBox Txt_tempo_previsto_lote_prep 
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
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Tempo previsto por lote."
         Top             =   990
         Width           =   1050
      End
      Begin VB.TextBox Txt_tempo_real_lote_prep 
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
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Tempo real por lote."
         Top             =   1320
         Width           =   1050
      End
      Begin VB.TextBox Txt_custo_previsto_peca_prep 
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
         Height          =   315
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor previsto por peça."
         Top             =   330
         Width           =   1140
      End
      Begin VB.TextBox Txt_custo_real_peca_prep 
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
         Height          =   315
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor real por peça."
         Top             =   660
         Width           =   1140
      End
      Begin VB.TextBox Txt_custo_previsto_lote_prep 
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
         Height          =   315
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor previsto por lote."
         Top             =   990
         Width           =   1140
      End
      Begin VB.TextBox Txt_custo_real_lote_prep 
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
         Height          =   315
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor real por lote."
         Top             =   1320
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Alinhar à Direita
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Previsto p/ item :"
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
         Left            =   225
         TabIndex        =   38
         Top             =   330
         Width           =   1230
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Alinhar à Direita
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Real p/ item :"
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
         Left            =   495
         TabIndex        =   37
         Top             =   660
         Width           =   960
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Alinhar à Direita
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Previsto do lote :"
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
         Left            =   225
         TabIndex        =   36
         Top             =   990
         Width           =   1230
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Alinhar à Direita
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Real do lote :"
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
         Left            =   495
         TabIndex        =   35
         Top             =   1320
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Custo final"
      ForeColor       =   &H00000080&
      Height          =   945
      Left            =   240
      TabIndex        =   49
      Top             =   2490
      Width           =   9990
      Begin VB.TextBox Txt_custo_outras 
         Alignment       =   1  'Alinhar à Direita
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
         Left            =   5370
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Custo final de outras despesas."
         Top             =   510
         Width           =   1440
      End
      Begin VB.TextBox Txt_custo_total_peca 
         Alignment       =   1  'Alinhar à Direita
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
         Left            =   8290
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Custo total por peça."
         Top             =   510
         Width           =   1515
      End
      Begin VB.CommandButton Cmd_terceiros 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4950
         Picture         =   "FrmProd_Resultados.frx":A14B
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Visualizar pedidos de compra e notas de terceiros."
         Top             =   510
         Width           =   315
      End
      Begin VB.CommandButton Cmd_material 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3075
         Picture         =   "FrmProd_Resultados.frx":A24D
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Visualizar movimentações no estoque."
         Top             =   510
         Width           =   315
      End
      Begin VB.TextBox Txt_custo_mao_de_obra 
         Alignment       =   1  'Alinhar à Direita
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
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Custo final de mão de obra."
         Top             =   510
         Width           =   1430
      End
      Begin VB.TextBox Txt_custo_material 
         Alignment       =   1  'Alinhar à Direita
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
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Custo final de material."
         Top             =   510
         Width           =   1440
      End
      Begin VB.TextBox Txt_custo_terceiros 
         Alignment       =   1  'Alinhar à Direita
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
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Custo final de terceiros."
         Top             =   510
         Width           =   1440
      End
      Begin VB.TextBox Txt_custo_total 
         Alignment       =   1  'Alinhar à Direita
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
         Left            =   6820
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Custo total."
         Top             =   510
         Width           =   1470
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   $"FrmProd_Resultados.frx":A34F
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
         Left            =   360
         TabIndex        =   50
         Top             =   300
         Width           =   9315
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Visual. apontam."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   945
      Left            =   10230
      TabIndex        =   51
      Top             =   2490
      Width           =   1860
      Begin DrawSuite2022.USButton Cmd_preparacao 
         Height          =   285
         Left            =   450
         TabIndex        =   32
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Preparação"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   1
      End
      Begin DrawSuite2022.USButton Cmd_execucao 
         Height          =   285
         Left            =   450
         TabIndex        =   33
         Top             =   570
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Execução"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   1
      End
   End
End
Attribute VB_Name = "FrmProd_Resultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_execucao_Click()
On Error GoTo tratar_erro

Sit_REG = 2
FrmProd_Resultados_Apontamentos.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_material_Click()
On Error GoTo tratar_erro

FrmProd_Resultados_Material.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_preparacao_Click()
On Error GoTo tratar_erro

Sit_REG = 1
FrmProd_Resultados_Apontamentos.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_terceiros_Click()
On Error GoTo tratar_erro

FrmProd_Resultados_Terceiros.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyEscape: Unload Me
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

TTE = 0 'Tempo de preparação previsto p/ peça
TPUSEG = 0 'Tempo de preparação real p/ peça
quantidade = 0 'Tempo de preparação previsto do lote
Encontrada = 0 'Tempo de execucao previsto p/ peça
TotalSegundos = 0 'Tempo de execucao previsto do lote
TEUSEG = 0 'Tempo de execucao real p/ peça
Valor1 = 0 'Tempo de execucao real do lote

QuantsolicitadoN10 = 0
Valor_Cofins_Prod = 0
Qtde = 0
Valor_Cofins_Serv = 0
QuantComprado = 0
Valor_CSLL_Prod = 0
QuantEmpenho = 0
Valor_CSLL_Serv = 0
quantestoque = 0
Valor_INSS_Serv = 0
quantnovo = 0
Valor_IPI = 0
QuantSolicitado = 0
Valor_IRPJ_Prod = 0
QuantsolicitadoN1 = 0
Valor_IRPJ_Serv = 0

Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select * from producao where Ordem = " & OF, Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
  With frmprod
     QTProd = .txtquantreal                                                'ORDEM         QTDE. PREVISTA                                QTDE. OK                                              QT. PROD.(OK+NC)                                                                                         CUSTO LOTE     CUSTO PEÇA  CUSTO TERCEIROS      CUSTO MATERIAL     CUSTO OUTRAS    ORDEM CONSIGNADA
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from ordemservico where Ordem = " & OF & " and Custos = 'True' order by IDProducao", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            valor = IIf(IsNull(TBAbrir!Valor_hs_prep), 0, TBAbrir!Valor_hs_prep)
            valor = valor / 3600
            
            'Tempo de preparação previsto p/ peça
            TempoTotalPrep = IIf(IsNull(TBAbrir!TempoPreparacao), 0, TBAbrir!TempoPreparacao)
            ElapsedTime (TempoTotalPrep)
            If TBOrdem!Quant <> 0 Then TTE = s / TBOrdem!Quant Else TTE = 0
            QuantsolicitadoN10 = QuantsolicitadoN10 + TTE
            s = QuantsolicitadoN10
            Txt_tempo_previsto_peca_prep = FormataTempo(s)
                        
            'Custo de preparação previsto p/ peça
            Valor_Cofins_Prod = Valor_Cofins_Prod + (TTE * valor)
            Txt_custo_previsto_peca_prep = Format(Valor_Cofins_Prod, "###,##0.00")
            
            'Tempo de preparação real p/ peça
            If TBAbrir!Totalprod <> 0 Then TPUSEG = IIf(IsNull(TBAbrir!TPUSEG), 0, TBAbrir!TPUSEG) / IIf(IsNull(TBAbrir!Totalprod), 0, TBAbrir!Totalprod)
            Qtde = Qtde + TPUSEG
            s = Qtde
            Txt_tempo_real_peca_prep = FormataTempo(s)
            
            'Custo de preparação real p/ peça
            Valor_Cofins_Serv = Valor_Cofins_Serv + (TPUSEG * valor)
            Txt_custo_real_peca_prep = Format(Valor_Cofins_Serv, "###,##0.00")
            
            'Tempo de preparação previsto do lote
            TempoTotalPrep = IIf(IsNull(TBAbrir!TempoPreparacao), 0, TBAbrir!TempoPreparacao)
            ElapsedTime (TempoTotalPrep)
            
            'Custo de preparação previsto do lote
            Valor_CSLL_Prod = Valor_CSLL_Prod + (s * valor)
            Txt_custo_previsto_lote_prep = Format(Valor_CSLL_Prod, "###,##0.00")
            
            QuantComprado = QuantComprado + s
            s = QuantComprado
            Txt_tempo_previsto_lote_prep = FormataTempo(s)
           
            
            'Tempo de preparação real do lote
            quantidade = IIf(IsNull(TBAbrir!TPUSEG), 0, TBAbrir!TPUSEG)
            QuantEmpenho = QuantEmpenho + quantidade
            s = QuantEmpenho
            Txt_tempo_real_lote_prep = FormataTempo(s)
            
            'Custo de preparação real do lote
            Valor_CSLL_Serv = Valor_CSLL_Serv + (quantidade * valor)
            Txt_custo_real_lote_prep = Format(Valor_CSLL_Serv, "###,##0.00")
            
            valor = IIf(IsNull(TBAbrir!Valor_hs_exec), 0, TBAbrir!Valor_hs_exec)
            valor = valor / 3600
            
            'Tempo de execução previsto p/ peça
            Encontrada = IIf(IsNull(TBAbrir!TESegundos), 0, TBAbrir!TESegundos)
            quantestoque = quantestoque + Encontrada
            s = quantestoque
            Txt_tempo_previsto_peca_exec = FormataTempo(s)
            
            'Custo de execução previsto p/ peça
            Valor_INSS_Serv = Valor_INSS_Serv + (Encontrada * valor)
            Txt_custo_previsto_peca_exec = Format(Valor_INSS_Serv, "###,##0.00")
            
            'Tempo de execução real p/ peça
            TEUSEG = IIf(IsNull(TBAbrir!TEUSEG), 0, TBAbrir!TEUSEG)
            quantnovo = quantnovo + TEUSEG
            s = quantnovo
            Txt_tempo_real_peca_exec = FormataTempo(s)
            
            'Custo de execução real p/ peça
            Valor_IPI = Valor_IPI + (TEUSEG * valor)
            Txt_custo_real_peca_exec = Format(TBOrdem!CPR, "###,##0.00")
            
            'Tempo de execução previsto do lote
            TotalSegundos = IIf(IsNull(TBAbrir!TESegundos), 0, TBAbrir!TESegundos) * TBOrdem!Quant
            QuantSolicitado = QuantSolicitado + TotalSegundos
            s = QuantSolicitado
            Txt_tempo_previsto_lote_exec = FormataTempo(s)
                        
            'Custo de execução previsto do lote
            Valor_IRPJ_Prod = Valor_IRPJ_Prod + (TotalSegundos * valor)
            Txt_custo_previsto_lote_exec = Format(Valor_IRPJ_Prod, "###,##0.00")
            
            'Tempo de execução real do lote
            Valor1 = IIf(IsNull(TBAbrir!TEUSEG), 0, TBAbrir!TEUSEG) * IIf(IsNull(TBAbrir!Totalprod), 0, TBAbrir!Totalprod)
            QuantsolicitadoN1 = QuantsolicitadoN1 + Valor1
            s = QuantsolicitadoN1
            Txt_tempo_real_lote_exec = FormataTempo(s)
            
            'Custo de execução real do lote
            Valor_IRPJ_Serv = Valor_IRPJ_Serv + (Valor1 * valor)
            Valor_IPI = Format(Valor_IPI, "###,##0.00")
            Txt_custo_real_lote_exec = Format(TBOrdem!CPR * QTProd, "###,##0.00") 'Format(Valor_IRPJ_Serv, "###,##0.00")
            
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
    
    'Total (Tempo/valor)
    Txt_tempo_previsto_peca_total = IIf(IsNull(TBOrdem!TPP), "00:00:00", TBOrdem!TPP)
    Txt_custo_previsto_peca_total = IIf(IsNull(TBOrdem!cpp), "0,00", Format(TBOrdem!cpp, "###,##0.00"))
    s = Qtde + quantnovo
    Txt_tempo_real_peca_total = FormataTempo(s)
    Txt_custo_real_peca_total = IIf(IsNull(TBOrdem!CPR), "0,00", Format(TBOrdem!CPR, "###,##0.00"))
    Txt_tempo_previsto_lote_total = IIf(IsNull(TBOrdem!TTTPrev), "00:00:00", TBOrdem!TTTPrev)
    Txt_custo_previsto_lote_total = IIf(IsNull(TBOrdem!CTTPrev), "0,00", Format(TBOrdem!CTTPrev, "###,##0.00"))
    Txt_tempo_real_lote_total = IIf(IsNull(TBOrdem!TTTReal), "00:00:00", TBOrdem!TTTReal)
    Txt_custo_real_lote_total = Format(TBOrdem!CPR * QTProd, "###,##0.00") 'IIf(IsNull(TBOrdem!CTTReal), "0,00", Format(TBOrdem!CTTReal, "###,##0.00"))
    
    'Custos final
    CustoPeca = IIf(IsNull(TBOrdem!CPR), "0,00", Format(TBOrdem!CPR, "###,##0.00"))
    TotalProduzido = QTProd
    CustototalMO = CustoPeca * TotalProduzido
    
    
    Txt_custo_mao_de_obra = Format(CustototalMO, "###,##0.00") 'IIf(IsNull(TBOrdem!CTTReal), "0,00", Format(TBOrdem!CTTReal, "###,##0.00"))
    
    Txt_custo_material = IIf(IsNull(TBOrdem!CTMaterial), "0,00", Format(TBOrdem!CTMaterial, "###,##0.00"))
    Txt_custo_terceiros = IIf(IsNull(TBOrdem!CTServico), "0,00", Format(TBOrdem!CTServico, "###,##0.00"))
    Txt_custo_outras = IIf(IsNull(TBOrdem!CTOutras), "0,00", Format(TBOrdem!CTOutras, "###,##0.00"))
    Valor1 = Txt_custo_mao_de_obra
    Valor2 = Txt_custo_material
    Valor3 = Txt_custo_terceiros
    ValorConta = Txt_custo_outras
    Txt_custo_total = Format(Valor1 + Valor2 + Valor3 + ValorConta, "###,##0.00")
    
    'Por peça
        Txt_custo_total_peca = (Valor1 + Valor2 + Valor3 + ValorConta) / QTProd 'FunCalculaValorUnitOrdem(TBOrdem!Ordem, IIf(IsNull(TBOrdem!Quant), 0, TBOrdem!Quant), IIf(IsNull(TBOrdem!QuantProd), 0, TBOrdem!QuantProd), IIf(IsNull(TBOrdem!QuantProd), 0, TBOrdem!QuantProd) + IIf(IsNull(TBOrdem!QuantNC), 0, TBOrdem!QuantNC), .txtcustoefet, .cpecareal, Txt_custo_terceiros, Txt_custo_material, Txt_custo_outras, TBOrdem!consignacao)
        Txt_custo_total_peca = Format(Txt_custo_total_peca, "###,##0.00")
    End With
End If
TBOrdem.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
