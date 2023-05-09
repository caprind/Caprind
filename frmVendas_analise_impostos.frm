VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_analise_impostos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Outros - Análise crítica - Fechamento"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   15390
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   15390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTotal 
      BackColor       =   &H00E0E0E0&
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
      Left            =   13770
      TabIndex        =   181
      Top             =   2400
      Width           =   195
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impostos total"
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
      Height          =   5895
      Index           =   4
      Left            =   12285
      TabIndex        =   308
      Top             =   2400
      Width           =   3030
      Begin VB.TextBox Txt_pdespesas_administrativas_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   213
         Text            =   "0"
         ToolTipText     =   "Porcentagem de desepesa administrativa."
         Top             =   3630
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_despesas_administrativas_total 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   214
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da despesa administrativa."
         Top             =   3630
         Width           =   750
      End
      Begin VB.TextBox Txt_pdespesas_comerciais_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   207
         Text            =   "0"
         ToolTipText     =   "Porcentagem de desepesa comercial."
         Top             =   2940
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_despesas_comerciais_total 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   208
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da despesa comercial."
         Top             =   2940
         Width           =   750
      End
      Begin VB.TextBox Txt_psimples_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   201
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   2256
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_simples_total 
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
         Left            =   2145
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   202
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   2256
         Width           =   750
      End
      Begin VB.TextBox Txt_pcofins_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   189
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   912
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_PIS_total 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   187
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   576
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_cofins_total 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   190
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   912
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_despesas_financeiras_total 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   211
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da despesa financeira."
         Top             =   3285
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_ICMS_total 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   184
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   240
         Width           =   750
      End
      Begin VB.TextBox Txt_pICMS_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   183
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   240
         Width           =   510
      End
      Begin VB.TextBox Txt_pPIS_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   186
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   576
         Width           =   510
      End
      Begin VB.TextBox Txt_pCSLL_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   192
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1248
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_CSLL_total 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   193
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   1248
         Width           =   750
      End
      Begin VB.TextBox Txt_pIRPJ_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   198
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1920
         Width           =   510
      End
      Begin VB.TextBox Txt_pISSQN_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   195
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1584
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_ISSQN_total 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   196
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   1584
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_IRPJ_total 
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
         Left            =   2145
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   199
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   1920
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_frete_total 
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
         Left            =   2145
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   217
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do frete."
         Top             =   3960
         Width           =   750
      End
      Begin VB.TextBox Txt_pdespesas_financeiras_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   210
         Text            =   "0"
         ToolTipText     =   "Porcentagem de desepesa financeira."
         Top             =   3285
         Width           =   510
      End
      Begin VB.TextBox Txt_pfrete_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   216
         Text            =   "0"
         ToolTipText     =   "Porcentagem do frete."
         Top             =   3960
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_venda_total 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   225
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da venda."
         Top             =   5430
         Width           =   1455
      End
      Begin VB.TextBox Txt_valor_total_total 
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
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   222
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total."
         Top             =   4710
         Width           =   930
      End
      Begin VB.TextBox Txt_pcomissao_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   204
         Text            =   "0"
         ToolTipText     =   "Porcentagem de comissão."
         Top             =   2592
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_comissao_total 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   205
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da comissão."
         Top             =   2592
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_margem_total 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   220
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da margem."
         Top             =   4290
         Width           =   750
      End
      Begin VB.TextBox Txt_pmargem_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   219
         Text            =   "0"
         ToolTipText     =   "Porcentagem de margem."
         Top             =   4290
         Width           =   510
      End
      Begin VB.TextBox Txt_ptotal_total 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   221
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Porcentagem total."
         Top             =   4710
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_reciploca_total 
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
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   224
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor recíproca/fator."
         Top             =   5070
         Width           =   930
      End
      Begin VB.TextBox Txt_preciploca_total 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   223
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Porcentagem recíproca/fator."
         Top             =   5070
         Width           =   510
      End
      Begin VB.CheckBox Chk_ISSQN_total 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ISSQN"
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
         TabIndex        =   194
         Top             =   1644
         Width           =   1275
      End
      Begin VB.CheckBox Chk_ICMS_total 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ICMS"
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
         TabIndex        =   182
         Top             =   300
         Width           =   1275
      End
      Begin VB.CheckBox Chk_despesas_financeiras_total 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desp. financ."
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
         TabIndex        =   209
         Top             =   3345
         Width           =   1275
      End
      Begin VB.CheckBox Chk_PIS_total 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PIS"
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
         TabIndex        =   185
         Top             =   636
         Width           =   1275
      End
      Begin VB.CheckBox Chk_cofins_total 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cofins"
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
         TabIndex        =   188
         Top             =   960
         Width           =   1275
      End
      Begin VB.CheckBox Chk_CSLL_total 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CSLL"
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
         TabIndex        =   191
         Top             =   1308
         Width           =   1275
      End
      Begin VB.CheckBox Chk_IRPJ_total 
         BackColor       =   &H00E0E0E0&
         Caption         =   "IRPJ"
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
         TabIndex        =   197
         Top             =   1980
         Width           =   1275
      End
      Begin VB.CheckBox Chk_frete_total 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frete"
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
         TabIndex        =   215
         Top             =   4020
         Width           =   1275
      End
      Begin VB.CheckBox Chk_simples_total 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Simples"
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
         TabIndex        =   200
         Top             =   2316
         Width           =   1275
      End
      Begin VB.CheckBox Chk_comissao_total 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comissão"
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
         TabIndex        =   203
         Top             =   2652
         Width           =   1275
      End
      Begin VB.CheckBox Chk_margem_total 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Margem"
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
         TabIndex        =   218
         Top             =   4350
         Width           =   1275
      End
      Begin VB.CheckBox Chk_despesas_comerciais_total 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desp. com."
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
         TabIndex        =   206
         Top             =   3000
         Width           =   1275
      End
      Begin VB.CheckBox Chk_despesas_administrativas_total 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desp. adm."
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
         TabIndex        =   212
         Top             =   3690
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   324
         Top             =   636
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   323
         Top             =   972
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   322
         Top             =   1308
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   82
         Left            =   1950
         TabIndex        =   321
         Top             =   1644
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   81
         Left            =   1950
         TabIndex        =   320
         Top             =   1980
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   319
         Top             =   2316
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   318
         Top             =   2652
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   317
         Top             =   3000
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   77
         Left            =   1950
         TabIndex        =   316
         Top             =   3345
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   315
         Top             =   3690
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   75
         Left            =   1950
         TabIndex        =   314
         Top             =   4020
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   74
         Left            =   1950
         TabIndex        =   313
         Top             =   4350
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   73
         Left            =   1950
         TabIndex        =   312
         Top             =   300
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor venda :"
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
         Index           =   72
         Left            =   285
         TabIndex        =   311
         Top             =   5430
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total : "
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
         Index           =   71
         Left            =   930
         TabIndex        =   310
         Top             =   4710
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recíproca/Fator : "
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
         Index           =   70
         Left            =   135
         TabIndex        =   309
         Top             =   5070
         Width           =   1305
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impostos outros"
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
      Height          =   5895
      Index           =   3
      Left            =   9226
      TabIndex        =   291
      Top             =   2400
      Width           =   3030
      Begin VB.TextBox Txt_pdespesas_administrativas_outros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   168
         Text            =   "0"
         ToolTipText     =   "Porcentagem de desepesa administrativa."
         Top             =   3630
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_despesas_administrativas_outros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   169
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da despesa administrativa."
         Top             =   3630
         Width           =   750
      End
      Begin VB.TextBox Txt_pdespesas_comerciais_outros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   162
         Text            =   "0"
         ToolTipText     =   "Porcentagem de desepesa comercial."
         Top             =   2940
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_despesas_comerciais_outros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   163
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da despesa comercial."
         Top             =   2940
         Width           =   750
      End
      Begin VB.TextBox Txt_psimples_outros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   156
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   2256
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_simples_outros 
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
         Left            =   2145
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   157
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   2256
         Width           =   750
      End
      Begin VB.TextBox Txt_pcofins_outros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   144
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   912
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_PIS_outros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   142
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   576
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_cofins_outros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   145
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   912
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_despesas_financeiras_outros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   166
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da despesa financeira."
         Top             =   3285
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_ICMS_outros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   139
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   240
         Width           =   750
      End
      Begin VB.TextBox Txt_pICMS_outros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   138
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   240
         Width           =   510
      End
      Begin VB.TextBox Txt_pPIS_outros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   141
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   576
         Width           =   510
      End
      Begin VB.TextBox Txt_pCSLL_outros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   147
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1248
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_CSLL_outros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   148
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   1248
         Width           =   750
      End
      Begin VB.TextBox Txt_pIRPJ_outros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   153
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1920
         Width           =   510
      End
      Begin VB.TextBox Txt_pISSQN_outros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   150
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1584
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_ISSQN_outros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   151
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   1584
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_IRPJ_outros 
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
         Left            =   2145
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   154
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   1920
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_frete_outros 
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
         Left            =   2145
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   172
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do frete."
         Top             =   3960
         Width           =   750
      End
      Begin VB.TextBox Txt_pdespesas_financeiras_outros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   165
         Text            =   "0"
         ToolTipText     =   "Porcentagem de desepesa financeira."
         Top             =   3285
         Width           =   510
      End
      Begin VB.TextBox Txt_pfrete_outros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   171
         Text            =   "0"
         ToolTipText     =   "Porcentagem do frete."
         Top             =   3960
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_venda_outros 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   180
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da venda."
         Top             =   5430
         Width           =   1455
      End
      Begin VB.TextBox Txt_valor_total_outros 
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
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   177
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total."
         Top             =   4710
         Width           =   930
      End
      Begin VB.TextBox Txt_pcomissao_outros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   159
         Text            =   "0"
         ToolTipText     =   "Porcentagem de comissão."
         Top             =   2592
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_comissao_outros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   160
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da comissão."
         Top             =   2592
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_margem_outros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   175
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da margem."
         Top             =   4290
         Width           =   750
      End
      Begin VB.TextBox Txt_pmargem_outros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   174
         Text            =   "0"
         ToolTipText     =   "Porcentagem de margem."
         Top             =   4290
         Width           =   510
      End
      Begin VB.TextBox Txt_ptotal_outros 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   176
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Porcentagem total."
         Top             =   4710
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_reciploca_outros 
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
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   179
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor recíproca/fator."
         Top             =   5070
         Width           =   930
      End
      Begin VB.TextBox Txt_preciploca_outros 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   178
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Porcentagem recíproca/fator."
         Top             =   5070
         Width           =   510
      End
      Begin VB.CheckBox Chk_ISSQN_outros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ISSQN"
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
         TabIndex        =   149
         Top             =   1644
         Width           =   1275
      End
      Begin VB.CheckBox Chk_ICMS_outros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ICMS"
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
         TabIndex        =   137
         Top             =   300
         Width           =   1275
      End
      Begin VB.CheckBox Chk_despesas_financeiras_outros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desp. financ."
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
         TabIndex        =   164
         Top             =   3345
         Width           =   1275
      End
      Begin VB.CheckBox Chk_PIS_outros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PIS"
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
         TabIndex        =   140
         Top             =   636
         Width           =   1275
      End
      Begin VB.CheckBox Chk_cofins_outros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cofins"
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
         TabIndex        =   143
         Top             =   960
         Width           =   1275
      End
      Begin VB.CheckBox Chk_CSLL_outros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CSLL"
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
         TabIndex        =   146
         Top             =   1308
         Width           =   1275
      End
      Begin VB.CheckBox Chk_IRPJ_outros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "IRPJ"
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
         TabIndex        =   152
         Top             =   1980
         Width           =   1275
      End
      Begin VB.CheckBox Chk_frete_outros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frete"
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
         TabIndex        =   170
         Top             =   4020
         Width           =   1275
      End
      Begin VB.CheckBox Chk_simples_outros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Simples"
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
         TabIndex        =   155
         Top             =   2316
         Width           =   1275
      End
      Begin VB.CheckBox Chk_comissao_outros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comissão"
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
         TabIndex        =   158
         Top             =   2652
         Width           =   1275
      End
      Begin VB.CheckBox Chk_margem_outros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Margem"
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
         TabIndex        =   173
         Top             =   4350
         Width           =   1275
      End
      Begin VB.CheckBox Chk_despesas_comerciais_outros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desp. com."
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
         TabIndex        =   161
         Top             =   3000
         Width           =   1275
      End
      Begin VB.CheckBox Chk_despesas_administrativas_outros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desp. adm."
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
         TabIndex        =   167
         Top             =   3690
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   54
         Left            =   1950
         TabIndex        =   307
         Top             =   636
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   55
         Left            =   1950
         TabIndex        =   306
         Top             =   972
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   56
         Left            =   1950
         TabIndex        =   305
         Top             =   1308
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   304
         Top             =   1644
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   303
         Top             =   1980
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   302
         Top             =   2316
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   301
         Top             =   2652
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   300
         Top             =   3000
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   62
         Left            =   1950
         TabIndex        =   299
         Top             =   3345
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   298
         Top             =   3690
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   297
         Top             =   4020
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   296
         Top             =   4350
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   295
         Top             =   300
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor venda :"
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
         Index           =   67
         Left            =   285
         TabIndex        =   294
         Top             =   5430
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total : "
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
         Left            =   930
         TabIndex        =   293
         Top             =   4710
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recíproca/Fator : "
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
         Index           =   69
         Left            =   135
         TabIndex        =   292
         Top             =   5070
         Width           =   1305
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impostos terceiros"
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
      Height          =   5895
      Index           =   2
      Left            =   6169
      TabIndex        =   274
      Top             =   2400
      Width           =   3030
      Begin VB.TextBox Txt_preciploca_terceiros 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   134
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Porcentagem recíproca/fator."
         Top             =   5070
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_reciploca_terceiros 
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
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   135
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor recíproca/fator."
         Top             =   5070
         Width           =   930
      End
      Begin VB.TextBox Txt_ptotal_terceiros 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   132
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Porcentagem total."
         Top             =   4710
         Width           =   510
      End
      Begin VB.TextBox Txt_pmargem_terceiros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   130
         Text            =   "0"
         ToolTipText     =   "Porcentagem de margem."
         Top             =   4290
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_margem_terceiros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   131
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da margem."
         Top             =   4290
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_comissao_terceiros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   116
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da comissão."
         Top             =   2592
         Width           =   750
      End
      Begin VB.TextBox Txt_pcomissao_terceiros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   115
         Text            =   "0"
         ToolTipText     =   "Porcentagem de comissão."
         Top             =   2592
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_total_terceiros 
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
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   133
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total."
         Top             =   4710
         Width           =   930
      End
      Begin VB.TextBox Txt_valor_venda_terceiros 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   136
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da venda."
         Top             =   5430
         Width           =   1455
      End
      Begin VB.TextBox Txt_pfrete_terceiros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   127
         Text            =   "0"
         ToolTipText     =   "Porcentagem do frete."
         Top             =   3960
         Width           =   510
      End
      Begin VB.TextBox Txt_pdespesas_financeiras_terceiros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   121
         Text            =   "0"
         ToolTipText     =   "Porcentagem de desepesa financeira."
         Top             =   3285
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_frete_terceiros 
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
         Left            =   2145
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   128
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do frete."
         Top             =   3960
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_IRPJ_terceiros 
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
         Left            =   2145
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   110
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   1920
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_ISSQN_terceiros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   107
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   1584
         Width           =   750
      End
      Begin VB.TextBox Txt_pISSQN_terceiros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   106
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1584
         Width           =   510
      End
      Begin VB.TextBox Txt_pIRPJ_terceiros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   109
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1920
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_CSLL_terceiros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   104
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   1248
         Width           =   750
      End
      Begin VB.TextBox Txt_pCSLL_terceiros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   103
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1248
         Width           =   510
      End
      Begin VB.TextBox Txt_pPIS_terceiros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   97
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   576
         Width           =   510
      End
      Begin VB.TextBox Txt_pICMS_terceiros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   94
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   240
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_ICMS_terceiros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   95
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   240
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_despesas_financeiras_terceiros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   122
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da despesa financeira."
         Top             =   3285
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_cofins_terceiros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   101
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   912
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_PIS_terceiros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   98
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   576
         Width           =   750
      End
      Begin VB.TextBox Txt_pcofins_terceiros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   100
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   912
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_simples_terceiros 
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
         Left            =   2145
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   113
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   2256
         Width           =   750
      End
      Begin VB.TextBox Txt_psimples_terceiros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   112
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   2256
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_despesas_comerciais_terceiros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   119
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da despesa comercial."
         Top             =   2940
         Width           =   750
      End
      Begin VB.TextBox Txt_pdespesas_comerciais_terceiros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   118
         Text            =   "0"
         ToolTipText     =   "Porcentagem de desepesa comercial."
         Top             =   2940
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_despesas_administrativas_terceiros 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   125
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da despesa administrativa."
         Top             =   3630
         Width           =   750
      End
      Begin VB.TextBox Txt_pdespesas_administrativas_terceiros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   124
         Text            =   "0"
         ToolTipText     =   "Porcentagem de desepesa administrativa."
         Top             =   3630
         Width           =   510
      End
      Begin VB.CheckBox Chk_ISSQN_terceiros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ISSQN"
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
         TabIndex        =   105
         Top             =   1644
         Width           =   1275
      End
      Begin VB.CheckBox Chk_ICMS_terceiros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ICMS"
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
         TabIndex        =   93
         Top             =   300
         Width           =   1275
      End
      Begin VB.CheckBox Chk_despesas_financeiras_terceiros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desp. financ."
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
         TabIndex        =   120
         Top             =   3345
         Width           =   1275
      End
      Begin VB.CheckBox Chk_PIS_terceiros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PIS"
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
         TabIndex        =   96
         Top             =   636
         Width           =   1275
      End
      Begin VB.CheckBox Chk_cofins_terceiros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cofins"
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
         TabIndex        =   99
         Top             =   960
         Width           =   1275
      End
      Begin VB.CheckBox Chk_CSLL_terceiros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CSLL"
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
         TabIndex        =   102
         Top             =   1308
         Width           =   1275
      End
      Begin VB.CheckBox Chk_IRPJ_terceiros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "IRPJ"
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
         TabIndex        =   108
         Top             =   1980
         Width           =   1275
      End
      Begin VB.CheckBox Chk_frete_terceiros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frete"
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
         TabIndex        =   126
         Top             =   4020
         Width           =   1275
      End
      Begin VB.CheckBox Chk_simples_terceiros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Simples"
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
         TabIndex        =   111
         Top             =   2316
         Width           =   1275
      End
      Begin VB.CheckBox Chk_comissao_terceiros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comissão"
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
         TabIndex        =   114
         Top             =   2652
         Width           =   1275
      End
      Begin VB.CheckBox Chk_margem_terceiros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Margem"
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
         TabIndex        =   129
         Top             =   4350
         Width           =   1275
      End
      Begin VB.CheckBox Chk_despesas_comerciais_terceiros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desp. com."
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
         TabIndex        =   117
         Top             =   3000
         Width           =   1275
      End
      Begin VB.CheckBox Chk_despesas_administrativas_terceiros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desp. adm."
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
         TabIndex        =   123
         Top             =   3690
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recíproca/Fator : "
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
         Index           =   53
         Left            =   135
         TabIndex        =   290
         Top             =   5070
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total : "
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
         Index           =   52
         Left            =   930
         TabIndex        =   289
         Top             =   4710
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor venda :"
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
         Left            =   285
         TabIndex        =   288
         Top             =   5430
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   287
         Top             =   300
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   286
         Top             =   4350
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   48
         Left            =   1950
         TabIndex        =   285
         Top             =   4020
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   47
         Left            =   1950
         TabIndex        =   284
         Top             =   3690
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   46
         Left            =   1950
         TabIndex        =   283
         Top             =   3345
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   45
         Left            =   1950
         TabIndex        =   282
         Top             =   3000
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   281
         Top             =   2652
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   280
         Top             =   2316
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   279
         Top             =   1980
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   278
         Top             =   1644
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   277
         Top             =   1308
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   276
         Top             =   972
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   275
         Top             =   636
         Width           =   165
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impostos materiais"
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
      Height          =   5895
      Index           =   1
      Left            =   3112
      TabIndex        =   257
      Top             =   2400
      Width           =   3030
      Begin VB.TextBox Txt_preciploca_materiais 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   90
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Porcentagem recíproca/fator."
         Top             =   5070
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_reciploca_materiais 
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
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   91
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor recíproca/fator."
         Top             =   5070
         Width           =   930
      End
      Begin VB.TextBox Txt_ptotal_materiais 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   88
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Porcentagem total."
         Top             =   4710
         Width           =   510
      End
      Begin VB.TextBox Txt_pmargem_materiais 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   86
         Text            =   "0"
         ToolTipText     =   "Porcentagem de margem."
         Top             =   4290
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_margem_materiais 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   87
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da margem."
         Top             =   4290
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_comissao_materiais 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   72
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da comissão."
         Top             =   2592
         Width           =   750
      End
      Begin VB.TextBox Txt_pcomissao_materiais 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   71
         Text            =   "0"
         ToolTipText     =   "Porcentagem de comissão."
         Top             =   2592
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_total_materiais 
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
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total."
         Top             =   4710
         Width           =   930
      End
      Begin VB.TextBox Txt_valor_venda_materiais 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   92
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da venda."
         Top             =   5430
         Width           =   1455
      End
      Begin VB.TextBox Txt_pfrete_materiais 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   83
         Text            =   "0"
         ToolTipText     =   "Porcentagem do frete."
         Top             =   3960
         Width           =   510
      End
      Begin VB.TextBox Txt_pdespesas_financeiras_materiais 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   77
         Text            =   "0"
         ToolTipText     =   "Porcentagem de desepesa financeira."
         Top             =   3285
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_frete_materiais 
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
         Left            =   2145
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   84
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do frete."
         Top             =   3960
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_IRPJ_materiais 
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
         Left            =   2145
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   66
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   1920
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_ISSQN_materiais 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   63
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   1584
         Width           =   750
      End
      Begin VB.TextBox Txt_pISSQN_materiais 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   62
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1584
         Width           =   510
      End
      Begin VB.TextBox Txt_pIRPJ_materiais 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   65
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1920
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_CSLL_materiais 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   1248
         Width           =   750
      End
      Begin VB.TextBox Txt_pCSLL_materiais 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   59
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1248
         Width           =   510
      End
      Begin VB.TextBox Txt_pPIS_materiais 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   53
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   576
         Width           =   510
      End
      Begin VB.TextBox Txt_pICMS_materiais 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   50
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   240
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_ICMS_materiais 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   240
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_despesas_financeiras_materiais 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   78
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da despesa financeira."
         Top             =   3285
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_cofins_materiais 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   912
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_PIS_materiais 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   576
         Width           =   750
      End
      Begin VB.TextBox Txt_pcofins_materiais 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   56
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   912
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_simples_materiais 
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
         Left            =   2145
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   69
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   2256
         Width           =   750
      End
      Begin VB.TextBox Txt_psimples_materiais 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   68
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   2256
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_despesas_comerciais_materiais 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   75
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da despesa comercial."
         Top             =   2940
         Width           =   750
      End
      Begin VB.TextBox Txt_pdespesas_comerciais_materiais 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   74
         Text            =   "0"
         ToolTipText     =   "Porcentagem de desepesa comercial."
         Top             =   2940
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_despesas_administrativas_materiais 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   81
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da despesa administrativa."
         Top             =   3630
         Width           =   750
      End
      Begin VB.TextBox Txt_pdespesas_administrativas_materiais 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   80
         Text            =   "0"
         ToolTipText     =   "Porcentagem de desepesa administrativa."
         Top             =   3630
         Width           =   510
      End
      Begin VB.CheckBox Chk_ICMS_materiais 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ICMS"
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
         TabIndex        =   49
         Top             =   300
         Width           =   1275
      End
      Begin VB.CheckBox Chk_PIS_materiais 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PIS"
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
         TabIndex        =   52
         Top             =   636
         Width           =   1275
      End
      Begin VB.CheckBox Chk_cofins_materiais 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cofins"
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
         TabIndex        =   55
         Top             =   960
         Width           =   1275
      End
      Begin VB.CheckBox Chk_CSLL_materiais 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CSLL"
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
         TabIndex        =   58
         Top             =   1308
         Width           =   1275
      End
      Begin VB.CheckBox Chk_ISSQN_materiais 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ISSQN"
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
         TabIndex        =   61
         Top             =   1644
         Width           =   1275
      End
      Begin VB.CheckBox Chk_IRPJ_materiais 
         BackColor       =   &H00E0E0E0&
         Caption         =   "IRPJ"
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
         TabIndex        =   64
         Top             =   1980
         Width           =   1275
      End
      Begin VB.CheckBox Chk_simples_materiais 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Simples"
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
         TabIndex        =   67
         Top             =   2316
         Width           =   1275
      End
      Begin VB.CheckBox Chk_comissao_materiais 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comissão"
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
         TabIndex        =   70
         Top             =   2652
         Width           =   1275
      End
      Begin VB.CheckBox Chk_despesas_comerciais_materiais 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desp. com."
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
         TabIndex        =   73
         Top             =   3000
         Width           =   1275
      End
      Begin VB.CheckBox Chk_despesas_financeiras_materiais 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desp. financ."
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
         TabIndex        =   76
         Top             =   3345
         Width           =   1275
      End
      Begin VB.CheckBox Chk_despesas_administrativas_materiais 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desp. adm."
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
         TabIndex        =   79
         Top             =   3690
         Width           =   1275
      End
      Begin VB.CheckBox Chk_frete_materiais 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frete"
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
         TabIndex        =   82
         Top             =   4020
         Width           =   1275
      End
      Begin VB.CheckBox Chk_margem_materiais 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Margem"
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
         TabIndex        =   85
         Top             =   4350
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recíproca/Fator : "
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
         Left            =   135
         TabIndex        =   273
         Top             =   5070
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total : "
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
         Left            =   930
         TabIndex        =   272
         Top             =   4710
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor venda :"
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
         Index           =   35
         Left            =   285
         TabIndex        =   271
         Top             =   5430
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   34
         Left            =   1950
         TabIndex        =   270
         Top             =   300
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   269
         Top             =   4350
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   268
         Top             =   4020
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   267
         Top             =   3690
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   266
         Top             =   3345
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   265
         Top             =   3000
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   264
         Top             =   2652
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   263
         Top             =   2316
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   262
         Top             =   1980
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   261
         Top             =   1644
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   260
         Top             =   1308
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   259
         Top             =   972
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   258
         Top             =   636
         Width           =   165
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Optante pelo regime tributário"
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
      Height          =   645
      Left            =   55
      TabIndex        =   229
      Top             =   990
      Width           =   15240
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Height          =   495
         Left            =   2760
         TabIndex        =   325
         Top             =   120
         Width           =   9840
         Begin VB.OptionButton Opt_real 
            BackColor       =   &H00E0E0E0&
            Caption         =   "3 - Lucro real"
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
            Left            =   3840
            TabIndex        =   329
            Top             =   210
            Width           =   1245
         End
         Begin VB.OptionButton Opt_presumido 
            BackColor       =   &H00E0E0E0&
            Caption         =   "2 - Lucro presumido"
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
            Left            =   2025
            TabIndex        =   328
            Top             =   210
            Width           =   1725
         End
         Begin VB.OptionButton Opt_simples 
            BackColor       =   &H00E0E0E0&
            Caption         =   "1 - Simples nacional"
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
            Left            =   150
            TabIndex        =   327
            Top             =   210
            Width           =   1725
         End
         Begin VB.OptionButton Opt_simples1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "4 - Simples nacional (excesso de sublimite de receita bruta)"
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
            Left            =   5220
            TabIndex        =   326
            Top             =   210
            Width           =   4575
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totais (processo + materiais + terceiros + outros = total geral)"
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
      Height          =   675
      Left            =   75
      TabIndex        =   230
      Top             =   1680
      Width           =   15240
      Begin VB.TextBox Txt_total_outros 
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
         Height          =   315
         Left            =   9045
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total de outros."
         Top             =   240
         Width           =   1725
      End
      Begin VB.TextBox Txt_total_terceiros 
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
         Height          =   315
         Left            =   6675
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "0,00"
         ToolTipText     =   "Valor total de terceiros."
         Top             =   240
         Width           =   1725
      End
      Begin VB.TextBox Txt_total_geral 
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
         Height          =   315
         Left            =   11430
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total geral."
         Top             =   240
         Width           =   1725
      End
      Begin VB.TextBox Txt_total_materiais 
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
         Height          =   315
         Left            =   4290
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "0,00"
         ToolTipText     =   "Valor total de materiais."
         Top             =   240
         Width           =   1725
      End
      Begin VB.TextBox Txt_total_processo 
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
         Height          =   315
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   0
         Text            =   "0,00"
         ToolTipText     =   "Valor total do processo."
         Top             =   240
         Width           =   1725
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
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
         Left            =   8655
         TabIndex        =   243
         Top             =   300
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
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
         Left            =   6285
         TabIndex        =   241
         Top             =   300
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "="
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
         Left            =   11025
         TabIndex        =   232
         Top             =   300
         Width           =   135
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
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
         Left            =   3900
         TabIndex        =   231
         Top             =   300
         Width           =   135
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impostos processo"
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
      Height          =   5895
      Index           =   0
      Left            =   55
      TabIndex        =   233
      Top             =   2400
      Width           =   3030
      Begin VB.TextBox Txt_pdespesas_administrativas_processo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   36
         Text            =   "0"
         ToolTipText     =   "Porcentagem de desepesa administrativa."
         Top             =   3630
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_despesas_administrativas_processo 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da despesa administrativa."
         Top             =   3630
         Width           =   750
      End
      Begin VB.TextBox Txt_pdespesas_comerciais_processo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   30
         Text            =   "0"
         ToolTipText     =   "Porcentagem de desepesa comercial."
         Top             =   2940
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_despesas_comerciais_processo 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da despesa comercial."
         Top             =   2940
         Width           =   750
      End
      Begin VB.TextBox Txt_psimples_processo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   24
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   2256
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_simples_processo 
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
         Left            =   2145
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   2256
         Width           =   750
      End
      Begin VB.TextBox Txt_pcofins_processo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   12
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   912
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_PIS_processo 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   576
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_cofins_processo 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   912
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_despesas_financeiras_processo 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da despesa financeira."
         Top             =   3285
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_ICMS_processo 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   240
         Width           =   750
      End
      Begin VB.TextBox Txt_pICMS_processo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   6
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   240
         Width           =   510
      End
      Begin VB.TextBox Txt_pPIS_processo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   9
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   576
         Width           =   510
      End
      Begin VB.TextBox Txt_pCSLL_processo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   15
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1248
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_CSLL_processo 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   1248
         Width           =   750
      End
      Begin VB.TextBox Txt_pIRPJ_processo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   21
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1920
         Width           =   510
      End
      Begin VB.TextBox Txt_pISSQN_processo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   18
         Text            =   "0"
         ToolTipText     =   "Porcentagem de imposto."
         Top             =   1584
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_ISSQN_processo 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   1584
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_IRPJ_processo 
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
         Left            =   2145
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do imposto."
         Top             =   1920
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_frete_processo 
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
         Left            =   2145
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do frete."
         Top             =   3960
         Width           =   750
      End
      Begin VB.TextBox Txt_pdespesas_financeiras_processo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   33
         Text            =   "0"
         ToolTipText     =   "Porcentagem de desepesa financeira."
         Top             =   3285
         Width           =   510
      End
      Begin VB.TextBox Txt_pfrete_processo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   39
         Text            =   "0"
         ToolTipText     =   "Porcentagem do frete."
         Top             =   3960
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_venda_processo 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da venda."
         Top             =   5430
         Width           =   1455
      End
      Begin VB.TextBox Txt_valor_total_processo 
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
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total."
         Top             =   4710
         Width           =   930
      End
      Begin VB.TextBox Txt_pcomissao_processo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   27
         Text            =   "0"
         ToolTipText     =   "Porcentagem de comissão."
         Top             =   2592
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_comissao_processo 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da comissão."
         Top             =   2592
         Width           =   750
      End
      Begin VB.TextBox Txt_valor_margem_processo 
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
         Left            =   2145
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor da margem."
         Top             =   4290
         Width           =   750
      End
      Begin VB.TextBox Txt_pmargem_processo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   1440
         TabIndex        =   42
         Text            =   "0"
         ToolTipText     =   "Porcentagem de margem."
         Top             =   4290
         Width           =   510
      End
      Begin VB.TextBox Txt_ptotal_processo 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Porcentagem total."
         Top             =   4710
         Width           =   510
      End
      Begin VB.TextBox Txt_valor_reciploca_processo 
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
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor recíproca/fator."
         Top             =   5070
         Width           =   930
      End
      Begin VB.TextBox Txt_preciploca_processo 
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Porcentagem recíproca/fator."
         Top             =   5070
         Width           =   510
      End
      Begin VB.CheckBox Chk_ISSQN_processo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ISSQN"
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
         TabIndex        =   17
         Top             =   1644
         Width           =   1275
      End
      Begin VB.CheckBox Chk_ICMS_processo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ICMS"
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
         TabIndex        =   5
         Top             =   300
         Width           =   1275
      End
      Begin VB.CheckBox Chk_despesas_financeiras_processo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desp. financ."
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
         TabIndex        =   32
         Top             =   3345
         Width           =   1275
      End
      Begin VB.CheckBox Chk_PIS_processo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PIS"
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
         TabIndex        =   8
         Top             =   636
         Width           =   1275
      End
      Begin VB.CheckBox Chk_cofins_processo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cofins"
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
         TabIndex        =   11
         Top             =   960
         Width           =   1275
      End
      Begin VB.CheckBox Chk_CSLL_processo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CSLL"
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
         TabIndex        =   14
         Top             =   1308
         Width           =   1275
      End
      Begin VB.CheckBox Chk_IRPJ_processo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "IRPJ"
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
         TabIndex        =   20
         Top             =   1980
         Width           =   1275
      End
      Begin VB.CheckBox Chk_frete_processo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frete"
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
         TabIndex        =   38
         Top             =   4020
         Width           =   1275
      End
      Begin VB.CheckBox Chk_simples_processo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Simples"
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
         TabIndex        =   23
         Top             =   2316
         Width           =   1275
      End
      Begin VB.CheckBox Chk_comissao_processo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comissão"
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
         TabIndex        =   26
         Top             =   2652
         Width           =   1275
      End
      Begin VB.CheckBox Chk_margem_processo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Margem"
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
         TabIndex        =   41
         Top             =   4350
         Width           =   1275
      End
      Begin VB.CheckBox Chk_despesas_comerciais_processo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desp. com."
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
         TabIndex        =   29
         Top             =   3000
         Width           =   1275
      End
      Begin VB.CheckBox Chk_despesas_administrativas_processo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desp. adm."
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
         TabIndex        =   35
         Top             =   3690
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   256
         Top             =   636
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   255
         Top             =   972
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   254
         Top             =   1308
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   253
         Top             =   1644
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   252
         Top             =   1980
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   251
         Top             =   2316
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   250
         Top             =   2652
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   249
         Top             =   3000
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   248
         Top             =   3345
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   247
         Top             =   3690
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   246
         Top             =   4020
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   245
         Top             =   4350
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   1950
         TabIndex        =   244
         Top             =   300
         Width           =   165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor venda :"
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
         Index           =   20
         Left            =   285
         TabIndex        =   236
         Top             =   5430
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total : "
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
         Left            =   930
         TabIndex        =   235
         Top             =   4710
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recíproca/Fator : "
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
         Left            =   135
         TabIndex        =   234
         Top             =   5070
         Width           =   1305
      End
   End
   Begin VB.Frame Frame15 
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
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   4005
      TabIndex        =   237
      Top             =   8310
      Width           =   7380
      Begin VB.TextBox Txt_margem_de_lucro 
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   226
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Margem de lucro."
         Top             =   350
         Width           =   2340
      End
      Begin VB.TextBox Txt_ultimo_valor 
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
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   227
         TabStop         =   0   'False
         Text            =   "0,00000"
         ToolTipText     =   "Último valor total da venda."
         Top             =   350
         Width           =   2340
      End
      Begin VB.TextBox Txt_valor_total_venda 
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
         Left            =   4920
         TabIndex        =   228
         Text            =   "0,00000"
         ToolTipText     =   "Valor total da venda."
         Top             =   350
         Width           =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Margem de lucro"
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
         Left            =   758
         TabIndex        =   240
         Top             =   150
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Último valor total da venda"
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
         Left            =   2753
         TabIndex        =   239
         Top             =   150
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total da venda"
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
         Left            =   5235
         TabIndex        =   238
         Top             =   150
         Width           =   1710
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   13410
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmVendas_analise_impostos.frx":0000
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   242
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
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
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonAlignment3=   2
      ButtonType3     =   1
      ButtonStyle3    =   -1
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState3    =   -1
      ButtonLeft3     =   83
      ButtonTop3      =   4
      ButtonWidth3    =   2
      ButtonHeight3   =   54
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Ajuda"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Ajuda (F1)"
      ButtonKey4      =   "3"
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
      ButtonWidth4    =   36
      ButtonHeight4   =   21
      ButtonCaption5  =   "Sair"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Sair (Esc)"
      ButtonKey5      =   "4"
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
      ButtonLeft5     =   125
      ButtonTop5      =   2
      ButtonWidth5    =   26
      ButtonHeight5   =   21
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState6    =   5
      ButtonLeft6     =   153
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
   End
End
Attribute VB_Name = "frmVendas_analise_impostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Margem_Reciproca As Boolean

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_cofins_outros_Click()
On Error GoTo tratar_erro

With Txt_pcofins_outros
    If Chk_cofins_outros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_cofins_total_Click()
On Error GoTo tratar_erro

With Txt_pcofins_total
    If Chk_cofins_total.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_comissao_outros_Click()
On Error GoTo tratar_erro

With Txt_pcomissao_outros
    If Chk_comissao_outros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_comissao_total_Click()
On Error GoTo tratar_erro

With Txt_pcomissao_total
    If Chk_comissao_total.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_CSLL_outros_Click()
On Error GoTo tratar_erro

With Txt_pCSLL_outros
    If Chk_CSLL_outros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_CSLL_total_Click()
On Error GoTo tratar_erro

With Txt_pCSLL_total
    If Chk_CSLL_total.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_despesas_administrativas_outros_Click()
On Error GoTo tratar_erro

With Txt_pdespesas_administrativas_outros
    If Chk_despesas_administrativas_outros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_despesas_administrativas_total_Click()
On Error GoTo tratar_erro

With Txt_pdespesas_administrativas_total
    If Chk_despesas_administrativas_total.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_despesas_comerciais_outros_Click()
On Error GoTo tratar_erro

With Txt_pdespesas_comerciais_outros
    If Chk_despesas_comerciais_outros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_despesas_comerciais_total_Click()
On Error GoTo tratar_erro

With Txt_pdespesas_comerciais_total
    If Chk_despesas_comerciais_total.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_despesas_financeiras_outros_Click()
On Error GoTo tratar_erro

With Txt_pdespesas_financeiras_outros
    If Chk_despesas_financeiras_outros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_despesas_financeiras_total_Click()
On Error GoTo tratar_erro

With Txt_pdespesas_financeiras_total
    If Chk_despesas_financeiras_total.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_frete_outros_Click()
On Error GoTo tratar_erro

With Txt_pfrete_outros
    If Chk_frete_outros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_frete_total_Click()
On Error GoTo tratar_erro

With Txt_pfrete_total
    If Chk_frete_total.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ICMS_outros_Click()
On Error GoTo tratar_erro

With Txt_pICMS_outros
    If Chk_ICMS_outros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ICMS_total_Click()
On Error GoTo tratar_erro

With Txt_pICMS_total
    If Chk_ICMS_total.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_IRPJ_outros_Click()
On Error GoTo tratar_erro

With Txt_pIRPJ_outros
    If Chk_IRPJ_outros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_IRPJ_total_Click()
On Error GoTo tratar_erro

With Txt_pIRPJ_total
    If Chk_IRPJ_total.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ISSQN_outros_Click()
On Error GoTo tratar_erro

With Txt_pISSQN_outros
    If Chk_ISSQN_outros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ISSQN_total_Click()
On Error GoTo tratar_erro

With Txt_pISSQN_total
    If Chk_ISSQN_total.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_margem_outros_Click()
On Error GoTo tratar_erro

With Txt_pmargem_outros
    If Chk_margem_outros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_margem_total_Click()
On Error GoTo tratar_erro

With Txt_pmargem_total
    If Chk_margem_total.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_PIS_outros_Click()
On Error GoTo tratar_erro

With Txt_pPIS_outros
    If Chk_PIS_outros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_PIS_total_Click()
On Error GoTo tratar_erro

With Txt_pPIS_total
    If Chk_PIS_total.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_simples_outros_Click()
On Error GoTo tratar_erro

With Txt_psimples_outros
    If Chk_simples_outros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_simples_total_Click()
On Error GoTo tratar_erro

With Txt_psimples_total
    If Chk_simples_total.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkTotal_Click()
On Error GoTo tratar_erro
  
ProcLiberaBloqueia

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF3: ProcGravar
    Case vbKeyF4: ProcExcluir
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15192, 6, True

Set TBCiclo = CreateObject("adodb.recordset")
TBCiclo.Open "Select MargemAnalise from empresa where MargemAnalise = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBCiclo.EOF = False Then
    Margem_Reciproca = True
Else
    Margem_Reciproca = False
End If

ProcLimpaCampos
ProcPuxaDadosImpostos
ProcPuxaDadosValores

'Verifica último valor total da venda
With frmVendas_analise
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Valor_total from Vendas_analise where codinterno = '" & .txtdesenho & "' and ID <> " & .txtId & " and Fechada = 'True' order by ID desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Txt_ultimo_valor = Format(TBAbrir!Valor_total, "###,##0.0000000000")
    End If
    TBAbrir.Close
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Opt_simples.Value = False
Opt_presumido.Value = False
Opt_real.Value = False

With frmVendas_analise
    QuantSolicitado = IIf(.Txt_qtde_sol = "", 0, .Txt_qtde_sol)
    
    valor = 0
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select Sum(VlrTotal) as Valor from Vendas_analise_setores where IDanalise = " & .txtId & " and Setor = 'PROCESSOS'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        valor = IIf(IsNull(TBCFOP!valor), 0, TBCFOP!valor)
    End If
    If valor = 0 Then
        Txt_total_processo.Locked = False
        Txt_total_processo.TabStop = True
        Txt_total_processo = "0,00"
    Else
        Txt_total_processo = Format(valor, "###,##0.00")
        Txt_total_processo.Locked = True
        Txt_total_processo.TabStop = False
    End If
    
    valor = 0
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select Sum(VlrTotal) as Valor from Vendas_analise_setores where IDanalise = " & .txtId & " and (Setor = 'ENGENHARIA' or Setor = 'QUALIDADE' or Setor = 'FERRAMENTAS') and tipo = 'M'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        valor = IIf(IsNull(TBCFOP!valor), 0, TBCFOP!valor)
    End If
    If valor = 0 Then
        'Txt_total_materiais.Locked = False
        'Txt_total_materiais.TabStop = True
        Txt_total_materiais = "0,00"
    Else
        If QuantSolicitado <> 0 Then Txt_total_materiais = Format(valor / QuantSolicitado, "###,##0.00") Else Txt_total_materiais = Format(valor, "###,##0.00")
        'Txt_total_materiais.Locked = True
        'Txt_total_materiais.TabStop = False
    End If
    
    valor = 0
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select Sum(VlrTotal) as Valor from Vendas_analise_setores where IDanalise = " & .txtId & " and Setor = 'ENGENHARIA' and tipo = 'T'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        valor = IIf(IsNull(TBCFOP!valor), 0, TBCFOP!valor)
    End If
    If valor = 0 Then
        Txt_total_terceiros.Locked = False
        Txt_total_terceiros.TabStop = True
        Txt_total_terceiros = "0,00"
    Else
        If QuantSolicitado <> 0 Then Txt_total_terceiros = Format(valor / QuantSolicitado, "###,##0.00") Else Txt_total_terceiros = Format(valor, "###,##0.00")
        Txt_total_terceiros.Locked = True
        Txt_total_terceiros.TabStop = False
    End If
    
    valor = 0
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select Sum(VlrTotal) as Valor from Vendas_analise_setores where IDanalise = " & .txtId & " and Setor = 'ENGENHARIA' and tipo = 'O'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        valor = IIf(IsNull(TBCFOP!valor), 0, TBCFOP!valor)
    End If
    TBCFOP.Close
    If valor = 0 Then
        Txt_total_outros.Locked = False
        Txt_total_outros.TabStop = True
        Txt_total_outros = "0,00"
    Else
        If QuantSolicitado <> 0 Then Txt_total_outros = Format(valor / QuantSolicitado, "###,##0.00") Else Txt_total_outros = Format(valor, "###,##0.00")
        Txt_total_outros.Locked = True
        Txt_total_outros.TabStop = False
    End If
    ProcCalculaTotal
End With

'Processo
Chk_ICMS_processo.Value = 0
Chk_PIS_processo.Value = 0
Chk_cofins_processo.Value = 0
Chk_CSLL_processo.Value = 0
Chk_ISSQN_processo.Value = 0
Chk_IRPJ_processo.Value = 0
Chk_simples_processo.Value = 0
Chk_comissao_processo.Value = 0
Chk_despesas_comerciais_processo.Value = 0
Chk_despesas_financeiras_processo.Value = 0
Chk_despesas_administrativas_processo.Value = 0
Chk_frete_processo.Value = 0
Chk_margem_processo.Value = 0

Txt_pICMS_processo = 0
Txt_pPIS_processo = 0
Txt_pcofins_processo = 0
Txt_pCSLL_processo = 0
Txt_pISSQN_processo = 0
Txt_pIRPJ_processo = 0
Txt_psimples_processo = 0
Txt_pcomissao_processo = 0
Txt_pdespesas_comerciais_processo = 0
Txt_pdespesas_financeiras_processo = 0
Txt_pdespesas_administrativas_processo = 0
Txt_pfrete_processo = 0
Txt_pmargem_processo = 0
Txt_ptotal_processo = 0
Txt_preciploca_processo = 0

Txt_valor_ICMS_processo = "0,00"
Txt_valor_PIS_processo = "0,00"
Txt_valor_cofins_processo = "0,00"
Txt_valor_CSLL_processo = "0,00"
Txt_valor_ISSQN_processo = "0,00"
Txt_valor_IRPJ_processo = "0,00"
Txt_valor_simples_processo = "0,00"
Txt_valor_comissao_processo = "0,00"
Txt_valor_despesas_comerciais_processo = "0,00"
Txt_valor_despesas_financeiras_processo = "0,00"
Txt_valor_despesas_administrativas_processo = "0,00"
Txt_valor_frete_processo = "0,00"
Txt_valor_margem_processo = "0,00"
Txt_valor_total_processo = "0,00"
Txt_valor_reciploca_processo = "0,00"
Txt_valor_venda_processo = "0,00"

'Materiais
Chk_ICMS_materiais.Value = 0
Chk_PIS_materiais.Value = 0
Chk_cofins_materiais.Value = 0
Chk_CSLL_materiais.Value = 0
Chk_ISSQN_materiais.Value = 0
Chk_IRPJ_materiais.Value = 0
Chk_simples_materiais.Value = 0
Chk_comissao_materiais.Value = 0
Chk_despesas_comerciais_materiais.Value = 0
Chk_despesas_financeiras_materiais.Value = 0
Chk_despesas_administrativas_materiais.Value = 0
Chk_frete_materiais.Value = 0
Chk_margem_materiais.Value = 0

Txt_pICMS_materiais = 0
Txt_pPIS_materiais = 0
Txt_pcofins_materiais = 0
Txt_pCSLL_materiais = 0
Txt_pISSQN_materiais = 0
Txt_pIRPJ_materiais = 0
Txt_psimples_materiais = 0
Txt_pcomissao_materiais = 0
Txt_pdespesas_comerciais_materiais = 0
Txt_pdespesas_financeiras_materiais = 0
Txt_pdespesas_administrativas_materiais = 0
Txt_pfrete_materiais = 0
Txt_pmargem_materiais = 0
Txt_ptotal_materiais = 0
Txt_preciploca_materiais = 0

Txt_valor_ICMS_materiais = "0,00"
Txt_valor_PIS_materiais = "0,00"
Txt_valor_cofins_materiais = "0,00"
Txt_valor_CSLL_materiais = "0,00"
Txt_valor_ISSQN_materiais = "0,00"
Txt_valor_IRPJ_materiais = "0,00"
Txt_valor_simples_materiais = "0,00"
Txt_valor_comissao_materiais = "0,00"
Txt_valor_despesas_comerciais_materiais = "0,00"
Txt_valor_despesas_financeiras_materiais = "0,00"
Txt_valor_despesas_administrativas_materiais = "0,00"
Txt_valor_frete_materiais = "0,00"
Txt_valor_margem_materiais = "0,00"
Txt_valor_total_materiais = "0,00"
Txt_valor_reciploca_materiais = "0,00"
Txt_valor_venda_materiais = "0,00"

'Terceiros
Chk_ICMS_terceiros.Value = 0
Chk_PIS_terceiros.Value = 0
Chk_cofins_terceiros.Value = 0
Chk_CSLL_terceiros.Value = 0
Chk_ISSQN_terceiros.Value = 0
Chk_IRPJ_terceiros.Value = 0
Chk_simples_terceiros.Value = 0
Chk_comissao_terceiros.Value = 0
Chk_despesas_comerciais_terceiros.Value = 0
Chk_despesas_financeiras_terceiros.Value = 0
Chk_despesas_administrativas_terceiros.Value = 0
Chk_frete_terceiros.Value = 0
Chk_margem_terceiros.Value = 0

Txt_pICMS_terceiros = 0
Txt_pPIS_terceiros = 0
Txt_pcofins_terceiros = 0
Txt_pCSLL_terceiros = 0
Txt_pISSQN_terceiros = 0
Txt_pIRPJ_terceiros = 0
Txt_psimples_terceiros = 0
Txt_pcomissao_terceiros = 0
Txt_pdespesas_comerciais_terceiros = 0
Txt_pdespesas_financeiras_terceiros = 0
Txt_pdespesas_administrativas_terceiros = 0
Txt_pfrete_terceiros = 0
Txt_pmargem_terceiros = 0
Txt_ptotal_terceiros = 0
Txt_preciploca_terceiros = 0

Txt_valor_ICMS_terceiros = "0,00"
Txt_valor_PIS_terceiros = "0,00"
Txt_valor_cofins_terceiros = "0,00"
Txt_valor_CSLL_terceiros = "0,00"
Txt_valor_ISSQN_terceiros = "0,00"
Txt_valor_IRPJ_terceiros = "0,00"
Txt_valor_simples_terceiros = "0,00"
Txt_valor_comissao_terceiros = "0,00"
Txt_valor_despesas_comerciais_terceiros = "0,00"
Txt_valor_despesas_financeiras_terceiros = "0,00"
Txt_valor_despesas_administrativas_terceiros = "0,00"
Txt_valor_frete_terceiros = "0,00"
Txt_valor_margem_terceiros = "0,00"
Txt_valor_total_terceiros = "0,00"
Txt_valor_reciploca_terceiros = "0,00"
Txt_valor_venda_terceiros = "0,00"

'Outros
Chk_ICMS_outros.Value = 0
Chk_PIS_outros.Value = 0
Chk_cofins_outros.Value = 0
Chk_CSLL_outros.Value = 0
Chk_ISSQN_outros.Value = 0
Chk_IRPJ_outros.Value = 0
Chk_simples_outros.Value = 0
Chk_comissao_outros.Value = 0
Chk_despesas_comerciais_outros.Value = 0
Chk_despesas_financeiras_outros.Value = 0
Chk_despesas_administrativas_outros.Value = 0
Chk_frete_outros.Value = 0
Chk_margem_outros.Value = 0

Txt_pICMS_outros = 0
Txt_pPIS_outros = 0
Txt_pcofins_outros = 0
Txt_pCSLL_outros = 0
Txt_pISSQN_outros = 0
Txt_pIRPJ_outros = 0
Txt_psimples_outros = 0
Txt_pcomissao_outros = 0
Txt_pdespesas_comerciais_outros = 0
Txt_pdespesas_financeiras_outros = 0
Txt_pdespesas_administrativas_outros = 0
Txt_pfrete_outros = 0
Txt_pmargem_outros = 0
Txt_ptotal_outros = 0
Txt_preciploca_outros = 0

Txt_valor_ICMS_outros = "0,00"
Txt_valor_PIS_outros = "0,00"
Txt_valor_cofins_outros = "0,00"
Txt_valor_CSLL_outros = "0,00"
Txt_valor_ISSQN_outros = "0,00"
Txt_valor_IRPJ_outros = "0,00"
Txt_valor_simples_outros = "0,00"
Txt_valor_comissao_outros = "0,00"
Txt_valor_despesas_comerciais_outros = "0,00"
Txt_valor_despesas_financeiras_outros = "0,00"
Txt_valor_despesas_administrativas_outros = "0,00"
Txt_valor_frete_outros = "0,00"
Txt_valor_margem_outros = "0,00"
Txt_valor_total_outros = "0,00"
Txt_valor_reciploca_outros = "0,00"
Txt_valor_venda_outros = "0,00"

'Total
Chk_ICMS_total.Value = 0
Chk_PIS_total.Value = 0
Chk_cofins_total.Value = 0
Chk_CSLL_total.Value = 0
Chk_ISSQN_total.Value = 0
Chk_IRPJ_total.Value = 0
Chk_simples_total.Value = 0
Chk_comissao_total.Value = 0
Chk_despesas_comerciais_total.Value = 0
Chk_despesas_financeiras_total.Value = 0
Chk_despesas_administrativas_total.Value = 0
Chk_frete_total.Value = 0
Chk_margem_total.Value = 0

Txt_pICMS_total = 0
Txt_pPIS_total = 0
Txt_pcofins_total = 0
Txt_pCSLL_total = 0
Txt_pISSQN_total = 0
Txt_pIRPJ_total = 0
Txt_psimples_total = 0
Txt_pcomissao_total = 0
Txt_pdespesas_comerciais_total = 0
Txt_pdespesas_financeiras_total = 0
Txt_pdespesas_administrativas_total = 0
Txt_pfrete_total = 0
Txt_pmargem_total = 0
Txt_ptotal_total = 0
Txt_preciploca_total = 0

Txt_valor_ICMS_total = "0,00"
Txt_valor_PIS_total = "0,00"
Txt_valor_cofins_total = "0,00"
Txt_valor_CSLL_total = "0,00"
Txt_valor_ISSQN_total = "0,00"
Txt_valor_IRPJ_total = "0,00"
Txt_valor_simples_total = "0,00"
Txt_valor_comissao_total = "0,00"
Txt_valor_despesas_comerciais_total = "0,00"
Txt_valor_despesas_financeiras_total = "0,00"
Txt_valor_despesas_administrativas_total = "0,00"
Txt_valor_frete_total = "0,00"
Txt_valor_margem_total = "0,00"
Txt_valor_total_total = "0,00"
Txt_valor_reciploca_total = "0,00"
Txt_valor_venda_total = "0,00"

Set TBCiclo = CreateObject("adodb.recordset")
TBCiclo.Open "Select * from Empresa", Conexao, adOpenKeyset, adLockOptimistic
If TBCiclo.EOF = False Then
    If TBCiclo!Simples = True Then Opt_simples.Value = True
    If TBCiclo!Presumido = True Then Opt_presumido.Value = True
    If TBCiclo!Real = True Then Opt_real.Value = True
    If TBCiclo!Simples1 = True Then Opt_simples1.Value = True
End If
    
Set TBCiclo = CreateObject("adodb.recordset")
TBCiclo.Open "Select * from impostos", Conexao, adOpenKeyset, adLockOptimistic
If TBCiclo.EOF = False Then
    If IsNull(TBCiclo!PIS_produtos) = False Then
        If chkTotal.Value = 0 Then
            Chk_PIS_processo.Value = 1
            Chk_PIS_materiais.Value = 1
            Chk_PIS_terceiros.Value = 1
            Chk_PIS_outros.Value = 1
            Txt_pPIS_processo.Enabled = True
            Txt_pPIS_materiais.Enabled = True
            Txt_pPIS_terceiros.Enabled = True
            Txt_pPIS_outros.Enabled = True
            Txt_pPIS_processo = TBCiclo!PIS_produtos
            Txt_pPIS_materiais = TBCiclo!PIS_produtos
            Txt_pPIS_terceiros = TBCiclo!PIS_produtos
            Txt_pPIS_outros = TBCiclo!PIS_produtos
        Else
            Chk_PIS_total.Value = 1
            Txt_pPIS_total.Enabled = True
            Txt_pPIS_total = TBCiclo!PIS_produtos
        End If
    End If
    If IsNull(TBCiclo!Cofins_produtos) = False Then
        If chkTotal.Value = 0 Then
            Chk_cofins_processo.Value = 1
            Chk_cofins_materiais.Value = 1
            Chk_cofins_terceiros.Value = 1
            Chk_cofins_outros.Value = 1
            Txt_pcofins_processo.Enabled = True
            Txt_pcofins_materiais.Enabled = True
            Txt_pcofins_terceiros.Enabled = True
            Txt_pcofins_outros.Enabled = True
            Txt_pcofins_processo = TBCiclo!Cofins_produtos
            Txt_pcofins_materiais = TBCiclo!Cofins_produtos
            Txt_pcofins_terceiros = TBCiclo!Cofins_produtos
            Txt_pcofins_outros = TBCiclo!Cofins_produtos
        Else
            Chk_cofins_total.Value = 1
            Txt_pcofins_total.Enabled = True
            Txt_pcofins_total = TBCiclo!Cofins_produtos
        End If
    End If
    If IsNull(TBCiclo!CSLL_produtos) = False Then
        If chkTotal.Value = 0 Then
            Chk_CSLL_processo.Value = 1
            Chk_CSLL_materiais.Value = 1
            Chk_CSLL_terceiros.Value = 1
            Chk_CSLL_outros.Value = 1
            Txt_pCSLL_processo.Enabled = True
            Txt_pCSLL_materiais.Enabled = True
            Txt_pCSLL_terceiros.Enabled = True
            Txt_pCSLL_outros.Enabled = True
            Txt_pCSLL_processo = TBCiclo!CSLL_produtos
            Txt_pCSLL_materiais = TBCiclo!CSLL_produtos
            Txt_pCSLL_terceiros = TBCiclo!CSLL_produtos
            Txt_pCSLL_outros = TBCiclo!CSLL_produtos
        Else
            Chk_CSLL_total.Value = 1
            Txt_pCSLL_total.Enabled = True
            Txt_pCSLL_total = TBCiclo!CSLL_produtos
        End If
    End If
    If IsNull(TBCiclo!IRPJ_produtos) = False Then
        If chkTotal.Value = 0 Then
            Chk_IRPJ_processo.Value = 1
            Chk_IRPJ_materiais.Value = 1
            Chk_IRPJ_terceiros.Value = 1
            Chk_IRPJ_outros.Value = 1
            Txt_pIRPJ_processo.Enabled = True
            Txt_pIRPJ_materiais.Enabled = True
            Txt_pIRPJ_terceiros.Enabled = True
            Txt_pIRPJ_outros.Enabled = True
            Txt_pIRPJ_processo = TBCiclo!IRPJ_produtos
            Txt_pIRPJ_materiais = TBCiclo!IRPJ_produtos
            Txt_pIRPJ_terceiros = TBCiclo!IRPJ_produtos
            Txt_pIRPJ_outros = TBCiclo!IRPJ_produtos
        Else
            Chk_IRPJ_total.Value = 1
            Txt_pIRPJ_total.Enabled = True
            Txt_pIRPJ_total = TBCiclo!IRPJ_produtos
        End If
    End If
    ProcCalculaImpostos
End If
TBCiclo.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDadosImpostos()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Vendas_analise where ID = " & frmVendas_analise.txtId & " and Fechada = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If TBAbrir!chkTotal = False Or IsNull(TBAbrir!chkTotal) = True Then
        chkTotal.Value = 0
        'Processo
        If TBAbrir!Opt_ICMS_processo = True Then Chk_ICMS_processo.Value = 1 Else Chk_ICMS_processo.Value = 0
        If TBAbrir!Opt_PIS_processo = True Then Chk_PIS_processo.Value = 1 Else Chk_PIS_processo.Value = 0
        If TBAbrir!Opt_cofins_processo = True Then Chk_cofins_processo.Value = 1 Else Chk_cofins_processo.Value = 0
        If TBAbrir!Opt_CSLL_processo = True Then Chk_CSLL_processo.Value = 1 Else Chk_CSLL_processo.Value = 0
        If TBAbrir!Opt_ISSQN_processo = True Then Chk_ISSQN_processo.Value = 1 Else Chk_ISSQN_processo.Value = 0
        If TBAbrir!Opt_IRPJ_processo = True Then Chk_IRPJ_processo.Value = 1 Else Chk_IRPJ_processo.Value = 0
        If TBAbrir!Opt_simples_processo = True Then Chk_simples_processo.Value = 1 Else Chk_simples_processo.Value = 0
        If TBAbrir!Opt_comissao_processo = True Then Chk_comissao_processo.Value = 1 Else Chk_comissao_processo.Value = 0
        If TBAbrir!Opt_despesas_comerciais_processo = True Then Chk_despesas_comerciais_processo.Value = 1 Else Chk_despesas_comerciais_processo.Value = 0
        If TBAbrir!Opt_despesas_financeiras_processo = True Then Chk_despesas_financeiras_processo.Value = 1 Else Chk_despesas_financeiras_processo.Value = 0
        If TBAbrir!Opt_despesas_administrativas_processo = True Then Chk_despesas_administrativas_processo.Value = 1 Else Chk_despesas_administrativas_processo.Value = 0
        If TBAbrir!Opt_frete_processo = True Then Chk_frete_processo.Value = 1 Else Chk_frete_processo.Value = 0
        If TBAbrir!Opt_margem_processo = True Then Chk_margem_processo.Value = 1 Else Chk_margem_processo.Value = 0
        Txt_pICMS_processo = IIf(IsNull(TBAbrir!ICMS_processo), 0, TBAbrir!ICMS_processo)
        Txt_pPIS_processo = IIf(IsNull(TBAbrir!PIS_processo), 0, TBAbrir!PIS_processo)
        Txt_pcofins_processo = IIf(IsNull(TBAbrir!Cofins_processo), 0, TBAbrir!Cofins_processo)
        Txt_pCSLL_processo = IIf(IsNull(TBAbrir!CSLL_processo), 0, TBAbrir!CSLL_processo)
        Txt_pISSQN_processo = IIf(IsNull(TBAbrir!ISSQN_processo), 0, TBAbrir!ISSQN_processo)
        Txt_pIRPJ_processo = IIf(IsNull(TBAbrir!IRPJ_processo), 0, TBAbrir!IRPJ_processo)
        Txt_psimples_processo = IIf(IsNull(TBAbrir!Simples_processo), 0, TBAbrir!Simples_processo)
        Txt_pcomissao_processo = IIf(IsNull(TBAbrir!Comissao_processo), 0, TBAbrir!Comissao_processo)
        Txt_pdespesas_comerciais_processo = IIf(IsNull(TBAbrir!Despesas_comerciais_processo), 0, TBAbrir!Despesas_comerciais_processo)
        Txt_pdespesas_financeiras_processo = IIf(IsNull(TBAbrir!Despesas_financeiras_processo), 0, TBAbrir!Despesas_financeiras_processo)
        Txt_pdespesas_administrativas_processo = IIf(IsNull(TBAbrir!Despesas_administrativas_processo), 0, TBAbrir!Despesas_administrativas_processo)
        Txt_pfrete_processo = IIf(IsNull(TBAbrir!Frete_processo), 0, TBAbrir!Frete_processo)
        Txt_pmargem_processo = IIf(IsNull(TBAbrir!Margem_processo), 0, TBAbrir!Margem_processo)
        
        'Materiais
        If TBAbrir!Opt_ICMS_materiais = True Then Chk_ICMS_materiais.Value = 1 Else Chk_ICMS_materiais.Value = 0
        If TBAbrir!Opt_PIS_materiais = True Then Chk_PIS_materiais.Value = 1 Else Chk_PIS_materiais.Value = 0
        If TBAbrir!Opt_cofins_materiais = True Then Chk_cofins_materiais.Value = 1 Else Chk_cofins_materiais.Value = 0
        If TBAbrir!Opt_CSLL_materiais = True Then Chk_CSLL_materiais.Value = 1 Else Chk_CSLL_materiais.Value = 0
        If TBAbrir!Opt_ISSQN_materiais = True Then Chk_ISSQN_materiais.Value = 1 Else Chk_ISSQN_materiais.Value = 0
        If TBAbrir!Opt_IRPJ_materiais = True Then Chk_IRPJ_materiais.Value = 1 Else Chk_IRPJ_materiais.Value = 0
        If TBAbrir!Opt_simples_materiais = True Then Chk_simples_materiais.Value = 1 Else Chk_simples_materiais.Value = 0
        If TBAbrir!Opt_comissao_materiais = True Then Chk_comissao_materiais.Value = 1 Else Chk_comissao_materiais.Value = 0
        If TBAbrir!Opt_despesas_comerciais_materiais = True Then Chk_despesas_comerciais_materiais.Value = 1 Else Chk_despesas_comerciais_materiais.Value = 0
        If TBAbrir!Opt_despesas_financeiras_materiais = True Then Chk_despesas_financeiras_materiais.Value = 1 Else Chk_despesas_financeiras_materiais.Value = 0
        If TBAbrir!Opt_despesas_administrativas_materiais = True Then Chk_despesas_administrativas_materiais.Value = 1 Else Chk_despesas_administrativas_materiais.Value = 0
        If TBAbrir!Opt_frete_materiais = True Then Chk_frete_materiais.Value = 1 Else Chk_frete_materiais.Value = 0
        If TBAbrir!Opt_margem_materiais = True Then Chk_margem_materiais.Value = 1 Else Chk_margem_materiais.Value = 0
        Txt_pICMS_materiais = IIf(IsNull(TBAbrir!ICMS_materiais), 0, TBAbrir!ICMS_materiais)
        Txt_pPIS_materiais = IIf(IsNull(TBAbrir!PIS_materiais), 0, TBAbrir!PIS_materiais)
        Txt_pcofins_materiais = IIf(IsNull(TBAbrir!Cofins_materiais), 0, TBAbrir!Cofins_materiais)
        Txt_pCSLL_materiais = IIf(IsNull(TBAbrir!CSLL_materiais), 0, TBAbrir!CSLL_materiais)
        Txt_pISSQN_materiais = IIf(IsNull(TBAbrir!ISSQN_materiais), 0, TBAbrir!ISSQN_materiais)
        Txt_pIRPJ_materiais = IIf(IsNull(TBAbrir!IRPJ_materiais), 0, TBAbrir!IRPJ_materiais)
        Txt_psimples_materiais = IIf(IsNull(TBAbrir!Simples_materiais), 0, TBAbrir!Simples_materiais)
        Txt_pcomissao_materiais = IIf(IsNull(TBAbrir!Comissao_materiais), 0, TBAbrir!Comissao_materiais)
        Txt_pdespesas_comerciais_materiais = IIf(IsNull(TBAbrir!Despesas_comerciais_materiais), 0, TBAbrir!Despesas_comerciais_materiais)
        Txt_pdespesas_financeiras_materiais = IIf(IsNull(TBAbrir!Despesas_financeiras_materiais), 0, TBAbrir!Despesas_financeiras_materiais)
        Txt_pdespesas_administrativas_materiais = IIf(IsNull(TBAbrir!Despesas_administrativas_materiais), 0, TBAbrir!Despesas_administrativas_materiais)
        Txt_pfrete_materiais = IIf(IsNull(TBAbrir!Frete_materiais), 0, TBAbrir!Frete_materiais)
        Txt_pmargem_materiais = IIf(IsNull(TBAbrir!Margem_materiais), 0, TBAbrir!Margem_materiais)
        
        'Terceiros
        If TBAbrir!Opt_ICMS_terceiros = True Then Chk_ICMS_terceiros.Value = 1 Else Chk_ICMS_terceiros.Value = 0
        If TBAbrir!Opt_PIS_terceiros = True Then Chk_PIS_terceiros.Value = 1 Else Chk_PIS_terceiros.Value = 0
        If TBAbrir!Opt_cofins_terceiros = True Then Chk_cofins_terceiros.Value = 1 Else Chk_cofins_terceiros.Value = 0
        If TBAbrir!Opt_CSLL_terceiros = True Then Chk_CSLL_terceiros.Value = 1 Else Chk_CSLL_terceiros.Value = 0
        If TBAbrir!Opt_ISSQN_terceiros = True Then Chk_ISSQN_terceiros.Value = 1 Else Chk_ISSQN_terceiros.Value = 0
        If TBAbrir!Opt_IRPJ_terceiros = True Then Chk_IRPJ_terceiros.Value = 1 Else Chk_IRPJ_terceiros.Value = 0
        If TBAbrir!Opt_simples_terceiros = True Then Chk_simples_terceiros.Value = 1 Else Chk_simples_terceiros.Value = 0
        If TBAbrir!Opt_comissao_terceiros = True Then Chk_comissao_terceiros.Value = 1 Else Chk_comissao_terceiros.Value = 0
        If TBAbrir!Opt_despesas_comerciais_terceiros = True Then Chk_despesas_comerciais_terceiros.Value = 1 Else Chk_despesas_comerciais_terceiros.Value = 0
        If TBAbrir!Opt_despesas_financeiras_terceiros = True Then Chk_despesas_financeiras_terceiros.Value = 1 Else Chk_despesas_financeiras_terceiros.Value = 0
        If TBAbrir!Opt_despesas_administrativas_terceiros = True Then Chk_despesas_administrativas_terceiros.Value = 1 Else Chk_despesas_administrativas_terceiros.Value = 0
        If TBAbrir!Opt_frete_terceiros = True Then Chk_frete_terceiros.Value = 1 Else Chk_frete_terceiros.Value = 0
        If TBAbrir!Opt_margem_terceiros = True Then Chk_margem_terceiros.Value = 1 Else Chk_margem_terceiros.Value = 0
        Txt_pICMS_terceiros = IIf(IsNull(TBAbrir!ICMS_terceiros), 0, TBAbrir!ICMS_terceiros)
        Txt_pPIS_terceiros = IIf(IsNull(TBAbrir!PIS_terceiros), 0, TBAbrir!PIS_terceiros)
        Txt_pcofins_terceiros = IIf(IsNull(TBAbrir!Cofins_terceiros), 0, TBAbrir!Cofins_terceiros)
        Txt_pCSLL_terceiros = IIf(IsNull(TBAbrir!CSLL_terceiros), 0, TBAbrir!CSLL_terceiros)
        Txt_pISSQN_terceiros = IIf(IsNull(TBAbrir!ISSQN_terceiros), 0, TBAbrir!ISSQN_terceiros)
        Txt_pIRPJ_terceiros = IIf(IsNull(TBAbrir!IRPJ_terceiros), 0, TBAbrir!IRPJ_terceiros)
        Txt_psimples_terceiros = IIf(IsNull(TBAbrir!Simples_terceiros), 0, TBAbrir!Simples_terceiros)
        Txt_pcomissao_terceiros = IIf(IsNull(TBAbrir!Comissao_terceiros), 0, TBAbrir!Comissao_terceiros)
        Txt_pdespesas_comerciais_terceiros = IIf(IsNull(TBAbrir!Despesas_comerciais_terceiros), 0, TBAbrir!Despesas_comerciais_terceiros)
        Txt_pdespesas_financeiras_terceiros = IIf(IsNull(TBAbrir!Despesas_financeiras_terceiros), 0, TBAbrir!Despesas_financeiras_terceiros)
        Txt_pdespesas_administrativas_terceiros = IIf(IsNull(TBAbrir!Despesas_administrativas_terceiros), 0, TBAbrir!Despesas_administrativas_terceiros)
        Txt_pfrete_terceiros = IIf(IsNull(TBAbrir!Frete_terceiros), 0, TBAbrir!Frete_terceiros)
        Txt_pmargem_terceiros = IIf(IsNull(TBAbrir!Margem_terceiros), 0, TBAbrir!Margem_terceiros)
        
        'Outros
        If TBAbrir!Opt_ICMS_outros = True Then Chk_ICMS_outros.Value = 1 Else Chk_ICMS_outros.Value = 0
        If TBAbrir!Opt_PIS_outros = True Then Chk_PIS_outros.Value = 1 Else Chk_PIS_outros.Value = 0
        If TBAbrir!Opt_cofins_outros = True Then Chk_cofins_outros.Value = 1 Else Chk_cofins_outros.Value = 0
        If TBAbrir!Opt_CSLL_outros = True Then Chk_CSLL_outros.Value = 1 Else Chk_CSLL_outros.Value = 0
        If TBAbrir!Opt_ISSQN_outros = True Then Chk_ISSQN_outros.Value = 1 Else Chk_ISSQN_outros.Value = 0
        If TBAbrir!Opt_IRPJ_outros = True Then Chk_IRPJ_outros.Value = 1 Else Chk_IRPJ_outros.Value = 0
        If TBAbrir!Opt_simples_outros = True Then Chk_simples_outros.Value = 1 Else Chk_simples_outros.Value = 0
        If TBAbrir!Opt_comissao_outros = True Then Chk_comissao_outros.Value = 1 Else Chk_comissao_outros.Value = 0
        If TBAbrir!Opt_despesas_comerciais_outros = True Then Chk_despesas_comerciais_outros.Value = 1 Else Chk_despesas_comerciais_outros.Value = 0
        If TBAbrir!Opt_despesas_financeiras_outros = True Then Chk_despesas_financeiras_outros.Value = 1 Else Chk_despesas_financeiras_outros.Value = 0
        If TBAbrir!Opt_despesas_administrativas_outros = True Then Chk_despesas_administrativas_outros.Value = 1 Else Chk_despesas_administrativas_outros.Value = 0
        If TBAbrir!Opt_frete_outros = True Then Chk_frete_outros.Value = 1 Else Chk_frete_outros.Value = 0
        If TBAbrir!Opt_margem_outros = True Then Chk_margem_outros.Value = 1 Else Chk_margem_outros.Value = 0
        Txt_pICMS_outros = IIf(IsNull(TBAbrir!ICMS_outros), 0, TBAbrir!ICMS_outros)
        Txt_pPIS_outros = IIf(IsNull(TBAbrir!PIS_outros), 0, TBAbrir!PIS_outros)
        Txt_pcofins_outros = IIf(IsNull(TBAbrir!Cofins_outros), 0, TBAbrir!Cofins_outros)
        Txt_pCSLL_outros = IIf(IsNull(TBAbrir!CSLL_outros), 0, TBAbrir!CSLL_outros)
        Txt_pISSQN_outros = IIf(IsNull(TBAbrir!ISSQN_outros), 0, TBAbrir!ISSQN_outros)
        Txt_pIRPJ_outros = IIf(IsNull(TBAbrir!IRPJ_outros), 0, TBAbrir!IRPJ_outros)
        Txt_psimples_outros = IIf(IsNull(TBAbrir!Simples_outros), 0, TBAbrir!Simples_outros)
        Txt_pcomissao_outros = IIf(IsNull(TBAbrir!Comissao_outros), 0, TBAbrir!Comissao_outros)
        Txt_pdespesas_comerciais_outros = IIf(IsNull(TBAbrir!Despesas_comerciais_outros), 0, TBAbrir!Despesas_comerciais_outros)
        Txt_pdespesas_financeiras_outros = IIf(IsNull(TBAbrir!Despesas_financeiras_outros), 0, TBAbrir!Despesas_financeiras_outros)
        Txt_pdespesas_administrativas_outros = IIf(IsNull(TBAbrir!Despesas_administrativas_outros), 0, TBAbrir!Despesas_administrativas_outros)
        Txt_pfrete_outros = IIf(IsNull(TBAbrir!Frete_outros), 0, TBAbrir!Frete_outros)
        Txt_pmargem_outros = IIf(IsNull(TBAbrir!Margem_outros), 0, TBAbrir!Margem_outros)
    Else
        'Total
        chkTotal.Value = 1
        If TBAbrir!Opt_ICMS_total = True Then Chk_ICMS_total.Value = 1 Else Chk_ICMS_total.Value = 0
        If TBAbrir!Opt_PIS_total = True Then Chk_PIS_total.Value = 1 Else Chk_PIS_total.Value = 0
        If TBAbrir!Opt_cofins_total = True Then Chk_cofins_total.Value = 1 Else Chk_cofins_total.Value = 0
        If TBAbrir!Opt_CSLL_total = True Then Chk_CSLL_total.Value = 1 Else Chk_CSLL_total.Value = 0
        If TBAbrir!Opt_ISSQN_total = True Then Chk_ISSQN_total.Value = 1 Else Chk_ISSQN_total.Value = 0
        If TBAbrir!Opt_IRPJ_total = True Then Chk_IRPJ_total.Value = 1 Else Chk_IRPJ_total.Value = 0
        If TBAbrir!Opt_simples_total = True Then Chk_simples_total.Value = 1 Else Chk_simples_total.Value = 0
        If TBAbrir!Opt_comissao_total = True Then Chk_comissao_total.Value = 1 Else Chk_comissao_total.Value = 0
        If TBAbrir!Opt_despesas_comerciais_total = True Then Chk_despesas_comerciais_total.Value = 1 Else Chk_despesas_comerciais_total.Value = 0
        If TBAbrir!Opt_despesas_financeiras_total = True Then Chk_despesas_financeiras_total.Value = 1 Else Chk_despesas_financeiras_total.Value = 0
        If TBAbrir!Opt_despesas_administrativas_total = True Then Chk_despesas_administrativas_total.Value = 1 Else Chk_despesas_administrativas_total.Value = 0
        If TBAbrir!Opt_frete_total = True Then Chk_frete_total.Value = 1 Else Chk_frete_total.Value = 0
        If TBAbrir!Opt_margem_total = True Then Chk_margem_total.Value = 1 Else Chk_margem_total.Value = 0
        Txt_pICMS_total = IIf(IsNull(TBAbrir!ICMS_total), 0, TBAbrir!ICMS_total)
        Txt_pPIS_total = IIf(IsNull(TBAbrir!PIS_total), 0, TBAbrir!PIS_total)
        Txt_pcofins_total = IIf(IsNull(TBAbrir!Cofins_total), 0, TBAbrir!Cofins_total)
        Txt_pCSLL_total = IIf(IsNull(TBAbrir!CSLL_total), 0, TBAbrir!CSLL_total)
        Txt_pISSQN_total = IIf(IsNull(TBAbrir!ISSQN_total), 0, TBAbrir!ISSQN_total)
        Txt_pIRPJ_total = IIf(IsNull(TBAbrir!IRPJ_total), 0, TBAbrir!IRPJ_total)
        Txt_psimples_total = IIf(IsNull(TBAbrir!Simples_total), 0, TBAbrir!Simples_total)
        Txt_pcomissao_total = IIf(IsNull(TBAbrir!Comissao_total), 0, TBAbrir!Comissao_total)
        Txt_pdespesas_comerciais_total = IIf(IsNull(TBAbrir!Despesas_comerciais_total), 0, TBAbrir!Despesas_comerciais_total)
        Txt_pdespesas_financeiras_total = IIf(IsNull(TBAbrir!Despesas_financeiras_total), 0, TBAbrir!Despesas_financeiras_total)
        Txt_pdespesas_administrativas_total = IIf(IsNull(TBAbrir!Despesas_administrativas_total), 0, TBAbrir!Despesas_administrativas_total)
        Txt_pfrete_total = IIf(IsNull(TBAbrir!Frete_total), 0, TBAbrir!Frete_total)
        Txt_pmargem_total = IIf(IsNull(TBAbrir!Margem_total), 0, TBAbrir!Margem_total)
    End If
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDadosValores()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Valor_total, Valor_total_processo, Valor_total_materiais, Valor_total_terceiros, Valor_total_outros from Vendas_analise where ID = " & frmVendas_analise.txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If Txt_total_processo = "0,00" And Txt_total_processo.Enabled = True Then Txt_total_processo = IIf(IsNull(TBAbrir!Valor_total_processo), "0,00", Format(TBAbrir!Valor_total_processo, "###,##0.00"))
    'If Txt_total_materiais = "0,00" And Txt_total_materiais.Enabled = True Then Txt_total_materiais = IIf(IsNull(TBAbrir!Valor_total_materiais), "0,00", Format(TBAbrir!Valor_total_materiais, "###,##0.00"))
    Txt_total_materiais = IIf(IsNull(TBAbrir!Valor_total_materiais), "0,00", Format(TBAbrir!Valor_total_materiais, "###,##0.00"))
    If Txt_total_terceiros = "0,00" And Txt_total_terceiros.Enabled = True Then Txt_total_terceiros = IIf(IsNull(TBAbrir!Valor_total_terceiros), "0,00", Format(TBAbrir!Valor_total_terceiros, "###,##0.00"))
    If Txt_total_outros = "0,00" And Txt_total_outros.Enabled = True Then Txt_total_outros = IIf(IsNull(TBAbrir!Valor_total_outros), "0,00", Format(TBAbrir!Valor_total_outros, "###,##0.00"))
    If IsNull(TBAbrir!Valor_total) = False And TBAbrir!Valor_total <> "" And TBAbrir!Valor_total <> "0" Then Txt_valor_total_venda = IIf(IsNull(TBAbrir!Valor_total), "0,00000", Format(TBAbrir!Valor_total, "###,##0.0000000000"))
End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If frmVendas_analise.Txt_status <> "ABERTA EM ANALISE" Then
    USMsgBox ("Só é permitido alterar o fechamento se a análise estiver com o status aberta em análise."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from vendas_analise WHERE ID = " & frmVendas_analise.txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    ProcEnviaDados
    TBGravar!Fechada = True
    TBGravar.Update
End If
TBGravar.Close

USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
Evento = "Alterar impostos"
With frmVendas_analise
    '==================================
    Modulo = "Outros/Análise crítica"
    With frmVendas_analise
        ID_documento = .txtId
        Documento = "Nº análise: " & .Txt_analise & " - Rev.: " & .Txt_rev_analise
    End With
    Documento1 = ""
    ProcGravaEvento
    '==================================
    .Lista.ListItems.Clear
    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDados()
On Error GoTo tratar_erro

'Processo
If Chk_ICMS_processo.Value = 1 Then TBGravar!Opt_ICMS_processo = True Else TBGravar!Opt_ICMS_processo = False
If Chk_PIS_processo.Value = 1 Then TBGravar!Opt_PIS_processo = True Else TBGravar!Opt_PIS_processo = False
If Chk_cofins_processo.Value = 1 Then TBGravar!Opt_cofins_processo = True Else TBGravar!Opt_cofins_processo = False
If Chk_CSLL_processo.Value = 1 Then TBGravar!Opt_CSLL_processo = True Else TBGravar!Opt_CSLL_processo = False
If Chk_ISSQN_processo.Value = 1 Then TBGravar!Opt_ISSQN_processo = True Else TBGravar!Opt_ISSQN_processo = False
If Chk_IRPJ_processo.Value = 1 Then TBGravar!Opt_IRPJ_processo = True Else TBGravar!Opt_IRPJ_processo = False
If Chk_simples_processo.Value = 1 Then TBGravar!Opt_simples_processo = True Else TBGravar!Opt_simples_processo = False
If Chk_comissao_processo.Value = 1 Then TBGravar!Opt_comissao_processo = True Else TBGravar!Opt_comissao_processo = False
If Chk_despesas_comerciais_processo.Value = 1 Then TBGravar!Opt_despesas_comerciais_processo = True Else TBGravar!Opt_despesas_comerciais_processo = False
If Chk_despesas_financeiras_processo.Value = 1 Then TBGravar!Opt_despesas_financeiras_processo = True Else TBGravar!Opt_despesas_financeiras_processo = False
If Chk_despesas_administrativas_processo.Value = 1 Then TBGravar!Opt_despesas_administrativas_processo = True Else TBGravar!Opt_despesas_administrativas_processo = False
If Chk_frete_processo.Value = 1 Then TBGravar!Opt_frete_processo = True Else TBGravar!Opt_frete_processo = False
If Chk_margem_processo.Value = 1 Then TBGravar!Opt_margem_processo = True Else TBGravar!Opt_margem_processo = False
TBGravar!ICMS_processo = IIf(Txt_pICMS_processo = "", 0, Txt_pICMS_processo)
TBGravar!PIS_processo = IIf(Txt_pPIS_processo = "", 0, Txt_pPIS_processo)
TBGravar!Cofins_processo = IIf(Txt_pcofins_processo = "", 0, Txt_pcofins_processo)
TBGravar!CSLL_processo = IIf(Txt_pCSLL_processo = "", 0, Txt_pCSLL_processo)
TBGravar!ISSQN_processo = IIf(Txt_pISSQN_processo = "", 0, Txt_pISSQN_processo)
TBGravar!IRPJ_processo = IIf(Txt_pIRPJ_processo = "", 0, Txt_pIRPJ_processo)
TBGravar!Simples_processo = IIf(Txt_psimples_processo = "", 0, Txt_psimples_processo)
TBGravar!Comissao_processo = IIf(Txt_pcomissao_processo = "", 0, Txt_pcomissao_processo)
TBGravar!Despesas_comerciais_processo = IIf(Txt_pdespesas_comerciais_processo = "", 0, Txt_pdespesas_comerciais_processo)
TBGravar!Despesas_financeiras_processo = IIf(Txt_pdespesas_financeiras_processo = "", 0, Txt_pdespesas_financeiras_processo)
TBGravar!Despesas_administrativas_processo = IIf(Txt_pdespesas_administrativas_processo = "", 0, Txt_pdespesas_administrativas_processo)
TBGravar!Frete_processo = IIf(Txt_pfrete_processo = "", 0, Txt_pfrete_processo)
TBGravar!Margem_processo = IIf(Txt_pmargem_processo = "", 0, Txt_pmargem_processo)

'Materiais
If Chk_ICMS_materiais.Value = 1 Then TBGravar!Opt_ICMS_materiais = True Else TBGravar!Opt_ICMS_materiais = False
If Chk_PIS_materiais.Value = 1 Then TBGravar!Opt_PIS_materiais = True Else TBGravar!Opt_PIS_materiais = False
If Chk_cofins_materiais.Value = 1 Then TBGravar!Opt_cofins_materiais = True Else TBGravar!Opt_cofins_materiais = False
If Chk_CSLL_materiais.Value = 1 Then TBGravar!Opt_CSLL_materiais = True Else TBGravar!Opt_CSLL_materiais = False
If Chk_ISSQN_materiais.Value = 1 Then TBGravar!Opt_ISSQN_materiais = True Else TBGravar!Opt_ISSQN_materiais = False
If Chk_IRPJ_materiais.Value = 1 Then TBGravar!Opt_IRPJ_materiais = True Else TBGravar!Opt_IRPJ_materiais = False
If Chk_simples_materiais.Value = 1 Then TBGravar!Opt_simples_materiais = True Else TBGravar!Opt_simples_materiais = False
If Chk_comissao_materiais.Value = 1 Then TBGravar!Opt_comissao_materiais = True Else TBGravar!Opt_comissao_materiais = False
If Chk_despesas_comerciais_materiais.Value = 1 Then TBGravar!Opt_despesas_comerciais_materiais = True Else TBGravar!Opt_despesas_comerciais_materiais = False
If Chk_despesas_financeiras_materiais.Value = 1 Then TBGravar!Opt_despesas_financeiras_materiais = True Else TBGravar!Opt_despesas_financeiras_materiais = False
If Chk_despesas_administrativas_materiais.Value = 1 Then TBGravar!Opt_despesas_administrativas_materiais = True Else TBGravar!Opt_despesas_administrativas_materiais = False
If Chk_frete_materiais.Value = 1 Then TBGravar!Opt_frete_materiais = True Else TBGravar!Opt_frete_materiais = False
If Chk_margem_materiais.Value = 1 Then TBGravar!Opt_margem_materiais = True Else TBGravar!Opt_margem_materiais = False
TBGravar!ICMS_materiais = IIf(Txt_pICMS_materiais = "", 0, Txt_pICMS_materiais)
TBGravar!PIS_materiais = IIf(Txt_pPIS_materiais = "", 0, Txt_pPIS_materiais)
TBGravar!Cofins_materiais = IIf(Txt_pcofins_materiais = "", 0, Txt_pcofins_materiais)
TBGravar!CSLL_materiais = IIf(Txt_pCSLL_materiais = "", 0, Txt_pCSLL_materiais)
TBGravar!ISSQN_materiais = IIf(Txt_pISSQN_materiais = "", 0, Txt_pISSQN_materiais)
TBGravar!IRPJ_materiais = IIf(Txt_pIRPJ_materiais = "", 0, Txt_pIRPJ_materiais)
TBGravar!Simples_materiais = IIf(Txt_psimples_materiais = "", 0, Txt_psimples_materiais)
TBGravar!Comissao_materiais = IIf(Txt_pcomissao_materiais = "", 0, Txt_pcomissao_materiais)
TBGravar!Despesas_comerciais_materiais = IIf(Txt_pdespesas_comerciais_materiais = "", 0, Txt_pdespesas_comerciais_materiais)
TBGravar!Despesas_financeiras_materiais = IIf(Txt_pdespesas_financeiras_materiais = "", 0, Txt_pdespesas_financeiras_materiais)
TBGravar!Despesas_administrativas_materiais = IIf(Txt_pdespesas_administrativas_materiais = "", 0, Txt_pdespesas_administrativas_materiais)
TBGravar!Frete_materiais = IIf(Txt_pfrete_materiais = "", 0, Txt_pfrete_materiais)
TBGravar!Margem_materiais = IIf(Txt_pmargem_materiais = "", 0, Txt_pmargem_materiais)

'Terceiros
If Chk_ICMS_terceiros.Value = 1 Then TBGravar!Opt_ICMS_terceiros = True Else TBGravar!Opt_ICMS_terceiros = False
If Chk_PIS_terceiros.Value = 1 Then TBGravar!Opt_PIS_terceiros = True Else TBGravar!Opt_PIS_terceiros = False
If Chk_cofins_terceiros.Value = 1 Then TBGravar!Opt_cofins_terceiros = True Else TBGravar!Opt_cofins_terceiros = False
If Chk_CSLL_terceiros.Value = 1 Then TBGravar!Opt_CSLL_terceiros = True Else TBGravar!Opt_CSLL_terceiros = False
If Chk_ISSQN_terceiros.Value = 1 Then TBGravar!Opt_ISSQN_terceiros = True Else TBGravar!Opt_ISSQN_terceiros = False
If Chk_IRPJ_terceiros.Value = 1 Then TBGravar!Opt_IRPJ_terceiros = True Else TBGravar!Opt_IRPJ_terceiros = False
If Chk_simples_terceiros.Value = 1 Then TBGravar!Opt_simples_terceiros = True Else TBGravar!Opt_simples_terceiros = False
If Chk_comissao_terceiros.Value = 1 Then TBGravar!Opt_comissao_terceiros = True Else TBGravar!Opt_comissao_terceiros = False
If Chk_despesas_comerciais_terceiros.Value = 1 Then TBGravar!Opt_despesas_comerciais_terceiros = True Else TBGravar!Opt_despesas_comerciais_terceiros = False
If Chk_despesas_financeiras_terceiros.Value = 1 Then TBGravar!Opt_despesas_financeiras_terceiros = True Else TBGravar!Opt_despesas_financeiras_terceiros = False
If Chk_despesas_administrativas_terceiros.Value = 1 Then TBGravar!Opt_despesas_administrativas_terceiros = True Else TBGravar!Opt_despesas_administrativas_terceiros = False
If Chk_frete_terceiros.Value = 1 Then TBGravar!Opt_frete_terceiros = True Else TBGravar!Opt_frete_terceiros = False
If Chk_margem_terceiros.Value = 1 Then TBGravar!Opt_margem_terceiros = True Else TBGravar!Opt_margem_terceiros = False
TBGravar!ICMS_terceiros = IIf(Txt_pICMS_terceiros = "", 0, Txt_pICMS_terceiros)
TBGravar!PIS_terceiros = IIf(Txt_pPIS_terceiros = "", 0, Txt_pPIS_terceiros)
TBGravar!Cofins_terceiros = IIf(Txt_pcofins_terceiros = "", 0, Txt_pcofins_terceiros)
TBGravar!CSLL_terceiros = IIf(Txt_pCSLL_terceiros = "", 0, Txt_pCSLL_terceiros)
TBGravar!ISSQN_terceiros = IIf(Txt_pISSQN_terceiros = "", 0, Txt_pISSQN_terceiros)
TBGravar!IRPJ_terceiros = IIf(Txt_pIRPJ_terceiros = "", 0, Txt_pIRPJ_terceiros)
TBGravar!Simples_terceiros = IIf(Txt_psimples_terceiros = "", 0, Txt_psimples_terceiros)
TBGravar!Comissao_terceiros = IIf(Txt_pcomissao_terceiros = "", 0, Txt_pcomissao_terceiros)
TBGravar!Despesas_comerciais_terceiros = IIf(Txt_pdespesas_comerciais_terceiros = "", 0, Txt_pdespesas_comerciais_terceiros)
TBGravar!Despesas_financeiras_terceiros = IIf(Txt_pdespesas_financeiras_terceiros = "", 0, Txt_pdespesas_financeiras_terceiros)
TBGravar!Despesas_administrativas_terceiros = IIf(Txt_pdespesas_administrativas_terceiros = "", 0, Txt_pdespesas_administrativas_terceiros)
TBGravar!Frete_terceiros = IIf(Txt_pfrete_terceiros = "", 0, Txt_pfrete_terceiros)
TBGravar!Margem_terceiros = IIf(Txt_pmargem_terceiros = "", 0, Txt_pmargem_terceiros)

'Outros
If Chk_ICMS_outros.Value = 1 Then TBGravar!Opt_ICMS_outros = True Else TBGravar!Opt_ICMS_outros = False
If Chk_PIS_outros.Value = 1 Then TBGravar!Opt_PIS_outros = True Else TBGravar!Opt_PIS_outros = False
If Chk_cofins_outros.Value = 1 Then TBGravar!Opt_cofins_outros = True Else TBGravar!Opt_cofins_outros = False
If Chk_CSLL_outros.Value = 1 Then TBGravar!Opt_CSLL_outros = True Else TBGravar!Opt_CSLL_outros = False
If Chk_ISSQN_outros.Value = 1 Then TBGravar!Opt_ISSQN_outros = True Else TBGravar!Opt_ISSQN_outros = False
If Chk_IRPJ_outros.Value = 1 Then TBGravar!Opt_IRPJ_outros = True Else TBGravar!Opt_IRPJ_outros = False
If Chk_simples_outros.Value = 1 Then TBGravar!Opt_simples_outros = True Else TBGravar!Opt_simples_outros = False
If Chk_comissao_outros.Value = 1 Then TBGravar!Opt_comissao_outros = True Else TBGravar!Opt_comissao_outros = False
If Chk_despesas_comerciais_outros.Value = 1 Then TBGravar!Opt_despesas_comerciais_outros = True Else TBGravar!Opt_despesas_comerciais_outros = False
If Chk_despesas_financeiras_outros.Value = 1 Then TBGravar!Opt_despesas_financeiras_outros = True Else TBGravar!Opt_despesas_financeiras_outros = False
If Chk_despesas_administrativas_outros.Value = 1 Then TBGravar!Opt_despesas_administrativas_outros = True Else TBGravar!Opt_despesas_administrativas_outros = False
If Chk_frete_outros.Value = 1 Then TBGravar!Opt_frete_outros = True Else TBGravar!Opt_frete_outros = False
If Chk_margem_outros.Value = 1 Then TBGravar!Opt_margem_outros = True Else TBGravar!Opt_margem_outros = False
TBGravar!ICMS_outros = IIf(Txt_pICMS_outros = "", 0, Txt_pICMS_outros)
TBGravar!PIS_outros = IIf(Txt_pPIS_outros = "", 0, Txt_pPIS_outros)
TBGravar!Cofins_outros = IIf(Txt_pcofins_outros = "", 0, Txt_pcofins_outros)
TBGravar!CSLL_outros = IIf(Txt_pCSLL_outros = "", 0, Txt_pCSLL_outros)
TBGravar!ISSQN_outros = IIf(Txt_pISSQN_outros = "", 0, Txt_pISSQN_outros)
TBGravar!IRPJ_outros = IIf(Txt_pIRPJ_outros = "", 0, Txt_pIRPJ_outros)
TBGravar!Simples_outros = IIf(Txt_psimples_outros = "", 0, Txt_psimples_outros)
TBGravar!Comissao_outros = IIf(Txt_pcomissao_outros = "", 0, Txt_pcomissao_outros)
TBGravar!Despesas_comerciais_outros = IIf(Txt_pdespesas_comerciais_outros = "", 0, Txt_pdespesas_comerciais_outros)
TBGravar!Despesas_financeiras_outros = IIf(Txt_pdespesas_financeiras_outros = "", 0, Txt_pdespesas_financeiras_outros)
TBGravar!Despesas_administrativas_outros = IIf(Txt_pdespesas_administrativas_outros = "", 0, Txt_pdespesas_administrativas_outros)
TBGravar!Frete_outros = IIf(Txt_pfrete_outros = "", 0, Txt_pfrete_outros)
TBGravar!Margem_outros = IIf(Txt_pmargem_outros = "", 0, Txt_pmargem_outros)

'total
If chkTotal.Value = 1 Then TBGravar!chkTotal = True Else TBGravar!chkTotal = False
If Chk_ICMS_total.Value = 1 Then TBGravar!Opt_ICMS_total = True Else TBGravar!Opt_ICMS_total = False
If Chk_PIS_total.Value = 1 Then TBGravar!Opt_PIS_total = True Else TBGravar!Opt_PIS_total = False
If Chk_cofins_total.Value = 1 Then TBGravar!Opt_cofins_total = True Else TBGravar!Opt_cofins_total = False
If Chk_CSLL_total.Value = 1 Then TBGravar!Opt_CSLL_total = True Else TBGravar!Opt_CSLL_total = False
If Chk_ISSQN_total.Value = 1 Then TBGravar!Opt_ISSQN_total = True Else TBGravar!Opt_ISSQN_total = False
If Chk_IRPJ_total.Value = 1 Then TBGravar!Opt_IRPJ_total = True Else TBGravar!Opt_IRPJ_total = False
If Chk_simples_total.Value = 1 Then TBGravar!Opt_simples_total = True Else TBGravar!Opt_simples_total = False
If Chk_comissao_total.Value = 1 Then TBGravar!Opt_comissao_total = True Else TBGravar!Opt_comissao_total = False
If Chk_despesas_comerciais_total.Value = 1 Then TBGravar!Opt_despesas_comerciais_total = True Else TBGravar!Opt_despesas_comerciais_total = False
If Chk_despesas_financeiras_total.Value = 1 Then TBGravar!Opt_despesas_financeiras_total = True Else TBGravar!Opt_despesas_financeiras_total = False
If Chk_despesas_administrativas_total.Value = 1 Then TBGravar!Opt_despesas_administrativas_total = True Else TBGravar!Opt_despesas_administrativas_total = False
If Chk_frete_total.Value = 1 Then TBGravar!Opt_frete_total = True Else TBGravar!Opt_frete_total = False
If Chk_margem_total.Value = 1 Then TBGravar!Opt_margem_total = True Else TBGravar!Opt_margem_total = False
TBGravar!ICMS_total = IIf(Txt_pICMS_total = "", 0, Txt_pICMS_total)
TBGravar!PIS_total = IIf(Txt_pPIS_total = "", 0, Txt_pPIS_total)
TBGravar!Cofins_total = IIf(Txt_pcofins_total = "", 0, Txt_pcofins_total)
TBGravar!CSLL_total = IIf(Txt_pCSLL_total = "", 0, Txt_pCSLL_total)
TBGravar!ISSQN_total = IIf(Txt_pISSQN_total = "", 0, Txt_pISSQN_total)
TBGravar!IRPJ_total = IIf(Txt_pIRPJ_total = "", 0, Txt_pIRPJ_total)
TBGravar!Simples_total = IIf(Txt_psimples_total = "", 0, Txt_psimples_total)
TBGravar!Comissao_total = IIf(Txt_pcomissao_total = "", 0, Txt_pcomissao_total)
TBGravar!Despesas_comerciais_total = IIf(Txt_pdespesas_comerciais_total = "", 0, Txt_pdespesas_comerciais_total)
TBGravar!Despesas_financeiras_total = IIf(Txt_pdespesas_financeiras_total = "", 0, Txt_pdespesas_financeiras_total)
TBGravar!Despesas_administrativas_total = IIf(Txt_pdespesas_administrativas_total = "", 0, Txt_pdespesas_administrativas_total)
TBGravar!Frete_total = IIf(Txt_pfrete_total = "", 0, Txt_pfrete_total)
TBGravar!Margem_total = IIf(Txt_pmargem_total = "", 0, Txt_pmargem_total)

TBGravar!Valor_total_processo = IIf(Txt_total_processo = "", 0, Txt_total_processo)
TBGravar!Valor_total_materiais = IIf(Txt_total_materiais = "", 0, Txt_total_materiais)
TBGravar!Valor_total_terceiros = IIf(Txt_total_terceiros = "", 0, Txt_total_terceiros)
TBGravar!Valor_total_outros = IIf(Txt_total_outros = "", 0, Txt_total_outros)
TBGravar!Valor_total = IIf(Txt_valor_total_venda = "", 0, Txt_valor_total_venda)

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
If USMsgBox("Deseja realmente excluir os registros?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    ProcLimpaCampos
    Set TBGravar = CreateObject("adodb.recordset")
    TBGravar.Open "Select * from Vendas_analise where ID = " & frmVendas_analise.txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBGravar.EOF = False Then
        ProcEnviaDados
        TBGravar!Fechada = False
        TBGravar.Update
    End If
    USMsgBox ("Registros excluídos com sucesso."), vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Outros/Análise crítica"
    Evento = "Excluir"
    With frmVendas_analise
        ID_documento = .txtId
        Documento = "Nº análise: " & .Txt_analise & " - Rev.: " & .Txt_rev_analise
    End With
    Documento1 = ""
    ProcGravaEvento
    '==================================
    ProcPuxaDadosValores
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ICMS_processo_Click()
On Error GoTo tratar_erro

With Txt_pICMS_processo
    If Chk_ICMS_processo.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_PIS_processo_Click()
On Error GoTo tratar_erro

With Txt_pPIS_processo
    If Chk_PIS_processo.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_cofins_processo_Click()
On Error GoTo tratar_erro

With Txt_pcofins_processo
    If Chk_cofins_processo.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_CSLL_processo_Click()
On Error GoTo tratar_erro

With Txt_pCSLL_processo
    If Chk_CSLL_processo.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ISSQN_processo_Click()
On Error GoTo tratar_erro

With Txt_pISSQN_processo
    If Chk_ISSQN_processo.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_IRPJ_processo_Click()
On Error GoTo tratar_erro

With Txt_pIRPJ_processo
    If Chk_IRPJ_processo.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_Simples_processo_Click()
On Error GoTo tratar_erro

With Txt_psimples_processo
    If Chk_simples_processo.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_comissao_processo_Click()
On Error GoTo tratar_erro

With Txt_pcomissao_processo
    If Chk_comissao_processo.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_despesas_comerciais_processo_Click()
On Error GoTo tratar_erro

With Txt_pdespesas_comerciais_processo
    If Chk_despesas_comerciais_processo.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_despesas_financeiras_processo_Click()
On Error GoTo tratar_erro

With Txt_pdespesas_financeiras_processo
    If Chk_despesas_financeiras_processo.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_despesas_administrativas_processo_Click()
On Error GoTo tratar_erro

With Txt_pdespesas_administrativas_processo
    If Chk_despesas_administrativas_processo.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_frete_processo_Click()
On Error GoTo tratar_erro

With Txt_pfrete_processo
    If Chk_frete_processo.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_margem_processo_Click()
On Error GoTo tratar_erro

With Txt_pmargem_processo
    If Chk_margem_processo.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcofins_outros_Change()
On Error GoTo tratar_erro

If Txt_pcofins_outros <> "" Then
    VerifNumero = Txt_pcofins_outros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pcofins_outros = ""
        Txt_pcofins_outros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_outros = "", 0, Txt_pICMS_outros), IIf(Txt_pPIS_outros = "", 0, Txt_pPIS_outros), IIf(Txt_pcofins_outros = "", 0, Txt_pcofins_outros), IIf(Txt_pCSLL_outros = "", 0, Txt_pCSLL_outros), IIf(Txt_pISSQN_outros = "", 0, Txt_pISSQN_outros), IIf(Txt_pIRPJ_outros = "", 0, Txt_pIRPJ_outros), IIf(Txt_psimples_outros = "", 0, Txt_psimples_outros), IIf(Txt_pcomissao_outros = "", 0, Txt_pcomissao_outros), IIf(Txt_pdespesas_comerciais_outros = "", 0, Txt_pdespesas_comerciais_outros), IIf(Txt_pdespesas_financeiras_outros = "", 0, Txt_pdespesas_financeiras_outros), IIf(Txt_pdespesas_administrativas_outros = "", 0, Txt_pdespesas_administrativas_outros), IIf(Txt_pfrete_outros = "", 0, Txt_pfrete_outros), IIf(Txt_pmargem_outros = "", 0, Txt_pmargem_outros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pcofins_outros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcofins_outros_GotFocus()
On Error GoTo tratar_erro

If Txt_pcofins_outros = 0 Then Txt_pcofins_outros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcofins_total_Change()
On Error GoTo tratar_erro

If Txt_pcofins_total <> "" Then
    VerifNumero = Txt_pcofins_total
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pcofins_total = ""
        Txt_pcofins_total.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_total = "", 0, Txt_pICMS_total), IIf(Txt_pPIS_total = "", 0, Txt_pPIS_total), IIf(Txt_pcofins_total = "", 0, Txt_pcofins_total), IIf(Txt_pCSLL_total = "", 0, Txt_pCSLL_total), IIf(Txt_pISSQN_total = "", 0, Txt_pISSQN_total), IIf(Txt_pIRPJ_total = "", 0, Txt_pIRPJ_total), IIf(Txt_psimples_total = "", 0, Txt_psimples_total), IIf(Txt_pcomissao_total = "", 0, Txt_pcomissao_total), IIf(Txt_pdespesas_comerciais_total = "", 0, Txt_pdespesas_comerciais_total), IIf(Txt_pdespesas_financeiras_total = "", 0, Txt_pdespesas_financeiras_total), IIf(Txt_pdespesas_administrativas_total = "", 0, Txt_pdespesas_administrativas_total), IIf(Txt_pfrete_total = "", 0, Txt_pfrete_total), IIf(Txt_pmargem_total = "", 0, Txt_pmargem_total)) = True Then
    ProcCalculaImpostos
Else
    Txt_pcofins_total = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcofins_total_GotFocus()
On Error GoTo tratar_erro

If Txt_pcofins_total = "0" Then Txt_pcofins_total = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcomissao_outros_Change()
On Error GoTo tratar_erro

If Txt_pcomissao_outros <> "" Then
    VerifNumero = Txt_pcomissao_outros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pcomissao_outros = ""
        Txt_pcomissao_outros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_outros = "", 0, Txt_pICMS_outros), IIf(Txt_pPIS_outros = "", 0, Txt_pPIS_outros), IIf(Txt_pcofins_outros = "", 0, Txt_pcofins_outros), IIf(Txt_pCSLL_outros = "", 0, Txt_pCSLL_outros), IIf(Txt_pISSQN_outros = "", 0, Txt_pISSQN_outros), IIf(Txt_pIRPJ_outros = "", 0, Txt_pIRPJ_outros), IIf(Txt_psimples_outros = "", 0, Txt_psimples_outros), IIf(Txt_pcomissao_outros = "", 0, Txt_pcomissao_outros), IIf(Txt_pdespesas_comerciais_outros = "", 0, Txt_pdespesas_comerciais_outros), IIf(Txt_pdespesas_financeiras_outros = "", 0, Txt_pdespesas_financeiras_outros), IIf(Txt_pdespesas_administrativas_outros = "", 0, Txt_pdespesas_administrativas_outros), IIf(Txt_pfrete_outros = "", 0, Txt_pfrete_outros), IIf(Txt_pmargem_outros = "", 0, Txt_pmargem_outros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pcomissao_outros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcomissao_outros_GotFocus()
On Error GoTo tratar_erro

If Txt_pcomissao_outros = 0 Then Txt_pcomissao_outros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcomissao_total_Change()
On Error GoTo tratar_erro

If Txt_pcomissao_total <> "" Then
    VerifNumero = Txt_pcomissao_total
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pcomissao_total = ""
        Txt_pcomissao_total.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_total = "", 0, Txt_pICMS_total), IIf(Txt_pPIS_total = "", 0, Txt_pPIS_total), IIf(Txt_pcofins_total = "", 0, Txt_pcofins_total), IIf(Txt_pCSLL_total = "", 0, Txt_pCSLL_total), IIf(Txt_pISSQN_total = "", 0, Txt_pISSQN_total), IIf(Txt_pIRPJ_total = "", 0, Txt_pIRPJ_total), IIf(Txt_psimples_total = "", 0, Txt_psimples_total), IIf(Txt_pcomissao_total = "", 0, Txt_pcomissao_total), IIf(Txt_pdespesas_comerciais_total = "", 0, Txt_pdespesas_comerciais_total), IIf(Txt_pdespesas_financeiras_total = "", 0, Txt_pdespesas_financeiras_total), IIf(Txt_pdespesas_administrativas_total = "", 0, Txt_pdespesas_administrativas_total), IIf(Txt_pfrete_total = "", 0, Txt_pfrete_total), IIf(Txt_pmargem_total = "", 0, Txt_pmargem_total)) = True Then
    ProcCalculaImpostos
Else
    Txt_pcomissao_total = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcomissao_total_GotFocus()
On Error GoTo tratar_erro

If Txt_pcomissao_total = 0 Then Txt_pcomissao_total = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pCSLL_outros_Change()
On Error GoTo tratar_erro

If Txt_pCSLL_outros <> "" Then
    VerifNumero = Txt_pCSLL_outros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pCSLL_outros = ""
        Txt_pCSLL_outros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_outros = "", 0, Txt_pICMS_outros), IIf(Txt_pPIS_outros = "", 0, Txt_pPIS_outros), IIf(Txt_pcofins_outros = "", 0, Txt_pcofins_outros), IIf(Txt_pCSLL_outros = "", 0, Txt_pCSLL_outros), IIf(Txt_pISSQN_outros = "", 0, Txt_pISSQN_outros), IIf(Txt_pIRPJ_outros = "", 0, Txt_pIRPJ_outros), IIf(Txt_psimples_outros = "", 0, Txt_psimples_outros), IIf(Txt_pcomissao_outros = "", 0, Txt_pcomissao_outros), IIf(Txt_pdespesas_comerciais_outros = "", 0, Txt_pdespesas_comerciais_outros), IIf(Txt_pdespesas_financeiras_outros = "", 0, Txt_pdespesas_financeiras_outros), IIf(Txt_pdespesas_administrativas_outros = "", 0, Txt_pdespesas_administrativas_outros), IIf(Txt_pfrete_outros = "", 0, Txt_pfrete_outros), IIf(Txt_pmargem_outros = "", 0, Txt_pmargem_outros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pCSLL_outros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pCSLL_outros_GotFocus()
On Error GoTo tratar_erro

If Txt_pCSLL_outros = 0 Then Txt_pCSLL_outros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pCSLL_total_Change()
On Error GoTo tratar_erro

If Txt_pCSLL_total <> "" Then
    VerifNumero = Txt_pCSLL_total
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pCSLL_total = ""
        Txt_pCSLL_total.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_total = "", 0, Txt_pICMS_total), IIf(Txt_pPIS_total = "", 0, Txt_pPIS_total), IIf(Txt_pcofins_total = "", 0, Txt_pcofins_total), IIf(Txt_pCSLL_total = "", 0, Txt_pCSLL_total), IIf(Txt_pISSQN_total = "", 0, Txt_pISSQN_total), IIf(Txt_pIRPJ_total = "", 0, Txt_pIRPJ_total), IIf(Txt_psimples_total = "", 0, Txt_psimples_total), IIf(Txt_pcomissao_total = "", 0, Txt_pcomissao_total), IIf(Txt_pdespesas_comerciais_total = "", 0, Txt_pdespesas_comerciais_total), IIf(Txt_pdespesas_financeiras_total = "", 0, Txt_pdespesas_financeiras_total), IIf(Txt_pdespesas_administrativas_total = "", 0, Txt_pdespesas_administrativas_total), IIf(Txt_pfrete_total = "", 0, Txt_pfrete_total), IIf(Txt_pmargem_total = "", 0, Txt_pmargem_total)) = True Then
    ProcCalculaImpostos
Else
    Txt_pCSLL_total = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pCSLL_total_GotFocus()
On Error GoTo tratar_erro

If Txt_pCSLL_total = "0" Then Txt_pCSLL_total = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_administrativas_outros_Change()
On Error GoTo tratar_erro

If Txt_pdespesas_administrativas_outros <> "" Then
    VerifNumero = Txt_pdespesas_administrativas_outros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pdespesas_administrativas_outros = ""
        Txt_pdespesas_administrativas_outros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_outros = "", 0, Txt_pICMS_outros), IIf(Txt_pPIS_outros = "", 0, Txt_pPIS_outros), IIf(Txt_pcofins_outros = "", 0, Txt_pcofins_outros), IIf(Txt_pCSLL_outros = "", 0, Txt_pCSLL_outros), IIf(Txt_pISSQN_outros = "", 0, Txt_pISSQN_outros), IIf(Txt_pIRPJ_outros = "", 0, Txt_pIRPJ_outros), IIf(Txt_psimples_outros = "", 0, Txt_psimples_outros), IIf(Txt_pcomissao_outros = "", 0, Txt_pcomissao_outros), IIf(Txt_pdespesas_comerciais_outros = "", 0, Txt_pdespesas_comerciais_outros), IIf(Txt_pdespesas_financeiras_outros = "", 0, Txt_pdespesas_financeiras_outros), IIf(Txt_pdespesas_administrativas_outros = "", 0, Txt_pdespesas_administrativas_outros), IIf(Txt_pfrete_outros = "", 0, Txt_pfrete_outros), IIf(Txt_pmargem_outros = "", 0, Txt_pmargem_outros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pdespesas_administrativas_outros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_administrativas_outros_GotFocus()
On Error GoTo tratar_erro

If Txt_pdespesas_administrativas_outros = "0" Then Txt_pdespesas_administrativas_outros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_administrativas_total_Change()
On Error GoTo tratar_erro

If Txt_pdespesas_administrativas_total <> "" Then
    VerifNumero = Txt_pdespesas_administrativas_total
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pdespesas_administrativas_total = ""
        Txt_pdespesas_administrativas_total.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_total = "", 0, Txt_pICMS_total), IIf(Txt_pPIS_total = "", 0, Txt_pPIS_total), IIf(Txt_pcofins_total = "", 0, Txt_pcofins_total), IIf(Txt_pCSLL_total = "", 0, Txt_pCSLL_total), IIf(Txt_pISSQN_total = "", 0, Txt_pISSQN_total), IIf(Txt_pIRPJ_total = "", 0, Txt_pIRPJ_total), IIf(Txt_psimples_total = "", 0, Txt_psimples_total), IIf(Txt_pcomissao_total = "", 0, Txt_pcomissao_total), IIf(Txt_pdespesas_comerciais_total = "", 0, Txt_pdespesas_comerciais_total), IIf(Txt_pdespesas_financeiras_total = "", 0, Txt_pdespesas_financeiras_total), IIf(Txt_pdespesas_administrativas_total = "", 0, Txt_pdespesas_administrativas_total), IIf(Txt_pfrete_total = "", 0, Txt_pfrete_total), IIf(Txt_pmargem_total = "", 0, Txt_pmargem_total)) = True Then
    ProcCalculaImpostos
Else
    Txt_pdespesas_administrativas_total = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_administrativas_total_GotFocus()
On Error GoTo tratar_erro

If Txt_pdespesas_administrativas_total = "0" Then Txt_pdespesas_administrativas_total = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_comerciais_outros_Change()
On Error GoTo tratar_erro

If Txt_pdespesas_comerciais_outros <> "" Then
    VerifNumero = Txt_pdespesas_comerciais_outros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pdespesas_comerciais_outros = ""
        Txt_pdespesas_comerciais_outros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_outros = "", 0, Txt_pICMS_outros), IIf(Txt_pPIS_outros = "", 0, Txt_pPIS_outros), IIf(Txt_pcofins_outros = "", 0, Txt_pcofins_outros), IIf(Txt_pCSLL_outros = "", 0, Txt_pCSLL_outros), IIf(Txt_pISSQN_outros = "", 0, Txt_pISSQN_outros), IIf(Txt_pIRPJ_outros = "", 0, Txt_pIRPJ_outros), IIf(Txt_psimples_outros = "", 0, Txt_psimples_outros), IIf(Txt_pcomissao_outros = "", 0, Txt_pcomissao_outros), IIf(Txt_pdespesas_comerciais_outros = "", 0, Txt_pdespesas_comerciais_outros), IIf(Txt_pdespesas_financeiras_outros = "", 0, Txt_pdespesas_financeiras_outros), IIf(Txt_pdespesas_administrativas_outros = "", 0, Txt_pdespesas_administrativas_outros), IIf(Txt_pfrete_outros = "", 0, Txt_pfrete_outros), IIf(Txt_pmargem_outros = "", 0, Txt_pmargem_outros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pdespesas_comerciais_outros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_comerciais_outros_GotFocus()
On Error GoTo tratar_erro

If Txt_pdespesas_comerciais_outros = "0" Then Txt_pdespesas_comerciais_outros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_comerciais_total_Change()
On Error GoTo tratar_erro

If Txt_pdespesas_comerciais_total <> "" Then
    VerifNumero = Txt_pdespesas_comerciais_total
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pdespesas_comerciais_total = ""
        Txt_pdespesas_comerciais_total.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_total = "", 0, Txt_pICMS_total), IIf(Txt_pPIS_total = "", 0, Txt_pPIS_total), IIf(Txt_pcofins_total = "", 0, Txt_pcofins_total), IIf(Txt_pCSLL_total = "", 0, Txt_pCSLL_total), IIf(Txt_pISSQN_total = "", 0, Txt_pISSQN_total), IIf(Txt_pIRPJ_total = "", 0, Txt_pIRPJ_total), IIf(Txt_psimples_total = "", 0, Txt_psimples_total), IIf(Txt_pcomissao_total = "", 0, Txt_pcomissao_total), IIf(Txt_pdespesas_comerciais_total = "", 0, Txt_pdespesas_comerciais_total), IIf(Txt_pdespesas_financeiras_total = "", 0, Txt_pdespesas_financeiras_total), IIf(Txt_pdespesas_administrativas_total = "", 0, Txt_pdespesas_administrativas_total), IIf(Txt_pfrete_total = "", 0, Txt_pfrete_total), IIf(Txt_pmargem_total = "", 0, Txt_pmargem_total)) = True Then
    ProcCalculaImpostos
Else
    Txt_pdespesas_comerciais_total = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_comerciais_total_GotFocus()
On Error GoTo tratar_erro

If Txt_pdespesas_comerciais_total = "0" Then Txt_pdespesas_comerciais_total = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_financeiras_outros_Change()
On Error GoTo tratar_erro

If Txt_pdespesas_financeiras_outros <> "" Then
    VerifNumero = Txt_pdespesas_financeiras_outros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pdespesas_financeiras_outros = ""
        Txt_pdespesas_financeiras_outros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_outros = "", 0, Txt_pICMS_outros), IIf(Txt_pPIS_outros = "", 0, Txt_pPIS_outros), IIf(Txt_pcofins_outros = "", 0, Txt_pcofins_outros), IIf(Txt_pCSLL_outros = "", 0, Txt_pCSLL_outros), IIf(Txt_pISSQN_outros = "", 0, Txt_pISSQN_outros), IIf(Txt_pIRPJ_outros = "", 0, Txt_pIRPJ_outros), IIf(Txt_psimples_outros = "", 0, Txt_psimples_outros), IIf(Txt_pcomissao_outros = "", 0, Txt_pcomissao_outros), IIf(Txt_pdespesas_comerciais_outros = "", 0, Txt_pdespesas_comerciais_outros), IIf(Txt_pdespesas_financeiras_outros = "", 0, Txt_pdespesas_financeiras_outros), IIf(Txt_pdespesas_administrativas_outros = "", 0, Txt_pdespesas_administrativas_outros), IIf(Txt_pfrete_outros = "", 0, Txt_pfrete_outros), IIf(Txt_pmargem_outros = "", 0, Txt_pmargem_outros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pdespesas_financeiras_outros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_financeiras_outros_GotFocus()
On Error GoTo tratar_erro

If Txt_pdespesas_financeiras_outros = "0" Then Txt_pdespesas_financeiras_outros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_financeiras_total_Change()
On Error GoTo tratar_erro

If Txt_pdespesas_financeiras_total <> "" Then
    VerifNumero = Txt_pdespesas_financeiras_total
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pdespesas_financeiras_total = ""
        Txt_pdespesas_financeiras_total.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_total = "", 0, Txt_pICMS_total), IIf(Txt_pPIS_total = "", 0, Txt_pPIS_total), IIf(Txt_pcofins_total = "", 0, Txt_pcofins_total), IIf(Txt_pCSLL_total = "", 0, Txt_pCSLL_total), IIf(Txt_pISSQN_total = "", 0, Txt_pISSQN_total), IIf(Txt_pIRPJ_total = "", 0, Txt_pIRPJ_total), IIf(Txt_psimples_total = "", 0, Txt_psimples_total), IIf(Txt_pcomissao_total = "", 0, Txt_pcomissao_total), IIf(Txt_pdespesas_comerciais_total = "", 0, Txt_pdespesas_comerciais_total), IIf(Txt_pdespesas_financeiras_total = "", 0, Txt_pdespesas_financeiras_total), IIf(Txt_pdespesas_administrativas_total = "", 0, Txt_pdespesas_administrativas_total), IIf(Txt_pfrete_total = "", 0, Txt_pfrete_total), IIf(Txt_pmargem_total = "", 0, Txt_pmargem_total)) = True Then
    ProcCalculaImpostos
Else
    Txt_pdespesas_financeiras_total = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_financeiras_total_GotFocus()
On Error GoTo tratar_erro

If Txt_pdespesas_financeiras_total = "0" Then Txt_pdespesas_financeiras_total = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pfrete_outros_Change()
On Error GoTo tratar_erro

If Txt_pfrete_outros <> "" Then
    VerifNumero = Txt_pfrete_outros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pfrete_outros = ""
        Txt_pfrete_outros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_outros = "", 0, Txt_pICMS_outros), IIf(Txt_pPIS_outros = "", 0, Txt_pPIS_outros), IIf(Txt_pcofins_outros = "", 0, Txt_pcofins_outros), IIf(Txt_pCSLL_outros = "", 0, Txt_pCSLL_outros), IIf(Txt_pISSQN_outros = "", 0, Txt_pISSQN_outros), IIf(Txt_pIRPJ_outros = "", 0, Txt_pIRPJ_outros), IIf(Txt_psimples_outros = "", 0, Txt_psimples_outros), IIf(Txt_pcomissao_outros = "", 0, Txt_pcomissao_outros), IIf(Txt_pdespesas_comerciais_outros = "", 0, Txt_pdespesas_comerciais_outros), IIf(Txt_pdespesas_financeiras_outros = "", 0, Txt_pdespesas_financeiras_outros), IIf(Txt_pdespesas_administrativas_outros = "", 0, Txt_pdespesas_administrativas_outros), IIf(Txt_pfrete_outros = "", 0, Txt_pfrete_outros), IIf(Txt_pmargem_outros = "", 0, Txt_pmargem_outros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pfrete_outros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pfrete_outros_GotFocus()
On Error GoTo tratar_erro

If Txt_pfrete_outros = "0" Then Txt_pfrete_outros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pfrete_total_Change()
On Error GoTo tratar_erro

If Txt_pfrete_total <> "" Then
    VerifNumero = Txt_pfrete_total
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pfrete_total = ""
        Txt_pfrete_total.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_total = "", 0, Txt_pICMS_total), IIf(Txt_pPIS_total = "", 0, Txt_pPIS_total), IIf(Txt_pcofins_total = "", 0, Txt_pcofins_total), IIf(Txt_pCSLL_total = "", 0, Txt_pCSLL_total), IIf(Txt_pISSQN_total = "", 0, Txt_pISSQN_total), IIf(Txt_pIRPJ_total = "", 0, Txt_pIRPJ_total), IIf(Txt_psimples_total = "", 0, Txt_psimples_total), IIf(Txt_pcomissao_total = "", 0, Txt_pcomissao_total), IIf(Txt_pdespesas_comerciais_total = "", 0, Txt_pdespesas_comerciais_total), IIf(Txt_pdespesas_financeiras_total = "", 0, Txt_pdespesas_financeiras_total), IIf(Txt_pdespesas_administrativas_total = "", 0, Txt_pdespesas_administrativas_total), IIf(Txt_pfrete_total = "", 0, Txt_pfrete_total), IIf(Txt_pmargem_total = "", 0, Txt_pmargem_total)) = True Then
    ProcCalculaImpostos
Else
    Txt_pfrete_total = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pfrete_total_GotFocus()
On Error GoTo tratar_erro

If Txt_pfrete_total = "0" Then Txt_pfrete_total = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pICMS_outros_Change()
On Error GoTo tratar_erro

If Txt_pICMS_outros <> "" Then
    VerifNumero = Txt_pICMS_outros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pICMS_outros = ""
        Txt_pICMS_outros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_outros = "", 0, Txt_pICMS_outros), IIf(Txt_pPIS_outros = "", 0, Txt_pPIS_outros), IIf(Txt_pcofins_outros = "", 0, Txt_pcofins_outros), IIf(Txt_pCSLL_outros = "", 0, Txt_pCSLL_outros), IIf(Txt_pISSQN_outros = "", 0, Txt_pISSQN_outros), IIf(Txt_pIRPJ_outros = "", 0, Txt_pIRPJ_outros), IIf(Txt_psimples_outros = "", 0, Txt_psimples_outros), IIf(Txt_pcomissao_outros = "", 0, Txt_pcomissao_outros), IIf(Txt_pdespesas_comerciais_outros = "", 0, Txt_pdespesas_comerciais_outros), IIf(Txt_pdespesas_financeiras_outros = "", 0, Txt_pdespesas_financeiras_outros), IIf(Txt_pdespesas_administrativas_outros = "", 0, Txt_pdespesas_administrativas_outros), IIf(Txt_pfrete_outros = "", 0, Txt_pfrete_outros), IIf(Txt_pmargem_outros = "", 0, Txt_pmargem_outros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pICMS_outros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pICMS_outros_GotFocus()
On Error GoTo tratar_erro

If Txt_pICMS_outros = "0" Then Txt_pICMS_outros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pICMS_processo_Change()
On Error GoTo tratar_erro

If Txt_pICMS_processo <> "" Then
    VerifNumero = Txt_pICMS_processo
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pICMS_processo = ""
        Txt_pICMS_processo.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_processo = "", 0, Txt_pICMS_processo), IIf(Txt_pPIS_processo = "", 0, Txt_pPIS_processo), IIf(Txt_pcofins_processo = "", 0, Txt_pcofins_processo), IIf(Txt_pCSLL_processo = "", 0, Txt_pCSLL_processo), IIf(Txt_pISSQN_processo = "", 0, Txt_pISSQN_processo), IIf(Txt_pIRPJ_processo = "", 0, Txt_pIRPJ_processo), IIf(Txt_psimples_processo = "", 0, Txt_psimples_processo), IIf(Txt_pcomissao_processo = "", 0, Txt_pcomissao_processo), IIf(Txt_pdespesas_comerciais_processo = "", 0, Txt_pdespesas_comerciais_processo), IIf(Txt_pdespesas_financeiras_processo = "", 0, Txt_pdespesas_financeiras_processo), IIf(Txt_pdespesas_administrativas_processo = "", 0, Txt_pdespesas_administrativas_processo), IIf(Txt_pfrete_processo = "", 0, Txt_pfrete_processo), IIf(Txt_pmargem_processo = "", 0, Txt_pmargem_processo)) = True Then
    ProcCalculaImpostos
Else
    Txt_pICMS_processo = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pICMS_total_Change()
On Error GoTo tratar_erro

If Txt_pICMS_total <> "" Then
    VerifNumero = Txt_pICMS_total
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pICMS_total = ""
        Txt_pICMS_total.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_total = "", 0, Txt_pICMS_total), IIf(Txt_pPIS_total = "", 0, Txt_pPIS_total), IIf(Txt_pcofins_total = "", 0, Txt_pcofins_total), IIf(Txt_pCSLL_total = "", 0, Txt_pCSLL_total), IIf(Txt_pISSQN_total = "", 0, Txt_pISSQN_total), IIf(Txt_pIRPJ_total = "", 0, Txt_pIRPJ_total), IIf(Txt_psimples_total = "", 0, Txt_psimples_total), IIf(Txt_pcomissao_total = "", 0, Txt_pcomissao_total), IIf(Txt_pdespesas_comerciais_total = "", 0, Txt_pdespesas_comerciais_total), IIf(Txt_pdespesas_financeiras_total = "", 0, Txt_pdespesas_financeiras_total), IIf(Txt_pdespesas_administrativas_total = "", 0, Txt_pdespesas_administrativas_total), IIf(Txt_pfrete_total = "", 0, Txt_pfrete_total), IIf(Txt_pmargem_total = "", 0, Txt_pmargem_total)) = True Then
    ProcCalculaImpostos
Else
    Txt_pICMS_total = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pICMS_total_GotFocus()
On Error GoTo tratar_erro

If Txt_pICMS_total = "0" Then Txt_pICMS_total = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pIRPJ_outros_Change()
On Error GoTo tratar_erro

If Txt_pIRPJ_outros <> "" Then
    VerifNumero = Txt_pIRPJ_outros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pIRPJ_outros = ""
        Txt_pIRPJ_outros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_outros = "", 0, Txt_pICMS_outros), IIf(Txt_pPIS_outros = "", 0, Txt_pPIS_outros), IIf(Txt_pcofins_outros = "", 0, Txt_pcofins_outros), IIf(Txt_pCSLL_outros = "", 0, Txt_pCSLL_outros), IIf(Txt_pISSQN_outros = "", 0, Txt_pISSQN_outros), IIf(Txt_pIRPJ_outros = "", 0, Txt_pIRPJ_outros), IIf(Txt_psimples_outros = "", 0, Txt_psimples_outros), IIf(Txt_pcomissao_outros = "", 0, Txt_pcomissao_outros), IIf(Txt_pdespesas_comerciais_outros = "", 0, Txt_pdespesas_comerciais_outros), IIf(Txt_pdespesas_financeiras_outros = "", 0, Txt_pdespesas_financeiras_outros), IIf(Txt_pdespesas_administrativas_outros = "", 0, Txt_pdespesas_administrativas_outros), IIf(Txt_pfrete_outros = "", 0, Txt_pfrete_outros), IIf(Txt_pmargem_outros = "", 0, Txt_pmargem_outros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pIRPJ_outros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pIRPJ_outros_GotFocus()
On Error GoTo tratar_erro

If Txt_pIRPJ_outros = "0" Then Txt_pIRPJ_outros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pIRPJ_total_Change()
On Error GoTo tratar_erro

If Txt_pIRPJ_total <> "" Then
    VerifNumero = Txt_pIRPJ_total
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pIRPJ_total = ""
        Txt_pIRPJ_total.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_total = "", 0, Txt_pICMS_total), IIf(Txt_pPIS_total = "", 0, Txt_pPIS_total), IIf(Txt_pcofins_total = "", 0, Txt_pcofins_total), IIf(Txt_pCSLL_total = "", 0, Txt_pCSLL_total), IIf(Txt_pISSQN_total = "", 0, Txt_pISSQN_total), IIf(Txt_pIRPJ_total = "", 0, Txt_pIRPJ_total), IIf(Txt_psimples_total = "", 0, Txt_psimples_total), IIf(Txt_pcomissao_total = "", 0, Txt_pcomissao_total), IIf(Txt_pdespesas_comerciais_total = "", 0, Txt_pdespesas_comerciais_total), IIf(Txt_pdespesas_financeiras_total = "", 0, Txt_pdespesas_financeiras_total), IIf(Txt_pdespesas_administrativas_total = "", 0, Txt_pdespesas_administrativas_total), IIf(Txt_pfrete_total = "", 0, Txt_pfrete_total), IIf(Txt_pmargem_total = "", 0, Txt_pmargem_total)) = True Then
    ProcCalculaImpostos
Else
    Txt_pIRPJ_total = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pIRPJ_total_GotFocus()
On Error GoTo tratar_erro

If Txt_pIRPJ_total = "0" Then Txt_pIRPJ_total = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pISSQN_outros_Change()
On Error GoTo tratar_erro

If Txt_pISSQN_outros <> "" Then
    VerifNumero = Txt_pISSQN_outros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pISSQN_outros = ""
        Txt_pISSQN_outros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_outros = "", 0, Txt_pICMS_outros), IIf(Txt_pPIS_outros = "", 0, Txt_pPIS_outros), IIf(Txt_pcofins_outros = "", 0, Txt_pcofins_outros), IIf(Txt_pCSLL_outros = "", 0, Txt_pCSLL_outros), IIf(Txt_pISSQN_outros = "", 0, Txt_pISSQN_outros), IIf(Txt_pIRPJ_outros = "", 0, Txt_pIRPJ_outros), IIf(Txt_psimples_outros = "", 0, Txt_psimples_outros), IIf(Txt_pcomissao_outros = "", 0, Txt_pcomissao_outros), IIf(Txt_pdespesas_comerciais_outros = "", 0, Txt_pdespesas_comerciais_outros), IIf(Txt_pdespesas_financeiras_outros = "", 0, Txt_pdespesas_financeiras_outros), IIf(Txt_pdespesas_administrativas_outros = "", 0, Txt_pdespesas_administrativas_outros), IIf(Txt_pfrete_outros = "", 0, Txt_pfrete_outros), IIf(Txt_pmargem_outros = "", 0, Txt_pmargem_outros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pISSQN_outros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pISSQN_outros_GotFocus()
On Error GoTo tratar_erro

If Txt_pISSQN_outros = "0" Then Txt_pISSQN_outros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pISSQN_total_Change()
On Error GoTo tratar_erro

If Txt_pISSQN_total <> "" Then
    VerifNumero = Txt_pISSQN_total
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pISSQN_total = ""
        Txt_pISSQN_total.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_total = "", 0, Txt_pICMS_total), IIf(Txt_pPIS_total = "", 0, Txt_pPIS_total), IIf(Txt_pcofins_total = "", 0, Txt_pcofins_total), IIf(Txt_pCSLL_total = "", 0, Txt_pCSLL_total), IIf(Txt_pISSQN_total = "", 0, Txt_pISSQN_total), IIf(Txt_pIRPJ_total = "", 0, Txt_pIRPJ_total), IIf(Txt_psimples_total = "", 0, Txt_psimples_total), IIf(Txt_pcomissao_total = "", 0, Txt_pcomissao_total), IIf(Txt_pdespesas_comerciais_total = "", 0, Txt_pdespesas_comerciais_total), IIf(Txt_pdespesas_financeiras_total = "", 0, Txt_pdespesas_financeiras_total), IIf(Txt_pdespesas_administrativas_total = "", 0, Txt_pdespesas_administrativas_total), IIf(Txt_pfrete_total = "", 0, Txt_pfrete_total), IIf(Txt_pmargem_total = "", 0, Txt_pmargem_total)) = True Then
    ProcCalculaImpostos
Else
    Txt_pISSQN_total = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pISSQN_total_GotFocus()
On Error GoTo tratar_erro

If Txt_pISSQN_total = "0" Then Txt_pISSQN_total = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pmargem_outros_Change()
On Error GoTo tratar_erro

If Txt_pmargem_outros <> "" Then
    VerifNumero = Txt_pmargem_outros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pmargem_outros = ""
        Txt_pmargem_outros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_outros = "", 0, Txt_pICMS_outros), IIf(Txt_pPIS_outros = "", 0, Txt_pPIS_outros), IIf(Txt_pcofins_outros = "", 0, Txt_pcofins_outros), IIf(Txt_pCSLL_outros = "", 0, Txt_pCSLL_outros), IIf(Txt_pISSQN_outros = "", 0, Txt_pISSQN_outros), IIf(Txt_pIRPJ_outros = "", 0, Txt_pIRPJ_outros), IIf(Txt_psimples_outros = "", 0, Txt_psimples_outros), IIf(Txt_pcomissao_outros = "", 0, Txt_pcomissao_outros), IIf(Txt_pdespesas_comerciais_outros = "", 0, Txt_pdespesas_comerciais_outros), IIf(Txt_pdespesas_financeiras_outros = "", 0, Txt_pdespesas_financeiras_outros), IIf(Txt_pdespesas_administrativas_outros = "", 0, Txt_pdespesas_administrativas_outros), IIf(Txt_pfrete_outros = "", 0, Txt_pfrete_outros), IIf(Txt_pmargem_outros = "", 0, Txt_pmargem_outros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pmargem_outros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pmargem_outros_GotFocus()
On Error GoTo tratar_erro

If Txt_pmargem_outros = "0" Then Txt_pmargem_outros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pmargem_total_Change()
On Error GoTo tratar_erro

If Txt_pmargem_total <> "" Then
    VerifNumero = Txt_pmargem_total
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pmargem_total = ""
        Txt_pmargem_total.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_total = "", 0, Txt_pICMS_total), IIf(Txt_pPIS_total = "", 0, Txt_pPIS_total), IIf(Txt_pcofins_total = "", 0, Txt_pcofins_total), IIf(Txt_pCSLL_total = "", 0, Txt_pCSLL_total), IIf(Txt_pISSQN_total = "", 0, Txt_pISSQN_total), IIf(Txt_pIRPJ_total = "", 0, Txt_pIRPJ_total), IIf(Txt_psimples_total = "", 0, Txt_psimples_total), IIf(Txt_pcomissao_total = "", 0, Txt_pcomissao_total), IIf(Txt_pdespesas_comerciais_total = "", 0, Txt_pdespesas_comerciais_total), IIf(Txt_pdespesas_financeiras_total = "", 0, Txt_pdespesas_financeiras_total), IIf(Txt_pdespesas_administrativas_total = "", 0, Txt_pdespesas_administrativas_total), IIf(Txt_pfrete_total = "", 0, Txt_pfrete_total), IIf(Txt_pmargem_total = "", 0, Txt_pmargem_total)) = True Then
    ProcCalculaImpostos
Else
    Txt_pmargem_total = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pmargem_total_GotFocus()
On Error GoTo tratar_erro

If Txt_pmargem_total = "0" Then Txt_pmargem_total = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pPIS_outros_Change()
On Error GoTo tratar_erro

If Txt_pPIS_outros <> "" Then
    VerifNumero = Txt_pPIS_outros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pPIS_outros = ""
        Txt_pPIS_outros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_outros = "", 0, Txt_pICMS_outros), IIf(Txt_pPIS_outros = "", 0, Txt_pPIS_outros), IIf(Txt_pcofins_outros = "", 0, Txt_pcofins_outros), IIf(Txt_pCSLL_outros = "", 0, Txt_pCSLL_outros), IIf(Txt_pISSQN_outros = "", 0, Txt_pISSQN_outros), IIf(Txt_pIRPJ_outros = "", 0, Txt_pIRPJ_outros), IIf(Txt_psimples_outros = "", 0, Txt_psimples_outros), IIf(Txt_pcomissao_outros = "", 0, Txt_pcomissao_outros), IIf(Txt_pdespesas_comerciais_outros = "", 0, Txt_pdespesas_comerciais_outros), IIf(Txt_pdespesas_financeiras_outros = "", 0, Txt_pdespesas_financeiras_outros), IIf(Txt_pdespesas_administrativas_outros = "", 0, Txt_pdespesas_administrativas_outros), IIf(Txt_pfrete_outros = "", 0, Txt_pfrete_outros), IIf(Txt_pmargem_outros = "", 0, Txt_pmargem_outros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pPIS_outros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pPIS_outros_GotFocus()
On Error GoTo tratar_erro

If Txt_pPIS_outros = "0" Then Txt_pPIS_outros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pPIS_processo_Change()
On Error GoTo tratar_erro

If Txt_pPIS_processo <> "" Then
    VerifNumero = Txt_pPIS_processo
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pPIS_processo = ""
        Txt_pPIS_processo.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_processo = "", 0, Txt_pICMS_processo), IIf(Txt_pPIS_processo = "", 0, Txt_pPIS_processo), IIf(Txt_pcofins_processo = "", 0, Txt_pcofins_processo), IIf(Txt_pCSLL_processo = "", 0, Txt_pCSLL_processo), IIf(Txt_pISSQN_processo = "", 0, Txt_pISSQN_processo), IIf(Txt_pIRPJ_processo = "", 0, Txt_pIRPJ_processo), IIf(Txt_psimples_processo = "", 0, Txt_psimples_processo), IIf(Txt_pcomissao_processo = "", 0, Txt_pcomissao_processo), IIf(Txt_pdespesas_comerciais_processo = "", 0, Txt_pdespesas_comerciais_processo), IIf(Txt_pdespesas_financeiras_processo = "", 0, Txt_pdespesas_financeiras_processo), IIf(Txt_pdespesas_administrativas_processo = "", 0, Txt_pdespesas_administrativas_processo), IIf(Txt_pfrete_processo = "", 0, Txt_pfrete_processo), IIf(Txt_pmargem_processo = "", 0, Txt_pmargem_processo)) = True Then
    ProcCalculaImpostos
Else
    Txt_pPIS_processo = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcofins_processo_Change()
On Error GoTo tratar_erro

If Txt_pcofins_processo <> "" Then
    VerifNumero = Txt_pcofins_processo
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pcofins_processo = ""
        Txt_pcofins_processo.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_processo = "", 0, Txt_pICMS_processo), IIf(Txt_pPIS_processo = "", 0, Txt_pPIS_processo), IIf(Txt_pcofins_processo = "", 0, Txt_pcofins_processo), IIf(Txt_pCSLL_processo = "", 0, Txt_pCSLL_processo), IIf(Txt_pISSQN_processo = "", 0, Txt_pISSQN_processo), IIf(Txt_pIRPJ_processo = "", 0, Txt_pIRPJ_processo), IIf(Txt_psimples_processo = "", 0, Txt_psimples_processo), IIf(Txt_pcomissao_processo = "", 0, Txt_pcomissao_processo), IIf(Txt_pdespesas_comerciais_processo = "", 0, Txt_pdespesas_comerciais_processo), IIf(Txt_pdespesas_financeiras_processo = "", 0, Txt_pdespesas_financeiras_processo), IIf(Txt_pdespesas_administrativas_processo = "", 0, Txt_pdespesas_administrativas_processo), IIf(Txt_pfrete_processo = "", 0, Txt_pfrete_processo), IIf(Txt_pmargem_processo = "", 0, Txt_pmargem_processo)) = True Then
    ProcCalculaImpostos
Else
    Txt_pcofins_processo = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pCSLL_processo_Change()
On Error GoTo tratar_erro

If Txt_pCSLL_processo <> "" Then
    VerifNumero = Txt_pCSLL_processo
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pCSLL_processo = ""
        Txt_pCSLL_processo.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_processo = "", 0, Txt_pICMS_processo), IIf(Txt_pPIS_processo = "", 0, Txt_pPIS_processo), IIf(Txt_pcofins_processo = "", 0, Txt_pcofins_processo), IIf(Txt_pCSLL_processo = "", 0, Txt_pCSLL_processo), IIf(Txt_pISSQN_processo = "", 0, Txt_pISSQN_processo), IIf(Txt_pIRPJ_processo = "", 0, Txt_pIRPJ_processo), IIf(Txt_psimples_processo = "", 0, Txt_psimples_processo), IIf(Txt_pcomissao_processo = "", 0, Txt_pcomissao_processo), IIf(Txt_pdespesas_comerciais_processo = "", 0, Txt_pdespesas_comerciais_processo), IIf(Txt_pdespesas_financeiras_processo = "", 0, Txt_pdespesas_financeiras_processo), IIf(Txt_pdespesas_administrativas_processo = "", 0, Txt_pdespesas_administrativas_processo), IIf(Txt_pfrete_processo = "", 0, Txt_pfrete_processo), IIf(Txt_pmargem_processo = "", 0, Txt_pmargem_processo)) = True Then
    ProcCalculaImpostos
Else
    Txt_pCSLL_processo = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pISSQN_processo_Change()
On Error GoTo tratar_erro

If Txt_pISSQN_processo <> "" Then
    VerifNumero = Txt_pISSQN_processo
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pISSQN_processo = ""
        Txt_pISSQN_processo.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_processo = "", 0, Txt_pICMS_processo), IIf(Txt_pPIS_processo = "", 0, Txt_pPIS_processo), IIf(Txt_pcofins_processo = "", 0, Txt_pcofins_processo), IIf(Txt_pCSLL_processo = "", 0, Txt_pCSLL_processo), IIf(Txt_pISSQN_processo = "", 0, Txt_pISSQN_processo), IIf(Txt_pIRPJ_processo = "", 0, Txt_pIRPJ_processo), IIf(Txt_psimples_processo = "", 0, Txt_psimples_processo), IIf(Txt_pcomissao_processo = "", 0, Txt_pcomissao_processo), IIf(Txt_pdespesas_comerciais_processo = "", 0, Txt_pdespesas_comerciais_processo), IIf(Txt_pdespesas_financeiras_processo = "", 0, Txt_pdespesas_financeiras_processo), IIf(Txt_pdespesas_administrativas_processo = "", 0, Txt_pdespesas_administrativas_processo), IIf(Txt_pfrete_processo = "", 0, Txt_pfrete_processo), IIf(Txt_pmargem_processo = "", 0, Txt_pmargem_processo)) = True Then
    ProcCalculaImpostos
Else
    Txt_pISSQN_processo = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pIRPJ_processo_Change()
On Error GoTo tratar_erro

If Txt_pIRPJ_processo <> "" Then
    VerifNumero = Txt_pIRPJ_processo
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pIRPJ_processo = ""
        Txt_pIRPJ_processo.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_processo = "", 0, Txt_pICMS_processo), IIf(Txt_pPIS_processo = "", 0, Txt_pPIS_processo), IIf(Txt_pcofins_processo = "", 0, Txt_pcofins_processo), IIf(Txt_pCSLL_processo = "", 0, Txt_pCSLL_processo), IIf(Txt_pISSQN_processo = "", 0, Txt_pISSQN_processo), IIf(Txt_pIRPJ_processo = "", 0, Txt_pIRPJ_processo), IIf(Txt_psimples_processo = "", 0, Txt_psimples_processo), IIf(Txt_pcomissao_processo = "", 0, Txt_pcomissao_processo), IIf(Txt_pdespesas_comerciais_processo = "", 0, Txt_pdespesas_comerciais_processo), IIf(Txt_pdespesas_financeiras_processo = "", 0, Txt_pdespesas_financeiras_processo), IIf(Txt_pdespesas_administrativas_processo = "", 0, Txt_pdespesas_administrativas_processo), IIf(Txt_pfrete_processo = "", 0, Txt_pfrete_processo), IIf(Txt_pmargem_processo = "", 0, Txt_pmargem_processo)) = True Then
    ProcCalculaImpostos
Else
    Txt_pIRPJ_processo = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pPIS_total_Change()
On Error GoTo tratar_erro

If Txt_pPIS_total <> "" Then
    VerifNumero = Txt_pPIS_total
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pPIS_total = ""
        Txt_pPIS_total.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_total = "", 0, Txt_pICMS_total), IIf(Txt_pPIS_total = "", 0, Txt_pPIS_total), IIf(Txt_pcofins_total = "", 0, Txt_pcofins_total), IIf(Txt_pCSLL_total = "", 0, Txt_pCSLL_total), IIf(Txt_pISSQN_total = "", 0, Txt_pISSQN_total), IIf(Txt_pIRPJ_total = "", 0, Txt_pIRPJ_total), IIf(Txt_psimples_total = "", 0, Txt_psimples_total), IIf(Txt_pcomissao_total = "", 0, Txt_pcomissao_total), IIf(Txt_pdespesas_comerciais_total = "", 0, Txt_pdespesas_comerciais_total), IIf(Txt_pdespesas_financeiras_total = "", 0, Txt_pdespesas_financeiras_total), IIf(Txt_pdespesas_administrativas_total = "", 0, Txt_pdespesas_administrativas_total), IIf(Txt_pfrete_total = "", 0, Txt_pfrete_total), IIf(Txt_pmargem_total = "", 0, Txt_pmargem_total)) = True Then
    ProcCalculaImpostos
Else
    Txt_pPIS_total = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pPIS_total_GotFocus()
On Error GoTo tratar_erro

If Txt_pPIS_total = "0" Then Txt_pPIS_total = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_psimples_outros_Change()
On Error GoTo tratar_erro

If Txt_psimples_outros <> "" Then
    VerifNumero = Txt_psimples_outros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_psimples_outros = ""
        Txt_psimples_outros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_outros = "", 0, Txt_pICMS_outros), IIf(Txt_pPIS_outros = "", 0, Txt_pPIS_outros), IIf(Txt_pcofins_outros = "", 0, Txt_pcofins_outros), IIf(Txt_pCSLL_outros = "", 0, Txt_pCSLL_outros), IIf(Txt_pISSQN_outros = "", 0, Txt_pISSQN_outros), IIf(Txt_pIRPJ_outros = "", 0, Txt_pIRPJ_outros), IIf(Txt_psimples_outros = "", 0, Txt_psimples_outros), IIf(Txt_pcomissao_outros = "", 0, Txt_pcomissao_outros), IIf(Txt_pdespesas_comerciais_outros = "", 0, Txt_pdespesas_comerciais_outros), IIf(Txt_pdespesas_financeiras_outros = "", 0, Txt_pdespesas_financeiras_outros), IIf(Txt_pdespesas_administrativas_outros = "", 0, Txt_pdespesas_administrativas_outros), IIf(Txt_pfrete_outros = "", 0, Txt_pfrete_outros), IIf(Txt_pmargem_outros = "", 0, Txt_pmargem_outros)) = True Then
    ProcCalculaImpostos
Else
    Txt_psimples_outros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_psimples_outros_GotFocus()
On Error GoTo tratar_erro

If Txt_psimples_outros = "0" Then Txt_psimples_outros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_psimples_processo_Change()
On Error GoTo tratar_erro

If Txt_psimples_processo <> "" Then
    VerifNumero = Txt_psimples_processo
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_psimples_processo = ""
        Txt_psimples_processo.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_processo = "", 0, Txt_pICMS_processo), IIf(Txt_pPIS_processo = "", 0, Txt_pPIS_processo), IIf(Txt_pcofins_processo = "", 0, Txt_pcofins_processo), IIf(Txt_pCSLL_processo = "", 0, Txt_pCSLL_processo), IIf(Txt_pISSQN_processo = "", 0, Txt_pISSQN_processo), IIf(Txt_pIRPJ_processo = "", 0, Txt_pIRPJ_processo), IIf(Txt_psimples_processo = "", 0, Txt_psimples_processo), IIf(Txt_pcomissao_processo = "", 0, Txt_pcomissao_processo), IIf(Txt_pdespesas_comerciais_processo = "", 0, Txt_pdespesas_comerciais_processo), IIf(Txt_pdespesas_financeiras_processo = "", 0, Txt_pdespesas_financeiras_processo), IIf(Txt_pdespesas_administrativas_processo = "", 0, Txt_pdespesas_administrativas_processo), IIf(Txt_pfrete_processo = "", 0, Txt_pfrete_processo), IIf(Txt_pmargem_processo = "", 0, Txt_pmargem_processo)) = True Then
    ProcCalculaImpostos
Else
    Txt_psimples_processo = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcomissao_processo_Change()
On Error GoTo tratar_erro

If Txt_pcomissao_processo <> "" Then
    VerifNumero = Txt_pcomissao_processo
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pcomissao_processo = ""
        Txt_pcomissao_processo.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_processo = "", 0, Txt_pICMS_processo), IIf(Txt_pPIS_processo = "", 0, Txt_pPIS_processo), IIf(Txt_pcofins_processo = "", 0, Txt_pcofins_processo), IIf(Txt_pCSLL_processo = "", 0, Txt_pCSLL_processo), IIf(Txt_pISSQN_processo = "", 0, Txt_pISSQN_processo), IIf(Txt_pIRPJ_processo = "", 0, Txt_pIRPJ_processo), IIf(Txt_psimples_processo = "", 0, Txt_psimples_processo), IIf(Txt_pcomissao_processo = "", 0, Txt_pcomissao_processo), IIf(Txt_pdespesas_comerciais_processo = "", 0, Txt_pdespesas_comerciais_processo), IIf(Txt_pdespesas_financeiras_processo = "", 0, Txt_pdespesas_financeiras_processo), IIf(Txt_pdespesas_administrativas_processo = "", 0, Txt_pdespesas_administrativas_processo), IIf(Txt_pfrete_processo = "", 0, Txt_pfrete_processo), IIf(Txt_pmargem_processo = "", 0, Txt_pmargem_processo)) = True Then
    ProcCalculaImpostos
Else
    Txt_pcomissao_processo = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_comerciais_processo_Change()
On Error GoTo tratar_erro

If Txt_pdespesas_comerciais_processo <> "" Then
    VerifNumero = Txt_pdespesas_comerciais_processo
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pdespesas_comerciais_processo = ""
        Txt_pdespesas_comerciais_processo.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_processo = "", 0, Txt_pICMS_processo), IIf(Txt_pPIS_processo = "", 0, Txt_pPIS_processo), IIf(Txt_pcofins_processo = "", 0, Txt_pcofins_processo), IIf(Txt_pCSLL_processo = "", 0, Txt_pCSLL_processo), IIf(Txt_pISSQN_processo = "", 0, Txt_pISSQN_processo), IIf(Txt_pIRPJ_processo = "", 0, Txt_pIRPJ_processo), IIf(Txt_psimples_processo = "", 0, Txt_psimples_processo), IIf(Txt_pcomissao_processo = "", 0, Txt_pcomissao_processo), IIf(Txt_pdespesas_comerciais_processo = "", 0, Txt_pdespesas_comerciais_processo), IIf(Txt_pdespesas_financeiras_processo = "", 0, Txt_pdespesas_financeiras_processo), IIf(Txt_pdespesas_administrativas_processo = "", 0, Txt_pdespesas_administrativas_processo), IIf(Txt_pfrete_processo = "", 0, Txt_pfrete_processo), IIf(Txt_pmargem_processo = "", 0, Txt_pmargem_processo)) = True Then
    ProcCalculaImpostos
Else
    Txt_pdespesas_comerciais_processo = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_financeiras_processo_Change()
On Error GoTo tratar_erro

If Txt_pdespesas_financeiras_processo <> "" Then
    VerifNumero = Txt_pdespesas_financeiras_processo
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pdespesas_financeiras_processo = ""
        Txt_pdespesas_financeiras_processo.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_processo = "", 0, Txt_pICMS_processo), IIf(Txt_pPIS_processo = "", 0, Txt_pPIS_processo), IIf(Txt_pcofins_processo = "", 0, Txt_pcofins_processo), IIf(Txt_pCSLL_processo = "", 0, Txt_pCSLL_processo), IIf(Txt_pISSQN_processo = "", 0, Txt_pISSQN_processo), IIf(Txt_pIRPJ_processo = "", 0, Txt_pIRPJ_processo), IIf(Txt_psimples_processo = "", 0, Txt_psimples_processo), IIf(Txt_pcomissao_processo = "", 0, Txt_pcomissao_processo), IIf(Txt_pdespesas_comerciais_processo = "", 0, Txt_pdespesas_comerciais_processo), IIf(Txt_pdespesas_financeiras_processo = "", 0, Txt_pdespesas_financeiras_processo), IIf(Txt_pdespesas_administrativas_processo = "", 0, Txt_pdespesas_administrativas_processo), IIf(Txt_pfrete_processo = "", 0, Txt_pfrete_processo), IIf(Txt_pmargem_processo = "", 0, Txt_pmargem_processo)) = True Then
    ProcCalculaImpostos
Else
    Txt_pdespesas_financeiras_processo = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_administrativas_processo_Change()
On Error GoTo tratar_erro

If Txt_pdespesas_administrativas_processo <> "" Then
    VerifNumero = Txt_pdespesas_administrativas_processo
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pdespesas_administrativas_processo = ""
        Txt_pdespesas_administrativas_processo.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_processo = "", 0, Txt_pICMS_processo), IIf(Txt_pPIS_processo = "", 0, Txt_pPIS_processo), IIf(Txt_pcofins_processo = "", 0, Txt_pcofins_processo), IIf(Txt_pCSLL_processo = "", 0, Txt_pCSLL_processo), IIf(Txt_pISSQN_processo = "", 0, Txt_pISSQN_processo), IIf(Txt_pIRPJ_processo = "", 0, Txt_pIRPJ_processo), IIf(Txt_psimples_processo = "", 0, Txt_psimples_processo), IIf(Txt_pcomissao_processo = "", 0, Txt_pcomissao_processo), IIf(Txt_pdespesas_comerciais_processo = "", 0, Txt_pdespesas_comerciais_processo), IIf(Txt_pdespesas_financeiras_processo = "", 0, Txt_pdespesas_financeiras_processo), IIf(Txt_pdespesas_administrativas_processo = "", 0, Txt_pdespesas_administrativas_processo), IIf(Txt_pfrete_processo = "", 0, Txt_pfrete_processo), IIf(Txt_pmargem_processo = "", 0, Txt_pmargem_processo)) = True Then
    ProcCalculaImpostos
Else
    Txt_pdespesas_administrativas_processo = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pfrete_processo_Change()
On Error GoTo tratar_erro

If Txt_pfrete_processo <> "" Then
    VerifNumero = Txt_pfrete_processo
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pfrete_processo = ""
        Txt_pfrete_processo.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_processo = "", 0, Txt_pICMS_processo), IIf(Txt_pPIS_processo = "", 0, Txt_pPIS_processo), IIf(Txt_pcofins_processo = "", 0, Txt_pcofins_processo), IIf(Txt_pCSLL_processo = "", 0, Txt_pCSLL_processo), IIf(Txt_pISSQN_processo = "", 0, Txt_pISSQN_processo), IIf(Txt_pIRPJ_processo = "", 0, Txt_pIRPJ_processo), IIf(Txt_psimples_processo = "", 0, Txt_psimples_processo), IIf(Txt_pcomissao_processo = "", 0, Txt_pcomissao_processo), IIf(Txt_pdespesas_comerciais_processo = "", 0, Txt_pdespesas_comerciais_processo), IIf(Txt_pdespesas_financeiras_processo = "", 0, Txt_pdespesas_financeiras_processo), IIf(Txt_pdespesas_administrativas_processo = "", 0, Txt_pdespesas_administrativas_processo), IIf(Txt_pfrete_processo = "", 0, Txt_pfrete_processo), IIf(Txt_pmargem_processo = "", 0, Txt_pmargem_processo)) = True Then
    ProcCalculaImpostos
Else
    Txt_pfrete_processo = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pmargem_processo_Change()
On Error GoTo tratar_erro

If Txt_pmargem_processo <> "" Then
    VerifNumero = Txt_pmargem_processo
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pmargem_processo = ""
        Txt_pmargem_processo.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_processo = "", 0, Txt_pICMS_processo), IIf(Txt_pPIS_processo = "", 0, Txt_pPIS_processo), IIf(Txt_pcofins_processo = "", 0, Txt_pcofins_processo), IIf(Txt_pCSLL_processo = "", 0, Txt_pCSLL_processo), IIf(Txt_pISSQN_processo = "", 0, Txt_pISSQN_processo), IIf(Txt_pIRPJ_processo = "", 0, Txt_pIRPJ_processo), IIf(Txt_psimples_processo = "", 0, Txt_psimples_processo), IIf(Txt_pcomissao_processo = "", 0, Txt_pcomissao_processo), IIf(Txt_pdespesas_comerciais_processo = "", 0, Txt_pdespesas_comerciais_processo), IIf(Txt_pdespesas_financeiras_processo = "", 0, Txt_pdespesas_financeiras_processo), IIf(Txt_pdespesas_administrativas_processo = "", 0, Txt_pdespesas_administrativas_processo), IIf(Txt_pfrete_processo = "", 0, Txt_pfrete_processo), IIf(Txt_pmargem_processo = "", 0, Txt_pmargem_processo)) = True Then
    ProcCalculaImpostos
Else
    Txt_pmargem_processo = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pICMS_processo_GotFocus()
On Error GoTo tratar_erro

If Txt_pICMS_processo = "0" Then Txt_pICMS_processo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pPIS_processo_GotFocus()
On Error GoTo tratar_erro

If Txt_pPIS_processo = "0" Then Txt_pPIS_processo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcofins_processo_GotFocus()
On Error GoTo tratar_erro

If Txt_pcofins_processo = "0" Then Txt_pcofins_processo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pCSLL_processo_GotFocus()
On Error GoTo tratar_erro

If Txt_pCSLL_processo = "0" Then Txt_pCSLL_processo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pISSQN_processo_GotFocus()
On Error GoTo tratar_erro

If Txt_pISSQN_processo = "0" Then Txt_pISSQN_processo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pIRPJ_processo_GotFocus()
On Error GoTo tratar_erro

If Txt_pIRPJ_processo = "0" Then Txt_pIRPJ_processo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_psimples_processo_GotFocus()
On Error GoTo tratar_erro

If Txt_psimples_processo = "0" Then Txt_psimples_processo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcomissao_processo_GotFocus()
On Error GoTo tratar_erro

If Txt_pcomissao_processo = "0" Then Txt_pcomissao_processo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_comerciais_processo_GotFocus()
On Error GoTo tratar_erro

If Txt_pdespesas_comerciais_processo = "0" Then Txt_pdespesas_comerciais_processo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_financeiras_processo_GotFocus()
On Error GoTo tratar_erro

If Txt_pdespesas_financeiras_processo = "0" Then Txt_pdespesas_financeiras_processo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_administrativas_processo_GotFocus()
On Error GoTo tratar_erro

If Txt_pdespesas_administrativas_processo = "0" Then Txt_pdespesas_administrativas_processo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pfrete_processo_GotFocus()
On Error GoTo tratar_erro

If Txt_pfrete_processo = "0" Then Txt_pfrete_processo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pmargem_processo_GotFocus()
On Error GoTo tratar_erro

If Txt_pmargem_processo = "0" Then Txt_pmargem_processo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

'MATERIAL
Private Sub Chk_ICMS_materiais_Click()
On Error GoTo tratar_erro

With Txt_pICMS_materiais
    If Chk_ICMS_materiais.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_PIS_materiais_Click()
On Error GoTo tratar_erro

With Txt_pPIS_materiais
    If Chk_PIS_materiais.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_cofins_materiais_Click()
On Error GoTo tratar_erro

With Txt_pcofins_materiais
    If Chk_cofins_materiais.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_CSLL_materiais_Click()
On Error GoTo tratar_erro

With Txt_pCSLL_materiais
    If Chk_CSLL_materiais.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ISSQN_materiais_Click()
On Error GoTo tratar_erro

With Txt_pISSQN_materiais
    If Chk_ISSQN_materiais.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_IRPJ_materiais_Click()
On Error GoTo tratar_erro

With Txt_pIRPJ_materiais
    If Chk_IRPJ_materiais.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_Simples_materiais_Click()
On Error GoTo tratar_erro

With Txt_psimples_materiais
    If Chk_simples_materiais.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_comissao_materiais_Click()
On Error GoTo tratar_erro

With Txt_pcomissao_materiais
    If Chk_comissao_materiais.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_despesas_comerciais_materiais_Click()
On Error GoTo tratar_erro

With Txt_pdespesas_comerciais_materiais
    If Chk_despesas_comerciais_materiais.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_despesas_financeiras_materiais_Click()
On Error GoTo tratar_erro

With Txt_pdespesas_financeiras_materiais
    If Chk_despesas_financeiras_materiais.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_despesas_administrativas_materiais_Click()
On Error GoTo tratar_erro

With Txt_pdespesas_administrativas_materiais
    If Chk_despesas_administrativas_materiais.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_frete_materiais_Click()
On Error GoTo tratar_erro

With Txt_pfrete_materiais
    If Chk_frete_materiais.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_margem_materiais_Click()
On Error GoTo tratar_erro

With Txt_pmargem_materiais
    If Chk_margem_materiais.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pICMS_materiais_Change()
On Error GoTo tratar_erro

If Txt_pICMS_materiais <> "" Then
    VerifNumero = Txt_pICMS_materiais
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pICMS_materiais = ""
        Txt_pICMS_materiais.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_materiais = "", 0, Txt_pICMS_materiais), IIf(Txt_pPIS_materiais = "", 0, Txt_pPIS_materiais), IIf(Txt_pcofins_materiais = "", 0, Txt_pcofins_materiais), IIf(Txt_pCSLL_materiais = "", 0, Txt_pCSLL_materiais), IIf(Txt_pISSQN_materiais = "", 0, Txt_pISSQN_materiais), IIf(Txt_pIRPJ_materiais = "", 0, Txt_pIRPJ_materiais), IIf(Txt_psimples_materiais = "", 0, Txt_psimples_materiais), IIf(Txt_pcomissao_materiais = "", 0, Txt_pcomissao_materiais), IIf(Txt_pdespesas_comerciais_materiais = "", 0, Txt_pdespesas_comerciais_materiais), IIf(Txt_pdespesas_financeiras_materiais = "", 0, Txt_pdespesas_financeiras_materiais), IIf(Txt_pdespesas_administrativas_materiais = "", 0, Txt_pdespesas_administrativas_materiais), IIf(Txt_pfrete_materiais = "", 0, Txt_pfrete_materiais), IIf(Txt_pmargem_materiais = "", 0, Txt_pmargem_materiais)) = True Then
    ProcCalculaImpostos
Else
    Txt_pICMS_materiais = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pPIS_materiais_Change()
On Error GoTo tratar_erro

If Txt_pPIS_materiais <> "" Then
    VerifNumero = Txt_pPIS_materiais
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pPIS_materiais = ""
        Txt_pPIS_materiais.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_materiais = "", 0, Txt_pICMS_materiais), IIf(Txt_pPIS_materiais = "", 0, Txt_pPIS_materiais), IIf(Txt_pcofins_materiais = "", 0, Txt_pcofins_materiais), IIf(Txt_pCSLL_materiais = "", 0, Txt_pCSLL_materiais), IIf(Txt_pISSQN_materiais = "", 0, Txt_pISSQN_materiais), IIf(Txt_pIRPJ_materiais = "", 0, Txt_pIRPJ_materiais), IIf(Txt_psimples_materiais = "", 0, Txt_psimples_materiais), IIf(Txt_pcomissao_materiais = "", 0, Txt_pcomissao_materiais), IIf(Txt_pdespesas_comerciais_materiais = "", 0, Txt_pdespesas_comerciais_materiais), IIf(Txt_pdespesas_financeiras_materiais = "", 0, Txt_pdespesas_financeiras_materiais), IIf(Txt_pdespesas_administrativas_materiais = "", 0, Txt_pdespesas_administrativas_materiais), IIf(Txt_pfrete_materiais = "", 0, Txt_pfrete_materiais), IIf(Txt_pmargem_materiais = "", 0, Txt_pmargem_materiais)) = True Then
    ProcCalculaImpostos
Else
    Txt_pPIS_materiais = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcofins_materiais_Change()
On Error GoTo tratar_erro

If Txt_pcofins_materiais <> "" Then
    VerifNumero = Txt_pcofins_materiais
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pcofins_materiais = ""
        Txt_pcofins_materiais.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_materiais = "", 0, Txt_pICMS_materiais), IIf(Txt_pPIS_materiais = "", 0, Txt_pPIS_materiais), IIf(Txt_pcofins_materiais = "", 0, Txt_pcofins_materiais), IIf(Txt_pCSLL_materiais = "", 0, Txt_pCSLL_materiais), IIf(Txt_pISSQN_materiais = "", 0, Txt_pISSQN_materiais), IIf(Txt_pIRPJ_materiais = "", 0, Txt_pIRPJ_materiais), IIf(Txt_psimples_materiais = "", 0, Txt_psimples_materiais), IIf(Txt_pcomissao_materiais = "", 0, Txt_pcomissao_materiais), IIf(Txt_pdespesas_comerciais_materiais = "", 0, Txt_pdespesas_comerciais_materiais), IIf(Txt_pdespesas_financeiras_materiais = "", 0, Txt_pdespesas_financeiras_materiais), IIf(Txt_pdespesas_administrativas_materiais = "", 0, Txt_pdespesas_administrativas_materiais), IIf(Txt_pfrete_materiais = "", 0, Txt_pfrete_materiais), IIf(Txt_pmargem_materiais = "", 0, Txt_pmargem_materiais)) = True Then
    ProcCalculaImpostos
Else
    Txt_pcofins_materiais = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pCSLL_materiais_Change()
On Error GoTo tratar_erro

If Txt_pCSLL_materiais <> "" Then
    VerifNumero = Txt_pCSLL_materiais
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pCSLL_materiais = ""
        Txt_pCSLL_materiais.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_materiais = "", 0, Txt_pICMS_materiais), IIf(Txt_pPIS_materiais = "", 0, Txt_pPIS_materiais), IIf(Txt_pcofins_materiais = "", 0, Txt_pcofins_materiais), IIf(Txt_pCSLL_materiais = "", 0, Txt_pCSLL_materiais), IIf(Txt_pISSQN_materiais = "", 0, Txt_pISSQN_materiais), IIf(Txt_pIRPJ_materiais = "", 0, Txt_pIRPJ_materiais), IIf(Txt_psimples_materiais = "", 0, Txt_psimples_materiais), IIf(Txt_pcomissao_materiais = "", 0, Txt_pcomissao_materiais), IIf(Txt_pdespesas_comerciais_materiais = "", 0, Txt_pdespesas_comerciais_materiais), IIf(Txt_pdespesas_financeiras_materiais = "", 0, Txt_pdespesas_financeiras_materiais), IIf(Txt_pdespesas_administrativas_materiais = "", 0, Txt_pdespesas_administrativas_materiais), IIf(Txt_pfrete_materiais = "", 0, Txt_pfrete_materiais), IIf(Txt_pmargem_materiais = "", 0, Txt_pmargem_materiais)) = True Then
    ProcCalculaImpostos
Else
    Txt_pCSLL_materiais = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pISSQN_materiais_Change()
On Error GoTo tratar_erro

If Txt_pISSQN_materiais <> "" Then
    VerifNumero = Txt_pISSQN_materiais
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pISSQN_materiais = ""
        Txt_pISSQN_materiais.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_materiais = "", 0, Txt_pICMS_materiais), IIf(Txt_pPIS_materiais = "", 0, Txt_pPIS_materiais), IIf(Txt_pcofins_materiais = "", 0, Txt_pcofins_materiais), IIf(Txt_pCSLL_materiais = "", 0, Txt_pCSLL_materiais), IIf(Txt_pISSQN_materiais = "", 0, Txt_pISSQN_materiais), IIf(Txt_pIRPJ_materiais = "", 0, Txt_pIRPJ_materiais), IIf(Txt_psimples_materiais = "", 0, Txt_psimples_materiais), IIf(Txt_pcomissao_materiais = "", 0, Txt_pcomissao_materiais), IIf(Txt_pdespesas_comerciais_materiais = "", 0, Txt_pdespesas_comerciais_materiais), IIf(Txt_pdespesas_financeiras_materiais = "", 0, Txt_pdespesas_financeiras_materiais), IIf(Txt_pdespesas_administrativas_materiais = "", 0, Txt_pdespesas_administrativas_materiais), IIf(Txt_pfrete_materiais = "", 0, Txt_pfrete_materiais), IIf(Txt_pmargem_materiais = "", 0, Txt_pmargem_materiais)) = True Then
    ProcCalculaImpostos
Else
    Txt_pISSQN_materiais = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pIRPJ_materiais_Change()
On Error GoTo tratar_erro

If Txt_pIRPJ_materiais <> "" Then
    VerifNumero = Txt_pIRPJ_materiais
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pIRPJ_materiais = ""
        Txt_pIRPJ_materiais.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_materiais = "", 0, Txt_pICMS_materiais), IIf(Txt_pPIS_materiais = "", 0, Txt_pPIS_materiais), IIf(Txt_pcofins_materiais = "", 0, Txt_pcofins_materiais), IIf(Txt_pCSLL_materiais = "", 0, Txt_pCSLL_materiais), IIf(Txt_pISSQN_materiais = "", 0, Txt_pISSQN_materiais), IIf(Txt_pIRPJ_materiais = "", 0, Txt_pIRPJ_materiais), IIf(Txt_psimples_materiais = "", 0, Txt_psimples_materiais), IIf(Txt_pcomissao_materiais = "", 0, Txt_pcomissao_materiais), IIf(Txt_pdespesas_comerciais_materiais = "", 0, Txt_pdespesas_comerciais_materiais), IIf(Txt_pdespesas_financeiras_materiais = "", 0, Txt_pdespesas_financeiras_materiais), IIf(Txt_pdespesas_administrativas_materiais = "", 0, Txt_pdespesas_administrativas_materiais), IIf(Txt_pfrete_materiais = "", 0, Txt_pfrete_materiais), IIf(Txt_pmargem_materiais = "", 0, Txt_pmargem_materiais)) = True Then
    ProcCalculaImpostos
Else
    Txt_pIRPJ_materiais = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_psimples_materiais_Change()
On Error GoTo tratar_erro

If Txt_psimples_materiais <> "" Then
    VerifNumero = Txt_psimples_materiais
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_psimples_materiais = ""
        Txt_psimples_materiais.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_materiais = "", 0, Txt_pICMS_materiais), IIf(Txt_pPIS_materiais = "", 0, Txt_pPIS_materiais), IIf(Txt_pcofins_materiais = "", 0, Txt_pcofins_materiais), IIf(Txt_pCSLL_materiais = "", 0, Txt_pCSLL_materiais), IIf(Txt_pISSQN_materiais = "", 0, Txt_pISSQN_materiais), IIf(Txt_pIRPJ_materiais = "", 0, Txt_pIRPJ_materiais), IIf(Txt_psimples_materiais = "", 0, Txt_psimples_materiais), IIf(Txt_pcomissao_materiais = "", 0, Txt_pcomissao_materiais), IIf(Txt_pdespesas_comerciais_materiais = "", 0, Txt_pdespesas_comerciais_materiais), IIf(Txt_pdespesas_financeiras_materiais = "", 0, Txt_pdespesas_financeiras_materiais), IIf(Txt_pdespesas_administrativas_materiais = "", 0, Txt_pdespesas_administrativas_materiais), IIf(Txt_pfrete_materiais = "", 0, Txt_pfrete_materiais), IIf(Txt_pmargem_materiais = "", 0, Txt_pmargem_materiais)) = True Then
    ProcCalculaImpostos
Else
    Txt_psimples_materiais = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcomissao_materiais_Change()
On Error GoTo tratar_erro

If Txt_pcomissao_materiais <> "" Then
    VerifNumero = Txt_pcomissao_materiais
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pcomissao_materiais = ""
        Txt_pcomissao_materiais.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_materiais = "", 0, Txt_pICMS_materiais), IIf(Txt_pPIS_materiais = "", 0, Txt_pPIS_materiais), IIf(Txt_pcofins_materiais = "", 0, Txt_pcofins_materiais), IIf(Txt_pCSLL_materiais = "", 0, Txt_pCSLL_materiais), IIf(Txt_pISSQN_materiais = "", 0, Txt_pISSQN_materiais), IIf(Txt_pIRPJ_materiais = "", 0, Txt_pIRPJ_materiais), IIf(Txt_psimples_materiais = "", 0, Txt_psimples_materiais), IIf(Txt_pcomissao_materiais = "", 0, Txt_pcomissao_materiais), IIf(Txt_pdespesas_comerciais_materiais = "", 0, Txt_pdespesas_comerciais_materiais), IIf(Txt_pdespesas_financeiras_materiais = "", 0, Txt_pdespesas_financeiras_materiais), IIf(Txt_pdespesas_administrativas_materiais = "", 0, Txt_pdespesas_administrativas_materiais), IIf(Txt_pfrete_materiais = "", 0, Txt_pfrete_materiais), IIf(Txt_pmargem_materiais = "", 0, Txt_pmargem_materiais)) = True Then
    ProcCalculaImpostos
Else
    Txt_pcomissao_materiais = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_comerciais_materiais_Change()
On Error GoTo tratar_erro

If Txt_pdespesas_comerciais_materiais <> "" Then
    VerifNumero = Txt_pdespesas_comerciais_materiais
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pdespesas_comerciais_materiais = ""
        Txt_pdespesas_comerciais_materiais.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_materiais = "", 0, Txt_pICMS_materiais), IIf(Txt_pPIS_materiais = "", 0, Txt_pPIS_materiais), IIf(Txt_pcofins_materiais = "", 0, Txt_pcofins_materiais), IIf(Txt_pCSLL_materiais = "", 0, Txt_pCSLL_materiais), IIf(Txt_pISSQN_materiais = "", 0, Txt_pISSQN_materiais), IIf(Txt_pIRPJ_materiais = "", 0, Txt_pIRPJ_materiais), IIf(Txt_psimples_materiais = "", 0, Txt_psimples_materiais), IIf(Txt_pcomissao_materiais = "", 0, Txt_pcomissao_materiais), IIf(Txt_pdespesas_comerciais_materiais = "", 0, Txt_pdespesas_comerciais_materiais), IIf(Txt_pdespesas_financeiras_materiais = "", 0, Txt_pdespesas_financeiras_materiais), IIf(Txt_pdespesas_administrativas_materiais = "", 0, Txt_pdespesas_administrativas_materiais), IIf(Txt_pfrete_materiais = "", 0, Txt_pfrete_materiais), IIf(Txt_pmargem_materiais = "", 0, Txt_pmargem_materiais)) = True Then
    ProcCalculaImpostos
Else
    Txt_pdespesas_comerciais_materiais = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_financeiras_materiais_Change()
On Error GoTo tratar_erro

If Txt_pdespesas_financeiras_materiais <> "" Then
    VerifNumero = Txt_pdespesas_financeiras_materiais
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pdespesas_financeiras_materiais = ""
        Txt_pdespesas_financeiras_materiais.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_materiais = "", 0, Txt_pICMS_materiais), IIf(Txt_pPIS_materiais = "", 0, Txt_pPIS_materiais), IIf(Txt_pcofins_materiais = "", 0, Txt_pcofins_materiais), IIf(Txt_pCSLL_materiais = "", 0, Txt_pCSLL_materiais), IIf(Txt_pISSQN_materiais = "", 0, Txt_pISSQN_materiais), IIf(Txt_pIRPJ_materiais = "", 0, Txt_pIRPJ_materiais), IIf(Txt_psimples_materiais = "", 0, Txt_psimples_materiais), IIf(Txt_pcomissao_materiais = "", 0, Txt_pcomissao_materiais), IIf(Txt_pdespesas_comerciais_materiais = "", 0, Txt_pdespesas_comerciais_materiais), IIf(Txt_pdespesas_financeiras_materiais = "", 0, Txt_pdespesas_financeiras_materiais), IIf(Txt_pdespesas_administrativas_materiais = "", 0, Txt_pdespesas_administrativas_materiais), IIf(Txt_pfrete_materiais = "", 0, Txt_pfrete_materiais), IIf(Txt_pmargem_materiais = "", 0, Txt_pmargem_materiais)) = True Then
    ProcCalculaImpostos
Else
    Txt_pdespesas_financeiras_materiais = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_administrativas_materiais_Change()
On Error GoTo tratar_erro

If Txt_pdespesas_administrativas_materiais <> "" Then
    VerifNumero = Txt_pdespesas_administrativas_materiais
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pdespesas_administrativas_materiais = ""
        Txt_pdespesas_administrativas_materiais.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_materiais = "", 0, Txt_pICMS_materiais), IIf(Txt_pPIS_materiais = "", 0, Txt_pPIS_materiais), IIf(Txt_pcofins_materiais = "", 0, Txt_pcofins_materiais), IIf(Txt_pCSLL_materiais = "", 0, Txt_pCSLL_materiais), IIf(Txt_pISSQN_materiais = "", 0, Txt_pISSQN_materiais), IIf(Txt_pIRPJ_materiais = "", 0, Txt_pIRPJ_materiais), IIf(Txt_psimples_materiais = "", 0, Txt_psimples_materiais), IIf(Txt_pcomissao_materiais = "", 0, Txt_pcomissao_materiais), IIf(Txt_pdespesas_comerciais_materiais = "", 0, Txt_pdespesas_comerciais_materiais), IIf(Txt_pdespesas_financeiras_materiais = "", 0, Txt_pdespesas_financeiras_materiais), IIf(Txt_pdespesas_administrativas_materiais = "", 0, Txt_pdespesas_administrativas_materiais), IIf(Txt_pfrete_materiais = "", 0, Txt_pfrete_materiais), IIf(Txt_pmargem_materiais = "", 0, Txt_pmargem_materiais)) = True Then
    ProcCalculaImpostos
Else
    Txt_pdespesas_administrativas_materiais = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pfrete_materiais_Change()
On Error GoTo tratar_erro

If Txt_pfrete_materiais <> "" Then
    VerifNumero = Txt_pfrete_materiais
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pfrete_materiais = ""
        Txt_pfrete_materiais.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_materiais = "", 0, Txt_pICMS_materiais), IIf(Txt_pPIS_materiais = "", 0, Txt_pPIS_materiais), IIf(Txt_pcofins_materiais = "", 0, Txt_pcofins_materiais), IIf(Txt_pCSLL_materiais = "", 0, Txt_pCSLL_materiais), IIf(Txt_pISSQN_materiais = "", 0, Txt_pISSQN_materiais), IIf(Txt_pIRPJ_materiais = "", 0, Txt_pIRPJ_materiais), IIf(Txt_psimples_materiais = "", 0, Txt_psimples_materiais), IIf(Txt_pcomissao_materiais = "", 0, Txt_pcomissao_materiais), IIf(Txt_pdespesas_comerciais_materiais = "", 0, Txt_pdespesas_comerciais_materiais), IIf(Txt_pdespesas_financeiras_materiais = "", 0, Txt_pdespesas_financeiras_materiais), IIf(Txt_pdespesas_administrativas_materiais = "", 0, Txt_pdespesas_administrativas_materiais), IIf(Txt_pfrete_materiais = "", 0, Txt_pfrete_materiais), IIf(Txt_pmargem_materiais = "", 0, Txt_pmargem_materiais)) = True Then
    ProcCalculaImpostos
Else
    Txt_pfrete_materiais = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pmargem_materiais_Change()
On Error GoTo tratar_erro

If Txt_pmargem_materiais <> "" Then
    VerifNumero = Txt_pmargem_materiais
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pmargem_materiais = ""
        Txt_pmargem_materiais.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_materiais = "", 0, Txt_pICMS_materiais), IIf(Txt_pPIS_materiais = "", 0, Txt_pPIS_materiais), IIf(Txt_pcofins_materiais = "", 0, Txt_pcofins_materiais), IIf(Txt_pCSLL_materiais = "", 0, Txt_pCSLL_materiais), IIf(Txt_pISSQN_materiais = "", 0, Txt_pISSQN_materiais), IIf(Txt_pIRPJ_materiais = "", 0, Txt_pIRPJ_materiais), IIf(Txt_psimples_materiais = "", 0, Txt_psimples_materiais), IIf(Txt_pcomissao_materiais = "", 0, Txt_pcomissao_materiais), IIf(Txt_pdespesas_comerciais_materiais = "", 0, Txt_pdespesas_comerciais_materiais), IIf(Txt_pdespesas_financeiras_materiais = "", 0, Txt_pdespesas_financeiras_materiais), IIf(Txt_pdespesas_administrativas_materiais = "", 0, Txt_pdespesas_administrativas_materiais), IIf(Txt_pfrete_materiais = "", 0, Txt_pfrete_materiais), IIf(Txt_pmargem_materiais = "", 0, Txt_pmargem_materiais)) = True Then
    ProcCalculaImpostos
Else
    Txt_pmargem_materiais = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pICMS_materiais_GotFocus()
On Error GoTo tratar_erro

If Txt_pICMS_materiais = "0" Then Txt_pICMS_materiais = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pPIS_materiais_GotFocus()
On Error GoTo tratar_erro

If Txt_pPIS_materiais = "0" Then Txt_pPIS_materiais = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcofins_materiais_GotFocus()
On Error GoTo tratar_erro

If Txt_pcofins_materiais = "0" Then Txt_pcofins_materiais = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pCSLL_materiais_GotFocus()
On Error GoTo tratar_erro

If Txt_pCSLL_materiais = "0" Then Txt_pCSLL_materiais = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pISSQN_materiais_GotFocus()
On Error GoTo tratar_erro

If Txt_pISSQN_materiais = "0" Then Txt_pISSQN_materiais = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pIRPJ_materiais_GotFocus()
On Error GoTo tratar_erro

If Txt_pIRPJ_materiais = "0" Then Txt_pIRPJ_materiais = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_psimples_materiais_GotFocus()
On Error GoTo tratar_erro

If Txt_psimples_materiais = "0" Then Txt_psimples_materiais = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcomissao_materiais_GotFocus()
On Error GoTo tratar_erro

If Txt_pcomissao_materiais = "0" Then Txt_pcomissao_materiais = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_comerciais_materiais_GotFocus()
On Error GoTo tratar_erro

If Txt_pdespesas_comerciais_materiais = "0" Then Txt_pdespesas_comerciais_materiais = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_financeiras_materiais_GotFocus()
On Error GoTo tratar_erro

If Txt_pdespesas_financeiras_materiais = "0" Then Txt_pdespesas_financeiras_materiais = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_administrativas_materiais_GotFocus()
On Error GoTo tratar_erro

If Txt_pdespesas_administrativas_materiais = "0" Then Txt_pdespesas_administrativas_materiais = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pfrete_materiais_GotFocus()
On Error GoTo tratar_erro

If Txt_pfrete_materiais = "0" Then Txt_pfrete_materiais = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pmargem_materiais_GotFocus()
On Error GoTo tratar_erro

If Txt_pmargem_materiais = "0" Then Txt_pmargem_materiais = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

'TERCEIROS
Private Sub Chk_ICMS_terceiros_Click()
On Error GoTo tratar_erro

With Txt_pICMS_terceiros
    If Chk_ICMS_terceiros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_PIS_terceiros_Click()
On Error GoTo tratar_erro

With Txt_pPIS_terceiros
    If Chk_PIS_terceiros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_cofins_terceiros_Click()
On Error GoTo tratar_erro

With Txt_pcofins_terceiros
    If Chk_cofins_terceiros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_CSLL_terceiros_Click()
On Error GoTo tratar_erro

With Txt_pCSLL_terceiros
    If Chk_CSLL_terceiros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ISSQN_terceiros_Click()
On Error GoTo tratar_erro

With Txt_pISSQN_terceiros
    If Chk_ISSQN_terceiros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_IRPJ_terceiros_Click()
On Error GoTo tratar_erro

With Txt_pIRPJ_terceiros
    If Chk_IRPJ_terceiros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_Simples_terceiros_Click()
On Error GoTo tratar_erro

With Txt_psimples_terceiros
    If Chk_simples_terceiros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_comissao_terceiros_Click()
On Error GoTo tratar_erro

With Txt_pcomissao_terceiros
    If Chk_comissao_terceiros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_despesas_comerciais_terceiros_Click()
On Error GoTo tratar_erro

With Txt_pdespesas_comerciais_terceiros
    If Chk_despesas_comerciais_terceiros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_despesas_financeiras_terceiros_Click()
On Error GoTo tratar_erro

With Txt_pdespesas_financeiras_terceiros
    If Chk_despesas_financeiras_terceiros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_despesas_administrativas_terceiros_Click()
On Error GoTo tratar_erro

With Txt_pdespesas_administrativas_terceiros
    If Chk_despesas_administrativas_terceiros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_frete_terceiros_Click()
On Error GoTo tratar_erro

With Txt_pfrete_terceiros
    If Chk_frete_terceiros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_margem_terceiros_Click()
On Error GoTo tratar_erro

With Txt_pmargem_terceiros
    If Chk_margem_terceiros.Value = 1 Then
        .Enabled = True
        If .Visible = True Then .SetFocus
    Else
        .Enabled = False
        .Text = 0
        ProcCalculaImpostos
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pICMS_terceiros_Change()
On Error GoTo tratar_erro

If Txt_pICMS_terceiros <> "" Then
    VerifNumero = Txt_pICMS_terceiros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pICMS_terceiros = ""
        Txt_pICMS_terceiros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_terceiros = "", 0, Txt_pICMS_terceiros), IIf(Txt_pPIS_terceiros = "", 0, Txt_pPIS_terceiros), IIf(Txt_pcofins_terceiros = "", 0, Txt_pcofins_terceiros), IIf(Txt_pCSLL_terceiros = "", 0, Txt_pCSLL_terceiros), IIf(Txt_pISSQN_terceiros = "", 0, Txt_pISSQN_terceiros), IIf(Txt_pIRPJ_terceiros = "", 0, Txt_pIRPJ_terceiros), IIf(Txt_psimples_terceiros = "", 0, Txt_psimples_terceiros), IIf(Txt_pcomissao_terceiros = "", 0, Txt_pcomissao_terceiros), IIf(Txt_pdespesas_comerciais_terceiros = "", 0, Txt_pdespesas_comerciais_terceiros), IIf(Txt_pdespesas_financeiras_terceiros = "", 0, Txt_pdespesas_financeiras_terceiros), IIf(Txt_pdespesas_administrativas_terceiros = "", 0, Txt_pdespesas_administrativas_terceiros), IIf(Txt_pfrete_terceiros = "", 0, Txt_pfrete_terceiros), IIf(Txt_pmargem_terceiros = "", 0, Txt_pmargem_terceiros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pICMS_terceiros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pPIS_terceiros_Change()
On Error GoTo tratar_erro

If Txt_pPIS_terceiros <> "" Then
    VerifNumero = Txt_pPIS_terceiros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pPIS_terceiros = ""
        Txt_pPIS_terceiros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_terceiros = "", 0, Txt_pICMS_terceiros), IIf(Txt_pPIS_terceiros = "", 0, Txt_pPIS_terceiros), IIf(Txt_pcofins_terceiros = "", 0, Txt_pcofins_terceiros), IIf(Txt_pCSLL_terceiros = "", 0, Txt_pCSLL_terceiros), IIf(Txt_pISSQN_terceiros = "", 0, Txt_pISSQN_terceiros), IIf(Txt_pIRPJ_terceiros = "", 0, Txt_pIRPJ_terceiros), IIf(Txt_psimples_terceiros = "", 0, Txt_psimples_terceiros), IIf(Txt_pcomissao_terceiros = "", 0, Txt_pcomissao_terceiros), IIf(Txt_pdespesas_comerciais_terceiros = "", 0, Txt_pdespesas_comerciais_terceiros), IIf(Txt_pdespesas_financeiras_terceiros = "", 0, Txt_pdespesas_financeiras_terceiros), IIf(Txt_pdespesas_administrativas_terceiros = "", 0, Txt_pdespesas_administrativas_terceiros), IIf(Txt_pfrete_terceiros = "", 0, Txt_pfrete_terceiros), IIf(Txt_pmargem_terceiros = "", 0, Txt_pmargem_terceiros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pPIS_terceiros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcofins_terceiros_Change()
On Error GoTo tratar_erro

If Txt_pcofins_terceiros <> "" Then
    VerifNumero = Txt_pcofins_terceiros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pcofins_terceiros = ""
        Txt_pcofins_terceiros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_terceiros = "", 0, Txt_pICMS_terceiros), IIf(Txt_pPIS_terceiros = "", 0, Txt_pPIS_terceiros), IIf(Txt_pcofins_terceiros = "", 0, Txt_pcofins_terceiros), IIf(Txt_pCSLL_terceiros = "", 0, Txt_pCSLL_terceiros), IIf(Txt_pISSQN_terceiros = "", 0, Txt_pISSQN_terceiros), IIf(Txt_pIRPJ_terceiros = "", 0, Txt_pIRPJ_terceiros), IIf(Txt_psimples_terceiros = "", 0, Txt_psimples_terceiros), IIf(Txt_pcomissao_terceiros = "", 0, Txt_pcomissao_terceiros), IIf(Txt_pdespesas_comerciais_terceiros = "", 0, Txt_pdespesas_comerciais_terceiros), IIf(Txt_pdespesas_financeiras_terceiros = "", 0, Txt_pdespesas_financeiras_terceiros), IIf(Txt_pdespesas_administrativas_terceiros = "", 0, Txt_pdespesas_administrativas_terceiros), IIf(Txt_pfrete_terceiros = "", 0, Txt_pfrete_terceiros), IIf(Txt_pmargem_terceiros = "", 0, Txt_pmargem_terceiros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pcofins_terceiros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pCSLL_terceiros_Change()
On Error GoTo tratar_erro

If Txt_pCSLL_terceiros <> "" Then
    VerifNumero = Txt_pCSLL_terceiros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pCSLL_terceiros = ""
        Txt_pCSLL_terceiros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_terceiros = "", 0, Txt_pICMS_terceiros), IIf(Txt_pPIS_terceiros = "", 0, Txt_pPIS_terceiros), IIf(Txt_pcofins_terceiros = "", 0, Txt_pcofins_terceiros), IIf(Txt_pCSLL_terceiros = "", 0, Txt_pCSLL_terceiros), IIf(Txt_pISSQN_terceiros = "", 0, Txt_pISSQN_terceiros), IIf(Txt_pIRPJ_terceiros = "", 0, Txt_pIRPJ_terceiros), IIf(Txt_psimples_terceiros = "", 0, Txt_psimples_terceiros), IIf(Txt_pcomissao_terceiros = "", 0, Txt_pcomissao_terceiros), IIf(Txt_pdespesas_comerciais_terceiros = "", 0, Txt_pdespesas_comerciais_terceiros), IIf(Txt_pdespesas_financeiras_terceiros = "", 0, Txt_pdespesas_financeiras_terceiros), IIf(Txt_pdespesas_administrativas_terceiros = "", 0, Txt_pdespesas_administrativas_terceiros), IIf(Txt_pfrete_terceiros = "", 0, Txt_pfrete_terceiros), IIf(Txt_pmargem_terceiros = "", 0, Txt_pmargem_terceiros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pCSLL_terceiros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pISSQN_terceiros_Change()
On Error GoTo tratar_erro

If Txt_pISSQN_terceiros <> "" Then
    VerifNumero = Txt_pISSQN_terceiros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pISSQN_terceiros = ""
        Txt_pISSQN_terceiros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_terceiros = "", 0, Txt_pICMS_terceiros), IIf(Txt_pPIS_terceiros = "", 0, Txt_pPIS_terceiros), IIf(Txt_pcofins_terceiros = "", 0, Txt_pcofins_terceiros), IIf(Txt_pCSLL_terceiros = "", 0, Txt_pCSLL_terceiros), IIf(Txt_pISSQN_terceiros = "", 0, Txt_pISSQN_terceiros), IIf(Txt_pIRPJ_terceiros = "", 0, Txt_pIRPJ_terceiros), IIf(Txt_psimples_terceiros = "", 0, Txt_psimples_terceiros), IIf(Txt_pcomissao_terceiros = "", 0, Txt_pcomissao_terceiros), IIf(Txt_pdespesas_comerciais_terceiros = "", 0, Txt_pdespesas_comerciais_terceiros), IIf(Txt_pdespesas_financeiras_terceiros = "", 0, Txt_pdespesas_financeiras_terceiros), IIf(Txt_pdespesas_administrativas_terceiros = "", 0, Txt_pdespesas_administrativas_terceiros), IIf(Txt_pfrete_terceiros = "", 0, Txt_pfrete_terceiros), IIf(Txt_pmargem_terceiros = "", 0, Txt_pmargem_terceiros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pISSQN_terceiros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pIRPJ_terceiros_Change()
On Error GoTo tratar_erro

If Txt_pIRPJ_terceiros <> "" Then
    VerifNumero = Txt_pIRPJ_terceiros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pIRPJ_terceiros = ""
        Txt_pIRPJ_terceiros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_terceiros = "", 0, Txt_pICMS_terceiros), IIf(Txt_pPIS_terceiros = "", 0, Txt_pPIS_terceiros), IIf(Txt_pcofins_terceiros = "", 0, Txt_pcofins_terceiros), IIf(Txt_pCSLL_terceiros = "", 0, Txt_pCSLL_terceiros), IIf(Txt_pISSQN_terceiros = "", 0, Txt_pISSQN_terceiros), IIf(Txt_pIRPJ_terceiros = "", 0, Txt_pIRPJ_terceiros), IIf(Txt_psimples_terceiros = "", 0, Txt_psimples_terceiros), IIf(Txt_pcomissao_terceiros = "", 0, Txt_pcomissao_terceiros), IIf(Txt_pdespesas_comerciais_terceiros = "", 0, Txt_pdespesas_comerciais_terceiros), IIf(Txt_pdespesas_financeiras_terceiros = "", 0, Txt_pdespesas_financeiras_terceiros), IIf(Txt_pdespesas_administrativas_terceiros = "", 0, Txt_pdespesas_administrativas_terceiros), IIf(Txt_pfrete_terceiros = "", 0, Txt_pfrete_terceiros), IIf(Txt_pmargem_terceiros = "", 0, Txt_pmargem_terceiros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pIRPJ_terceiros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_psimples_terceiros_Change()
On Error GoTo tratar_erro

If Txt_psimples_terceiros <> "" Then
    VerifNumero = Txt_psimples_terceiros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_psimples_terceiros = ""
        Txt_psimples_terceiros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_terceiros = "", 0, Txt_pICMS_terceiros), IIf(Txt_pPIS_terceiros = "", 0, Txt_pPIS_terceiros), IIf(Txt_pcofins_terceiros = "", 0, Txt_pcofins_terceiros), IIf(Txt_pCSLL_terceiros = "", 0, Txt_pCSLL_terceiros), IIf(Txt_pISSQN_terceiros = "", 0, Txt_pISSQN_terceiros), IIf(Txt_pIRPJ_terceiros = "", 0, Txt_pIRPJ_terceiros), IIf(Txt_psimples_terceiros = "", 0, Txt_psimples_terceiros), IIf(Txt_pcomissao_terceiros = "", 0, Txt_pcomissao_terceiros), IIf(Txt_pdespesas_comerciais_terceiros = "", 0, Txt_pdespesas_comerciais_terceiros), IIf(Txt_pdespesas_financeiras_terceiros = "", 0, Txt_pdespesas_financeiras_terceiros), IIf(Txt_pdespesas_administrativas_terceiros = "", 0, Txt_pdespesas_administrativas_terceiros), IIf(Txt_pfrete_terceiros = "", 0, Txt_pfrete_terceiros), IIf(Txt_pmargem_terceiros = "", 0, Txt_pmargem_terceiros)) = True Then
    ProcCalculaImpostos
Else
    Txt_psimples_terceiros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcomissao_terceiros_Change()
On Error GoTo tratar_erro

If Txt_pcomissao_terceiros <> "" Then
    VerifNumero = Txt_pcomissao_terceiros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pcomissao_terceiros = ""
        Txt_pcomissao_terceiros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_terceiros = "", 0, Txt_pICMS_terceiros), IIf(Txt_pPIS_terceiros = "", 0, Txt_pPIS_terceiros), IIf(Txt_pcofins_terceiros = "", 0, Txt_pcofins_terceiros), IIf(Txt_pCSLL_terceiros = "", 0, Txt_pCSLL_terceiros), IIf(Txt_pISSQN_terceiros = "", 0, Txt_pISSQN_terceiros), IIf(Txt_pIRPJ_terceiros = "", 0, Txt_pIRPJ_terceiros), IIf(Txt_psimples_terceiros = "", 0, Txt_psimples_terceiros), IIf(Txt_pcomissao_terceiros = "", 0, Txt_pcomissao_terceiros), IIf(Txt_pdespesas_comerciais_terceiros = "", 0, Txt_pdespesas_comerciais_terceiros), IIf(Txt_pdespesas_financeiras_terceiros = "", 0, Txt_pdespesas_financeiras_terceiros), IIf(Txt_pdespesas_administrativas_terceiros = "", 0, Txt_pdespesas_administrativas_terceiros), IIf(Txt_pfrete_terceiros = "", 0, Txt_pfrete_terceiros), IIf(Txt_pmargem_terceiros = "", 0, Txt_pmargem_terceiros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pcomissao_terceiros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_comerciais_terceiros_Change()
On Error GoTo tratar_erro

If Txt_pdespesas_comerciais_terceiros <> "" Then
    VerifNumero = Txt_pdespesas_comerciais_terceiros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pdespesas_comerciais_terceiros = ""
        Txt_pdespesas_comerciais_terceiros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_terceiros = "", 0, Txt_pICMS_terceiros), IIf(Txt_pPIS_terceiros = "", 0, Txt_pPIS_terceiros), IIf(Txt_pcofins_terceiros = "", 0, Txt_pcofins_terceiros), IIf(Txt_pCSLL_terceiros = "", 0, Txt_pCSLL_terceiros), IIf(Txt_pISSQN_terceiros = "", 0, Txt_pISSQN_terceiros), IIf(Txt_pIRPJ_terceiros = "", 0, Txt_pIRPJ_terceiros), IIf(Txt_psimples_terceiros = "", 0, Txt_psimples_terceiros), IIf(Txt_pcomissao_terceiros = "", 0, Txt_pcomissao_terceiros), IIf(Txt_pdespesas_comerciais_terceiros = "", 0, Txt_pdespesas_comerciais_terceiros), IIf(Txt_pdespesas_financeiras_terceiros = "", 0, Txt_pdespesas_financeiras_terceiros), IIf(Txt_pdespesas_administrativas_terceiros = "", 0, Txt_pdespesas_administrativas_terceiros), IIf(Txt_pfrete_terceiros = "", 0, Txt_pfrete_terceiros), IIf(Txt_pmargem_terceiros = "", 0, Txt_pmargem_terceiros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pdespesas_comerciais_terceiros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_financeiras_terceiros_Change()
On Error GoTo tratar_erro

If Txt_pdespesas_financeiras_terceiros <> "" Then
    VerifNumero = Txt_pdespesas_financeiras_terceiros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pdespesas_financeiras_terceiros = ""
        Txt_pdespesas_financeiras_terceiros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_terceiros = "", 0, Txt_pICMS_terceiros), IIf(Txt_pPIS_terceiros = "", 0, Txt_pPIS_terceiros), IIf(Txt_pcofins_terceiros = "", 0, Txt_pcofins_terceiros), IIf(Txt_pCSLL_terceiros = "", 0, Txt_pCSLL_terceiros), IIf(Txt_pISSQN_terceiros = "", 0, Txt_pISSQN_terceiros), IIf(Txt_pIRPJ_terceiros = "", 0, Txt_pIRPJ_terceiros), IIf(Txt_psimples_terceiros = "", 0, Txt_psimples_terceiros), IIf(Txt_pcomissao_terceiros = "", 0, Txt_pcomissao_terceiros), IIf(Txt_pdespesas_comerciais_terceiros = "", 0, Txt_pdespesas_comerciais_terceiros), IIf(Txt_pdespesas_financeiras_terceiros = "", 0, Txt_pdespesas_financeiras_terceiros), IIf(Txt_pdespesas_administrativas_terceiros = "", 0, Txt_pdespesas_administrativas_terceiros), IIf(Txt_pfrete_terceiros = "", 0, Txt_pfrete_terceiros), IIf(Txt_pmargem_terceiros = "", 0, Txt_pmargem_terceiros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pdespesas_financeiras_terceiros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_administrativas_terceiros_Change()
On Error GoTo tratar_erro

If Txt_pdespesas_administrativas_terceiros <> "" Then
    VerifNumero = Txt_pdespesas_administrativas_terceiros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pdespesas_administrativas_terceiros = ""
        Txt_pdespesas_administrativas_terceiros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_terceiros = "", 0, Txt_pICMS_terceiros), IIf(Txt_pPIS_terceiros = "", 0, Txt_pPIS_terceiros), IIf(Txt_pcofins_terceiros = "", 0, Txt_pcofins_terceiros), IIf(Txt_pCSLL_terceiros = "", 0, Txt_pCSLL_terceiros), IIf(Txt_pISSQN_terceiros = "", 0, Txt_pISSQN_terceiros), IIf(Txt_pIRPJ_terceiros = "", 0, Txt_pIRPJ_terceiros), IIf(Txt_psimples_terceiros = "", 0, Txt_psimples_terceiros), IIf(Txt_pcomissao_terceiros = "", 0, Txt_pcomissao_terceiros), IIf(Txt_pdespesas_comerciais_terceiros = "", 0, Txt_pdespesas_comerciais_terceiros), IIf(Txt_pdespesas_financeiras_terceiros = "", 0, Txt_pdespesas_financeiras_terceiros), IIf(Txt_pdespesas_administrativas_terceiros = "", 0, Txt_pdespesas_administrativas_terceiros), IIf(Txt_pfrete_terceiros = "", 0, Txt_pfrete_terceiros), IIf(Txt_pmargem_terceiros = "", 0, Txt_pmargem_terceiros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pdespesas_administrativas_terceiros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pfrete_terceiros_Change()
On Error GoTo tratar_erro

If Txt_pfrete_terceiros <> "" Then
    VerifNumero = Txt_pfrete_terceiros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pfrete_terceiros = ""
        Txt_pfrete_terceiros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_terceiros = "", 0, Txt_pICMS_terceiros), IIf(Txt_pPIS_terceiros = "", 0, Txt_pPIS_terceiros), IIf(Txt_pcofins_terceiros = "", 0, Txt_pcofins_terceiros), IIf(Txt_pCSLL_terceiros = "", 0, Txt_pCSLL_terceiros), IIf(Txt_pISSQN_terceiros = "", 0, Txt_pISSQN_terceiros), IIf(Txt_pIRPJ_terceiros = "", 0, Txt_pIRPJ_terceiros), IIf(Txt_psimples_terceiros = "", 0, Txt_psimples_terceiros), IIf(Txt_pcomissao_terceiros = "", 0, Txt_pcomissao_terceiros), IIf(Txt_pdespesas_comerciais_terceiros = "", 0, Txt_pdespesas_comerciais_terceiros), IIf(Txt_pdespesas_financeiras_terceiros = "", 0, Txt_pdespesas_financeiras_terceiros), IIf(Txt_pdespesas_administrativas_terceiros = "", 0, Txt_pdespesas_administrativas_terceiros), IIf(Txt_pfrete_terceiros = "", 0, Txt_pfrete_terceiros), IIf(Txt_pmargem_terceiros = "", 0, Txt_pmargem_terceiros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pfrete_terceiros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pmargem_terceiros_Change()
On Error GoTo tratar_erro

If Txt_pmargem_terceiros <> "" Then
    VerifNumero = Txt_pmargem_terceiros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_pmargem_terceiros = ""
        Txt_pmargem_terceiros.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_terceiros = "", 0, Txt_pICMS_terceiros), IIf(Txt_pPIS_terceiros = "", 0, Txt_pPIS_terceiros), IIf(Txt_pcofins_terceiros = "", 0, Txt_pcofins_terceiros), IIf(Txt_pCSLL_terceiros = "", 0, Txt_pCSLL_terceiros), IIf(Txt_pISSQN_terceiros = "", 0, Txt_pISSQN_terceiros), IIf(Txt_pIRPJ_terceiros = "", 0, Txt_pIRPJ_terceiros), IIf(Txt_psimples_terceiros = "", 0, Txt_psimples_terceiros), IIf(Txt_pcomissao_terceiros = "", 0, Txt_pcomissao_terceiros), IIf(Txt_pdespesas_comerciais_terceiros = "", 0, Txt_pdespesas_comerciais_terceiros), IIf(Txt_pdespesas_financeiras_terceiros = "", 0, Txt_pdespesas_financeiras_terceiros), IIf(Txt_pdespesas_administrativas_terceiros = "", 0, Txt_pdespesas_administrativas_terceiros), IIf(Txt_pfrete_terceiros = "", 0, Txt_pfrete_terceiros), IIf(Txt_pmargem_terceiros = "", 0, Txt_pmargem_terceiros)) = True Then
    ProcCalculaImpostos
Else
    Txt_pmargem_terceiros = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pICMS_terceiros_GotFocus()
On Error GoTo tratar_erro

If Txt_pICMS_terceiros = "0" Then Txt_pICMS_terceiros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pPIS_terceiros_GotFocus()
On Error GoTo tratar_erro

If Txt_pPIS_terceiros = "0" Then Txt_pPIS_terceiros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcofins_terceiros_GotFocus()
On Error GoTo tratar_erro

If Txt_pcofins_terceiros = "0" Then Txt_pcofins_terceiros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pCSLL_terceiros_GotFocus()
On Error GoTo tratar_erro

If Txt_pCSLL_terceiros = "0" Then Txt_pCSLL_terceiros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pISSQN_terceiros_GotFocus()
On Error GoTo tratar_erro

If Txt_pISSQN_terceiros = "0" Then Txt_pISSQN_terceiros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pIRPJ_terceiros_GotFocus()
On Error GoTo tratar_erro

If Txt_pIRPJ_terceiros = "0" Then Txt_pIRPJ_terceiros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_psimples_terceiros_GotFocus()
On Error GoTo tratar_erro

If Txt_psimples_terceiros = "0" Then Txt_psimples_terceiros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pcomissao_terceiros_GotFocus()
On Error GoTo tratar_erro

If Txt_pcomissao_terceiros = "0" Then Txt_pcomissao_terceiros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_comerciais_terceiros_GotFocus()
On Error GoTo tratar_erro

If Txt_pdespesas_comerciais_terceiros = "0" Then Txt_pdespesas_comerciais_terceiros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_financeiras_terceiros_GotFocus()
On Error GoTo tratar_erro

If Txt_pdespesas_financeiras_terceiros = "0" Then Txt_pdespesas_financeiras_terceiros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pdespesas_administrativas_terceiros_GotFocus()
On Error GoTo tratar_erro

If Txt_pdespesas_administrativas_terceiros = "0" Then Txt_pdespesas_administrativas_terceiros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pfrete_terceiros_GotFocus()
On Error GoTo tratar_erro

If Txt_pfrete_terceiros = "0" Then Txt_pfrete_terceiros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_pmargem_terceiros_GotFocus()
On Error GoTo tratar_erro

If Txt_pmargem_terceiros = "0" Then Txt_pmargem_terceiros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaImpostos()
On Error GoTo tratar_erro

If chkTotal.Value = 0 Then
    'Processo
    valor = IIf(Txt_total_processo = "", 0, Txt_total_processo)
    
    ICMS = IIf(Txt_pICMS_processo = "", 0, Txt_pICMS_processo)
    PIS_Prod = IIf(Txt_pPIS_processo = "", 0, Txt_pPIS_processo)
    Cofins_Prod = IIf(Txt_pcofins_processo = "", 0, Txt_pcofins_processo)
    CSLL_Prod = IIf(Txt_pCSLL_processo = "", 0, Txt_pCSLL_processo)
    ICMSOUTROS = IIf(Txt_pISSQN_processo = "", 0, Txt_pISSQN_processo)
    IRPJ_Prod = IIf(Txt_pIRPJ_processo = "", 0, Txt_pIRPJ_processo)
    Qtde = IIf(Txt_psimples_processo = "", 0, Txt_psimples_processo)
    Quant = IIf(Txt_pcomissao_processo = "", 0, Txt_pcomissao_processo)
    qtdeliberada = IIf(Txt_pdespesas_comerciais_processo = "", 0, Txt_pdespesas_comerciais_processo)
    QuantComprado = IIf(Txt_pdespesas_financeiras_processo = "", 0, Txt_pdespesas_financeiras_processo)
    qtdeliberar = IIf(Txt_pdespesas_administrativas_processo = "", 0, Txt_pdespesas_administrativas_processo)
    QuantEmpenho = IIf(Txt_pfrete_processo = "", 0, Txt_pfrete_processo)
    quantestoque = IIf(Txt_pmargem_processo = "", 0, Txt_pmargem_processo)
    
    'Soma total porcentagem
    If Margem_Reciproca = True Then
        Txt_ptotal_processo = ICMS + PIS_Prod + Cofins_Prod + CSLL_Prod + ICMSOUTROS + IRPJ_Prod + Qtde + Quant + qtdeliberada + QuantComprado + qtdeliberar + QuantEmpenho + quantestoque
    Else
        Txt_ptotal_processo = ICMS + PIS_Prod + Cofins_Prod + CSLL_Prod + ICMSOUTROS + IRPJ_Prod + Qtde + Quant + qtdeliberada + QuantComprado + qtdeliberar + QuantEmpenho
        
        'Calcula valor da margem
        Valorparcela = Format((valor * quantestoque) / 100, "###,##0.00")
        Txt_valor_margem_processo = Format(Valorparcela, "###,##0.00")
        valor = valor + Valorparcela
    End If
    
    'Calcula porcentagem reciploca
    quantidade = Txt_ptotal_processo
    Txt_preciploca_processo = 100 - quantidade

    'Calcula valor da venda
    quantnovo = IIf(Txt_preciploca_processo = 0, 1, Txt_preciploca_processo)
    Txt_valor_venda_processo = Format((valor / quantnovo) * 100, "###,##0.00")
    
    'Calcula valor total
    quantnovo = Txt_ptotal_processo
    ValorICMS = Txt_valor_venda_processo
    Txt_valor_total_processo = Format((ValorICMS * quantnovo) / 100, "###,##0.00")
    
    'Calcula valor reciploca
    Txt_valor_reciploca_processo = IIf(Txt_total_processo = "", "0,00", Txt_total_processo)
    
    'Valor da venda
    valor = Txt_valor_venda_processo
    
    'Calcula valor dos impostos
    Txt_valor_ICMS_processo = Format((valor * ICMS) / 100, "###,##0.00")
    Txt_valor_PIS_processo = Format((valor * PIS_Prod) / 100, "###,##0.00")
    Txt_valor_cofins_processo = Format((valor * Cofins_Prod) / 100, "###,##0.00")
    Txt_valor_CSLL_processo = Format((valor * CSLL_Prod) / 100, "###,##0.00")
    Txt_valor_ISSQN_processo = Format((valor * ICMSOUTROS) / 100, "###,##0.00")
    Txt_valor_IRPJ_processo = Format((valor * IRPJ_Prod) / 100, "###,##0.00")
    Txt_valor_simples_processo = Format((valor * Qtde) / 100, "###,##0.00")
    Txt_valor_comissao_processo = Format((valor * Quant) / 100, "###,##0.00")
    Txt_valor_despesas_comerciais_processo = Format((valor * qtdeliberada) / 100, "###,##0.00")
    Txt_valor_despesas_financeiras_processo = Format((valor * QuantComprado) / 100, "###,##0.00")
    Txt_valor_despesas_administrativas_processo = Format((valor * qtdeliberar) / 100, "###,##0.00")
    Txt_valor_frete_processo = Format((valor * QuantEmpenho) / 100, "###,##0.00")
    If Margem_Reciproca = True Then Txt_valor_margem_processo = Format((valor * quantestoque) / 100, "###,##0.00")
    '=================================================================================================================
    
    'Material
    valor = IIf(Txt_total_materiais = "", 0, Txt_total_materiais)
    
    ICMS = IIf(Txt_pICMS_materiais = "", 0, Txt_pICMS_materiais)
    PIS_Prod = IIf(Txt_pPIS_materiais = "", 0, Txt_pPIS_materiais)
    Cofins_Prod = IIf(Txt_pcofins_materiais = "", 0, Txt_pcofins_materiais)
    CSLL_Prod = IIf(Txt_pCSLL_materiais = "", 0, Txt_pCSLL_materiais)
    ICMSOUTROS = IIf(Txt_pISSQN_materiais = "", 0, Txt_pISSQN_materiais)
    IRPJ_Prod = IIf(Txt_pIRPJ_materiais = "", 0, Txt_pIRPJ_materiais)
    Qtde = IIf(Txt_psimples_materiais = "", 0, Txt_psimples_materiais)
    Quant = IIf(Txt_pcomissao_materiais = "", 0, Txt_pcomissao_materiais)
    qtdeliberada = IIf(Txt_pdespesas_comerciais_materiais = "", 0, Txt_pdespesas_comerciais_materiais)
    QuantComprado = IIf(Txt_pdespesas_financeiras_materiais = "", 0, Txt_pdespesas_financeiras_materiais)
    qtdeliberar = IIf(Txt_pdespesas_administrativas_materiais = "", 0, Txt_pdespesas_administrativas_materiais)
    QuantEmpenho = IIf(Txt_pfrete_materiais = "", 0, Txt_pfrete_materiais)
    quantestoque = IIf(Txt_pmargem_materiais = "", 0, Txt_pmargem_materiais)
    
    'Soma total porcentagem
    If Margem_Reciproca = True Then
        Txt_ptotal_materiais = ICMS + PIS_Prod + Cofins_Prod + CSLL_Prod + ICMSOUTROS + IRPJ_Prod + Qtde + Quant + qtdeliberada + QuantComprado + qtdeliberar + QuantEmpenho + quantestoque
    Else
        Txt_ptotal_materiais = ICMS + PIS_Prod + Cofins_Prod + CSLL_Prod + ICMSOUTROS + IRPJ_Prod + Qtde + Quant + qtdeliberada + QuantComprado + qtdeliberar + QuantEmpenho
        
        'Calcula valor da margem
        Valorparcela = Format((valor * quantestoque) / 100, "###,##0.00")
        Txt_valor_margem_materiais = Format(Valorparcela, "###,##0.00")
        valor = valor + Valorparcela
    End If

    'Calcula porcentagem reciploca
    quantidade = Txt_ptotal_materiais
    Txt_preciploca_materiais = 100 - quantidade
    
    'Calcula valor da venda
    quantnovo = IIf(Txt_preciploca_materiais = 0, 1, Txt_preciploca_materiais)
    Txt_valor_venda_materiais = Format((valor / quantnovo) * 100, "###,##0.00")
    
    'Calcula valor total
    quantnovo = Txt_ptotal_materiais
    ValorICMS = Txt_valor_venda_materiais
    Txt_valor_total_materiais = Format((ValorICMS * quantnovo) / 100, "###,##0.00")
    
    'Calcula valor reciploca
    Txt_valor_reciploca_materiais = IIf(Txt_total_materiais = "", "0,00", Txt_total_materiais)
    
    'Valor da venda
    valor = Txt_valor_venda_materiais
    
    'Calcula valor dos impostos
    Txt_valor_ICMS_materiais = Format((valor * ICMS) / 100, "###,##0.00")
    Txt_valor_PIS_materiais = Format((valor * PIS_Prod) / 100, "###,##0.00")
    Txt_valor_cofins_materiais = Format((valor * Cofins_Prod) / 100, "###,##0.00")
    Txt_valor_CSLL_materiais = Format((valor * CSLL_Prod) / 100, "###,##0.00")
    Txt_valor_ISSQN_materiais = Format((valor * ICMSOUTROS) / 100, "###,##0.00")
    Txt_valor_IRPJ_materiais = Format((valor * IRPJ_Prod) / 100, "###,##0.00")
    Txt_valor_simples_materiais = Format((valor * Qtde) / 100, "###,##0.00")
    Txt_valor_comissao_materiais = Format((valor * Quant) / 100, "###,##0.00")
    Txt_valor_despesas_comerciais_materiais = Format((valor * qtdeliberada) / 100, "###,##0.00")
    Txt_valor_despesas_financeiras_materiais = Format((valor * QuantComprado) / 100, "###,##0.00")
    Txt_valor_despesas_administrativas_materiais = Format((valor * qtdeliberar) / 100, "###,##0.00")
    Txt_valor_frete_materiais = Format((valor * QuantEmpenho) / 100, "###,##0.00")
    If Margem_Reciproca = True Then Txt_valor_margem_materiais = Format((valor * quantestoque) / 100, "###,##0.00")
    '=================================================================================================================
    
    'Terceiros
    valor = IIf(Txt_total_terceiros = "", 0, Txt_total_terceiros)
    
    ICMS = IIf(Txt_pICMS_terceiros = "", 0, Txt_pICMS_terceiros)
    PIS_Prod = IIf(Txt_pPIS_terceiros = "", 0, Txt_pPIS_terceiros)
    Cofins_Prod = IIf(Txt_pcofins_terceiros = "", 0, Txt_pcofins_terceiros)
    CSLL_Prod = IIf(Txt_pCSLL_terceiros = "", 0, Txt_pCSLL_terceiros)
    ICMSOUTROS = IIf(Txt_pISSQN_terceiros = "", 0, Txt_pISSQN_terceiros)
    IRPJ_Prod = IIf(Txt_pIRPJ_terceiros = "", 0, Txt_pIRPJ_terceiros)
    Qtde = IIf(Txt_psimples_terceiros = "", 0, Txt_psimples_terceiros)
    Quant = IIf(Txt_pcomissao_terceiros = "", 0, Txt_pcomissao_terceiros)
    qtdeliberada = IIf(Txt_pdespesas_comerciais_terceiros = "", 0, Txt_pdespesas_comerciais_terceiros)
    QuantComprado = IIf(Txt_pdespesas_financeiras_terceiros = "", 0, Txt_pdespesas_financeiras_terceiros)
    qtdeliberar = IIf(Txt_pdespesas_administrativas_terceiros = "", 0, Txt_pdespesas_administrativas_terceiros)
    QuantEmpenho = IIf(Txt_pfrete_terceiros = "", 0, Txt_pfrete_terceiros)
    quantestoque = IIf(Txt_pmargem_terceiros = "", 0, Txt_pmargem_terceiros)
    
    'Soma total porcentagem
    If Margem_Reciproca = True Then
        Txt_ptotal_terceiros = ICMS + PIS_Prod + Cofins_Prod + CSLL_Prod + ICMSOUTROS + IRPJ_Prod + Qtde + Quant + qtdeliberada + QuantComprado + qtdeliberar + QuantEmpenho + quantestoque
    Else
        Txt_ptotal_terceiros = ICMS + PIS_Prod + Cofins_Prod + CSLL_Prod + ICMSOUTROS + IRPJ_Prod + Qtde + Quant + qtdeliberada + QuantComprado + qtdeliberar + QuantEmpenho
        
        'Calcula valor da margem
        Valorparcela = Format((valor * quantestoque) / 100, "###,##0.00")
        Txt_valor_margem_terceiros = Format(Valorparcela, "###,##0.00")
        valor = valor + Valorparcela
    End If
    
    'Calcula porcentagem reciploca
    quantidade = Txt_ptotal_terceiros
    Txt_preciploca_terceiros = 100 - quantidade
    
    'Calcula valor da venda
    quantnovo = IIf(Txt_preciploca_terceiros = 0, 1, Txt_preciploca_terceiros)
    Txt_valor_venda_terceiros = Format((valor / quantnovo) * 100, "###,##0.00")
    
    'Calcula valor total
    quantnovo = Txt_ptotal_terceiros
    ValorICMS = Txt_valor_venda_terceiros
    Txt_valor_total_terceiros = Format((ValorICMS * quantnovo) / 100, "###,##0.00")
    
    'Calcula valor reciploca
    Txt_valor_reciploca_terceiros = IIf(Txt_total_terceiros = "", "0,00", Txt_total_terceiros)
    
    'Valor da venda
    valor = Txt_valor_venda_terceiros
    
    'Calcula valor dos impostos
    Txt_valor_ICMS_terceiros = Format((valor * ICMS) / 100, "###,##0.00")
    Txt_valor_PIS_terceiros = Format((valor * PIS_Prod) / 100, "###,##0.00")
    Txt_valor_cofins_terceiros = Format((valor * Cofins_Prod) / 100, "###,##0.00")
    Txt_valor_CSLL_terceiros = Format((valor * CSLL_Prod) / 100, "###,##0.00")
    Txt_valor_ISSQN_terceiros = Format((valor * ICMSOUTROS) / 100, "###,##0.00")
    Txt_valor_IRPJ_terceiros = Format((valor * IRPJ_Prod) / 100, "###,##0.00")
    Txt_valor_simples_terceiros = Format((valor * Qtde) / 100, "###,##0.00")
    Txt_valor_comissao_terceiros = Format((valor * Quant) / 100, "###,##0.00")
    Txt_valor_despesas_comerciais_terceiros = Format((valor * qtdeliberada) / 100, "###,##0.00")
    Txt_valor_despesas_financeiras_terceiros = Format((valor * QuantComprado) / 100, "###,##0.00")
    Txt_valor_despesas_administrativas_terceiros = Format((valor * qtdeliberar) / 100, "###,##0.00")
    Txt_valor_frete_terceiros = Format((valor * QuantEmpenho) / 100, "###,##0.00")
    If Margem_Reciproca = True Then Txt_valor_margem_terceiros = Format((valor * quantestoque) / 100, "###,##0.00")
    '=================================================================================================================
    
    'Outros
    valor = IIf(Txt_total_outros = "", 0, Txt_total_outros)
    
    ICMS = IIf(Txt_pICMS_outros = "", 0, Txt_pICMS_outros)
    PIS_Prod = IIf(Txt_pPIS_outros = "", 0, Txt_pPIS_outros)
    Cofins_Prod = IIf(Txt_pcofins_outros = "", 0, Txt_pcofins_outros)
    CSLL_Prod = IIf(Txt_pCSLL_outros = "", 0, Txt_pCSLL_outros)
    ICMSOUTROS = IIf(Txt_pISSQN_outros = "", 0, Txt_pISSQN_outros)
    IRPJ_Prod = IIf(Txt_pIRPJ_outros = "", 0, Txt_pIRPJ_outros)
    Qtde = IIf(Txt_psimples_outros = "", 0, Txt_psimples_outros)
    Quant = IIf(Txt_pcomissao_outros = "", 0, Txt_pcomissao_outros)
    qtdeliberada = IIf(Txt_pdespesas_comerciais_outros = "", 0, Txt_pdespesas_comerciais_outros)
    QuantComprado = IIf(Txt_pdespesas_financeiras_outros = "", 0, Txt_pdespesas_financeiras_outros)
    qtdeliberar = IIf(Txt_pdespesas_administrativas_outros = "", 0, Txt_pdespesas_administrativas_outros)
    QuantEmpenho = IIf(Txt_pfrete_outros = "", 0, Txt_pfrete_outros)
    quantestoque = IIf(Txt_pmargem_outros = "", 0, Txt_pmargem_outros)
    
    'Soma total porcentagem
    If Margem_Reciproca = False Then
        Txt_ptotal_outros = ICMS + PIS_Prod + Cofins_Prod + CSLL_Prod + ICMSOUTROS + IRPJ_Prod + Qtde + Quant + qtdeliberada + QuantComprado + qtdeliberar + QuantEmpenho + quantestoque
    Else
        Txt_ptotal_outros = ICMS + PIS_Prod + Cofins_Prod + CSLL_Prod + ICMSOUTROS + IRPJ_Prod + Qtde + Quant + qtdeliberada + QuantComprado + qtdeliberar + QuantEmpenho
        
        'Calcula valor da margem
        Valorparcela = Format((valor * quantestoque) / 100, "###,##0.00")
        Txt_valor_margem_outros = Format(Valorparcela, "###,##0.00")
        valor = valor + Valorparcela
    End If
    
    'Calcula porcentagem reciploca
    quantidade = Txt_ptotal_outros
    Txt_preciploca_outros = 100 - quantidade
    
    'Calcula valor da venda
    quantnovo = IIf(Txt_preciploca_outros = 0, 1, Txt_preciploca_outros)
    Txt_valor_venda_outros = Format((valor / quantnovo) * 100, "###,##0.00")
    
    'Calcula valor total
    quantnovo = Txt_ptotal_outros
    ValorICMS = Txt_valor_venda_outros
    Txt_valor_total_outros = Format((ValorICMS * quantnovo) / 100, "###,##0.00")
    
    'Calcula valor reciploca
    Txt_valor_reciploca_outros = IIf(Txt_total_outros = "", "0,00", Txt_total_outros)
    
    'Valor da venda
    valor = Txt_valor_venda_outros
    
    'Calcula valor dos impostos
    Txt_valor_ICMS_outros = Format((valor * ICMS) / 100, "###,##0.00")
    Txt_valor_PIS_outros = Format((valor * PIS_Prod) / 100, "###,##0.00")
    Txt_valor_cofins_outros = Format((valor * Cofins_Prod) / 100, "###,##0.00")
    Txt_valor_CSLL_outros = Format((valor * CSLL_Prod) / 100, "###,##0.00")
    Txt_valor_ISSQN_outros = Format((valor * ICMSOUTROS) / 100, "###,##0.00")
    Txt_valor_IRPJ_outros = Format((valor * IRPJ_Prod) / 100, "###,##0.00")
    Txt_valor_simples_outros = Format((valor * Qtde) / 100, "###,##0.00")
    Txt_valor_comissao_outros = Format((valor * Quant) / 100, "###,##0.00")
    Txt_valor_despesas_comerciais_outros = Format((valor * qtdeliberada) / 100, "###,##0.00")
    Txt_valor_despesas_financeiras_outros = Format((valor * QuantComprado) / 100, "###,##0.00")
    Txt_valor_despesas_administrativas_outros = Format((valor * qtdeliberar) / 100, "###,##0.00")
    Txt_valor_frete_outros = Format((valor * QuantEmpenho) / 100, "###,##0.00")
    If Margem_Reciproca = True Then Txt_valor_margem_outros = Format((valor * quantestoque) / 100, "###,##0.00")
Else
    '=================================================================================================================
    
    'Total
    valor = IIf(Txt_total_geral = "", 0, Txt_total_geral)
    
    ICMS = IIf(Txt_pICMS_total = "", 0, Txt_pICMS_total)
    PIS_Prod = IIf(Txt_pPIS_total = "", 0, Txt_pPIS_total)
    Cofins_Prod = IIf(Txt_pcofins_total = "", 0, Txt_pcofins_total)
    CSLL_Prod = IIf(Txt_pCSLL_total = "", 0, Txt_pCSLL_total)
    ICMSOUTROS = IIf(Txt_pISSQN_total = "", 0, Txt_pISSQN_total)
    IRPJ_Prod = IIf(Txt_pIRPJ_total = "", 0, Txt_pIRPJ_total)
    Qtde = IIf(Txt_psimples_total = "", 0, Txt_psimples_total)
    Quant = IIf(Txt_pcomissao_total = "", 0, Txt_pcomissao_total)
    qtdeliberada = IIf(Txt_pdespesas_comerciais_total = "", 0, Txt_pdespesas_comerciais_total)
    QuantComprado = IIf(Txt_pdespesas_financeiras_total = "", 0, Txt_pdespesas_financeiras_total)
    qtdeliberar = IIf(Txt_pdespesas_administrativas_total = "", 0, Txt_pdespesas_administrativas_total)
    QuantEmpenho = IIf(Txt_pfrete_total = "", 0, Txt_pfrete_total)
    quantestoque = IIf(Txt_pmargem_total = "", 0, Txt_pmargem_total)
    
    'Soma total porcentagem
    If Margem_Reciproca = True Then
        Txt_ptotal_total = ICMS + PIS_Prod + Cofins_Prod + CSLL_Prod + ICMSOUTROS + IRPJ_Prod + Qtde + Quant + qtdeliberada + QuantComprado + qtdeliberar + QuantEmpenho + quantestoque
    Else
        Txt_ptotal_total = ICMS + PIS_Prod + Cofins_Prod + CSLL_Prod + ICMSOUTROS + IRPJ_Prod + Qtde + Quant + qtdeliberada + QuantComprado + qtdeliberar + QuantEmpenho
        
        'Calcula valor da margem
        Valorparcela = Format((valor * quantestoque) / 100, "###,##0.00")
        Txt_valor_margem_total = Format(Valorparcela, "###,##0.00")
        valor = valor + Valorparcela
    End If
    
    'Calcula porcentagem reciploca
    quantidade = Txt_ptotal_total
    Txt_preciploca_total = 100 - quantidade
    
    'Calcula valor da venda
    quantnovo = IIf(Txt_preciploca_total = 0, 1, Txt_preciploca_total)
    Txt_valor_venda_total = Format(((valor / quantnovo) * 100), "###,##0.00")
    
    'Calcula valor total
    quantnovo = Txt_ptotal_total
    ValorICMS = Txt_valor_venda_total
    Txt_valor_total_total = Format((ValorICMS * quantnovo) / 100, "###,##0.00")
    
    'Calcula valor reciploca
    Txt_valor_reciploca_total = IIf(Txt_total_geral = "", "0,00", Txt_total_geral)
    
    'Valor da venda
    valor = Txt_valor_venda_total
    
    'Calcula valor dos impostos
    Txt_valor_ICMS_total = Format((valor * ICMS) / 100, "###,##0.00")
    Txt_valor_PIS_total = Format((valor * PIS_Prod) / 100, "###,##0.00")
    Txt_valor_cofins_total = Format((valor * Cofins_Prod) / 100, "###,##0.00")
    Txt_valor_CSLL_total = Format((valor * CSLL_Prod) / 100, "###,##0.00")
    Txt_valor_ISSQN_total = Format((valor * ICMSOUTROS) / 100, "###,##0.00")
    Txt_valor_IRPJ_total = Format((valor * IRPJ_Prod) / 100, "###,##0.00")
    Txt_valor_simples_total = Format((valor * Qtde) / 100, "###,##0.00")
    Txt_valor_comissao_total = Format((valor * Quant) / 100, "###,##0.00")
    Txt_valor_despesas_comerciais_total = Format((valor * qtdeliberada) / 100, "###,##0.00")
    Txt_valor_despesas_financeiras_total = Format((valor * QuantComprado) / 100, "###,##0.00")
    Txt_valor_despesas_administrativas_total = Format((valor * qtdeliberar) / 100, "###,##0.00")
    Txt_valor_frete_total = Format((valor * QuantEmpenho) / 100, "###,##0.00")
    If Margem_Reciproca = True Then Txt_valor_margem_total = Format((valor * quantestoque) / 100, "###,##0.00")
End If

'Valor total
If chkTotal.Value = 0 Then
    'Margem de lucro
    Valor1 = Txt_valor_margem_processo
    Valor2 = Txt_valor_margem_materiais
    Valor3 = Txt_valor_margem_terceiros
    Qtde = Txt_valor_margem_outros
    Txt_margem_de_lucro = Format(Valor1 + Valor2 + Valor3 + Qtde, "###,##0.00")
    
    'Venda
    Valor1 = Txt_valor_venda_processo
    Valor2 = Txt_valor_venda_materiais
    Valor3 = Txt_valor_venda_terceiros
    Valores = Txt_valor_venda_outros
    Txt_valor_total_venda = Format(Valor1 + Valor2 + Valor3 + Valores, "###,##0.0000000000")
Else
    'Margem de lucro
    Txt_margem_de_lucro = Format(Txt_valor_margem_total, "###,##0.00")
    'Venda
    Txt_valor_total_venda = Format(Txt_valor_venda_total, "###,##0.0000000000")
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_psimples_total_Change()
On Error GoTo tratar_erro

If Txt_psimples_total <> "" Then
    VerifNumero = Txt_psimples_total
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_psimples_total = ""
        Txt_psimples_total.SetFocus
        Exit Sub
    End If
End If
If FunVerificaPorcentagem(IIf(Txt_pICMS_total = "", 0, Txt_pICMS_total), IIf(Txt_pPIS_total = "", 0, Txt_pPIS_total), IIf(Txt_pcofins_total = "", 0, Txt_pcofins_total), IIf(Txt_pCSLL_total = "", 0, Txt_pCSLL_total), IIf(Txt_pISSQN_total = "", 0, Txt_pISSQN_total), IIf(Txt_pIRPJ_total = "", 0, Txt_pIRPJ_total), IIf(Txt_psimples_total = "", 0, Txt_psimples_total), IIf(Txt_pcomissao_total = "", 0, Txt_pcomissao_total), IIf(Txt_pdespesas_comerciais_total = "", 0, Txt_pdespesas_comerciais_total), IIf(Txt_pdespesas_financeiras_total = "", 0, Txt_pdespesas_financeiras_total), IIf(Txt_pdespesas_administrativas_total = "", 0, Txt_pdespesas_administrativas_total), IIf(Txt_pfrete_total = "", 0, Txt_pfrete_total), IIf(Txt_pmargem_total = "", 0, Txt_pmargem_total)) = True Then
    ProcCalculaImpostos
Else
    Txt_psimples_total = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_psimples_total_GotFocus()
On Error GoTo tratar_erro

If Txt_psimples_total = "0" Then Txt_psimples_total = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_total_outros_Change()
On Error GoTo tratar_erro

If Txt_total_outros <> "" Then
    VerifNumero = Txt_total_outros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_total_outros = ""
        Txt_total_outros.SetFocus
        Exit Sub
    End If
End If
ProcCalculaTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_total_outros_GotFocus()
On Error GoTo tratar_erro

If Txt_total_outros = "0,00" Then Txt_total_outros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_total_outros_LostFocus()
On Error GoTo tratar_erro

Txt_total_outros = IIf(Txt_total_outros = "", "0,00", Format(Txt_total_outros, "###,##0.00"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_total_processo_Change()
On Error GoTo tratar_erro

If Txt_total_processo <> "" Then
    VerifNumero = Txt_total_processo
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_total_processo = ""
        Txt_total_processo.SetFocus
        Exit Sub
    End If
End If
ProcCalculaTotal
ProcCalculaImpostos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_total_processo_GotFocus()
On Error GoTo tratar_erro

If Txt_total_processo = "0,00" Then Txt_total_processo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_total_processo_LostFocus()
On Error GoTo tratar_erro

Txt_total_processo = IIf(Txt_total_processo = "", "0,00", Format(Txt_total_processo, "###,##0.00"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_total_materiais_Change()
On Error GoTo tratar_erro

If Txt_total_materiais <> "" Then
    VerifNumero = Txt_total_materiais
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_total_materiais = ""
        Txt_total_materiais.SetFocus
        Exit Sub
    End If
End If
ProcCalculaTotal
ProcCalculaImpostos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_total_materiais_GotFocus()
On Error GoTo tratar_erro

If Txt_total_materiais = "0,00" Then Txt_total_materiais = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_total_materiais_LostFocus()
On Error GoTo tratar_erro

Txt_total_materiais = IIf(Txt_total_materiais = "", "0,00", Format(Txt_total_materiais, "###,##0.00"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_total_terceiros_Change()
On Error GoTo tratar_erro

If Txt_total_terceiros <> "" Then
    VerifNumero = Txt_total_terceiros
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_total_terceiros = ""
        Txt_total_terceiros.SetFocus
        Exit Sub
    End If
End If
ProcCalculaTotal
ProcCalculaImpostos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_total_terceiros_GotFocus()
On Error GoTo tratar_erro

If Txt_total_terceiros = "0,00" Then Txt_total_terceiros = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_total_terceiros_LostFocus()
On Error GoTo tratar_erro

Txt_total_terceiros = IIf(Txt_total_terceiros = "", "0,00", Format(Txt_total_terceiros, "###,##0.00"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaTotal()
On Error GoTo tratar_erro

valor = IIf(Txt_total_processo = "", 0, Txt_total_processo)
Valor1 = IIf(Txt_total_materiais = "", 0, Txt_total_materiais)
Valor2 = IIf(Txt_total_terceiros = "", 0, Txt_total_terceiros)
Valor3 = IIf(Txt_total_outros = "", 0, Txt_total_outros)
Txt_total_geral = Format(valor + Valor1 + Valor2 + Valor3, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_valor_total_venda_Change()
On Error GoTo tratar_erro

If Txt_valor_total_venda <> "" Then
    VerifNumero = Txt_valor_total_venda
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor_total_venda = ""
        Txt_valor_total_venda.SetFocus
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
    Case 1: ProcGravar
    Case 2: ProcExcluir
    'Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcLiberaBloqueia()
On Error GoTo tratar_erro

If chkTotal.Value = 0 Then
    Frame16(0).Enabled = True
    Frame16(1).Enabled = True
    Frame16(2).Enabled = True
    Frame16(3).Enabled = True
    Frame16(4).Enabled = False
Else
    Frame16(0).Enabled = False
    Frame16(1).Enabled = False
    Frame16(2).Enabled = False
    Frame16(3).Enabled = False
    Frame16(4).Enabled = True
End If
ProcLimpaCampos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerificaPorcentagem(pICMS As Double, pPIS As Double, pCofins As Double, pCSLL As Double, pISSQN As Double, pIRPJ As Double, pSimples As Double, pComissao As Double, pComercias As Double, pFinanceiro As Double, pAdm As Double, pFrete As Double, pMargem As Double) As Boolean
On Error GoTo tratar_erro
Dim QtdePorcentagem As Double

If Margem_Reciproca = True Then
    QtdePorcentagem = pICMS + pPIS + pCofins + pCSLL + pISSQN + pIRPJ + pSimples + pComissao + pComercias + pFinanceiro + pAdm + pFrete + pMargem
Else
    QtdePorcentagem = pICMS + pPIS + pCofins + pCSLL + pISSQN + pIRPJ + pSimples + pComissao + pComercias + pFinanceiro + pAdm + pFrete
End If

If QtdePorcentagem > 100 Then
    FunVerificaPorcentagem = False
    USMsgBox "A soma de porcentagens não pode ser maior que 100, favor revisar.", vbInformation, "CAPRIND v5.0"
Else
    FunVerificaPorcentagem = True
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function
