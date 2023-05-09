VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmContas_pagar_localizar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Financeiro - Contas a pagar - Localizar"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8895
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkIncluir_contaFixa 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Incluir conta fixa"
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
      Left            =   3210
      TabIndex        =   14
      Top             =   1568
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkConta_fixa 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Apenas conta fixa"
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
      Left            =   3210
      TabIndex        =   15
      Top             =   1815
      Width           =   1815
   End
   Begin VB.ComboBox Cmb_agendado 
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
      ItemData        =   "frmContas_pagar_localizar.frx":0000
      Left            =   7770
      List            =   "frmContas_pagar_localizar.frx":000D
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      ToolTipText     =   "Agendado pagamento."
      Top             =   1500
      Width           =   885
   End
   Begin VB.CheckBox Chk_bloqueada 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Bloqueada"
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
      Left            =   5130
      TabIndex        =   16
      Top             =   1568
      Width           =   1215
   End
   Begin VB.CheckBox Chk_devolucao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Incluir contas de devolução"
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
      Left            =   300
      TabIndex        =   13
      Top             =   1815
      Value           =   1  'Checked
      Width           =   2805
   End
   Begin VB.CheckBox Chk_antecipacao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Incluir contas de antecipação"
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
      Left            =   300
      TabIndex        =   12
      Top             =   1568
      Value           =   1  'Checked
      Width           =   2805
   End
   Begin VB.CheckBox Chk_transacao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Transação"
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
      Left            =   270
      TabIndex        =   17
      Top             =   3870
      Width           =   1185
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   4680
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmContas_pagar_localizar.frx":0022
      Count           =   1
   End
   Begin VB.ComboBox Cmb_empresa 
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
      ItemData        =   "frmContas_pagar_localizar.frx":220A
      Left            =   1170
      List            =   "frmContas_pagar_localizar.frx":220C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1110
      Width           =   5115
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   25
      Top             =   2070
      Width           =   8805
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3810
         TabIndex        =   32
         Top             =   210
         Width           =   4785
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3930
            TabIndex        =   8
            Top             =   180
            Width           =   705
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1470
            TabIndex        =   6
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   5
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   7
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.ComboBox Cmb_ordenar 
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
         ItemData        =   "frmContas_pagar_localizar.frx":220E
         Left            =   6330
         List            =   "frmContas_pagar_localizar.frx":2221
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Ordenar por."
         Top             =   1050
         Width           =   2265
      End
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
         ItemData        =   "frmContas_pagar_localizar.frx":2264
         Left            =   180
         List            =   "frmContas_pagar_localizar.frx":2280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3555
      End
      Begin VB.TextBox txtTexto 
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
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   6135
      End
      Begin VB.ComboBox cmbTexto 
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
         ItemData        =   "frmContas_pagar_localizar.frx":22F3
         Left            =   180
         List            =   "frmContas_pagar_localizar.frx":22F5
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   6135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar por"
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
         Left            =   6952
         TabIndex        =   30
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label45 
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
         Left            =   1537
         TabIndex        =   27
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label3 
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
         Left            =   2512
         TabIndex        =   26
         Top             =   840
         Width           =   1470
      End
   End
   Begin VB.CheckBox chkVencimento 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vencimento"
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
      Left            =   2820
      TabIndex        =   19
      Top             =   3870
      Width           =   1305
   End
   Begin VB.CheckBox chkEmissao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Emissão"
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
      Left            =   1650
      TabIndex        =   18
      Top             =   3870
      Width           =   1005
   End
   Begin VB.ComboBox cmbtipo_conta 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "frmContas_pagar_localizar.frx":22F7
      Left            =   7770
      List            =   "frmContas_pagar_localizar.frx":22F9
      Style           =   2  'Dropdown List
      TabIndex        =   10
      ToolTipText     =   "Tipo do documento."
      Top             =   1140
      Width           =   885
   End
   Begin VB.CheckBox Optclassificacao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo docto. :"
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
      Left            =   6390
      TabIndex        =   9
      Top             =   1140
      Width           =   1335
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   28
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
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
      ButtonUseMaskColor2=   0   'False
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
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   675
      Left            =   55
      TabIndex        =   22
      Top             =   3600
      Width           =   8805
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   7320
         TabIndex        =   21
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
         Format          =   489816065
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   5430
         TabIndex        =   20
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
         Format          =   487456769
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   285
         Left            =   5070
         TabIndex        =   24
         Top             =   240
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
         Height          =   285
         Left            =   6915
         TabIndex        =   23
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agendado :"
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
      Left            =   6720
      TabIndex        =   31
      Top             =   1500
      Width           =   945
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa :"
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
      Left            =   270
      TabIndex        =   29
      Top             =   1110
      Width           =   825
   End
End
Attribute VB_Name = "frmContas_pagar_localizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Todas_contas As Boolean 'OK

Private Sub chkConta_fixa_Click()
On Error GoTo tratar_erro

If chkConta_fixa.Value = 1 Then chkIncluir_contaFixa.Value = 0
ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_antecipacao_Click()
On Error GoTo tratar_erro

ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_bloqueada_Click()
On Error GoTo tratar_erro

ProcCarregaComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_transacao_Click()
On Error GoTo tratar_erro

If Chk_transacao.Value = 1 Then
    chkEmissao.Value = 0
    chkVencimento.Value = 0
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkEmissao_Click()
On Error GoTo tratar_erro

If chkEmissao.Value = 1 Then
    Chk_transacao.Value = 0
    chkVencimento.Value = 0
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkIncluir_contaFixa_Click()
On Error GoTo tratar_erro

If chkIncluir_contaFixa.Value = 1 Then chkConta_fixa.Value = 0
ProcCarregaComboTexto

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

If Chk_bloqueada.Value = 1 Then TextoFiltroBloq = "CP.Bloqueado = 'True'" Else TextoFiltroBloq = "CP.Bloqueado = 'False'"
If chkConta_fixa.Value = 1 Then TextoFiltroFixa = " and CP.Conta_fixa = 'True'" Else TextoFiltroFixa = ""

Optinicio.Value = True
If cmbfiltrarpor = "Status" Or cmbfiltrarpor = "Conta contábil" Or cmbfiltrarpor = "Fornecedor" Or cmbfiltrarpor = "Código da conta contábil" Then
    txtTexto.Visible = False
    cmbTexto.Visible = True
    
    With cmbTexto
        .Clear
        .AddItem ""
        Set TBLISTA = CreateObject("adodb.recordset")
        If cmbfiltrarpor = "Conta contábil" Or cmbfiltrarpor = "Código da conta contábil" Then
            If Chk_antecipacao.Value = 1 Then TextoFiltro = "(CP.LogSit = 'N' or CP.Status = 'TÍTULO LIQUIDADO ANTECIPADO')" Else TextoFiltro = "CP.LogSit = 'N'"
            Set TBLISTA = CreateObject("adodb.recordset")
            TBLISTA.Open "Select F.int_codfamilia, F.Codigo, F.txt_descricao from (tbl_ContasPagar CP INNER JOIN familia_financeiro FF ON CP.IdIntConta = FF.IDConta) INNER JOIN tbl_familia F ON F.int_codfamilia = FF.ID_PC where FF.tipoconta = 'P' and FF.Deposito_transf <> 'True' and " & TextoFiltro & " and CP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TextoFiltroBloq & TextoFiltroFixa & " and " & frmContas_Pagar.Filtro_Contas_Pagar_Func & " Group by F.int_codfamilia, F.Codigo, F.txt_descricao", Conexao, adOpenKeyset, adLockOptimistic
            If TBLISTA.EOF = False Then
                Do While TBLISTA.EOF = False
                    If cmbfiltrarpor = "Conta contábil" Then .AddItem TBLISTA!Txt_descricao & " - " & TBLISTA!CODIGO Else .AddItem TBLISTA!CODIGO & " - " & TBLISTA!Txt_descricao
                    .ItemData(cmbTexto.NewIndex) = TBLISTA!int_codfamilia
                    TBLISTA.MoveNext
                Loop
            End If
            TBLISTA.Close
        ElseIf cmbfiltrarpor = "Fornecedor" Then
                If Chk_antecipacao.Value = 1 Then TextoFiltro = "(LogSit = 'N' or Status = 'TÍTULO LIQUIDADO ANTECIPADO')" Else TextoFiltro = "LogSit = 'N'"
                Set TBLISTA = CreateObject("adodb.recordset")
                TBLISTA.Open "Select int_codforn, txt_Fornecedor from tbl_ContasPagar CP where txt_Fornecedor is not null and " & TextoFiltro & " and ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & TextoFiltroBloq & TextoFiltroFixa & " and " & frmContas_Pagar.Filtro_Contas_Pagar_Func & " Group by int_codforn, txt_Fornecedor", Conexao, adOpenKeyset, adLockOptimistic
                If TBLISTA.EOF = False Then
                    Do While TBLISTA.EOF = False
                        .AddItem TBLISTA!Txt_fornecedor
                        .ItemData(cmbTexto.NewIndex) = TBLISTA!int_codforn
                        TBLISTA.MoveNext
                    Loop
                End If
                TBLISTA.Close
            ElseIf cmbfiltrarpor = "Status" Then
                    .AddItem "TÍTULO EM ABERTO"
                    If Chk_antecipacao.Value = 1 Then
                        .AddItem "TÍTULO LIQUIDADO ANTECIPADO"
                    Else
                        .AddItem "TÍTULO PAGO PARCIAL"
                    End If
        End If
    End With
Else
    txtTexto.Visible = True
    cmbTexto.Visible = False
    If cmbfiltrarpor = "Documento" Then
        Optfim.Value = True
    ElseIf cmbfiltrarpor = "Valor" Then
            If txtTexto <> "" Then
                VerifNumero = txtTexto
                ProcVerificaNumero
                If VerifNumero = False Then
                    txtTexto = ""
                    txtTexto.SetFocus
                    Exit Sub
                End If
            End If
    End If
End If

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
With frmContas_Pagar
    .StrSql_Contas_Pagar_AntecTotal = ""
    .StrSql_Contas_Pagar_DevTotal = ""
   
    Todas_contas = False
    
    TextoFiltroClass = ""
    TextoFiltroClassRel = ""
    If Optclassificacao.Value = 1 Then
        TextoFiltroClass = " and CP.class_conta = '" & cmbtipo_conta & "'"
        TextoFiltroClassRel = " and {tbl_contaspagar.class_conta} = '" & cmbtipo_conta & "'"
    End If
    TextoFiltroAgendado = ""
    TextoFiltroAgendadoRel = ""
    If Cmb_agendado = "Sim" Then
        TextoFiltroAgendado = " and CP.Agendado = 'True'"
        TextoFiltroAgendadoRel = " and {tbl_contaspagar.Agendado} = True"
    ElseIf Cmb_agendado = "Não" Then
            TextoFiltroAgendado = " and CP.Agendado = 'False'"
            TextoFiltroAgendadoRel = " and {tbl_contaspagar.Agendado} = False"
    End If
    If Chk_bloqueada.Value = 1 Then
        TextoFiltroBloq = " and CP.Bloqueado = 'True'"
        TextoFiltroBloqRel = " and {tbl_ContasPagar.Bloqueado} = True"
    Else
        TextoFiltroBloq = " and CP.Bloqueado = 'False'"
        TextoFiltroBloqRel = " and {tbl_ContasPagar.Bloqueado} = False"
    End If
    If chkConta_fixa.Value = 1 Then
        TextoFiltroFixa = " and CP.Conta_fixa = 'True'"
        TextoFiltroFixaRel = " and {tbl_ContasPagar.Conta_fixa} = True"
    ElseIf chkIncluir_contaFixa.Value = 1 Then
        TextoFiltroFixa = ""
        TextoFiltroFixaRel = ""
    Else
        TextoFiltroFixa = " and CP.Conta_fixa = 'False'"
        TextoFiltroFixaRel = " and {tbl_ContasPagar.Conta_fixa} = False"
    End If
    
    If Chk_antecipacao.Value = 1 Then
        TextoFiltroAntecipacao = " and (CP.LogSit = 'N' or CP.Status = 'TÍTULO LIQUIDADO ANTECIPADO' and CP.Saldo_antecipacao > 0)"
        TextoFiltroAntecipacaoRel = " and ({tbl_contaspagar.LogSit} = 'N' or {tbl_contaspagar.Status} = 'TÍTULO LIQUIDADO ANTECIPADO' and {tbl_contaspagar.Saldo_antecipacao} > 0)"
        TextoFiltroAntecipacao1 = " and CP.Status = 'TÍTULO LIQUIDADO ANTECIPADO' and CP.Saldo_antecipacao > 0"
    Else
        TextoFiltroAntecipacao = " and CP.Antecipacao = 'False' and CP.LogSit = 'N'"
        TextoFiltroAntecipacaoRel = " and {tbl_ContasPagar.Antecipacao} = False and {tbl_contaspagar.LogSit} = 'N'"
        TextoFiltroAntecipacao1 = ""
    End If
    
    If Chk_devolucao.Value = 1 Then
        TextoFiltroDevolucao = " and (CP.LogSit = 'N' or CP.Devolucao = 'True')"
        TextoFiltroDevolucaoRel = " and ({tbl_contaspagar.LogSit} = 'N' or {tbl_ContasPagar.Devolucao} = True)"
        TextoFiltroDevolucao1 = " and CP.Devolucao = 'True' and CP.LogSit = 'N'"
    Else
        TextoFiltroDevolucao = " and CP.Devolucao = 'False' and CP.LogSit = 'N'"
        TextoFiltroDevolucaoRel = " and {tbl_ContasPagar.Devolucao} = False and {tbl_contaspagar.LogSit} = 'N'"
        TextoFiltroDevolucao1 = ""
    End If
    
    DataFiltro = ""
    DataFiltroRel = ""
    If Chk_transacao.Value = 1 Or chkVencimento.Value = 1 Or chkEmissao.Value = 1 Then
        If Chk_transacao.Value = 1 Then
            Data_pagar = "CP.Data_transacao"
        ElseIf chkEmissao.Value = 1 Then
                Data_pagar = "CP.dt_Emissao"
            Else
                Data_pagar = "CP.dt_Pagamento"
        End If
        DataFiltro = " and " & Data_pagar & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
        DataFiltroRel = " and {" & Replace(Data_pagar, "CP.", "tbl_ContasPagar.") & "} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {" & Replace(Data_pagar, "CP.", "tbl_ContasPagar.") & "} <= Date(" & _
                                Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
    End If
    Select Case Cmb_ordenar
        Case "Dt. emissão": Ordenar = "CP.dt_Emissao"
        Case "Dt. transação": Ordenar = "CP.Data_transacao"
        Case "Dt. vencimento": Ordenar = "CP.dt_Pagamento"
        Case "Fornecedor": Ordenar = "CP.txt_Fornecedor"
        Case "Valor": Ordenar = "CP.dbl_valorpagto"
    End Select
    
    NomeRel = "Contas_pagar.rpt"
    
    CamposFiltro = "CP.IDintconta, CP.Dt_emissao, CP.dt_Pagamento, CP.dbl_valorpagto, CP.Class_conta, CP.txt_ndocumento, CP.txt_Parcela, CP.Txt_fornecedor, CP.ID_empresa, CP.Responsavel, CP.Antecipacao, CP.Saldo_antecipacao"
    INNERJOINTEXTO = "Select " & CamposFiltro & " from tbl_ContasPagar CP"
    INNERJOINTEXTOSUM = "Select SUM(CP.dbl_valorpagto) AS TotContas from tbl_ContasPagar CP"
    INNERJOINTEXTOSUMANTEC = "Select SUM(CP.Saldo_antecipacao) AS TotContas1 from tbl_ContasPagar CP"
    OrdenarTexto = " group by " & CamposFiltro & " order by " & Ordenar & ", CP.IdIntConta"
    TextoFiltroPadrao1 = "CP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & TextoFiltroClass & TextoFiltroBloq & TextoFiltroFixa & DataFiltro & TextoFiltroAgendado & " and " & .Filtro_Contas_Pagar_Func
    TextoFiltroPadrao = TextoFiltroPadrao1 & TextoFiltroAntecipacao & TextoFiltroDevolucao
    TextoFiltroPadraoSOMA = TextoFiltroPadrao1 & " and CP.Logsit = 'N' and CP.Antecipacao = 'False' and CP.Devolucao = 'False'"
    TextoFiltroPadraoANTEC = TextoFiltroPadrao1 & TextoFiltroAntecipacao1
    TextoFiltroPadraoDEV = TextoFiltroPadrao1 & TextoFiltroDevolucao1
    TextoFiltroPadraoRel = "{tbl_contaspagar.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & TextoFiltroClassRel & TextoFiltroBloqRel & TextoFiltroFixaRel & DataFiltroRel & TextoFiltroAgendadoRel & TextoFiltroAntecipacaoRel & TextoFiltroDevolucaoRel & " and " & .Filtro_Contas_Pagar_FuncRel
    
    If txtTexto.Visible = True And txtTexto <> "" Or cmbTexto.Visible = True And cmbTexto <> "" Then
        If cmbTexto.Visible = True Then
            If cmbfiltrarpor = "Conta contábil" Or cmbfiltrarpor = "Código da conta contábil" Then
                NomeRel = "Contas_pagar_conta contabil.rpt"
                
                INNERJOINPADRAO = " from tbl_ContasPagar CP INNER JOIN familia_financeiro FF ON CP.IdIntConta = FF.idconta"
                INNERJOINTEXTO = "Select " & CamposFiltro & INNERJOINPADRAO
                INNERJOINTEXTOSUM = "Select SUM(CP.dbl_valorpagto) AS TotContas " & INNERJOINPADRAO
                INNERJOINTEXTOSUMANTEC = "Select SUM(CP.Saldo_antecipacao) AS TotContas1 " & INNERJOINPADRAO
                TextoFiltro = "FF.ID_PC = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and FF.tipoconta = 'P'"
                .StrSql_Contas_Pagar = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao & OrdenarTexto
                .StrSql_Contas_PagarTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoSOMA
                If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Pagar_AntecTotal = INNERJOINTEXTOSUMANTEC & " where " & TextoFiltro & " and " & TextoFiltroPadraoANTEC
                If Chk_devolucao.Value = 1 Then .StrSql_Contas_Pagar_DevTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoDEV
                .FormulaRel_Contas_Pagar = "{familia_financeiro.ID_PC} = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and {familia_financeiro.tipoconta} = 'P' and " & TextoFiltroPadraoRel
            Else
                If cmbfiltrarpor = "Status" Then TextoFiltro = "CP.Status" Else TextoFiltro = "CP.txt_Fornecedor"
                .StrSql_Contas_Pagar = INNERJOINTEXTO & " where " & TextoFiltro & " = '" & cmbTexto.Text & "' and " & TextoFiltroPadrao & OrdenarTexto
                .StrSql_Contas_PagarTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " = '" & cmbTexto.Text & "' and " & TextoFiltroPadraoSOMA
                If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Pagar_AntecTotal = INNERJOINTEXTOSUMANTEC & " where " & TextoFiltro & " = '" & cmbTexto.Text & "' and " & TextoFiltroPadraoANTEC
                If Chk_devolucao.Value = 1 Then .StrSql_Contas_Pagar_DevTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " = '" & cmbTexto.Text & "' and " & TextoFiltroPadraoDEV
                .FormulaRel_Contas_Pagar = "{" & Replace(TextoFiltro, "CP.", "tbl_ContasPagar.") & "} = '" & cmbTexto.Text & "' and " & TextoFiltroPadraoRel
            End If
        Else
            If cmbfiltrarpor = "Valor" Then
                valor = txtTexto
                NovoValor = Replace(valor, ",", ".")
                .StrSql_Contas_Pagar = INNERJOINTEXTO & " where CP.dbl_valorpagto = " & NovoValor & " and " & TextoFiltroPadrao & OrdenarTexto
                .StrSql_Contas_PagarTotal = INNERJOINTEXTOSUM & " where CP.dbl_valorpagto = " & NovoValor & " and " & TextoFiltroPadraoSOMA
                If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Pagar_AntecTotal = INNERJOINTEXTOSUMANTEC & " where CP.dbl_valorpagto = " & NovoValor & " and " & TextoFiltroPadraoANTEC
                If Chk_devolucao.Value = 1 Then .StrSql_Contas_Pagar_DevTotal = INNERJOINTEXTOSUM & " where CP.dbl_valorpagto = " & NovoValor & " and " & TextoFiltroPadraoDEV
                .FormulaRel_Contas_Pagar = "{tbl_ContasPagar.dbl_valorpagto} = " & NovoValor & " and " & TextoFiltroPadraoRel
            ElseIf cmbfiltrarpor = "Pedido de compra" Then
                    INNERJOINPADRAO = " from tbl_ContasPagar CP LEFT JOIN tbl_proposta_nota PN ON PN.Id_nota = CP.Id_nota"
                    INNERJOINTEXTO = "Select " & CamposFiltro & INNERJOINPADRAO
                    INNERJOINTEXTOSUM = "Select SUM(CP.dbl_valorpagto) AS TotContas " & INNERJOINPADRAO
                    INNERJOINTEXTOSUMANTEC = "Select SUM(CP.Saldo_antecipacao) AS TotContas1 " & INNERJOINPADRAO
                    TextoFiltro = "PN.proposta " & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto)
                    .StrSql_Contas_Pagar = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao & OrdenarTexto
                    .StrSql_Contas_PagarTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoSOMA
                    If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Pagar_AntecTotal = INNERJOINTEXTOSUMANTEC & " where " & TextoFiltro & " and " & TextoFiltroPadraoANTEC
                    If Chk_devolucao.Value = 1 Then .StrSql_Contas_Pagar_DevTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & " and " & TextoFiltroPadraoDEV
                    .FormulaRel_Contas_Pagar = "{tbl_proposta_nota.proposta}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " And " & TextoFiltroPadraoRel
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open .StrSql_Contas_Pagar, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = True Then
                        TextoFiltro = "CP.txt_pedido"
                        .StrSql_Contas_Pagar = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao & OrdenarTexto
                        .StrSql_Contas_PagarTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoSOMA
                        If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Pagar_AntecTotal = INNERJOINTEXTOSUMANTEC & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoANTEC
                        If Chk_devolucao.Value = 1 Then .StrSql_Contas_Pagar_DevTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoDEV
                        .FormulaRel_Contas_Pagar = "{" & Replace(TextoFiltro, "CP.", "tbl_ContasPagar.") & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
                    End If
                Else
                    If cmbfiltrarpor = "Documento" Then TextoFiltro = "CP.txt_NDocumento" Else TextoFiltro = "CP.Competencia"
                    .StrSql_Contas_Pagar = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao & OrdenarTexto
                    .StrSql_Contas_PagarTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " And " & TextoFiltroPadraoSOMA
                    If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Pagar_AntecTotal = INNERJOINTEXTOSUMANTEC & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoANTEC
                    If Chk_devolucao.Value = 1 Then .StrSql_Contas_Pagar_DevTotal = INNERJOINTEXTOSUM & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoDEV
                    .FormulaRel_Contas_Pagar = "{" & Replace(TextoFiltro, "CP.", "tbl_ContasPagar.") & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
            End If
        End If
    Else
        .StrSql_Contas_Pagar = INNERJOINTEXTO & " where " & TextoFiltroPadrao & OrdenarTexto
        .StrSql_Contas_PagarTotal = INNERJOINTEXTOSUM & " where " & TextoFiltroPadraoSOMA
        If Chk_antecipacao.Value = 1 Then .StrSql_Contas_Pagar_AntecTotal = INNERJOINTEXTOSUMANTEC & " where " & TextoFiltroPadraoANTEC
        If Chk_devolucao.Value = 1 Then .StrSql_Contas_Pagar_DevTotal = INNERJOINTEXTOSUM & " where " & TextoFiltroPadraoDEV
        .FormulaRel_Contas_Pagar = TextoFiltroPadraoRel
        Todas_contas = True
    End If
    .ProcSalvarDadosRel Chk_transacao, chkEmissao, chkVencimento, msk_fltInicio, msk_fltFim
    .ProcCarregaLista (1)
    .Novo_Pagar = False
End With

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: ProcSair
    Case vbKeyF2: ProcFiltrar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 8805, 5, True

ProcCarregaComboEmpresa Cmb_empresa, False
cmbfiltrarpor = "Documento"
Cmb_ordenar = "Dt. vencimento"
msk_fltFim.Value = Date
msk_fltInicio.Value = Date

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

Private Sub Optclassificacao_Click()
On Error GoTo tratar_erro

With cmbtipo_conta
    If Optclassificacao.Value = 1 Then
        ProcCarregaComboTipoDocto cmbtipo_conta, "Tipo = 'P'"
        .Enabled = True
        .SetFocus
    Else
        .Enabled = False
        .ListIndex = -1
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkVencimento_Click()
On Error GoTo tratar_erro

If chkVencimento.Value = 1 Then
    Chk_transacao.Value = 0
    chkEmissao.Value = 0
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Valor" And txtTexto <> "" Then
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

Private Sub txtTexto_LostFocus()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Valor" And txtTexto <> "" Then txtTexto = Format(txtTexto, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
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
