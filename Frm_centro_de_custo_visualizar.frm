VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form Frm_centro_de_custo_visualizar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Custos - Centro de custo - Visualizar lançamentos realizados"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11295
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar lançamentos"
      Height          =   570
      Left            =   60
      TabIndex        =   22
      Top             =   990
      Width           =   11190
      Begin VB.ComboBox cmbAno 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Frm_centro_de_custo_visualizar.frx":0000
         Left            =   10170
         List            =   "Frm_centro_de_custo_visualizar.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OptDomes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Do mês"
         Height          =   195
         Left            =   150
         TabIndex        =   0
         Top             =   270
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton OptAteomes 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Até o mês"
         Height          =   195
         Left            =   1020
         TabIndex        =   1
         Top             =   270
         Width           =   1035
      End
      Begin MSComctlLib.TabStrip TabFiltro 
         Height          =   345
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   609
         TabWidthStyle   =   1
         MultiRow        =   -1  'True
         TabMinWidth     =   1177
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   12
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Jan"
               Key             =   "Jan"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Fev"
               Key             =   "Fev"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Mar"
               Key             =   "Mar"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Abril"
               Key             =   "Abr"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Maio"
               Key             =   "Maio"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Jun"
               Key             =   "Jun"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Jul"
               Key             =   "Jul"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Ago"
               Key             =   "Ago"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Set"
               Key             =   "Set"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Out"
               Key             =   "Out"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Nov"
               Key             =   "Nov"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dez"
               Key             =   "Dez"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox TXT_ID 
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
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "0"
      ToolTipText     =   "ID."
      Top             =   4770
      Visible         =   0   'False
      Width           =   1125
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   4590
      Top             =   270
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "Frm_centro_de_custo_visualizar.frx":0004
      Count           =   1
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   12
      Top             =   8010
      Width           =   11190
      _ExtentX        =   19738
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   13
      Top             =   0
      Width           =   11190
      _ExtentX        =   19738
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
      ButtonLeft2     =   42
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
      ButtonLeft3     =   46
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
      ButtonLeft4     =   84
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
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
      ButtonState5    =   5
      ButtonLeft5     =   112
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
   End
   Begin VB.Frame Frame2 
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
      Height          =   1395
      Left            =   55
      TabIndex        =   11
      Top             =   1590
      Width           =   11190
      Begin VB.TextBox Txt_modulo 
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
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Módulo."
         Top             =   375
         Width           =   1155
      End
      Begin VB.TextBox Txt_valor 
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
         Left            =   9840
         TabIndex        =   9
         ToolTipText     =   "Valor."
         Top             =   930
         Width           =   1155
      End
      Begin VB.TextBox Txt_operacao 
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
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Operação."
         Top             =   930
         Width           =   1365
      End
      Begin VB.TextBox Txt_referencia 
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
         Left            =   5490
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Referência."
         Top             =   930
         Width           =   2955
      End
      Begin VB.ComboBox Cmb_centro 
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
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Centro de custo."
         Top             =   930
         Width           =   5295
      End
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
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Responsável."
         Top             =   375
         Width           =   8445
      End
      Begin MSComCtl2.DTPicker Txt_data 
         Height          =   315
         Left            =   180
         TabIndex        =   23
         ToolTipText     =   "Data de emissão."
         Top             =   375
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   489488387
         CurrentDate     =   39057
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Centro de custo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   2250
         TabIndex        =   21
         Top             =   735
         Width           =   1155
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operação"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   8790
         TabIndex        =   20
         Top             =   735
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   10237
         TabIndex        =   18
         Top             =   740
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referência"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   6577
         TabIndex        =   17
         Top             =   735
         Width           =   780
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Módulo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   10162
         TabIndex        =   16
         Top             =   180
         Width           =   510
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   570
         TabIndex        =   15
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   5115
         TabIndex        =   14
         Top             =   180
         Width           =   915
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4995
      Left            =   60
      TabIndex        =   10
      Top             =   3000
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   8811
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Responsável"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "CC origem"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Módulo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Referência"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
   End
End
Attribute VB_Name = "Frm_centro_de_custo_visualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Frame2.Enabled = False Then Exit Sub
Acao = "salvar"
If Cmb_centro = "" Then
    NomeCampo = "o centro de custo"
    ProcVerificaAcao
    Cmb_centro.SetFocus
    Exit Sub
End If

With Frm_centro_de_custo
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select * from CC_realizado where ID = " & Txt_ID, Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        If TBFIltro!Operacao = "Crédito" Then
            Operacao = "Débito"
            Operacao1 = "Crédito"
        Else
            Operacao = "Crédito"
            Operacao1 = "Débito"
        End If
        
        'Salva os dados e atualiza o saldo do centro de custo de origem
        Familiatext = Operacao
        ProcSalvarCCRealizado Txt_data, TBFIltro!ID_empresa, Familiatext, TBFIltro!ID_CC, IIf(IsNull(TBFIltro!Cod_produto), 0, TBFIltro!Cod_produto), IIf(IsNull(TBFIltro!ID_PC), 0, TBFIltro!ID_PC), IIf(IsNull(TBFIltro!ID_estoque), 0, TBFIltro!ID_estoque), IIf(IsNull(TBFIltro!ID_lista), 0, TBFIltro!ID_lista), IIf(IsNull(TBFIltro!ID_financeiro), 0, TBFIltro!ID_financeiro), Txt_valor, True, TBFIltro!ID
        
        'Grava movimentação no centro consolidado
        Set TBAfericao = CreateObject("adodb.recordset")
        TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBFIltro!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
        If TBAfericao.EOF = False Then
            Do While TBAfericao.EOF = False
                ProcSalvarCCRealizado Txt_data, TBFIltro!ID_empresa, Familiatext, TBAfericao!ID_CC, IIf(IsNull(TBFIltro!Cod_produto), 0, TBFIltro!Cod_produto), IIf(IsNull(TBFIltro!ID_PC), 0, TBFIltro!ID_PC), IIf(IsNull(TBFIltro!ID_estoque), 0, TBFIltro!ID_estoque), IIf(IsNull(TBFIltro!ID_lista), 0, TBFIltro!ID_lista), IIf(IsNull(TBFIltro!ID_financeiro), 0, TBFIltro!ID_financeiro), Txt_valor, True, TBFIltro!ID
                
                Set TBCiclo = CreateObject("adodb.recordset")
                TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                If TBCiclo.EOF = False Then
                    Do While TBCiclo.EOF = False
                        ProcSalvarCCRealizado Txt_data, TBFIltro!ID_empresa, Familiatext, TBCiclo!ID_CC, IIf(IsNull(TBFIltro!Cod_produto), 0, TBFIltro!Cod_produto), IIf(IsNull(TBFIltro!ID_PC), 0, TBFIltro!ID_PC), IIf(IsNull(TBFIltro!ID_estoque), 0, TBFIltro!ID_estoque), IIf(IsNull(TBFIltro!ID_lista), 0, TBFIltro!ID_lista), IIf(IsNull(TBFIltro!ID_financeiro), 0, TBFIltro!ID_financeiro), Txt_valor, True, TBFIltro!ID
                        TBCiclo.MoveNext
                    Loop
                End If
                TBCiclo.Close
                
                TBAfericao.MoveNext
            Loop
        End If
        
        'Salva os dados e atualiza o saldo do centro de custo de destino
        Familiatext = Operacao1
        ProcSalvarCCRealizado Txt_data, TBFIltro!ID_empresa, Familiatext, Cmb_centro.ItemData(Cmb_centro.ListIndex), IIf(IsNull(TBFIltro!Cod_produto), 0, TBFIltro!Cod_produto), IIf(IsNull(TBFIltro!ID_PC), 0, TBFIltro!ID_PC), IIf(IsNull(TBFIltro!ID_estoque), 0, TBFIltro!ID_estoque), IIf(IsNull(TBFIltro!ID_lista), 0, TBFIltro!ID_lista), IIf(IsNull(TBFIltro!ID_financeiro), 0, TBFIltro!ID_financeiro), Txt_valor, True, TBFIltro!ID
        
        'Grava movimentação no centro consolidado
        Set TBAfericao = CreateObject("adodb.recordset")
        TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & Cmb_centro.ItemData(Cmb_centro.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
        If TBAfericao.EOF = False Then
            Do While TBAfericao.EOF = False
                ProcSalvarCCRealizado Txt_data, TBFIltro!ID_empresa, Familiatext, TBAfericao!ID_CC, IIf(IsNull(TBFIltro!Cod_produto), 0, TBFIltro!Cod_produto), IIf(IsNull(TBFIltro!ID_PC), 0, TBFIltro!ID_PC), IIf(IsNull(TBFIltro!ID_estoque), 0, TBFIltro!ID_estoque), IIf(IsNull(TBFIltro!ID_lista), 0, TBFIltro!ID_lista), IIf(IsNull(TBFIltro!ID_financeiro), 0, TBFIltro!ID_financeiro), Txt_valor, True, TBFIltro!ID
                
                Set TBCiclo = CreateObject("adodb.recordset")
                TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                If TBCiclo.EOF = False Then
                    Do While TBCiclo.EOF = False
                        ProcSalvarCCRealizado Txt_data, TBFIltro!ID_empresa, Familiatext, TBCiclo!ID_CC, IIf(IsNull(TBFIltro!Cod_produto), 0, TBFIltro!Cod_produto), IIf(IsNull(TBFIltro!ID_PC), 0, TBFIltro!ID_PC), IIf(IsNull(TBFIltro!ID_estoque), 0, TBFIltro!ID_estoque), IIf(IsNull(TBFIltro!ID_lista), 0, TBFIltro!ID_lista), IIf(IsNull(TBFIltro!ID_financeiro), 0, TBFIltro!ID_financeiro), Txt_valor, True, TBFIltro!ID
                        TBCiclo.MoveNext
                    Loop
                End If
                TBCiclo.Close
                
                TBAfericao.MoveNext
            Loop
        End If
        TBAfericao.Close
        
        '==================================
        Modulo = "Custos/Centro de custo/Visualizar lançamentos realizados"
        Evento = "Alterar"
        ID_documento = Txt_ID
        If .Lista.SelectedItem.ListSubItems(4) <> "" Then Texto = .Lista.SelectedItem.ListSubItems(4) & " - " & .Lista.SelectedItem.ListSubItems(5) Else Texto = .Lista.SelectedItem.ListSubItems(5)
        Documento = "Data: " & Txt_data & " - Responsável: " & Txt_responsavel & " - Módulo: " & Txt_modulo & " - Operação: " & Txt_operacao & " - De: " & Texto & " - Para: " & Cmb_centro
        Documento1 = ""
        ProcGravaEvento
        '==================================
        
        ProcFiltrarMes
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
            Lista.SelectedItem = Lista.ListItems(CodigoLista)
            Lista.SetFocus
        End If
    End If
    TBFIltro.Close
    ProcLimpaCampos
    
    .ProcCarregaLista
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarCCRealizado(Data1 As Date, ID_empresa As Integer, Operacao As String, ID_CC As Long, Cod_produto As Long, ID_plano_contas As Long, ID_estoque As Long, ID_lista As Long, ID_financeiro As Long, valor As Double, Bloqueado As Boolean, ID_origem As Long)
On Error GoTo tratar_erro

NovoValor = Replace(valor, ",", ".")
ProcINSERTINTO "CC_realizado", "Data, Responsavel, ID_empresa, Operacao, ID_CC, Cod_produto, ID_PC, ID_estoque, ID_lista, ID_financeiro, Valor, Bloqueado, ID_origem", "'" & Data & "', '" & pubUsuario & "', " & ID_empresa & ", '" & Operacao & "', " & ID_CC & ", " & Cod_produto & ", " & ID_plano_contas & ", " & IIf(ID_estoque = 0, "NULL", ID_estoque) & ", " & ID_lista & ", " & IIf(ID_financeiro = 0, "NULL", ID_financeiro) & ", " & NovoValor & ", " & IIf(Bloqueado = True, 1, 0) & ", " & ID_origem & ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbAno_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

With Frm_centro_de_custo
    Caption = "Custos - Centro de custo - Visualizar lançamentos realizados - Centro: " & .Lista.SelectedItem.ListSubItems(5)
    ProcCarregaToolBar1 Me, 11190, 5, True
    ProcLimpaVariaveisPrincipais
    ProcCarregaComboAno cmbAno, "2011", 1
    ProcCarregaComboSetor Cmb_centro, "Setor is not null and ID <> " & .txtID & " and (Consolidacao = 'False' or Consolidacao is null)", "", False, False, False, "", True, False
    TabFiltro.Tabs(Month(Date)).Selected = True
End With

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

Sub ProcCarregaLista(TextoFiltro As String)
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
            If IsNull(TBLISTA!ID_origem) = False And TBLISTA!ID_origem <> 0 Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Usuarios_Setor.Setor from Usuarios_Setor INNER JOIN CC_realizado ON Usuarios_Setor.Id = CC_realizado.ID_CC where CC_realizado.Id = " & TBLISTA!ID_origem, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBAbrir!Setor), "", TBAbrir!Setor)
                End If
            End If
            If IsNull(TBLISTA!ID_estoque) = False And TBLISTA!ID_estoque <> 0 Then
                Modulo_texto = "Estoque"
                
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Entrada, Documento, Lote from Estoque_movimentacao where Idoperacao = " & TBLISTA!ID_estoque, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    If TBAbrir!Entrada > 0 Then DocumentoRef = "Ped. " & TBAbrir!LOTE Else DocumentoRef = TBAbrir!Documento
                End If
            Else
                Modulo_texto = "Financeiro"
                
                'Verifica número do documento e se a conta já foi paga
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select txt_NDocumento, Logsit from tbl_ContasPagar where Idintconta = " & IIf(IsNull(TBLISTA!ID_financeiro), 0, TBLISTA!ID_financeiro), Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    If TBAbrir!Logsit = "N" Then DocumentoRef = "Doc. " & TBAbrir!txt_ndocumento & " | N" Else DocumentoRef = "Doc. " & TBAbrir!txt_ndocumento & " | S"
                End If
            End If
            .Item(.Count).SubItems(4) = Modulo_texto
            .Item(.Count).SubItems(5) = DocumentoRef
            
            Select Case TBLISTA!Operacao
                Case "Crédito": ValorTexto = "-" & IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
                Case "Débito": ValorTexto = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
            End Select
            .Item(.Count).SubItems(6) = ValorTexto
        End With
        TBLISTA.MoveNext
        contador = contador + 1
        PBLista.Value = contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_ID = 0
Txt_data.Value = Date
Txt_responsavel = ""
Txt_modulo = ""
Cmb_centro.ListIndex = -1
Txt_referencia = ""
Txt_operacao = ""
Txt_valor = ""
CodigoLista = 0

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

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CC_realizado.* from CC_realizado INNER JOIN Usuarios_Setor ON Usuarios_Setor.ID = CC_realizado.ID_CC where CC_realizado.ID = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos
    ProcPuxaDados
    Frame2.Enabled = True
    CodigoLista = Lista.SelectedItem.index
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

Txt_ID = TBLISTA!ID
If IsNull(TBLISTA!Data) = False Then Txt_data = Format(TBLISTA!Data, "dd/mm/yy")
Txt_responsavel = IIf(IsNull(TBLISTA!Responsavel), "", TBLISTA!Responsavel)
Txt_modulo = Lista.SelectedItem.ListSubItems(4)
Txt_referencia = Lista.SelectedItem.ListSubItems(5)
Txt_operacao = TBLISTA!Operacao
Txt_valor = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptAteomes_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub OptDomes_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TabFiltro_Click()
On Error GoTo tratar_erro

ProcFiltrarMes

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrarMes()
On Error GoTo tratar_erro

M = FunVerificaMes(TabFiltro.SelectedItem.key)
If OptDomes.Value = True Then
    Familiatext = "Select CC_realizado.* from CC_realizado INNER JOIN Usuarios_Setor ON Usuarios_Setor.ID = CC_realizado.ID_CC where Usuarios_setor.ID = " & Frm_centro_de_custo.txtID & " and month(CC_realizado.Data) = '" & M & "' and Year(CC_realizado.Data) = '" & cmbAno & "' order by CC_realizado.data desc, CC_realizado.ID desc"
Else
    Familiatext = "Select CC_realizado.* from CC_realizado INNER JOIN Usuarios_Setor ON Usuarios_Setor.ID = CC_realizado.ID_CC where Usuarios_setor.ID = " & Frm_centro_de_custo.txtID & " and month(CC_realizado.Data) <= '" & M & "' and Year(CC_realizado.Data) = '" & cmbAno & "' order by CC_realizado.data desc, CC_realizado.ID desc"
End If
ProcCarregaLista Familiatext

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_Valor_Change()
On Error GoTo tratar_erro

If Txt_valor <> "" Then
    VerifNumero = Txt_valor
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_valor = ""
        Txt_valor.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_Valor_LostFocus()
On Error GoTo tratar_erro

Txt_valor = Format(Txt_valor, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
