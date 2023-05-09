VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frm_Boleto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CAPRIND - Títulos em carteira"
   ClientHeight    =   10005
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   15300
   ClipControls    =   0   'False
   Icon            =   "Frm_boleto.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10005
   ScaleWidth      =   15300
   WindowState     =   2  'Maximizado
   Begin VB.Frame Frame6 
      Caption         =   "Local para armazenamento do arquivo remessa"
      Height          =   645
      Left            =   10035
      TabIndex        =   54
      Top             =   2790
      Width           =   5190
      Begin VB.CommandButton cmdLocal 
         Caption         =   "..."
         Height          =   285
         Left            =   4500
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "Abrirl local de armazenamento dos arquivos de remessa."
         Top             =   270
         Width           =   510
      End
      Begin VB.TextBox Txtlocal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   270
         Width           =   4290
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Arquivo de configuração da carteira"
      Height          =   645
      Left            =   10035
      TabIndex        =   50
      Top             =   2125
      Width           =   5190
      Begin VB.CommandButton cmdArquivo 
         Caption         =   "..."
         Height          =   285
         Left            =   4500
         TabIndex        =   52
         TabStop         =   0   'False
         ToolTipText     =   "Abrirl local de armazenamento do arquivo de configuração da carteira."
         Top             =   270
         Width           =   510
      End
      Begin VB.TextBox txtcarteiraconf 
         Enabled         =   0   'False
         Height          =   285
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   270
         Width           =   4290
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Processar títulos (Ações)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   45
      TabIndex        =   43
      Top             =   9270
      Width           =   15180
      Begin VB.CheckBox chkAtualizar 
         Caption         =   "Atualizar boleto(s)"
         Enabled         =   0   'False
         Height          =   195
         Left            =   9675
         TabIndex        =   53
         ToolTipText     =   "Atualizar boleto(s) vencido(s)."
         Top             =   360
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.CommandButton cmdProcessar 
         Caption         =   "&Processar titulo(s)"
         Enabled         =   0   'False
         Height          =   420
         Left            =   13365
         TabIndex        =   48
         ToolTipText     =   "Processar o(s) título(s) selecionado(s) na carteira."
         Top             =   180
         Width           =   1635
      End
      Begin VB.CheckBox chkEmail 
         Caption         =   "Enviar boleto(s) por email"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5580
         TabIndex        =   47
         ToolTipText     =   "Enviar boleto por email para o cliente."
         Top             =   360
         Width           =   2085
      End
      Begin VB.CheckBox chkImprimir 
         Caption         =   "Visualizar boleto(s) para impressão"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2655
         TabIndex        =   46
         ToolTipText     =   "Visualizar boleto para impressão."
         Top             =   360
         Width           =   2895
      End
      Begin VB.CheckBox chkEmailcopia 
         Caption         =   "Enviar-me cópia"
         Enabled         =   0   'False
         Height          =   195
         Left            =   7875
         TabIndex        =   45
         ToolTipText     =   "Enviar uma cópia do boleto para meu email."
         Top             =   360
         Width           =   2085
      End
      Begin VB.CheckBox chkRemessa 
         Caption         =   "Gerar arquivo remessa"
         Enabled         =   0   'False
         Height          =   195
         Left            =   675
         TabIndex        =   44
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame FramePesquisa 
      Caption         =   "Carregar carteira de títulos"
      Enabled         =   0   'False
      Height          =   870
      Left            =   45
      TabIndex        =   35
      Top             =   3420
      Width           =   15180
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   240
         Left            =   4365
         OleObjectBlob   =   "Frm_boleto.frx":2B012
         TabIndex        =   58
         Top             =   405
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   240
         Left            =   2835
         OleObjectBlob   =   "Frm_boleto.frx":2B086
         TabIndex        =   57
         Top             =   405
         Width           =   240
      End
      Begin VB.CommandButton cmdRetorno 
         Caption         =   "Retorno"
         Height          =   420
         Left            =   13860
         TabIndex        =   49
         ToolTipText     =   "Receber arquivo retorno do banco."
         Top             =   270
         Width           =   1140
      End
      Begin VB.ComboBox cmbCliente 
         Height          =   315
         Left            =   5130
         TabIndex        =   42
         ToolTipText     =   "Escolha um cliente para pesquisa."
         Top             =   360
         Width           =   6225
      End
      Begin VB.CommandButton cmdProcessados 
         Caption         =   "Processados"
         Height          =   420
         Left            =   12645
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Filtrar título(s) processado(s)"
         Top             =   270
         Width           =   1140
      End
      Begin VB.CommandButton CmdAprocessar 
         Caption         =   "Á processar"
         Height          =   420
         Left            =   11430
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Filtrar título(s) não processado(s). "
         Top             =   270
         Width           =   1140
      End
      Begin MSComCtl2.DTPicker DTFim 
         Height          =   315
         Left            =   3090
         TabIndex        =   4
         ToolTipText     =   "Data de vencimento final para pesquisa."
         Top             =   360
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
         Format          =   136839169
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker DTINI 
         Height          =   315
         Left            =   1665
         TabIndex        =   3
         ToolTipText     =   "Data de vencimento de início para pesquisa."
         Top             =   360
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
         Format          =   136839171
         CurrentDate     =   39057
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   240
         Left            =   495
         OleObjectBlob   =   "Frm_boleto.frx":2B0E6
         TabIndex        =   41
         Top             =   405
         Width           =   6405
      End
      Begin ActiveResizeCtl.ActiveResize ActiveResize1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         Resolution      =   99
         ResizeFonts     =   0   'False
         ScreenHeight    =   768
         ScreenWidth     =   1360
         ScreenHeightDT  =   1080
         ScreenWidthDT   =   1920
         AutoResizeOnLoad=   0   'False
         ApplicationName =   "Active Resize Control Professional"
         FormHeightDT    =   10440
         FormWidthDT     =   15390
         FormScaleHeightDT=   10005
         FormScaleWidthDT=   15300
         ResizeFormBackground=   -1  'True
         ResizePictureBoxContents=   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Instruções á serem enviadas para o banco"
      Height          =   2040
      Left            =   10035
      TabIndex        =   25
      Top             =   90
      Width           =   5190
      Begin VB.CommandButton cmdSalvarInstrucoes 
         Caption         =   "Salvar"
         Height          =   285
         Left            =   4020
         TabIndex        =   59
         Top             =   540
         Width           =   915
      End
      Begin VB.TextBox txtAssunto 
         Height          =   330
         Left            =   225
         TabIndex        =   11
         Text            =   "Boleto Sistema Caprind"
         ToolTipText     =   "Assunto para email á ser enviado."
         Top             =   1620
         Width           =   4785
      End
      Begin VB.TextBox Txtpercentual_juros 
         Alignment       =   2  'Centralizar
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Text            =   "0,20"
         ToolTipText     =   "Percentual dos juros a serem aplicados por dia de atraso."
         Top             =   540
         Width           =   885
      End
      Begin VB.TextBox Txtdias_protesto 
         Alignment       =   2  'Centralizar
         Height          =   285
         Left            =   2835
         TabIndex        =   10
         Text            =   "30"
         ToolTipText     =   "Numero de dias do prazo antes do título ser protestado."
         Top             =   540
         Width           =   1185
      End
      Begin VB.TextBox Txtinstrucoes 
         Height          =   330
         Left            =   225
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "Frm_boleto.frx":2B164
         ToolTipText     =   "Instruções para o banco."
         Top             =   1080
         Width           =   4785
      End
      Begin VB.TextBox Txtpercentual_desconto 
         Alignment       =   2  'Centralizar
         Height          =   285
         Left            =   1140
         TabIndex        =   8
         Text            =   "0,00"
         ToolTipText     =   "Percentual de desconto a ser aplicado por dia de antecipação."
         Top             =   540
         Width           =   915
      End
      Begin VB.TextBox Txtpercentual_multa 
         Alignment       =   2  'Centralizar
         Height          =   285
         Left            =   2070
         TabIndex        =   9
         Text            =   "10,00"
         ToolTipText     =   "Percentual da multa a ser aplicado sobre o valor total do boleto."
         Top             =   540
         Width           =   750
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   240
         Left            =   315
         OleObjectBlob   =   "Frm_boleto.frx":2B183
         TabIndex        =   26
         Top             =   360
         Width           =   750
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   240
         Left            =   1995
         OleObjectBlob   =   "Frm_boleto.frx":2B1EF
         TabIndex        =   27
         Top             =   360
         Width           =   840
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   240
         Left            =   1050
         OleObjectBlob   =   "Frm_boleto.frx":2B25B
         TabIndex        =   28
         Top             =   360
         Width           =   1005
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   240
         Left            =   270
         OleObjectBlob   =   "Frm_boleto.frx":2B2CD
         TabIndex        =   29
         Top             =   900
         Width           =   2850
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   240
         Left            =   2925
         OleObjectBlob   =   "Frm_boleto.frx":2B359
         TabIndex        =   30
         Top             =   360
         Width           =   960
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   195
         Left            =   270
         OleObjectBlob   =   "Frm_boleto.frx":2B3D1
         TabIndex        =   31
         Top             =   1440
         Width           =   2040
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados instituição financeira"
      ForeColor       =   &H00000000&
      Height          =   3345
      Left            =   45
      TabIndex        =   13
      Top             =   90
      Width           =   9960
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   495
         OleObjectBlob   =   "Frm_boleto.frx":2B449
         Top             =   765
      End
      Begin VB.TextBox Txt_IDBanco 
         Alignment       =   2  'Centralizar
         Height          =   315
         Left            =   2670
         TabIndex        =   60
         Top             =   1140
         Width           =   855
      End
      Begin VB.TextBox txtCodigocedente 
         Alignment       =   2  'Centralizar
         Enabled         =   0   'False
         Height          =   285
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1140
      End
      Begin VB.PictureBox Logo_Banco 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Nenhum
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   270
         Picture         =   "Frm_boleto.frx":4DC1A
         ScaleHeight     =   750
         ScaleWidth      =   1500
         TabIndex        =   39
         Top             =   945
         Width           =   1500
      End
      Begin VB.ComboBox cmbCarteira 
         Height          =   315
         Left            =   2655
         TabIndex        =   0
         Top             =   1665
         Width           =   7125
      End
      Begin VB.TextBox txtNomecedente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2880
         Width           =   8430
      End
      Begin VB.ComboBox cmbempresa 
         Height          =   315
         Left            =   2655
         TabIndex        =   1
         Top             =   540
         Width           =   7125
      End
      Begin VB.TextBox txtNomeAgencia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2340
         Width           =   6090
      End
      Begin VB.ComboBox cmbBanco 
         Height          =   315
         Left            =   3525
         TabIndex        =   2
         Top             =   1135
         Width           =   6255
      End
      Begin VB.TextBox txtContacorrente 
         Alignment       =   2  'Centralizar
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1140
      End
      Begin VB.TextBox txtAgencia 
         Alignment       =   2  'Centralizar
         Enabled         =   0   'False
         Height          =   285
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1140
      End
      Begin VB.TextBox txtNBanco 
         Alignment       =   2  'Centralizar
         Enabled         =   0   'False
         Height          =   285
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1140
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   240
         Left            =   1395
         OleObjectBlob   =   "Frm_boleto.frx":516F4
         TabIndex        =   19
         Top             =   2160
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   240
         Left            =   2565
         OleObjectBlob   =   "Frm_boleto.frx":51760
         TabIndex        =   20
         Top             =   2160
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   240
         Left            =   180
         OleObjectBlob   =   "Frm_boleto.frx":517DA
         TabIndex        =   21
         Top             =   2700
         Width           =   1230
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   240
         Left            =   2700
         OleObjectBlob   =   "Frm_boleto.frx":51854
         TabIndex        =   22
         Top             =   945
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   240
         Left            =   3960
         OleObjectBlob   =   "Frm_boleto.frx":518C6
         TabIndex        =   23
         Top             =   2160
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   240
         Left            =   2700
         OleObjectBlob   =   "Frm_boleto.frx":5193C
         TabIndex        =   24
         Top             =   1485
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   240
         Left            =   225
         OleObjectBlob   =   "Frm_boleto.frx":519AA
         TabIndex        =   18
         Top             =   2160
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   240
         Left            =   2745
         OleObjectBlob   =   "Frm_boleto.frx":51A20
         TabIndex        =   33
         Top             =   360
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   240
         Left            =   1485
         OleObjectBlob   =   "Frm_boleto.frx":51A96
         TabIndex        =   38
         Top             =   2700
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   1455
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   540
         Width           =   2400
      End
      Begin VB.TextBox txtCodcedente 
         Alignment       =   2  'Centralizar
         Enabled         =   0   'False
         Height          =   285
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2880
         Width           =   690
      End
   End
   Begin MSComctlLib.ListView lst_Duplicata 
      Height          =   4905
      Left            =   45
      TabIndex        =   34
      Top             =   4320
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   8652
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Nota fiscal"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Sacado (Cliente)"
         Object.Width           =   10231
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "D"
         Text            =   "Vencimento"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Parcela"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Nosso número"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Remessa"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Env.Financeiro"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Env. email?"
         Object.Width           =   2293
      EndProperty
   End
End
Attribute VB_Name = "frm_Boleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StrCnpj As String
Public StrEndereco As String
Public StrBairro As String
Public StrCidade As String
Public StrEstado As String
Public StrCEP As String
Public StrEmailBoleto As String

Public Sub ProcBuscaArquivolicenca()
Agencia = txtAgencia
'Início dos parâmetros obrigatórios da ContaCorrente corrente
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Empresa where codigo = " & txtCodcedente, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Select Case txtNBanco
        Case "001": 'Banco do brasil
            Select Case cmbCarteira
                Case "11 - Simples - Com Registro":
                    txtcarteiraconf = TBAbrir!Registro_boleto & "-001-11.conf"
                    OutrosDadosConfiguracao1 = Carteira1
                Case "11 - Vinculada - Com Registro"
                     ArquivoLicensa = TBAbrir!Registro_boleto & "-001-11Vinculada.conf"
                     txtcarteiraconf = ArquivoLicensa
'                    OutrosDadosConfiguracao1 = txtCodigocedente
'                    OutrosDadosConfiguracao2 = txtCodigocedente
                Case "17 - Direta Especial - Com Registro":
                    txtcarteiraconf = TBAbrir!Registro_boleto & "-001-17.conf"
                    OutrosDadosConfiguracao1 = Carteira1
                Case "17Simples - Direta Especial Simples - Com Registro":
                    txtcarteiraconf = TBAbrir!Registro_boleto & "-001-17SIMPLES.conf"
                    OutrosDadosConfiguracao1 = Carteira1
                Case "17-7 - Direta Especial - Com Registro Convênio 7 dígitos":
                    txtcarteiraconf = TBAbrir!Registro_boleto & "-001-17-7.conf"
                    OutrosDadosConfiguracao1 = Carteira1
                    OutrosDadosConfiguracao2 = "0000000000"
                Case "18 - Simples - Sem Registro":
                    txtcarteiraconf = TBAbrir!Registro_boleto & "-001-18.conf"
                    OutrosDadosConfiguracao1 = Carteira1
                Case "18-7 - Simples - Sem Registro - Convênio 7 dígitos":
                    txtcarteiraconf = TBAbrir!Registro_boleto & "-001-18-7.conf"
                    OutrosDadosConfiguracao1 = Carteira1
            End Select
            Select Case Len(txtAgencia)
                Case 1: AgenciaBol = "0000-" & Agencia
                Case 2: AgenciaBol = "000" & Left(Agencia, 1) & "-" & Right(Agencia, 1)
                Case 3: AgenciaBol = "00" & Left(Agencia, 2) & "-" & Right(Agencia, 1)
                Case 4: AgenciaBol = "0" & Left(Agencia, 3) & "-" & Right(Agencia, 1)
                Case Is >= 5: AgenciaBol = Left(Agencia, 4) & "-" & Right(Agencia, 1)
            End Select
            Select Case Len(ContaCorrente)
                Case 1: ContaCorrenteBol = "00000000-" & ContaCorrente
                Case 2: ContaCorrenteBol = "0000000" & Left(ContaCorrente, 1) & "-" & Right(ContaCorrente, 1)
                Case 3: ContaCorrenteBol = "000000" & Left(ContaCorrente, 2) & "-" & Right(ContaCorrente, 1)
                Case 4: ContaCorrenteBol = "00000" & Left(ContaCorrente, 3) & "-" & Right(ContaCorrente, 1)
                Case 5: ContaCorrenteBol = "0000" & Left(ContaCorrente, 4) & "-" & Right(ContaCorrente, 1)
                Case 6: ContaCorrenteBol = "000" & Left(ContaCorrente, 5) & "-" & Right(ContaCorrente, 1)
                Case 7: ContaCorrenteBol = "00" & Left(ContaCorrente, 6) & "-" & Right(ContaCorrente, 1)
                Case 8: ContaCorrenteBol = "0" & Left(ContaCorrente, 7) & "-" & Right(ContaCorrente, 1)
                Case Is >= 9: ContaCorrenteBol = Left(ContaCorrente, 8) & "-" & Right(ContaCorrente, 1)
            End Select

            Debug.Print AgenciaBol
            
            
            If cmbCarteira = "17-7 - Direta Especial - Com Registro Convênio 7 dígitos" Or cmbCarteira = "18-7 - Simples - Sem Registro - Convênio 7 dígitos" Then
                'Codigocedente = FunTamanhoTextoZeroEsq(Left(Codigocedente, 7), 7)
            Else
                'Codigocedente = FunTamanhoTextoZeroEsq(Left(Codigocedente, 6), 6)
            End If
            
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Banco do brasil"
            Arquivo = "CBR" & Dia & Mes & "." & SeqRemessa
            Layout = "FEBRABAN240"
        Case "033": 'Santander
            If cmbCarteira = "CSR - Cobrança Simples Sem Registro" Or cmbCarteira = "ECR - Cobrança Simples Com Registro" Or cmbCarteira = "COBR-Nova - Cobrança Simples - Rápida Com Registro" Then
                Select Case cmbCarteira
                    Case "CSR - Cobrança Simples Sem Registro": txtcarteiraconf = TBAbrir!Registro_boleto & "-033-CSR.conf"
                    Case "ECR - Cobrança Simples Com Registro":
                        txtcarteiraconf = TBAbrir!Registro_boleto & "-033-ECR.conf"
                        'OutrosDadosConfiguracao1 = Left( txtAgencia, 4) & ProcTamanhoTextoZeroEsq(Codigocedente, 7) & Right(txtContaCorrente, 9) forma antiga com 11 digitos no nosso numero
                        OutrosDadosConfiguracao1 = Left(txtAgencia, 5) & FunTamanhoTextoZeroEsq(Codigocedente, 7) & Left(txtContacorrente, 9)
                    Case "COBR-Nova - Cobrança Simples - Rápida Com Registro":
                        txtcarteiraconf = TBAbrir!Registro_boleto & "-033-COBR-NOVA.conf"
                        OutrosDadosConfiguracao1 = Left(txtAgencia, 5) & FunTamanhoTextoZeroEsq(Codigocedente, 7) & Left(txtContacorrente, 9)
                End Select
                Select Case Len(txtAgencia)
                    Case 1: AgenciaBol = "0000-" & txtAgencia
                    Case 2: AgenciaBol = "000" & Left(txtAgencia, 1) & "-" & Right(txtAgencia, 1)
                    Case 3: AgenciaBol = "00" & Left(txtAgencia, 2) & "-" & Right(txtAgencia, 1)
                    Case 4: AgenciaBol = "0" & Left(txtAgencia, 3) & "-" & Right(txtAgencia, 1)
                    Case Is >= 5: AgenciaBol = Left(txtAgencia, 4) & "-" & Right(txtAgencia, 1)
                End Select
                Select Case Len(txtContacorrente)
                    Case 1: ContaCorrenteBol = "000000000-" & txtContacorrente
                    Case 2: ContaCorrenteBol = "00000000" & Left(txtContacorrente, 1) & "-" & Right(txtContacorrente, 1)
                    Case 3: ContaCorrenteBol = "0000000" & Left(txtContacorrente, 2) & "-" & Right(txtContacorrente, 1)
                    Case 4: ContaCorrenteBol = "000000" & Left(txtContacorrente, 3) & "-" & Right(txtContacorrente, 1)
                    Case 5: ContaCorrenteBol = "00000" & Left(txtContacorrente, 4) & "-" & Right(txtContacorrente, 1)
                    Case 6: ContaCorrenteBol = "0000" & Left(txtContacorrente, 5) & "-" & Right(txtContacorrente, 1)
                    Case 7: ContaCorrenteBol = "000" & Left(txtContacorrente, 6) & "-" & Right(txtContacorrente, 1)
                    Case 8: ContaCorrenteBol = "00" & Left(txtContacorrente, 7) & "-" & Right(txtContacorrente, 1)
                    Case 9: ContaCorrenteBol = "0" & Left(txtContacorrente, 8) & "-" & Right(txtContacorrente, 1)
                    Case Is >= 10: ContaCorrenteBol = Left(txtContacorrente, 9) & "-" & Right(txtContacorrente, 1)
                End Select
                Select Case Len(Codigocedente)
                    Case 1: Codigocedente = "000000-" & Codigocedente
                    Case 2: Codigocedente = "00000" & Left(Codigocedente, 1) & "-" & Right(Codigocedente, 1)
                    Case 3: Codigocedente = "0000" & Left(Codigocedente, 2) & "-" & Right(Codigocedente, 1)
                    Case 4: Codigocedente = "000" & Left(Codigocedente, 3) & "-" & Right(Codigocedente, 1)
                    Case 5: Codigocedente = "00" & Left(Codigocedente, 4) & "-" & Right(Codigocedente, 1)
                    Case 6: Codigocedente = "0" & Left(Codigocedente, 5) & "-" & Right(Codigocedente, 1)
                    Case Is >= 7: Codigocedente = Left(Codigocedente, 6) & "-" & Right(Codigocedente, 1)
                End Select
            Else
                Select Case cmbCarteira
                    Case "COB - Cobrança Simples": txtcarteiraconf = TBAbrir!Registro_boleto & "-033-COB.conf"
                    Case "COBR - Cobrança Simples - Rápida Com Registro": txtcarteiraconf = TBAbrir!Registro_boleto & "-033-COBR.conf"
                End Select
                AgenciaBol = Mid(txtAgencia, 2, 3)
                ContaCorrente = Codigocedente
                Select Case Len(txtContacorrente)
                    Case 1: ContaCorrenteBol = "00" & " " & "00000" & " " & txtContacorrente
                    Case 2: ContaCorrenteBol = "00" & " " & "0000" & Left(txtContacorrente, 1) & " " & Mid(txtContacorrente, 2, 1)
                    Case 3: ContaCorrenteBol = "00" & " " & "000" & Left(txtContacorrente, 2) & " " & Mid(txtContacorrente, 3, 1)
                    Case 4: ContaCorrenteBol = "00" & " " & "00" & Left(txtContacorrente, 3) & " " & Mid(txtContacorrente, 4, 1)
                    Case 5: ContaCorrenteBol = "00" & " " & "0" & Left(txtContacorrente, 4) & " " & Mid(txtContacorrente, 5, 1)
                    Case 6: ContaCorrenteBol = "00" & " " & Left(txtContacorrente, 5) & " " & Mid(txtContacorrente, 6, 1)
                    Case 7: ContaCorrenteBol = "0" & Left(txtContacorrente, 1) & " " & Mid(txtContacorrente, 2, 5) & " " & Mid(txtContacorrente, 7, 1)
                    Case Is >= 8: ContaCorrenteBol = Left(txtContacorrente, 2) & " " & Mid(txtContacorrente, 3, 5) & " " & Mid(txtContacorrente, 8, 1)
                End Select
                Codigocedente = FunTamanhoTextoVazioDir(Left(NomeAgencia, 20), 20)
            End If
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Santander"
            Arquivo = "DB" & Dia & Mes & Right(ano, 2) & "." & SeqRemessa
            Layout = "CNAB400"
        Case "104": 'Caixa
            If cmbCarteira = "SIG14 - SIG Com Registro - Emissão pelo Cedente" Then
                txtcarteiraconf = TBAbrir!Registro_boleto & "-104-SIG14.conf"
                AgenciaBol = FunTamanhoTextoZeroEsq(Left(txtAgencia, 4), 4)
                ContaCorrenteBol = ""
                Codigocedente = FunTamanhoTextoZeroEsq(Left(FunSóNumeros(txtCodigocedente), 6), 6)
                Layout = "SIGCB240"
            Else
                Select Case cmbCarteira
                    Case "CR - Cobrança Rápida": txtcarteiraconf = TBAbrir!Registro_boleto & "-104-CR.conf"
                    Case "SR - Cobrança Sem Registro": txtcarteiraconf = TBAbrir!Registro_boleto & "-104-SR.conf"
                End Select
                AgenciaBol = ""
                ContaCorrenteBol = ""
                Codigocedente = FunSóNumeros(txtCodigocedente)
                Codigocedente = Left(Codigocedente, 4) & "." & Mid(Codigocedente, 5, 3) & "." & Mid(Codigocedente, 8, 8) & "-" & Right(Codigocedente, 1)
                Layout = "CNAB400"
            End If
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Caixa"
            Arquivo = "CB" & Dia & Mes & "." & SeqRemessa
        Case "237": 'Bradesco
            Select Case cmbCarteira
                Case "06 - Sem Registro": txtcarteiraconf = TBAbrir!Registro_boleto & "-237-06.conf"
                Case "09 - Com Registro": txtcarteiraconf = TBAbrir!Registro_boleto & "-237-09.conf"
            End Select
            Select Case Len(txtAgencia)
                Case 1: AgenciaBol = "0000-" & txtAgencia
                Case 2: AgenciaBol = "000" & Left(txtAgencia, 1) & "-" & Right(txtAgencia, 1)
                Case 3: AgenciaBol = "00" & Left(txtAgencia, 2) & "-" & Right(txtAgencia, 1)
                Case 4: AgenciaBol = "0" & Left(txtAgencia, 3) & "-" & Right(txtAgencia, 1)
                Case Is >= 5: AgenciaBol = Left(txtAgencia, 4) & "-" & Right(txtAgencia, 1)
            End Select
            Select Case Len(txtContacorrente)
                Case 1: ContaCorrenteBol = "0000000-" & txtContacorrente
                Case 2: ContaCorrenteBol = "000000" & Left(txtContacorrente, 1) & "-" & Right(txtContacorrente, 1)
                Case 3: ContaCorrenteBol = "00000" & Left(txtContacorrente, 2) & "-" & Right(txtContacorrente, 1)
                Case 4: ContaCorrenteBol = "0000" & Left(txtContacorrente, 3) & "-" & Right(txtContacorrente, 1)
                Case 5: ContaCorrenteBol = "000" & Left(txtContacorrente, 4) & "-" & Right(txtContacorrente, 1)
                Case 6: ContaCorrenteBol = "00" & Left(txtContacorrente, 5) & "-" & Right(txtContacorrente, 1)
                Case 7: ContaCorrenteBol = "0" & Left(txtContacorrente, 6) & "-" & Right(txtContacorrente, 1)
                Case Is >= 8: ContaCorrenteBol = Left(txtContacorrente, 7) & "-" & Right(txtContacorrente, 1)
            End Select
            Codigocedente = FunTamanhoTextoZeroEsq(Left(Codigocedente, 15), 15)
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Bradesco"
            Arquivo = "CB" & Dia & Mes & "." & SeqRemessa
            Layout = "CNAB400"
        Case "341": 'Itaú
            Select Case cmbCarteira
                Case "109 - Direta Eletrônica Sem Emissão - Simples": txtcarteiraconf = TBAbrir!Registro_boleto & "-341-109.conf"
                Case "112 - Escritual Eletrônica - simples / contratual": txtcarteiraconf = TBAbrir!Registro_boleto & "-341-112.conf"
                Case "175 - Sem Registro Sem Emissão": txtcarteiraconf = TBAbrir!Registro_boleto & "-341-175.conf"
            End Select
            AgenciaBol = FunTamanhoTextoZeroEsq(Left(txtAgencia, 4), 4)
            Select Case Len(txtContacorrente)
                Case 1: ContaCorrenteBol = "00000-" & txtContacorrente
                Case 2: ContaCorrenteBol = "0000" & Left(txtContacorrente, 1) & "-" & Right(txtContacorrente, 1)
                Case 3: ContaCorrenteBol = "000" & Left(txtContacorrente, 2) & "-" & Right(txtContacorrente, 1)
                Case 4: ContaCorrenteBol = "00" & Left(txtContacorrente, 3) & "-" & Right(txtContacorrente, 1)
                Case 5: ContaCorrenteBol = "0" & Left(txtContacorrente, 4) & "-" & Right(txtContacorrente, 1)
                Case Is >= 6: ContaCorrenteBol = Left(txtContacorrente, 5) & "-" & Right(txtContacorrente, 1)
            End Select
            Codigocedente = txtContacorrente
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Itaú"
            'Arquivo = Dia & Mes & Right(Ano, 2) & Seq1
            Arquivo = Dia & Mes & Right(ano, 2) & SeqRemessa
            Layout = "CNAB400"
        Case "356": 'ABN e Real
            Select Case cmbCarteira
                Case "20 - Cobrança Simples": txtcarteiraconf = TBAbrir!Registro_boleto & "-356-20.conf"
            End Select
            AgenciaBol = FunTamanhoTextoZeroEsq(Left(txtAgencia, 4), 4)
            Select Case Len(txtContacorrente)
                Case 1: ContaCorrenteBol = "000000-" & txtContacorrente
                Case 2: ContaCorrenteBol = "00000" & Left(txtContacorrente, 1) & "-" & Right(txtContacorrente, 1)
                Case 3: ContaCorrenteBol = "0000" & Left(txtContacorrente, 2) & "-" & Right(txtContacorrente, 1)
                Case 4: ContaCorrenteBol = "000" & Left(txtContacorrente, 3) & "-" & Right(txtContacorrente, 1)
                Case 5: ContaCorrenteBol = "00" & Left(txtContacorrente, 4) & "-" & Right(txtContacorrente, 1)
                Case 6: ContaCorrenteBol = "0" & Left(txtContacorrente, 5) & "-" & Right(txtContacorrente, 1)
                Case Is > 7: ContaCorrenteBol = Left(txtContacorrente, 6) & "-" & Right(txtContacorrente, 1)
            End Select
            Codigocedente = FunTamanhoTextoZeroEsq(Left(Codigocedente, 9), 9)
            Diretorio = Localrel & "\Boletos\Arquivos remessa\ABN e Real"
            Arquivo = "CB" & Dia & Mes & "." & SeqRemessa
            Layout = "CNAB400"
        Case "399": 'HSBC
            Select Case cmbCarteira
                Case "CNR - Sem Registro": txtcarteiraconf = TBAbrir!Registro_boleto & "-399-CNR.conf"
            End Select
            Codigocedente = FunTamanhoTextoZeroEsq(Left(Codigocedente, 7), 7)
            Diretorio = Localrel & "\Boletos\Arquivos remessa\HSBC"
            Arquivo = "D" & Dia & Mes & ano & "." & SeqRemessa
            Layout = "CNAB400"
        Case "409": 'Unibanco
            Select Case cmbCarteira
                Case "Especial": txtcarteiraconf = TBAbrir!Registro_boleto & "-409-ESPECIAL.conf"
            End Select
            AgenciaBol = FunTamanhoTextoZeroEsq(Left(txtAgencia, 4), 4)
            Select Case Len(txtContacorrente)
                Case 1: ContaCorrenteBol = "000" & "." & "000" & "-" & txtContacorrente
                Case 2: ContaCorrenteBol = "000" & "." & "00" & Left(txtContacorrente, 1) & "-" & Mid(txtContacorrente, 2, 1)
                Case 3: ContaCorrenteBol = "000" & "." & "0" & Left(txtContacorrente, 2) & "-" & Mid(txtContacorrente, 3, 1)
                Case 4: ContaCorrenteBol = "000" & "." & Left(txtContacorrente, 3) & "-" & Mid(txtContacorrente, 4, 1)
                Case 5: ContaCorrenteBol = "00" & Left(txtContacorrente, 1) & "." & Mid(txtContacorrente, 2, 3) & "-" & Mid(txtContacorrente, 5, 1)
                Case 6: ContaCorrenteBol = "0" & Left(txtContacorrente, 2) & "." & Mid(txtContacorrente, 3, 3) & "-" & Mid(txtContacorrente, 6, 1)
                Case Is >= 7: ContaCorrenteBol = Left(txtContacorrente, 3) & "." & Mid(txtContacorrente, 4, 3) & "-" & Mid(txtContacorrente, 7, 1)
            End Select
            Codigocedente = ContaCorrente
            Diretorio = Localrel & "\Boletos\Arquivos remessa\Unibanco"
            Arquivo = "CBR" & Dia & Mes & "." & SeqRemessa
            Layout = "CNAB240"
    End Select
End If

End Sub


Private Sub ProcVerificaFinanceiro(Enviado As Boolean)
On Error GoTo tratar_erro


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcGerarNossoNumero()
On Error GoTo tratar_erro

'Verif. último nosso numero

TextoFiltro = "txt_Portador_Banco = '" & cmbBanco & "' and txt_Agencia = '" & txtAgencia & "' and txt_Conta = '" & txtContacorrente & "'"

Set TBCarteira = CreateObject("adodb.recordset")
TBCarteira.Open "Select * from tbl_Detalhes_Recebimento where " & TextoFiltro & " order by Nosso_numero desc", Conexao, adOpenKeyset, adLockOptimistic
If TBCarteira.EOF = False Then
    txtjuros = Format(TBCarteira!Juros, "###,##0.0000000")
    Txtmulta = Format(TBCarteira!Multa, "###,##0.00")
    txtdesconto = Format(TBCarteira!Desconto, "###,##0.0000000")
    Txt_dias_protesto = IIf(IsNull(TBCarteira!dias_protesto), "", TBCarteira!dias_protesto)
    Txt_instrucoes = IIf(IsNull(TBCarteira!Instrucoes), "", TBCarteira!Instrucoes)
    

If IsNull(TBCarteira!Nosso_numero) = False And TBCarteira!Nosso_numero <> "" Then
        Texto = TBCarteira!Nosso_numero + IIf((chkAtualizar.Value = 1), 0, 1) 'Se for atualização de boleto não soma
    Else
        Texto = "1"
    End If
Else
        Texto = "1"
End If


Select Case txtNBanco
    Case "001": 'Banco do brasil
        Select Case cmbCarteira.Text
            Case "11 - Simples - Com Registro":
                Texto = FunTamanhoTextoZeroEsq(Texto, 5)
                Especie = "DM"
            Case "11 - Vinculada - Com Registro":
                Texto = FunTamanhoTextoZeroEsq(Texto, 7)
                Especie = "DM"
            Case "17 - Direta Especial - Com Registro":
                Texto = FunTamanhoTextoZeroEsq(Texto, 5)
                Especie = "DM"
            Case "17Simples - Direta Especial Simples - Com Registro":
                Texto = FunTamanhoTextoZeroEsq(Texto, 5)
                Especie = "DM"
            Case "17-7 - Direta Especial - Com Registro Convênio 7 dígitos":
                Texto = FunTamanhoTextoZeroEsq(Texto, 10)
                Especie = "DM"
            Case "18 - Simples - Sem Registro":
                Texto = FunTamanhoTextoZeroEsq(Texto, 5)
                Especie = "RC"
            Case "18-7 - Simples - Sem Registro - Convênio 7 dígitos":
                Texto = FunTamanhoTextoZeroEsq(Texto, 10)
                Especie = "RC"
        End Select
   CobreBemX1.OutroDadoConfiguracao1 = "019"
   CobreBemX1.OutroDadoConfiguracao2 = "035778"
    Case "341": 'Itaú
        Select Case cmbCarteira.Text
            Case "109 - Direta Eletrônica Sem Emissão - Simples":
                Texto = FunTamanhoTextoZeroEsq(Texto, 8)
                Especie = "DM"
            Case "112 - Escritual Eletrônica - simples / contratual":
                Texto = FunTamanhoTextoZeroEsq(Texto, 8)
                Especie = "RC"
            Case "175 - Sem Registro Sem Emissão":
                Texto = FunTamanhoTextoZeroEsq(Texto, 8)
                Especie = "RC"
        End Select
        
End Select
Var = Texto
Var1 = Especie

TBCarteira.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub


Public Sub ProcNomeArquivo()
On Error GoTo tratar_erro

'Verifica o último sequencial no banco para gerar o próximo
Dia = Day(Date)
    If (Len(Dia) = 1) Then
    Dia = "0" & Dia
    End If

Mes = Month(Date)
    If (Len(Mes) = 1) Then
    Mes = "0" & Mes
    End If

ano = Year(Date)

    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Seq_remessa from tbl_Detalhes_Recebimento where txt_Portador_Banco = '" & frm_Boleto.cmbBanco & "' order by Seq_remessa desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If IsNull(TBAbrir!Seq_remessa) = False And TBAbrir!Seq_remessa <> "" Then Seq = TBAbrir!Seq_remessa + 1 Else Seq = 1
    End If
    TBAbrir.Close

    If Seq < 10 Then SeqRemessa = "0" & Seq & ".txt" Else SeqRemessa = Seq & ".txt"
    SeqRemessaTexto = Left(SeqRemessa, Len(SeqRemessa) - 4)
    Select Case Len(SeqRemessaTexto)
        Case 1: RemessaTexto = "0" & Right(SeqRemessaTexto, 1)
        Case 2: RemessaTexto = SeqRemessaTexto
        Case Is >= 3: RemessaTexto = Right(SeqRemessaTexto, 2)
    End Select

 If txtNBanco.Text = "017" Then 'Banco Itau
    Arquivo = Dia & Mes & Right(ano, 2) & RemessaTexto & ".txt"
    Layout = "CNAB400"
  End If
  
  If txtNBanco.Text = "001" Then 'Banco Itau
    Arquivo = "CBR" & Dia & Mes & "." & RemessaTexto & ".txt"
    Layout = "FEBRABAN240"
  End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcAtualizaNomeArquivo()
On Error GoTo tratar_erro

'Verifica o último sequencial no banco para gerar o próximo
Dia = Day(Date)
Mes = "0" & Month(Date)
ano = Year(Date)

    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Seq_remessa from tbl_Detalhes_Recebimento where txt_Portador_Banco = '" & frm_Boleto.cmbBanco & "' order by Seq_remessa desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If IsNull(TBAbrir!Seq_remessa) = False And TBAbrir!Seq_remessa <> "" Then Seq = TBAbrir!Seq_remessa + 1 Else Seq = 1
    End If
    TBAbrir.Close
    
    'O nome do arquivo remessa do Itaú só aceita no máximo 8 caracteres
    'seqremessa = Seq
    If Seq < 10 Then SeqRemessa = "0" & Seq & ".txt" Else SeqRemessa = Seq & ".txt"
    SeqRemessaTexto = Left(SeqRemessa, Len(SeqRemessa) - 4)
    Select Case Len(SeqRemessaTexto)
        Case 1: RemessaTexto = "0" & Right(SeqRemessaTexto, 1)
        Case 2: RemessaTexto = SeqRemessaTexto
        Case Is >= 3: RemessaTexto = Right(SeqRemessaTexto, 2)
    End Select
    Arquivo = Dia & Mes & Right(ano, 2) & RemessaTexto & ".txt"
    Layout = "CNAB400"
    CobreBemX1.ArquivoRemessa.Sequencia = Left(SeqRemessa, Len(SeqRemessa) - 4)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcPassadadosEmailCopiaParaCobrebemX1()
On Error GoTo tratar_erro

    If chkEmailcopia.Value = 1 Then
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.CopiaReply = True
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailReply.Endereco = EmailCopia
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailReply.Nome = txtAssunto.Text
    End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcPassadadosEmailParaCobrebemX1()
On Error GoTo tratar_erro


    CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.LayoutBoleto = "Padrao"
    
    'Utilizar para sair o endereço em outro campo
    
    CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.LayoutBoleto = "PadraoReciboPersonalizado"
    CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.HTMLReciboPersonalizado = TxtHTLM
    CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.MargemSuperior = 0
    CobreBemX1.LocalPagamento = "Preferencialmente nas Casas Lotéricas até o valor limite"

    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select EE.*, E.Empresa from Empresa E INNER JOIN Empresa_email EE ON EE.ID_empresa = E.Codigo where EE.ID_empresa = " & txtCodcedente & " and EE.Aplicacao = 'F'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
       'Início da configuração dos dados do Cedente para envio de boletos por email
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.SMTP.Servidor = TBAbrir!Servidor_SMTP ' Trocar para apontar para o seu servidor SMTP
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.SMTP.Porta = TBAbrir!Porta
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.SMTP.Usuario = TBAbrir!Usuario 'utilizar esta propriedade para acesso a servidores SMTP seguros
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.SMTP.Senha = TBAbrir!Senha 'utilizar esta propriedade para acesso a servidores SMTP seguros
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.URLImagensCodigoBarras = "http://www.bptob.com/imagenscbe/"
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.URLLogotipo = "http://www.thisf.com.br/banners/BannerCBE.gif"
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.Assunto = "Boleto Caprind" 'Assunto_email
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailFrom.Endereco = TBAbrir!Email
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailFrom.Nome = TBAbrir!Nome
        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.FormaEnvio = IIf(Tipo_Documento = "PDF", feeSMTPMensagemBoletoPDFAnexo, feeSMTPBoletoHTML)
    End If

'Logotipo do cedente na parte superior do boleto
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Logotipo from Empresa where Codigo = " & txtCodcedente & " and Logotipo <> 'Null'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    If TBAbrir!Logotipo <> "" Then CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.ArquivoLogotipo = TBAbrir!Logotipo
End If
TBAbrir.Close
       
CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.CaminhoImagensCodigoBarras = Localrel & "\Imagens\Bancos\"

'Utilize o parâmetro abaixo para efetuar ajustes na impressão do boleto subindo ou descendo o mesmo na folha de papel
'Os valores devem ser informados em milímetros e quanto maior o valor mais para baixo será iniciado o boleto
'Se este parâmetro não for passado será assumido o valor 15 que é o indicado para a maioria das impressoras Jato de Tinta }
CobreBemX1.PadroesBoleto.PadroesBoletoImpresso.MargemSuperior = 3

'A próxima linha é utilizada para exibir um texto do lado direito do logotipo nos boletos impressos ou enviados por email
'CobreBemX1.PadroesBoleto.IdentificacaoCedente =

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcPassaDadosContaCorrenteParaCobrebemX1(Carteira As String, Carteira1 As String, Codigocedente As String, ID_empresa As Integer, EmitirBoleto As Boolean, Assunto_email As String)
On Error GoTo tratar_erro

Diretorio = Txtlocal
ContaCorrente = txtContacorrente

If txtNBanco.Text = "053" Then
        
    Select Case Len(ContaCorrente)
        Case 1: ContaCorrenteBol = "00000-" & ContaCorrente
        Case 2: ContaCorrenteBol = "0000" & Left(ContaCorrente, 1) & "-" & Right(ContaCorrente, 1)
        Case 3: ContaCorrenteBol = "000" & Left(ContaCorrente, 2) & "-" & Right(ContaCorrente, 1)
        Case 4: ContaCorrenteBol = "00" & Left(ContaCorrente, 3) & "-" & Right(ContaCorrente, 1)
        Case 5: ContaCorrenteBol = "0" & Left(ContaCorrente, 4) & "-" & Right(ContaCorrente, 1)
        Case Is >= 6: ContaCorrenteBol = Left(ContaCorrente, 5) & "-" & Right(ContaCorrente, 1)
    End Select
End If

If txtNBanco.Text = "001" Then
        
            Select Case Len(txtAgencia)
                Case 1: AgenciaBol = "0000-" & txtAgencia
                Case 2: AgenciaBol = "000" & Left(txtAgencia, 1) & "-" & Right(txtAgencia, 1)
                Case 3: AgenciaBol = "00" & Left(txtAgencia, 2) & "-" & Right(txtAgencia, 1)
                Case 4: AgenciaBol = "0" & Left(txtAgencia, 3) & "-" & Right(txtAgencia, 1)
                Case Is >= 5: AgenciaBol = Left(txtAgencia, 4) & "-" & Right(txtAgencia, 1)
            End Select
            Select Case Len(ContaCorrente)
                Case 1: ContaCorrenteBol = "00000000-" & ContaCorrente
                Case 2: ContaCorrenteBol = "0000000" & Left(ContaCorrente, 1) & "-" & Right(ContaCorrente, 1)
                Case 3: ContaCorrenteBol = "000000" & Left(ContaCorrente, 2) & "-" & Right(ContaCorrente, 1)
                Case 4: ContaCorrenteBol = "00000" & Left(ContaCorrente, 3) & "-" & Right(ContaCorrente, 1)
                Case 5: ContaCorrenteBol = "0000" & Left(ContaCorrente, 4) & "-" & Right(ContaCorrente, 1)
                Case 6: ContaCorrenteBol = "000" & Left(ContaCorrente, 5) & "-" & Right(ContaCorrente, 1)
                Case 7: ContaCorrenteBol = "00" & Left(ContaCorrente, 6) & "-" & Right(ContaCorrente, 1)
                Case 8: ContaCorrenteBol = "0" & Left(ContaCorrente, 7) & "-" & Right(ContaCorrente, 1)
                Case Is >= 9: ContaCorrenteBol = Left(ContaCorrente, 8) & "-" & Right(ContaCorrente, 1)
            End Select
End If
Debug.Print ContaCorrenteBol
Debug.Print AgenciaBol

ArquivoLicensa = txtcarteiraconf
CobreBemX1.ArquivoLicenca = Localrel & "\Boletos\Carteiras\" & ArquivoLicensa

If CobreBemX1.ArquivoLicenca = "" Then
    txtcarteiraconf.Text = ""
    MsgBox "Arquivo " & ArquivoLicensa & " de configuração da carteira " & cmbCarteira & " do banco " & cmbBanco & ", não foi encontrado na pasta de carteiras do banco, favor verificar antes de prosseguir", vbExclamation
    chkRemessa.Value = 0
    chkRemessa.Enabled = False
    chkEmail.Value = 0
    chkEmail.Enabled = False
    chkEmailcopia.Value = 0
    chkEmailcopia.Enabled = False
    chkImprimir.Value = 0
    chkImprimir.Enabled = False
    Exit Sub
End If

CobreBemX1.CodigoAgencia = AgenciaBol
CobreBemX1.NumeroContaCorrente = ContaCorrenteBol
CobreBemX1.Codigocedente = "035778"
CobreBemX1.OutroDadoConfiguracao1 = "019"
CobreBemX1.OutroDadoConfiguracao2 = "019"

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcPassaDadosBoletosParaCobrebemX1()
On Error GoTo tratar_erro
Dim Boleto As Object
Dim Email As Object

Set TBAbrir = CreateObject("adodb.recordset")
Set TBContas = CreateObject("adodb.recordset")

CobreBemX1.DocumentosCobranca.Clear
   
    With frm_Boleto.lst_Duplicata
        For InitFor = 1 To .ListItems.Count
        Contador = .ListItems.Count
        Contador = 100 / Contador
            If .ListItems.Item(InitFor).Checked = True Then
                TBContas.Open "Select * from tbl_contas_receber where IdIntConta = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    ProcCarregadadosSacado (TBContas!Nome_Razao)
                    TBAbrir.Open "Select * from tbl_Detalhes_Recebimento where IdContaReceber = " & TBContas!IDintconta & "", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                                            
                                Set Boleto = CobreBemX1.DocumentosCobranca.Add
                                Boleto.TipoDocumentoCobranca = Especie
                                Boleto.NumeroDocumento = Right(TBContas!NFiscal, 6) & "/" & Left(TBContas!Parcela, 3)
                                Boleto.NomeSacado = TBContas!Nome_Razao
                                Boleto.CNPJSacado = StrCnpj
                                Boleto.EnderecoSacado = StrEndereco
                                Boleto.BairroSacado = StrBairro
                                Boleto.CidadeSacado = StrCidade
                                Boleto.EstadoSacado = StrEstado
                                Boleto.CepSacado = StrCEP
                                Boleto.DataDocumento = Format$(Date, "dd/mm/yyyy")
                                Boleto.DataVencimento = TBAbrir!dt_Vencimento
                                Boleto.DataProcessamento = Format$(Date, "dd/mm/yyyy")
                                Boleto.ValorDocumento = TBAbrir!dbl_valor
                                Boleto.PercentualJurosDiaAtraso = Txtpercentual_juros
                                Boleto.PercentualMultaAtraso = Txtpercentual_multa
                                Boleto.PercentualDesconto = Txtpercentual_desconto
                                Boleto.ValorOutrosAcrescimos = IIf(IsNull(TBAbrir!Acrescimos), 0, TBAbrir!Acrescimos)
                                Boleto.PadroesBoleto.Demonstrativo = Right(TBContas!NFiscal, 6) & "/" & Left(TBContas!Parcela, 3)
                                Boleto.PadroesBoleto.InstrucoesCaixa = Txtinstrucoes
                                Boleto.ControleProcessamentoDocumento.Imprime = scpExecutar
                                Boleto.ControleProcessamentoDocumento.EnviaEmail = scpExecutar
                                Boleto.ControleProcessamentoDocumento.GravaRemessa = scpExecutar
                                ProcGerarNossoNumero
                                Boleto.NossoNumero = Var
                                
                            'Passa dados boleto para o detalhe recebimento
                                TBAbrir!Seq_remessa = IIf(Seq > 0, Seq, 1)
                                TBAbrir!Nosso_numero = Var
                                TBAbrir!Juros = Txtpercentual_juros
                                TBAbrir!Multa = Txtpercentual_multa
                                TBAbrir!Desconto = Txtpercentual_desconto
                                TBAbrir!Instrucoes = Txtinstrucoes
                                TBAbrir!dias_protesto = Txtdias_protesto
                                TBAbrir!Carteira = cmbCarteira
                                TBAbrir!Data_emissao = Date
                                TBAbrir!Vencimento_boleto = Boleto.DataVencimento
                                TBAbrir!Valor_boleto = Boleto.ValorDocumento
                                TBAbrir!Numero_documento = Right(TBContas!NFiscal, 6) & "/" & Left(TBContas!Parcela, 3)
                                TBContas!txt_ndocumento = Right(TBContas!NFiscal, 6) & "/" & Left(TBContas!Parcela, 3)
                                TBContas.Update
                            'Envia email
                                    If chkEmail.Value = 1 Then
                                            Set TBFIltro = CreateObject("adodb.recordset")
                                                TBFIltro.Open "Select * from Clientes_Contatos where IDCliente = " & TBContas!IDCliente & " and Enviar_boleto = 'TRUE'", Conexao, adOpenKeyset, adLockOptimistic
                                             If TBFIltro.EOF = False Then
                                                Do While TBFIltro.EOF = False
                                                    Set Email = Boleto.EnderecosEmailSacado.Add
                                                        Email.Nome = Boleto.NomeSacado
                                                        Email.Endereco = TBFIltro!Email
                                                        CobreBemX1.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.FormaEnvio = IIf(Tipo_Documento = "PDF", feeSMTPMensagemBoletoPDFAnexo, feeSMTPBoletoHTML)
                                                TBFIltro.MoveNext
                                                Loop
                                             Else
                                                 Email.Endereco = "caprind@caprind.com.br"
                                             End If
                                        TBAbrir!Enviado = True
                                    End If
                                        TBAbrir!data_envio = Date
                                TBAbrir.Update
                                'Imprimir boleto(s)
                                    If chkImprimir.Value = 1 Then
                                        CobreBemX1.ImprimeBoletos
                                    End If

                            'Passa dados boleto para o N_Boleto
                               Set TBBoleto = CreateObject("adodb.recordset")
                                    TBBoleto.Open "Select * from tbl_Detalhes_Recebimento_Nboletos where IDContaReceber = " & TBContas!IDintconta & " and Nosso_numero = '" & Var & "'", Conexao, adOpenKeyset, adLockOptimistic
                                    If TBBoleto.EOF = True Then TBBoleto.AddNew
                                    TBBoleto!data = Date
                                    TBBoleto!Responsavel = "Emissor boleto Caprind" 'pubUsuario
                                    TBBoleto!IdContaReceber = TBContas!IDintconta
                                    TBBoleto!Nosso_numero = Var
                                    TBBoleto!ID_nota = TBContas!ID_nota
                                    TBBoleto.Update
                                    TBBoleto.Close
                                TBContas.Close
                        End If
                        TBAbrir.Close
                End If
            End If
        Next InitFor
    End With
    
   

If chkEmail.Value = 1 Then
    CobreBemX1.EnviaBoletosPorEmail
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaDadosBoletosParaCobrebemX1()
On Error GoTo tratar_erro
     
Set TBAbrir = CreateObject("adodb.recordset")
Set TBContas = CreateObject("adodb.recordset")

CobreBemX1.DocumentosCobranca.Clear

   
    With frm_Boleto.lst_Duplicata
        For InitFor = 1 To .ListItems.Count
        Contador = .ListItems.Count
        Contador = 100 / Contador
            If .ListItems.Item(InitFor).Checked = True Then
                TBContas.Open "Select * from tbl_contas_receber where IdIntConta = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    ProcCarregadadosSacado (TBContas!Nome_Razao)
                    TBAbrir.Open "Select * from tbl_Detalhes_Recebimento where IdContaReceber = " & TBContas!IDintconta & "", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            If chkAtualizar.Value = 1 Then
                                CobreBemX1.ArquivoRemessa.Arquivo = .ListItems(InitFor).ListSubItems.Item(6).Text
                            End If
                            'frm_Atualiza_boleto.Show 1
                
                            ProcGerarNossoNumero
                            
                               Set Boleto = CobreBemX1.DocumentosCobranca.Add
                                Boleto.TipoDocumentoCobranca = Especie
                                Boleto.NumeroDocumento = Right(TBContas!NFiscal, 6) & "/" & Left(TBContas!Parcela, 3)
                                Boleto.NomeSacado = TBContas!Nome_Razao
                                Boleto.CNPJSacado = StrCnpj
                                Boleto.EnderecoSacado = StrEndereco
                                Boleto.BairroSacado = StrBairro
                                Boleto.CidadeSacado = StrCidade
                                Boleto.EstadoSacado = StrEstado
                                Boleto.CepSacado = StrCEP
                                Boleto.DataDocumento = Format$(Date, "dd/mm/yyyy")
                                Boleto.DataVencimento = IIf((chkAtualizar.Value = 1), NovoVencimento, TBAbrir!dt_Vencimento)
                                Boleto.DataProcessamento = Format$(Date, "dd/mm/yyyy")
                                Boleto.ValorDocumento = IIf((chkAtualizar.Value = 1), VA, TBAbrir!dbl_valor)
                                Boleto.PercentualJurosDiaAtraso = Txtpercentual_juros
                                Boleto.PercentualMultaAtraso = Txtpercentual_multa
                                Boleto.PercentualDesconto = Txtpercentual_desconto
                                Boleto.ValorOutrosAcrescimos = IIf(IsNull(TBAbrir!Acrescimos), 0, TBAbrir!Acrescimos)
                                Boleto.PadroesBoleto.Demonstrativo = Right(TBContas!NFiscal, 6) & "/" & Left(TBContas!Parcela, 3)
                                Boleto.PadroesBoleto.InstrucoesCaixa = Txtinstrucoes
                                Boleto.ControleProcessamentoDocumento.Imprime = scpExecutar
                                Boleto.ControleProcessamentoDocumento.EnviaEmail = scpExecutar
                                Boleto.ControleProcessamentoDocumento.GravaRemessa = scpExecutar
                                Boleto.NossoNumero = Var
                            'Passa dados boleto para o detalhe recebimento
                                If chkAtualizar.Value = 0 Then TBAbrir!Seq_remessa = IIf(Seq > 0, Seq, 1)
                                TBAbrir!Nosso_numero = Var
                                TBAbrir!Juros = Txtpercentual_juros
                                TBAbrir!Multa = Txtpercentual_multa
                                TBAbrir!Desconto = Txtpercentual_desconto
                                TBAbrir!Instrucoes = Txtinstrucoes
                                TBAbrir!dias_protesto = Txtdias_protesto
                                TBAbrir!Carteira = cmbCarteira
                                TBAbrir!Data_emissao = Date
                                TBAbrir!Vencimento_boleto = Boleto.DataVencimento
                                TBAbrir!Valor_boleto = Boleto.ValorDocumento
                                TBContas!txt_ndocumento = Right(TBContas!NFiscal, 6) & "/" & Left(TBContas!Parcela, 3)
                                TBContas.Update
                                 'Envia email
                                    If chkEmail.Value = 1 Then
                                            Set TBFIltro = CreateObject("adodb.recordset")
                                                TBFIltro.Open "Select * from Clientes_Contatos where IDCliente = " & TBContas!IDCliente & " and Enviar_boleto = 'TRUE'", Conexao, adOpenKeyset, adLockOptimistic
                                             If TBFIltro.EOF = False Then
                                                Do While TBFIltro.EOF = False
                                                    Set Email = Boleto.EnderecosEmailSacado.Add
                                                        Email.Nome = Boleto.NomeSacado
                                                        Email.Endereco = TBFIltro!Email
                                                TBFIltro.MoveNext
                                                Loop
                                             Else
                                                 Email.Endereco = "caprind@caprind.com.br"
                                             End If
                                        TBAbrir!Enviado = True
                                    End If
                                        TBAbrir!data_envio = Date
                                TBAbrir.Update
                                'Imprimir boleto(s)
                                    If chkImprimir.Value = 1 Then
                                        CobreBemX1.ImprimeBoletos
                                    End If

                            'Passa dados boleto para o N_Boleto
                               Set TBBoleto = CreateObject("adodb.recordset")
                                    TBBoleto.Open "Select * from tbl_Detalhes_Recebimento_Nboletos where IDContaReceber = " & TBContas!IDintconta & " and Nosso_numero = '" & Var & "'", Conexao, adOpenKeyset, adLockOptimistic
                                    If TBBoleto.EOF = True Then TBBoleto.AddNew
                                    TBBoleto!data = Date
                                    TBBoleto!Responsavel = "Emissor boleto Caprind" 'pubUsuario
                                    TBBoleto!IdContaReceber = TBContas!IDintconta
                                    TBBoleto!Nosso_numero = Var
                                    TBBoleto!ID_nota = TBContas!ID_nota
                                    TBBoleto.Update
                                    TBBoleto.Close
                                TBContas.Close
                        End If
                        TBAbrir.Close
                End If
            End If
        Next InitFor
    End With
    
   

If chkEmail.Value = 1 Then
    CobreBemX1.EnviaBoletosPorEmail
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub


Private Sub chkAtualizar_Click()
On Error GoTo tratar_erro

    If chkAtualizar.Value = 1 Then
    chkRemessa.Value = 0
    Else
    chkRemessa.Value = 1
    End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub chkEmail_Click()
On Error GoTo tratar_erro

If chkEmail.Value = 1 Then
ProcPassadadosEmailParaCobrebemX1
chkEmailcopia.Enabled = chkEmail.Value
cmdProcessar.Enabled = chkEmail.Value
Else
chkEmailcopia.Enabled = chkEmail.Value
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub chkEmailcopia_Click()
On Error GoTo tratar_erro

    If chkEmailcopia.Value = 1 Then
    ProcPassadadosEmailCopiaParaCobrebemX1
    End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub chkImprimir_Click()
On Error GoTo tratar_erro

If chkImprimir.Value = 1 Then
    cmdProcessar.Enabled = chkImprimir.Value
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub chkRemessa_Click()
On Error GoTo tratar_erro

    If chkRemessa.Value = 1 Then
    cmdProcessar.Enabled = chkRemessa.Value
    chkEmail.Enabled = True
    chkImprimir.Enabled = True
    chkImprimir.Value = 0
    End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbCarteira_Change()
On Error GoTo tratar_erro

    ProcBuscaArquivolicenca
    ProcPassaDadosContaCorrenteParaCobrebemX1 cmbCarteira, "", txtCodigocedente, txtCodcedente, True, ""
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbCarteira_Click()
On Error GoTo tratar_erro

    ProcBuscaArquivolicenca
    Debug.Print Arquivo
    ProcPassaDadosContaCorrenteParaCobrebemX1 cmbCarteira, "", txtCodigocedente, txtCodcedente, True, ""
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdArquivo_Click()
On Error GoTo tratar_erro

    DS.DS_AbrirPasta Localrel & "\Boletos\Carteiras\"

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub CmdAprocessar_Click()
On Error GoTo tratar_erro
Dataini = DTINI.Value
Datafim = DTFim.Value
'----------------
If cmbBanco.Text <> "" And cmbCliente.Text <> "" Then
    StrSql = "SELECT TOP (100) PERCENT dbo.tbl_Detalhes_Recebimento.Enviado,dbo.tbl_Detalhes_Recebimento.Data_envio,dbo.tbl_Detalhes_Recebimento.seq_remessa, dbo.tbl_Detalhes_Recebimento.IDContaReceber,dbo.tbl_Detalhes_Recebimento.txt_Cond_Recebimento, dbo.tbl_Detalhes_Recebimento.Id," _
    & "dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco,dbo.tbl_Detalhes_Recebimento.dt_Vencimento," _
    & "dbo.tbl_Detalhes_Recebimento.txt_tipoPagto, dbo.tbl_Detalhes_Recebimento.dbl_Valor," _
    & "dbo.tbl_Detalhes_Recebimento.int_NotaFiscal,dbo.tbl_Detalhes_Recebimento.txt_parcela, dbo.tbl_Detalhes_Recebimento.Nosso_numero, dbo.tbl_Detalhes_Recebimento.Carteira, dbo.tbl_Detalhes_Recebimento.Data_emissao,dbo.tbl_contas_receber.Nome_Razao FROM dbo.tbl_Detalhes_Recebimento" _
    & " INNER JOIN dbo.tbl_contas_receber ON dbo.tbl_Detalhes_Recebimento.IDContaReceber = dbo.tbl_contas_receber.IDIntconta" _
    & " WHERE (dbo.tbl_contas_receber.Nome_Razao = '" & cmbCliente & "') AND (dbo.tbl_Detalhes_Recebimento.txt_tipoPagto = N'BOLETO') AND (dbo.tbl_Detalhes_Recebimento.Nosso_numero IS NULL) AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento >= '" & Dataini & "') AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento <= '" & Datafim & "') AND (dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco = '" & cmbBanco.Text & "') ORDER BY dbo.tbl_Detalhes_Recebimento.dt_Vencimento"
    ProcCarregaListaDuplicatas
Else
    StrSql = "SELECT TOP (100) PERCENT dbo.tbl_Detalhes_Recebimento.Enviado,dbo.tbl_Detalhes_Recebimento.Data_envio,dbo.tbl_Detalhes_Recebimento.seq_remessa, dbo.tbl_Detalhes_Recebimento.IDContaReceber,dbo.tbl_Detalhes_Recebimento.txt_Cond_Recebimento, dbo.tbl_Detalhes_Recebimento.Id," _
    & "dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco,dbo.tbl_Detalhes_Recebimento.dt_Vencimento," _
    & "dbo.tbl_Detalhes_Recebimento.txt_tipoPagto, dbo.tbl_Detalhes_Recebimento.dbl_Valor," _
    & "dbo.tbl_Detalhes_Recebimento.int_NotaFiscal,dbo.tbl_Detalhes_Recebimento.txt_parcela, dbo.tbl_Detalhes_Recebimento.Nosso_numero, dbo.tbl_Detalhes_Recebimento.Carteira, dbo.tbl_Detalhes_Recebimento.Data_emissao,dbo.tbl_contas_receber.Nome_Razao FROM dbo.tbl_Detalhes_Recebimento" _
    & " INNER JOIN dbo.tbl_contas_receber ON dbo.tbl_Detalhes_Recebimento.IDContaReceber = dbo.tbl_contas_receber.IDIntconta" _
    & " WHERE (dbo.tbl_Detalhes_Recebimento.txt_tipoPagto = N'BOLETO') AND (dbo.tbl_Detalhes_Recebimento.Nosso_numero IS NULL) AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento >= '" & Dataini & "') AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento <= '" & Datafim & "') AND (dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco = '" & cmbBanco.Text & "') ORDER BY dbo.tbl_Detalhes_Recebimento.dt_Vencimento"
    ProcCarregaListaDuplicatas
End If

chkRemessa.Visible = True
chkAtualizar.Visible = False
chkAtualizar.Value = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbBanco_Change()
On Error GoTo tratar_erro
   
    ProcCarregaInstituicaoBoleto
    ProcCarregacomboCarteira

    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbbanco_Click()
On Error GoTo tratar_erro
   
    ProcCarregaInstituicaoBoleto
    ProcCarregacomboCarteira

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbempresa_Click()
On Error GoTo tratar_erro
    
    ProcCarregaComboBancoBoleto
    ProcCarregadadosCedente

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub


Private Sub cmdProcessar_Click()
On Error GoTo tratar_erro

If cmbCarteira.Text = "" Then
USMsgBox "Favor informar a certeira a ser utilizada", vbInformation, "CAPRIND"
Exit Sub
End If


If MsgBox("Deseja realmente processar o(s) título(s) selecionado(s)?", vbYesNo) = vbNo Then Exit Sub

    'Verifica se existe o arquivo de configuração da carteira
    If DS_ArquivoExiste(Localrel & "\Boletos\Carteiras\" & txtcarteiraconf) = False Then
        MsgBox ("Não será possível gerar o arquivo remessa, pois não foi encontrado o arquivo licença " & ArquivoLicensa & " na pasta " & Localrel & "\Boletos\Carteiras."), vbExclamation
        Exit Sub
    End If
    
    If chkRemessa.Value = 1 Then
    'Atribui um nome para o arquivo remessa
    ProcNomeArquivo

    'Passa dados para salvar aquivo remessa
    CobreBemX1.ArquivoRemessa.Diretorio = Txtlocal
    CobreBemX1.ArquivoRemessa.Arquivo = Arquivo
    
    'Passa dados do layout do boleto
    CobreBemX1.ArquivoRemessa.Layout = Layout
    End If
    
    'Se for novo boleto
    'If chkAtualizar.Value = 0 Then
    ProcPassaDadosBoletosParaCobrebemX1
    'End If
    
    'Se for atualizar boleto
    'If chkAtualizar.Value = 1 Then
    'ProcAtualizaDadosBoletosParaCobreBemX1
    'End If
    
    'Verifica se é pra gravar arquivo remessa
    If chkRemessa.Value = 1 And chkAtualizar.Value = 0 Then CobreBemX1.GravaArquivoRemessa
    
    MsgBox "Arquivo(s) processado(s) com sucesso"
    'Atualiza a lista de titulos
    ProcCarregaListaDuplicatas

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCarregadadosSacado(NomeRazao As String)
On Error GoTo tratar_erro

Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select * from Clientes where NomeRazao = '" & NomeRazao & "'", Conexao, adOpenKeyset, adLockReadOnly
        If TBClientes.EOF = False Then
            StrCnpj = TBClientes!CPF_CNPJ
            StrEndereco = TBClientes!Tipo_endereco
            StrEndereco = StrEndereco & " " & TBClientes!Endereco & ", N° " & TBClientes!Numero
            StrEndereco = StrConv(StrEndereco, vbProperCase)
            StrBairro = StrConv(TBClientes!Bairro, vbProperCase)
            StrCidade = StrConv(TBClientes!Cidade, vbProperCase)
            StrEstado = TBClientes!UF
            StrCEP = DS.DS_RetornarNumeros(TBClientes!CEP)
            Tipo_Documento = IIf(IsNull(TBClientes!Tipo_doc) = False, TBClientes!Tipo_doc, "HTML")
        End If
    TBClientes.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCarregadadosCedente()
On Error GoTo tratar_erro

   Set TBAbrir = CreateObject("adodb.recordset")
    
    TBAbrir.Open "Select Codigo, Razao,email from Empresa where razao = '" & cmbempresa & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        txtCodcedente = TBAbrir!CODIGO
        txtNomecedente = TBAbrir!Razao
        EmailCopia = TBAbrir!Email
    End If
            
    TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdLocal_Click()
On Error GoTo tratar_erro

    DS.DS_AbrirPasta Txtlocal.Text

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub


Private Sub CmdProcessados_Click()
On Error GoTo tratar_erro
'---------------------------------
If cmbBanco <> "" And cmbCliente.Text <> "" Then
    StrSql = "SELECT TOP (100) PERCENT dbo.tbl_Detalhes_Recebimento.Enviado,dbo.tbl_Detalhes_Recebimento.Data_envio,dbo.tbl_Detalhes_Recebimento.seq_remessa,dbo.tbl_Detalhes_Recebimento.IDContaReceber,dbo.tbl_Detalhes_Recebimento.txt_Cond_Recebimento, dbo.tbl_Detalhes_Recebimento.Id," _
    & "dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco,dbo.tbl_Detalhes_Recebimento.dt_Vencimento," _
    & "dbo.tbl_Detalhes_Recebimento.txt_tipoPagto, dbo.tbl_Detalhes_Recebimento.dbl_Valor," _
    & "dbo.tbl_Detalhes_Recebimento.int_NotaFiscal,dbo.tbl_Detalhes_Recebimento.txt_parcela, dbo.tbl_Detalhes_Recebimento.Nosso_numero, dbo.tbl_Detalhes_Recebimento.Carteira, dbo.tbl_Detalhes_Recebimento.Data_emissao,dbo.tbl_contas_receber.Nome_Razao,dbo.tbl_contas_receber.Vencimento FROM dbo.tbl_Detalhes_Recebimento" _
    & " INNER JOIN dbo.tbl_contas_receber ON dbo.tbl_Detalhes_Recebimento.IDContaReceber = dbo.tbl_contas_receber.IDIntconta" _
    & " WHERE (dbo.tbl_contas_receber.Nome_Razao = '" & cmbCliente & "') AND (dbo.tbl_Detalhes_Recebimento.txt_tipoPagto = N'BOLETO') AND  (NOT(dbo.tbl_Detalhes_Recebimento.Nosso_numero IS NULL)) AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento >= '" & DTINI & "') AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento <= '" & DTFim & "') AND (dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco = '" & cmbBanco.Text & "') ORDER BY dbo.tbl_Detalhes_Recebimento.dt_Vencimento"
    ProcCarregaListaDuplicatas
Else
    StrSql = "SELECT TOP (100) PERCENT dbo.tbl_Detalhes_Recebimento.Enviado,dbo.tbl_Detalhes_Recebimento.Data_envio,dbo.tbl_Detalhes_Recebimento.seq_remessa,dbo.tbl_Detalhes_Recebimento.IDContaReceber,dbo.tbl_Detalhes_Recebimento.txt_Cond_Recebimento, dbo.tbl_Detalhes_Recebimento.Id," _
    & "dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco,dbo.tbl_Detalhes_Recebimento.dt_Vencimento," _
    & "dbo.tbl_Detalhes_Recebimento.txt_tipoPagto, dbo.tbl_Detalhes_Recebimento.dbl_Valor," _
    & "dbo.tbl_Detalhes_Recebimento.int_NotaFiscal,dbo.tbl_Detalhes_Recebimento.txt_parcela, dbo.tbl_Detalhes_Recebimento.Nosso_numero, dbo.tbl_Detalhes_Recebimento.Carteira, dbo.tbl_Detalhes_Recebimento.Data_emissao,dbo.tbl_contas_receber.Nome_Razao,dbo.tbl_contas_receber.Vencimento FROM dbo.tbl_Detalhes_Recebimento" _
    & " INNER JOIN dbo.tbl_contas_receber ON dbo.tbl_Detalhes_Recebimento.IDContaReceber = dbo.tbl_contas_receber.IDIntconta" _
    & " WHERE (dbo.tbl_Detalhes_Recebimento.txt_tipoPagto = N'BOLETO') AND  (NOT(dbo.tbl_Detalhes_Recebimento.Nosso_numero IS NULL)) AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento >= '" & DTINI & "') AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento <= '" & DTFim & "') AND (dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco = '" & cmbBanco.Text & "') ORDER BY dbo.tbl_Detalhes_Recebimento.dt_Vencimento"
    ProcCarregaListaDuplicatas
End If

chkRemessa.Visible = False
chkRemessa.Value = 0
chkAtualizar.Visible = True

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub


Private Sub cmdSalvarInstrucoes_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente salvar as instruções do boleto?", vbYesNo, "CAPRIND V5.0") = vbYes Then
    Set TBBoleto = CreateObject("adodb.recordset")
         TBBoleto.Open "Select * from tbl_Instituicoes_Instrucoes_Boleto where ID_Instituicao = " & Txt_IDBanco & "", Conexao, adOpenKeyset, adLockOptimistic
         If TBBoleto.EOF = True Then TBBoleto.AddNew
         TBBoleto!ID_instituicao = Txt_IDBanco
         TBBoleto!Juros = Txtpercentual_juros
         TBBoleto!Desconto = Txtpercentual_desconto
         TBBoleto!Multa = Txtpercentual_multa
         TBBoleto!dias_protesto = Txtdias_protesto
         TBBoleto!Instrucoes_protesto = Txtinstrucoes
         TBBoleto!AssuntoEmail = txtAssunto
         TBBoleto.Update
         TBBoleto.Close
USMsgBox "Dados salvos com sucesso!", vbInformation, "CAPRIND V5.0"
End If
ProcCarregaInstrucoesBoleto

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Me.Skin1.ApplySkin Me.hWnd
ProcCarregaComboEmpresaBoleto
ProcRemoveObjetosResize Me
    
    If cmbempresa.Text <> "" Then
        Logo_Banco.Picture = LoadPicture("")
        ProcCarregadadosCedente
        ProcCarregaComboBancoBoleto
    End If

DTINI.Value = Date
DTFim.Value = "31/12/" & Year(Date)
ProcCarregacomboCarteira
ProcCarregaComboCliente
ProcCarregaInstrucoesBoleto

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Function FunSóNumeros(X As String) As String
On Error GoTo tratar_erro
Dim Temp As String
Dim j As Integer


Temp = ""
For j = 1 To Len(X)
    If Mid(X, j, 1) = "0" Or _
        Mid(X, j, 1) = "1" Or _
        Mid(X, j, 1) = "2" Or _
        Mid(X, j, 1) = "3" Or _
        Mid(X, j, 1) = "4" Or _
        Mid(X, j, 1) = "5" Or _
        Mid(X, j, 1) = "6" Or _
        Mid(X, j, 1) = "7" Or _
        Mid(X, j, 1) = "8" Or _
        Mid(X, j, 1) = "9" Then
        Temp = Temp + Mid(X, j, 1)
    End If
Next
FunSóNumeros = Temp

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Private Sub lst_Duplicata_Click()
On Error GoTo tratar_erro
Titulosselecionados = 0

If ColumnHeader = "" Then
    Contador = 0
    With lst_Duplicata
    For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
            Titulosselecionados = Titulosselecionados + 1
            End If
        Next InitFor
    End With
End If

If Titulosselecionados > 0 Then
    chkRemessa.Enabled = True
    chkEmail.Enabled = True
    chkImprimir.Enabled = True
    chkAtualizar.Enabled = False
Else
    chkRemessa.Enabled = False
    chkEmail.Enabled = False
    chkEmailcopia.Enabled = False
    chkImprimir.Enabled = False
    chkAtualizar.Enabled = False
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub lst_Duplicata_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro
Titulosselecionados = 0

If ColumnHeader = "" Then
    With lst_Duplicata
    For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                Titulosselecionados = Titulosselecionados + 1
                .ListItems.Item(InitFor).Checked = True
            End If
        Next InitFor
    End With
End If

If Titulosselecionados > 0 Then
    chkRemessa.Enabled = True
    chkRemessa.Value = 1
    chkEmail.Enabled = True
    chkImprimir.Enabled = True
    chkAtualizar.Enabled = True
    chkAtualizar.Value = 0
Else
    chkRemessa.Value = 0
    chkRemessa.Enabled = False
    chkEmail.Enabled = False
    chkEmailcopia.Enabled = False
    chkImprimir.Enabled = False
    chkAtualizar.Enabled = False
    chkAtualizar.Value = 0
End If

ProcOrdenaListView lst_Duplicata, ColumnHeader

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtcarteiraconf_Change()
On Error GoTo tratar_erro

    If txtcarteiraconf.Text <> "" Then
    FramePesquisa.Enabled = True
    Else
    FramePesquisa.Enabled = False
    End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

