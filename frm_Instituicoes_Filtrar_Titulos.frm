VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frm_Instituicoes_Filtrar_Titulos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Carteira de titulos"
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
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
   ScaleHeight     =   2715
   ScaleWidth      =   4320
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   2310
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   741
      DibPicture      =   "frm_Instituicoes_Filtrar_Titulos.frx":0000
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
      Icon            =   "frm_Instituicoes_Filtrar_Titulos.frx":3650
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   390
      TabIndex        =   0
      Top             =   600
      Width           =   3525
      Begin MSComCtl2.DTPicker dtData_Envio 
         Height          =   315
         Left            =   2025
         TabIndex        =   2
         ToolTipText     =   "Data de vencimento de início para pesquisa."
         Top             =   390
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
         Format          =   509673475
         CurrentDate     =   42344
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparente
         Caption         =   "Remessa gerada em : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   450
         Width           =   1635
      End
   End
   Begin DrawSuite2022.USButton btnFiltrar 
      Height          =   375
      Left            =   2670
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      DibPicture      =   "frm_Instituicoes_Filtrar_Titulos.frx":396A
      BorderColor     =   4960354
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   49152
      Caption         =   "Filtrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ForeColorDown   =   16777215
      ForeColorOver   =   16777215
      GradientColor1  =   4960354
      GradientColor2  =   4960354
      GradientColor3  =   4960354
      GradientColor4  =   4960354
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorDown1=   32768
      GradientColorDown2=   32768
      GradientColorDown3=   32768
      GradientColorDown4=   32768
      GradientColorOver1=   49152
      GradientColorOver2=   49152
      GradientColorOver3=   49152
      GradientColorOver4=   49152
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   3
   End
End
Attribute VB_Name = "frm_Instituicoes_Filtrar_Titulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnFiltrar_Click()

     StrSql = "SELECT TOP (100) PERCENT dbo.tbl_Detalhes_Recebimento.Enviado,dbo.tbl_Detalhes_Recebimento.seq_remessa, dbo.tbl_Detalhes_Recebimento.Data_envio,dbo.tbl_Detalhes_Recebimento.seq_remessa,dbo.tbl_Detalhes_Recebimento.IDContaReceber,dbo.tbl_Detalhes_Recebimento.txt_Cond_Recebimento, dbo.tbl_Detalhes_Recebimento.Id," _
    & "dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco,dbo.tbl_Detalhes_Recebimento.dt_Vencimento," _
    & "dbo.tbl_Detalhes_Recebimento.txt_tipoPagto, dbo.tbl_Detalhes_Recebimento.dbl_Valor," _
    & "dbo.tbl_Detalhes_Recebimento.int_NotaFiscal,dbo.tbl_Detalhes_Recebimento.txt_parcela, dbo.tbl_Detalhes_Recebimento.Nosso_numero, dbo.tbl_Detalhes_Recebimento.Carteira, dbo.tbl_Detalhes_Recebimento.Data_emissao,dbo.tbl_contas_receber.Nome_Razao,dbo.tbl_contas_receber.Vencimento FROM dbo.tbl_Detalhes_Recebimento" _
    & " INNER JOIN dbo.tbl_contas_receber ON dbo.tbl_Detalhes_Recebimento.IDContaReceber = dbo.tbl_contas_receber.IDIntconta" _
    & " WHERE (dbo.tbl_Detalhes_Recebimento.seq_remessa IS NOT NULL) AND (dbo.tbl_Detalhes_Recebimento.txt_tipoPagto = N'BOLETO') AND  (NOT(dbo.tbl_Detalhes_Recebimento.Nosso_numero IS NULL)) AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento >= '" & frm_Instituicoes.DTINI.Value & "') AND (dbo.tbl_Detalhes_Recebimento.dt_Vencimento <= '" & frm_Instituicoes.DTFim.Value & "') AND (dbo.tbl_Detalhes_Recebimento.txt_Portador_Banco = '" & frm_Instituicoes.txtdescricao.Text & "') AND DATA_ENVIO = '" & dtData_Envio.Value & "' and Seq_remessa = '" & frm_Instituicoes.txtSeq.Text & "' ORDER BY dbo.tbl_Detalhes_Recebimento.dt_Vencimento"
Unload Me

End Sub

Private Sub Form_Load()

dtData_Envio.Value = Date

End Sub
