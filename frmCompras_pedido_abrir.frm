VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCompras_pedido_abrir 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Administrativo | Compras - Pedido | Localizar"
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6780
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USButton btnFiltrar 
      Height          =   645
      Left            =   5250
      TabIndex        =   20
      Top             =   2850
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1138
      DibPicture      =   "frmCompras_pedido_abrir.frx":0000
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Filtrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ForeColorDown   =   16777215
      ForeColorOver   =   16777215
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorDown1=   4013465
      GradientColorDown2=   4013465
      GradientColorDown3=   4013465
      GradientColorDown4=   4013465
      GradientColorOver1=   4408288
      GradientColorOver2=   4408288
      GradientColorOver3=   4408288
      GradientColorOver4=   4408288
      Theme           =   4
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   873
      DibPicture      =   "frmCompras_pedido_abrir.frx":3650
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmCompras_pedido_abrir.frx":6CA0
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   18
      Top             =   3675
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   7350
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmCompras_pedido_abrir.frx":6FBA
      Count           =   1
   End
   Begin VB.CheckBox Chk_emissao 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   450
      TabIndex        =   7
      Top             =   3090
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2205
      Left            =   210
      TabIndex        =   10
      Top             =   630
      Width           =   6375
      Begin VB.ComboBox Cmb_empresa 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         ItemData        =   "frmCompras_pedido_abrir.frx":91A7
         Left            =   210
         List            =   "frmCompras_pedido_abrir.frx":91A9
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         ToolTipText     =   "Empresa."
         Top             =   420
         Width           =   2715
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frase"
         Height          =   540
         Left            =   180
         TabIndex        =   17
         Top             =   810
         Width           =   6045
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4110
            TabIndex        =   26
            Top             =   240
            Width           =   705
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   25
            Top             =   240
            Width           =   1155
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   300
            TabIndex        =   24
            Top             =   240
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1590
            TabIndex        =   23
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.ComboBox Cmb_alteracao 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         ItemData        =   "frmCompras_pedido_abrir.frx":91AB
         Left            =   4740
         List            =   "frmCompras_pedido_abrir.frx":91B8
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Opções para filtro."
         Top             =   420
         Width           =   1455
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         ItemData        =   "frmCompras_pedido_abrir.frx":91C8
         Left            =   2940
         List            =   "frmCompras_pedido_abrir.frx":9205
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   420
         Width           =   1785
      End
      Begin MSMask.MaskEdBox txtCpf 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         ToolTipText     =   "Número do CPF."
         Top             =   1650
         Visible         =   0   'False
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtcnpj 
         Height          =   315
         Left            =   180
         TabIndex        =   4
         ToolTipText     =   "Número do CNPJ."
         Top             =   1650
         Visible         =   0   'False
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.###.###/####-##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtTexto 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   180
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1650
         Width           =   6015
      End
      Begin VB.ComboBox cmbfamilia 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   180
         MousePointer    =   99  'Custom
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1650
         Visible         =   0   'False
         Width           =   6015
      End
      Begin MSComCtl2.DTPicker txtInicio 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1650
         Visible         =   0   'False
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Format          =   104136705
         CurrentDate     =   39057
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Left            =   1260
         TabIndex        =   22
         Top             =   210
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Com alteração"
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
         Left            =   4845
         TabIndex        =   16
         Top             =   210
         Width           =   1035
      End
      Begin VB.Label Label1 
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
         Left            =   2452
         TabIndex        =   12
         Top             =   1410
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
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
         Left            =   3420
         TabIndex        =   11
         Top             =   210
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   675
      Left            =   210
      TabIndex        =   13
      Top             =   2790
      Width           =   5025
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   3690
         TabIndex        =   9
         ToolTipText     =   "Data final."
         Top             =   210
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
         Format          =   104136705
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   2010
         TabIndex        =   8
         ToolTipText     =   "Data inicio."
         Top             =   210
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
         Format          =   104136705
         CurrentDate     =   39057
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
         Left            =   3285
         TabIndex        =   15
         Top             =   240
         Width           =   360
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
         Left            =   1650
         TabIndex        =   14
         Top             =   240
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmCompras_pedido_abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcCarregaNomeFantasia()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
StrSql = "select  Distinct nomefantasia from Compras_pedido CP inner join Compras_fornecedores CF on CP.idfornecedor = CF.Id where nomefantasia is not null"
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then

Do While TBAbrir.EOF = False
With cmbfamilia
If IsNull(TBAbrir!NomeFantasia) = False Then
.AddItem TBAbrir!NomeFantasia
End If
End With
TBAbrir.MoveNext
Loop

End If
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnFiltrar_Click()
On Error GoTo tratar_erro

 ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_emissao_Click()
On Error GoTo tratar_erro

If Chk_emissao.Value = 1 Then
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

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

With cmbfamilia
    txtinicio.Value = Date
    
    If cmbfiltrarpor = "Nome Fantasia" Then
        txtTexto.Visible = False
        .Visible = True
        txtinicio.Visible = False
        txtcnpj.Visible = False
        txtCpf.Visible = False
        .Clear
     ProcCarregaNomeFantasia
     Exit Sub
   End If

    If cmbfiltrarpor = "Status" Or cmbfiltrarpor = "Família" Then
        txtTexto.Visible = False
        .Visible = True
        txtinicio.Visible = False
        txtcnpj.Visible = False
        txtCpf.Visible = False
        .Clear
        If cmbfiltrarpor = "Status" Then
            .AddItem "AGUARDANDO APROVAÇÃO"
            .AddItem "COMPRADO"
            .AddItem "RECEBIDO PARCIAL"
            .AddItem "RECEBIDO"
            .AddItem "CANCELADO"
        Else
            ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and compras = 'True'", True
        End If
    ElseIf cmbfiltrarpor = "Prazo entrega" Then
            txtTexto.Visible = False
            .Visible = False
            txtinicio.Visible = True
            txtcnpj.Visible = False
            txtCpf.Visible = False
        ElseIf cmbfiltrarpor = "CNPJ" Then
                txtTexto.Visible = False
                .Visible = False
                txtinicio.Visible = False
                txtcnpj.Visible = True
                txtCpf.Visible = False
            ElseIf cmbfiltrarpor = "CPF" Then
                    txtTexto.Visible = False
                    .Visible = False
                    txtinicio.Visible = False
                    txtcnpj.Visible = False
                    txtCpf.Visible = True
                Else
                    txtTexto.Visible = True
                    .Visible = False
                    txtinicio.Visible = False
                    txtcnpj.Visible = False
                    txtCpf.Visible = False
    End If
End With

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
DataFiltro = ""
DataFiltroRel = ""
If Chk_emissao.Value = 1 Then
    DataFiltro = " and CP.Idpedido <> 0 and CP.Data Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
    DataFiltroRel = " and {Compras_pedido.Idpedido} <> 0 and {Compras_pedido.Data} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {Compras_pedido.Data} <= Date(" & _
                                        Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
End If
CamposFiltro = "CP.IDpedido, CP.Data, CP.Pedido, CC.Cotacaotexto, CP.Fornecedor, CP.Status_pedido, CP.DtValidacao, CP.Data_aprovado, CP.dbl_valor_total"

TextoFiltroAlt = ""
TextoFiltroAltRel = ""
If Cmb_alteracao = "Sim" Then
    TextoFiltroAlt = " and VCA.ID IS NOT NULL and VCA.Tipo = 'CPE'"
    TextoFiltroAltRel = " and Not(IsNull({vendas_carteira_alteracoes.ID})) and {vendas_carteira_alteracoes.tipo} = 'CPE'"
ElseIf Cmb_alteracao = "Não" Then
        TextoFiltroAlt = " and VCA.ID IS NULL"
        TextoFiltroAltRel = " and ISNULL({vendas_carteira_alteracoes.ID}) = True"
End If

INNERJOINTEXTO = "Select " & CamposFiltro & " from ((((((Compras_pedido CP LEFT JOIN Compras_fornecedores CF ON CF.IDCliente = CP.idfornecedor) LEFT JOIN Compras_pedido_lista CPL ON CPL.Idpedido = CP.Idpedido) LEFT JOIN Compras_requisicao CR ON CR.ID_requisicao = CPL.ID_requisicao) LEFT JOIN Compras_cotacao CC ON CC.ID_cotacao = CP.IDcotacao) LEFT JOIN Producao_pedidos PP ON PP.Ordem = CPL.Ordem) LEFT JOIN vendas_carteira VC ON VC.Codigo = PP.IDcarteira) LEFT JOIN vendas_proposta VP ON VP.Cotacao = VC.Cotacao LEFT JOIN vendas_carteira_alteracoes VCA ON VCA.ID_carteira = CPL.IDlista"
TextoFiltroPadrao = "CP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & TextoFiltroAlt & " group by " & CamposFiltro & " order by CP.idpedido desc"
TextoFiltroPadraoRel = "{Compras_pedido.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltroRel & TextoFiltroAltRel

With frmCompras_Pedido
    .listapedido.ListItems.Clear
    Empresarel = Cmb_empresa
    If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Or txtinicio.Visible = True Or txtcnpj.Visible = True And txtcnpj <> "__.___.___/____-__" Or txtCpf.Visible = True And txtCpf <> "___.___.___-__" Then
        If cmbfiltrarpor = "Status" Then
            If cmbfamilia = "COMPRADO" Or cmbfamilia = "RECEBIDO" Or cmbfamilia = "RECEBIDO PARCIAL" Then
                Select Case cmbfamilia
                    Case "COMPRADO": Texto = "ABERTO"
                    Case "RECEBIDO": Texto = "ENCERRADO"
                    Case "RECEBIDO PARCIAL": Texto = "PARCIAL"
                End Select
            Else
                Texto = cmbfamilia
            End If
            .Sql_Pedido_Localizar = INNERJOINTEXTO & " where CP.Status_pedido = '" & Texto & "' and " & TextoFiltroPadrao
            .FormulaRel_Pedido = "{Compras_pedido.Status_pedido} = '" & Texto & "' and " & TextoFiltroPadraoRel
        ElseIf cmbfiltrarpor = "Família" Then
                .Sql_Pedido_Localizar = INNERJOINTEXTO & " where CPL.Familia = '" & cmbfamilia & "' and " & TextoFiltroPadrao
                .FormulaRel_Pedido = "{Compras_pedido_lista.Familia} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
'=================================================================================
         ElseIf cmbfiltrarpor = "Nome Fantasia" Then
                .Sql_Pedido_Localizar = INNERJOINTEXTO & " where CF.nomefantasia = '" & cmbfamilia & "' and " & TextoFiltroPadrao
'=================================================================================
           ElseIf cmbfiltrarpor = "Prazo entrega" Then
                    .Sql_Pedido_Localizar = INNERJOINTEXTO & " where CPL.Prazo = '" & Format(txtinicio.Value, "Short Date") & "' and " & TextoFiltroPadrao
                    .FormulaRel_Pedido = "{Compras_pedido_lista.Prazo} = Date(" & Year(txtinicio.Value) & "," & Month(txtinicio.Value) & "," & Day(txtinicio.Value) & ") and " & TextoFiltroPadraoRel
                ElseIf cmbfiltrarpor = "Ordem" Or cmbfiltrarpor = "OS" Then
                        If cmbfiltrarpor = "Ordem" Then TextoFiltro = "CPL.Ordem" Else TextoFiltro = "CPL.OS"
                        .Sql_Pedido_Localizar = INNERJOINTEXTO & " where " & TextoFiltro & " = " & txtTexto & " and " & TextoFiltroPadrao
                        TextoFiltroRel = Replace(TextoFiltro, "CP.", "Compras_pedido.")
                        .FormulaRel_Pedido = "{" & TextoFiltroRel & "} = " & txtTexto & " and " & TextoFiltroPadraoRel
                    ElseIf cmbfiltrarpor = "CNPJ" Or cmbfiltrarpor = "CPF" Then
                            If cmbfiltrarpor = "CNPJ" Then
                                TextoFiltro = "CF.CPF_CNPJ = '" & txtcnpj & "'"
                                TextoFiltroRel = "{Compras_Compras_fornecedores.CPF_CNPJ} = '" & txtcnpj & "'"
                            Else
                                TextoFiltro = "CF.CPF_CNPJ = '" & txtCpf & "'"
                                TextoFiltroRel = "{Compras_Compras_fornecedores.CPF_CNPJ} = '" & txtCpf & "'"
                            End If
                            .Sql_Pedido_Localizar = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao
                            .FormulaRel_Pedido = TextoFiltroRel & " and " & TextoFiltroPadraoRel
                        Else
                            TextoFiltroOrdem = ""
                            TextoFiltroOrdemRel = ""
                            Select Case cmbfiltrarpor
                                Case "Solicitação": TextoFiltro = "CR.Requisicaotexto"
                                Case "Cotação": TextoFiltro = "CC.Cotacaotexto"
                                Case "Pedido": TextoFiltro = "CP.pedido"
                                Case "Fornecedor": TextoFiltro = "CP.fornecedor"
                                Case "Referência": TextoFiltro = "CP.N_referencia"
                                Case "Código interno": TextoFiltro = "CPL.desenho"
                                Case "Descrição": TextoFiltro = "CPL.descricao"
                                Case "Descrição comercial": TextoFiltro = "CPL.descricao_comercial"
                                Case "Detalhe": TextoFiltro = "CPL.Detalheitem"
                                Case "Pedido interno":
                                    TextoFiltro = "VP.Ncotacao"
                                    TextoFiltroOrdem = " and VP.Ncotacao IS NOT NULL"
                                    TextoFiltroOrdemRel = " and {Vendas_proposta.Ncotacao} <> 'Null'"
                                Case "Cliente": TextoFiltro = "CPL.Cliente"
                            End Select
                            If Left(TextoFiltro, 3) = "CR." Then
                                TextoFiltroRel = Replace(TextoFiltro, "CR.", "Compras_requisicao.")
                            ElseIf Left(TextoFiltro, 3) = "CC." Then
                                    TextoFiltroRel = Replace(TextoFiltro, "CC.", "Compras_Cotacao.")
                                ElseIf Left(TextoFiltro, 3) = "CP." Then
                                        TextoFiltroRel = Replace(TextoFiltro, "CP.", "Compras_pedido.")
                                    ElseIf Left(TextoFiltro, 3) = "VP." Then
                                            TextoFiltroRel = Replace(TextoFiltro, "VP.", "Vendas_proposta.")
                                        Else
                                            TextoFiltroRel = Replace(TextoFiltro, "CPL.", "Compras_pedido_lista.")
                            End If
                            .Sql_Pedido_Localizar = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & TextoFiltroOrdem & " and " & TextoFiltroPadrao
                            .FormulaRel_Pedido = "{" & TextoFiltroRel & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & TextoFiltroOrdemRel & " and " & TextoFiltroPadraoRel
            End If
    Else
        .Sql_Pedido_Localizar = INNERJOINTEXTO & " where " & TextoFiltroPadrao
        .FormulaRel_Pedido = TextoFiltroPadraoRel
    End If
    .ProcAtualizalistapedido (1)
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
    Case vbKeyEscape: Unload Me
    Case vbKeyF2: ProcFiltrar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaComboEmpresa Cmb_empresa, False
cmbfiltrarpor = "Pedido"
msk_fltInicio = Date
msk_fltFim = Date

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If txtTexto <> "" Then
    cmbfamilia.ListIndex = -1
    If cmbfiltrarpor = "Ordem" Or cmbfiltrarpor = "OS" Then
        VerifNumero = txtTexto
        ProcVerificaNumero
        If VerifNumero = False Then
            txtTexto = ""
            txtTexto.SetFocus
            Exit Sub
        End If
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


