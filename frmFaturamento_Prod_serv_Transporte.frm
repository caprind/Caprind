VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Prod_serv_Transporte 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Faturamento | Informações adicionais ao frete"
   ClientHeight    =   4665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações para transporte"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3225
      Index           =   14
      Left            =   270
      TabIndex        =   4
      Top             =   810
      Width           =   8295
      Begin VB.TextBox txttransporte 
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
         Height          =   645
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Transporte."
         Top             =   2385
         Width           =   7755
      End
      Begin VB.TextBox txt_Tipo_Frete 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   555
         Width           =   7815
      End
      Begin VB.TextBox txtidTransportadora 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   2670
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox txtRedespacho 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1710
         Width           =   6585
      End
      Begin VB.TextBox txtTransportadora 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1140
         Width           =   6585
      End
      Begin VB.TextBox txtTipoTransp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1215
      End
      Begin VB.TextBox txtTipoTransp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   1
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1710
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Informações do transporte"
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
         Left            =   2880
         TabIndex        =   15
         Top             =   2190
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Frete por conta"
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
         Left            =   3600
         TabIndex        =   12
         Top             =   330
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
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
         Left            =   615
         TabIndex        =   11
         Top             =   1500
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Redespacho"
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
         Left            =   4290
         TabIndex        =   10
         Top             =   1500
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
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
         Left            =   660
         TabIndex        =   9
         Top             =   930
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Transportadora"
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
         Left            =   4170
         TabIndex        =   8
         Top             =   930
         Width           =   1125
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   4260
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   741
      DibPicture      =   "frmFaturamento_Prod_serv_Transporte.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      EnableMaximizeButton=   0   'False
      EnableMinimizeButton=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmFaturamento_Prod_serv_Transporte.frx":5A65
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmFaturamento_Prod_serv_Transporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo tratar_erro

ProcPuxaDadostransporte


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcPuxaDadostransporte()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
If Formulario = "Estoque/Ordem de faturamento" Then
    TBAbrir.Open "Select * FROM vendas_Proposta WHERE NCotacao = '" & frmEstoque_Ordem_Faturamento.txt_proposta.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
Else
    TBAbrir.Open "Select * FROM vendas_Proposta WHERE NCotacao = '" & frmFaturamento_Prod_Serv.txt_proposta.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
End If

If TBAbrir.EOF = False Then

IDpedido = TBAbrir!Cotacao

'ProcLimparComercial
Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select * FROM vendas_comercial WHERE cotacao = " & IDpedido, Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = False Then
'    If TBCotacao!analize = "Sim" Or TBCotacao!analize = "Não" Then txtAnalize.Text = TBCotacao!analize
'    txtcalculos = IIf(IsNull(TBCotacao!calculos), "", TBCotacao!calculos)
'    txtimpostos = IIf(IsNull(TBCotacao!impostos), "", TBCotacao!impostos)
'    txtCondicoes = IIf(IsNull(TBCotacao!condicoes), "", TBCotacao!condicoes)
'    txtgarantia = IIf(IsNull(TBCotacao!garantia), "", TBCotacao!garantia)
'    TxtObservacoes = IIf(IsNull(TBCotacao!observacoes), "", TBCotacao!observacoes)
'    txtReajuste = IIf(IsNull(TBCotacao!reajuste), "", TBCotacao!reajuste)
     txttransporte = IIf(IsNull(TBCotacao!transporte), "", TBCotacao!transporte)
'     txtValidade = IIf(IsNull(TBCotacao!validade), "", TBCotacao!validade)
'     txtcalculos = IIf(IsNull(TBCotacao!calculos), "", TBCotacao!calculos)
    
'===========================================================================================
        If IsNull(TBCotacao!Tipo_transp2) = False And TBCotacao!Tipo_transp2 <> "" Then
        Select Case TBCotacao!Tipo_transp2
            Case "C": txtTipoTransp(1).Text = "Cliente"
            Case "F": txtTipoTransp(1).Text = "Fornecedor"
            Case "E": txtTipoTransp(1).Text = "Empresa"
        End Select
    End If
   ' NomeCampo = "a transportadora"
    
    If IsNull(TBCotacao!Redespacho) = False And TBCotacao!Redespacho <> "" Then
    txtRedespacho = TBCotacao!Redespacho
    Else
    txtRedespacho = "SEM INFORMAÇÕES"
    txtTipoTransp(1).Text = "n/a"
    End If
    
   If IsNull(TBCotacao!Tipo_Frete) = False And TBCotacao!Tipo_Frete <> "" Then
   txt_Tipo_Frete.Text = TBCotacao!Tipo_Frete
   End If
   
'
'        If IsNull(TBCotacao!Tipo_transp) = False And TBCotacao!Tipo_transp <> "" Then
'        Select Case TBCotacao!Tipo_transp
'            Case "C": txtTipoTransp(0).Text = "Cliente"
'            Case "F": txtTipoTransp(0).Text = "Fornecedor"
'            Case "E": txtTipoTransp(0).Text = "Empresa"
'        End Select
'    End If
    
'===========================================================================================
    
'    With txtlocal_entrega
'        .AddItem ""
'        If IsNull(TBCotacao!Local_entrega) = False And TBCotacao!Local_entrega <> "" Then
'            .AddItem TBCotacao!Local_entrega
'            .Text = TBCotacao!Local_entrega
'            Txt_ID_entrega = IIf(IsNull(TBCotacao!ID_entrega), 0, TBCotacao!ID_entrega)
'        End If
'    End With
'    With txtlocal_cobranca
'        .AddItem ""
'        If IsNull(TBCotacao!Local_cobranca) = False And TBCotacao!Local_cobranca <> "" Then
'            .AddItem TBCotacao!Local_cobranca
'            .Text = TBCotacao!Local_cobranca
'            Txt_ID_cobranca = IIf(IsNull(TBCotacao!ID_Cobranca), 0, TBCotacao!ID_Cobranca)
'        End If
'    End With
'    NomeCampo = "a transportadora"
    If IsNull(TBCotacao!Transportadora) = False And TBCotacao!Transportadora <> "" Then
    txtTransportadora = TBCotacao!Transportadora
    txtidTransportadora.Text = TBCotacao!IdIntTransp
    End If

    If IsNull(TBCotacao!Tipo_transp) = False And TBCotacao!Tipo_transp <> "" Then
        Select Case TBCotacao!Tipo_transp
            Case "C": txtTipoTransp(0).Text = "Cliente"
            Case "F": txtTipoTransp(0).Text = "Fornecedor"
            Case "E": txtTipoTransp(0).Text = "Empresa"
        End Select
    End If
    
'    If IsNull(TBCotacao!Moeda) = False And TBCotacao!Moeda <> "" Then cmbMoeda = TBCotacao!Moeda
'    Txt_valor_moeda = IIf(IsNull(TBCotacao!Valor_moeda), "", Format(TBCotacao!Valor_moeda, "###,##0.0000"))
End If
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
