VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{8C1279ED-044C-4258-A3E3-0D5514B899FC}#1.44#0"; "ControlesUteis.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{D2B08629-3629-406E-B7BD-0CBED5F2C38F}#63.0#0"; "kmail.ocx"
Begin VB.Form FrmEnviarEmail 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Enviar documento por email - CAPRIND  v5.0"
   ClientHeight    =   7440
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   9255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmEnviarEmail.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   22
      Top             =   7035
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   714
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados para envio do documento por email"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   210
      TabIndex        =   4
      Top             =   720
      Width           =   8780
      Begin VB.CheckBox chkCopia 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enviar-me uma cópia"
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
         Height          =   315
         Left            =   6690
         TabIndex        =   17
         Top             =   4410
         Width           =   1935
      End
      Begin DrawSuite2022.USButton Cmd_anexo 
         Height          =   435
         Left            =   270
         TabIndex        =   10
         Top             =   4800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   767
         DibPicture      =   "FrmEnviarEmail.frx":1CCA
         Caption         =   "Localizar anexo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         Theme           =   3
      End
      Begin VB.TextBox Txt_descricao 
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
         Height          =   2205
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         ToolTipText     =   "Descrição."
         Top             =   1980
         Width           =   7335
      End
      Begin VB.TextBox Txt_assunto 
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
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   0
         ToolTipText     =   "Assunto."
         Top             =   1530
         Width           =   7335
      End
      Begin VB.TextBox Txt_anexo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Anexo."
         Top             =   3000
         Visible         =   0   'False
         Width           =   5535
      End
      Begin ControlesUteis.txt TxtEmail 
         Height          =   360
         Left            =   1200
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "E-mail."
         Top             =   750
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   635
         Tamanho         =   7335
         Tipo            =   2
         Text            =   ""
         FocusColor      =   16777215
         ShowCaption     =   0   'False
         Caption         =   ""
         MaxLength       =   255
         Negative        =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
      End
      Begin DrawSuite2022.USButton Cmd_limpar_caminho 
         Height          =   435
         Left            =   2130
         TabIndex        =   11
         Top             =   4800
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   767
         DibPicture      =   "FrmEnviarEmail.frx":531A
         Caption         =   "Excluir anexo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         GradientColor1  =   5263559
         GradientColor2  =   5263559
         GradientColor3  =   5263559
         GradientColor4  =   5263559
         GradientColorDisabled1=   13160660
         GradientColorDisabled2=   13160660
         GradientColorDisabled3=   13160660
         GradientColorDisabled4=   13160660
         GradientColorOver1=   4408288
         GradientColorOver2=   4408288
         GradientColorOver3=   4408288
         GradientColorOver4=   4408288
         GradientColorDown1=   4013465
         GradientColorDown2=   4013465
         GradientColorDown3=   4013465
         GradientColorDown4=   4013465
         Theme           =   4
      End
      Begin DrawSuite2022.USButton Cmd_visualizar_arquivo 
         Height          =   435
         Left            =   3960
         TabIndex        =   12
         Top             =   4800
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   767
         DibPicture      =   "FrmEnviarEmail.frx":EB66
         Caption         =   "Abrir anexo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         GradientColor1  =   1154291
         GradientColor2  =   1154291
         GradientColor3  =   1154291
         GradientColor4  =   1154291
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   8438015
         GradientColorOver2=   8438015
         GradientColorOver3=   8438015
         GradientColorOver4=   8438015
         GradientColorDown1=   16576
         GradientColorDown2=   16576
         GradientColorDown3=   16576
         GradientColorDown4=   16576
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         Theme           =   5
      End
      Begin DrawSuite2022.USButton btnEnviar 
         Height          =   435
         Left            =   6660
         TabIndex        =   13
         Top             =   4800
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   767
         DibPicture      =   "FrmEnviarEmail.frx":103BA
         Caption         =   "Enviar email"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   8421376
         BorderColorDisabled=   8421376
         BorderColorDown =   8421376
         BorderColorOver =   8421376
         ForeColorOver   =   128
         ForeColorDown   =   128
         GradientColor1  =   8421376
         GradientColor2  =   8421376
         GradientColor3  =   8421376
         GradientColor4  =   8421376
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   12632064
         GradientColorOver2=   16776960
         GradientColorOver3=   16776960
         GradientColorOver4=   16776960
         GradientColorDown1=   8421376
         GradientColorDown2=   8421376
         GradientColorDown3=   8421376
         GradientColorDown4=   12632064
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         Theme           =   3
      End
      Begin ControlesUteis.txt txtDe 
         Height          =   360
         Left            =   1200
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "E-mail."
         Top             =   360
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   635
         Tamanho         =   7335
         Tipo            =   2
         Text            =   ""
         FocusColor      =   16777215
         ShowCaption     =   0   'False
         Caption         =   ""
         MaxLength       =   255
         Negative        =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
      End
      Begin ControlesUteis.txt txtCopia 
         Height          =   360
         Left            =   1200
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "E-mail."
         Top             =   1140
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   635
         Tamanho         =   7335
         Tipo            =   2
         Text            =   ""
         FocusColor      =   16777215
         ShowCaption     =   0   'False
         Caption         =   ""
         MaxLength       =   255
         Negative        =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
      End
      Begin DrawSuite2022.USCheckBox chkEmailenviado 
         Height          =   315
         Left            =   6690
         TabIndex        =   20
         Top             =   5340
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BackColor       =   -2147483633
         Caption         =   "Email enviado"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8388608
         ShowFocusRect   =   -1  'True
      End
      Begin DrawSuite2022.USCheckBox ChkCopiaEnviada 
         Height          =   315
         Left            =   6690
         TabIndex        =   21
         Top             =   5610
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         BackColor       =   -2147483633
         Caption         =   "Cópia enviada"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8388608
         ShowFocusRect   =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Com cópia:"
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
         Left            =   300
         TabIndex        =   19
         Top             =   1125
         Width           =   795
      End
      Begin DrawSuite2022.USAlphaImage USAlphaImage1 
         Height          =   360
         Left            =   1200
         TabIndex        =   24
         Top             =   4260
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         Image           =   "FrmEnviarEmail.frx":17DBE
         Props           =   5
      End
      Begin VB.Label lblanexo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anexo"
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
         Left            =   1680
         TabIndex        =   16
         Top             =   4290
         Width           =   6705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De:"
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
         Left            =   840
         TabIndex        =   14
         Top             =   330
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assunto:"
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
         Left            =   450
         TabIndex        =   6
         Top             =   1530
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Para:"
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
         Left            =   705
         TabIndex        =   8
         Top             =   735
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corpo email:"
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
         Left            =   225
         TabIndex        =   7
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anexo:"
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
         Left            =   570
         TabIndex        =   5
         Top             =   4260
         Width           =   525
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   953
      DibPicture      =   "FrmEnviarEmail.frx":2128F
      CaptionOnCenter =   -1  'True
      EnableMaximizeButton=   0   'False
      EnableMinimizeButton=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "FrmEnviarEmail.frx":28C93
      IconSize        =   1
      IconSizeX       =   24
      IconSizeY       =   24
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   4710
      Top             =   7110
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   5310
      Top             =   7110
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2190
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   2700
      Top             =   7080
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "FrmEnviarEmail.frx":2A96D
      Count           =   1
   End
   Begin KmailProject.kmail kmail 
      Height          =   615
      Left            =   300
      TabIndex        =   23
      Top             =   7350
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   1085
   End
End
Attribute VB_Name = "FrmEnviarEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnEnviar_Click()
On Error GoTo tratar_erro

chkEmailenviado.Value = 0
ChkCopiaEnviada.Value = 0
ProcEnviarEmail

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkCopia_Click()
On Error GoTo tratar_erro

If chkCopia.Value = 1 Then
txtCopia.Text = txtDe.Text
Else
txtCopia.Text = ""
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_anexo_Click()
On Error GoTo tratar_erro

ProcCarregaCaminhoNomeArquivo CommonDialog1, "*.*", "*.*"
Txt_anexo = caminho

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_limpar_caminho_Click()
On Error GoTo tratar_erro

Txt_anexo = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If Txt_anexo <> "" Then ProcAbrirArquivo Txt_anexo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcLimpaVariaveisPrincipais
ProcLimpaCampos
chkEmailenviado.Value = 0
ChkCopiaEnviada.Value = 0
   
'kmail.Visible = False

If Compras_Pedido = True Then
    Caption = "Compras - Pedido - Enviar e-mail"
    NomeTabela = "Compras_pedido"
    TextoFiltro = "IDpedido = " & frmCompras_Pedido.txtIDPedido
End If

If Vendas_Proposta = True Then
 With frmVendas_proposta

 Set TBClientes = CreateObject("adodb.recordset")
    TBClientes.Open "Select email from Clientes_Contatos where NomeContato ='" & .txtRemetente.Text & "' and IDCliente = " & .txtIDcliente.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBClientes.EOF = False And IsNull(TBClientes!Email) = False Then
    txtEmail.Text = TBClientes!Email
    Else
    If .txtEmail.Text = "" Then
    USMsgBox "Não foi encontrado um email valido cadastrado para envio da proposta comercial", vbCritical, "CAPRIND v5.0"
    Exit Sub
    Unload Me
    txtEmail.Text = .txtEmail.Text
    End If
    End If
 End With
End If


If NomeTabela <> "" And TextoFiltro <> "" Then
Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Assunto, Descricao, Anexo, Nome_anexo, Email from " & NomeTabela & " where " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then

If Compras_Pedido = True Then
 Txt_assunto = "Pedido de compras n° " & frmCompras_Pedido.txtPedido.Text
End If

If Vendas_Proposta = True Then
 Txt_assunto = "Proposta comercial n° " & frmVendas_proposta.txtCotacao.Text
End If

    'If IsNull(TBFI!Assunto) = False And TBFI!Assunto <> "" Then Txt_assunto = TBFI!Assunto
    If IsNull(TBFI!Descricao) = False And TBFI!Descricao <> "" Then Txt_descricao = TBFI!Descricao
    If TBFI!Anexo <> "" Then
        Txt_anexo = IIf(IsNull(TBFI!Anexo), "", TBFI!Anexo)
        Nome_anexo = IIf(IsNull(TBFI!Nome_anexo), "", TBFI!Nome_anexo)
    End If
    txtEmail.Text = IIf(IsNull(TBFI!Email), "", TBFI!Email)
End If
TBFI.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

TextoEmpresa = ""
TextoTel = ""
TextoFax = ""
Set TBFI = CreateObject("adodb.recordset")
If Compras_Pedido = True Then
    With frmCompras_Pedido
        TBFI.Open "Select Razao, Telefone, Fax from Empresa where codigo = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            TextoEmpresa = FunPrimeiraLetraMaiuscula(IIf(IsNull(TBFI!Razao), "", TBFI!Razao))
            If IsNull(TBFI!telefone) = False And TBFI!telefone <> "" Then TextoTel = "Tel: " & IIf(IsNull(TBFI!telefone), "", TBFI!telefone)
            If IsNull(TBFI!Fax) = False And TBFI!Fax <> "" Then
                If TextoTel <> "" Then TextoFax = " - Fax: " Else TextoFax = "Fax: "
                TextoFax = TextoFax & IIf(IsNull(TBFI!Fax), "", TBFI!Fax)
            End If
        End If
        TBFI.Close
        Txt_assunto = "Pedido de compra n. " & .txtPedido
        Txt_descricao = "Segue pedido de compra em anexo." & vbCrLf & "Favor acusar o recebimento deste pedido." & vbCrLf & vbCrLf & "Atenciosamente," & vbCrLf & vbCrLf & FunPrimeiraLetraMaiuscula(pubUsuario) & vbCrLf & TextoEmpresa & vbCrLf & TextoTel & TextoFax
    End With
End If

If Vendas_Proposta = True Then
    With frmVendas_proposta
        TBFI.Open "Select Razao, Telefone, Fax from Empresa where codigo = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            TextoEmpresa = FunPrimeiraLetraMaiuscula(IIf(IsNull(TBFI!Razao), "", TBFI!Razao))
            If IsNull(TBFI!telefone) = False And TBFI!telefone <> "" Then TextoTel = "Tel: " & IIf(IsNull(TBFI!telefone), "", TBFI!telefone)
            If IsNull(TBFI!Fax) = False And TBFI!Fax <> "" Then
                If TextoTel <> "" Then TextoFax = " - Fax: " Else TextoFax = "Fax: "
                TextoFax = TextoFax & IIf(IsNull(TBFI!Fax), "", TBFI!Fax)
            End If
        End If
        TBFI.Close
        Txt_assunto = "Proposta comercial n° " & .txtCotacao
        Txt_descricao = "Segue proposta comercial em anexo." & vbCrLf & "Favor acusar o recebimento deste email." & vbCrLf & vbCrLf & "Atenciosamente," & vbCrLf & vbCrLf & FunPrimeiraLetraMaiuscula(pubUsuario) & vbCrLf & TextoEmpresa & vbCrLf & TextoTel & TextoFax
    End With
End If

'If Custos_justificativa = True Then
'        With frmRelatorios_Custos_Prev_Real_Just
'            Txt_assunto = "Justificativa de gastos do centro de custo " & .txtCentroCusto
'            txt_Descricao = "Segue justificativa em anexo." & vbCrLf & "Favor acusar o recebimento." & vbCrLf & vbCrLf & "Atenciosamente," & vbCrLf & vbCrLf & FunPrimeiraLetraMaiuscula(pubUsuario)
'        End With
'End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviarEmail()
On Error GoTo tratar_erro

Acao = "enviar e-mail"
If Txt_assunto = "" Then
    NomeCampo = "o assunto"
    ProcVerificaAcao
    Txt_assunto.SetFocus
    Exit Sub
End If
If Txt_descricao = "" Then
    NomeCampo = "a descrição"
    ProcVerificaAcao
    Txt_descricao.SetFocus
    Exit Sub
End If
If txtEmail.Text = "" Then
    NomeCampo = "o e-mail"
    ProcVerificaAcao
    txtEmail.SetFocus
    Exit Sub
End If
Familiatext = "Assunto = '" & Txt_assunto & "', Descricao = '" & Txt_descricao & "', Anexo = '" & Txt_anexo & "', Nome_anexo = '" & Nome_anexo & "', Email = '" & txtEmail.Text & "'"
Permitido = True
'================================================
'Enviar pedido de compras
'================================================
If Compras_Pedido = True Then
ProcEnviarEmailCompras
End If
'================================================
'Enviar proposta comercial
'================================================
If Vendas_Proposta = True Then
ProcEnviarEmailVendas
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviarEmailVendas()
On Error GoTo tratar_erro

Dim MailSMTP As New MailSMTP.Email
Dim feitoEnvio As Boolean


'Cria o ArrayList de anexos
Dim anexos As Object
Set anexos = CreateObject("System.Collections.ArrayList")


'================================================
'Enviar proposta comercial
'================================================
If Vendas_Proposta = True Then
    With frmVendas_proposta
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select Razao, Telefone, Fax from Empresa where codigo = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            TextoEmpresa = FunPrimeiraLetraMaiuscula(IIf(IsNull(TBFI!Razao), "", TBFI!Razao))
            If IsNull(TBFI!telefone) = False And TBFI!telefone <> "" Then TextoTel = "Tel: " & IIf(IsNull(TBFI!telefone), "", TBFI!telefone)
            If IsNull(TBFI!Fax) = False And TBFI!Fax <> "" Then
                If TextoTel <> "" Then TextoFax = " - Fax: " Else TextoFax = "Fax: "
                TextoFax = TextoFax & IIf(IsNull(TBFI!Fax), "", TBFI!Fax)
            End If
        End If
        
        If FunVerifEnviarEmailOutlook(.Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)) = True Then
            ProcEnviarEmailAutomatico MAPISession1, MAPIMessages1, txtEmail.Text, Txt_assunto, Txt_descricao, Txt_anexo, Nome_anexo
        Else
        
        
If FunVerifConfigEmail("ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), "V", pubUsuario) = False Then Exit Sub
       
 If Left(Servidor_SMTP, 4) = "smtp" Then
            kmail.Charset = 0
            kmail.Priority = NORMAL_PRIORITY
            
            If kmail.sendEmail(Servidor_SMTP, pubUsuario, Usuario_email, txtCopia.Text, txtCopia.Text, txtEmail.Text, txtEmail.Text, Txt_assunto, Txt_descricao, Txt_anexo, False, "non HTML text", False, Porta_email, True, Usuario_email, Senha_email) Then
            
            If chkCopia.Value = 1 And txtCopia.Text <> "" Then
               If kmail.sendEmail(Servidor_SMTP, pubUsuario, Usuario_email, "", "", txtCopia.Text, txtCopia.Text, Txt_assunto, Txt_descricao, Txt_anexo, False, "non HTML text", False, Porta_email, True, Usuario_email, Senha_email) Then
               USMsgBox "Cópia do email enviado com sucesso!", vbInformation, "CAPRIND v5.0"
               ChkCopiaEnviada.Value = Checked
               End If
            End If
            chkEmailenviado.Value = Checked
            Else
                Permitido = False
                USMsgBox ("Ocorreu um erro ao enviar o e-mail."), vbExclamation, "CAPRIND v5.0"
            End If
            kmail.abort
 Else
        
        
        'Cria o ArrayList de Destinatarios
           Dim dests As Object
           Set dests = CreateObject("System.Collections.ArrayList")
           dests.Add txtEmail.Text & ";" & frmVendas_proposta.txtRemetente.Text
           dests.Add txtCopia.Text & ";" & pubUsuario

            
            anexos.Add Txt_anexo.Text
            'anexos.Add "C:\exemplo.txt"
            'Servidor_SMTP = "email-ssl.com.br"
            feitoEnvio = MailSMTP.sendEmail(Servidor_SMTP, Porta_email, Usuario_email, Usuario_email, dests, Txt_assunto, Txt_descricao, True, Usuario_email, Senha_email, anexos)
            
            If feitoEnvio = True Then
            Permitido = True
            chkEmailenviado.Value = 1

            '====================================================
            'Envia a cópia
            '====================================================
            If chkCopia.Value = 1 Then
            ChkCopiaEnviada.Value = 1
            End If
            Else
            Permitido = False
            USMsgBox "Não foi possivel enviar o email, por favor tente novamente.", vbCritical, "CAPRIND v5.0"
                Exit Sub
            End If
        End If
End If


        If Permitido = True Then
            USMsgBox ("E-mail enviado com sucesso."), vbInformation, "CAPRIND v5.0"
            '==================================
            Modulo = "Vendas/Proposta"
            Evento = "Enviar email"
            ID_documento = .txtId
            Documento = "Nº pedido: " & .txtCotacao
            Documento1 = ""
            ProcGravaEvento
            '==================================
            TextoCampo = ", Email_enviado = 'True'"
        Else
            TextoCampo = ""
        End If
        'Conexao.Execute "Update vendas_proposta Set " & Familiatext & TextoCampo & " where cotacao = " & .txtID
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviarEmailCompras()
On Error GoTo tratar_erro

Dim MailSMTP As New MailSMTP.Email
Dim feitoEnvio As Boolean


'Cria o ArrayList de anexos
Dim anexos As Object
Set anexos = CreateObject("System.Collections.ArrayList")

'================================================
'Enviar pedido de compras
'================================================
If Compras_Pedido = True Then
    With frmCompras_Pedido
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select Razao, Telefone, Fax from Empresa where codigo = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = False Then
            TextoEmpresa = FunPrimeiraLetraMaiuscula(IIf(IsNull(TBFI!Razao), "", TBFI!Razao))
            If IsNull(TBFI!telefone) = False And TBFI!telefone <> "" Then TextoTel = "Tel: " & IIf(IsNull(TBFI!telefone), "", TBFI!telefone)
            If IsNull(TBFI!Fax) = False And TBFI!Fax <> "" Then
                If TextoTel <> "" Then TextoFax = " - Fax: " Else TextoFax = "Fax: "
                TextoFax = TextoFax & IIf(IsNull(TBFI!Fax), "", TBFI!Fax)
            End If
        End If
        
        If FunVerifEnviarEmailOutlook(.Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)) = True Then
            ProcEnviarEmailAutomatico MAPISession1, MAPIMessages1, txtEmail.Text, Txt_assunto, Txt_descricao, Txt_anexo, Nome_anexo
        Else
        
        
If FunVerifConfigEmail("ID_empresa = " & .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), "C", pubUsuario) = False Then Exit Sub
       
If Left(Servidor_SMTP, 4) = "smtp" Then
            kmail.Charset = 0
            kmail.Priority = NORMAL_PRIORITY
            
            If kmail.sendEmail(Servidor_SMTP, pubUsuario, Usuario_email, txtCopia.Text, txtCopia.Text, txtEmail.Text, txtEmail.Text, Txt_assunto, Txt_descricao, Txt_anexo, False, "non HTML text", False, Porta_email, True, Usuario_email, Senha_email) Then
            
            If chkCopia.Value = 1 And txtCopia.Text <> "" Then
               If kmail.sendEmail(Servidor_SMTP, pubUsuario, Usuario_email, "", "", txtCopia.Text, txtCopia.Text, Txt_assunto, Txt_descricao, Txt_anexo, False, "non HTML text", False, Porta_email, True, Usuario_email, Senha_email) Then
               USMsgBox "Cópia do email enviado com sucesso!", vbInformation, "CAPRIND v5.0"
               ChkCopiaEnviada.Value = Checked
               End If
            End If
            chkEmailenviado.Value = Checked
            Else
                Permitido = False
                USMsgBox ("Ocorreu um erro ao enviar o e-mail."), vbExclamation, "CAPRIND v5.0"
            End If
            kmail.abort
 Else
        
        'Cria o ArrayList de Destinatarios
           Dim dests As Object
           Set dests = CreateObject("System.Collections.ArrayList")
           dests.Add txtEmail.Text & ";" & frmCompras_Pedido.txtContato.Text
           dests.Add txtCopia.Text & ";" & pubUsuario
        
            
            anexos.Add Txt_anexo.Text
            '=====================================================
            ' Envia o email
            '=====================================================
            feitoEnvio = MailSMTP.sendEmail(Servidor_SMTP, Porta_email, Usuario_email, Usuario_email, dests, Txt_assunto, Txt_descricao, True, Usuario_email, Senha_email, anexos)
            
            'Debug.print feitoEnvio
            '=====================================================
            ' Verifica resultado do envio
            '=====================================================
            If feitoEnvio = True Then
            
            
            Permitido = True
            chkEmailenviado.Value = 1
            '====================================================
            'Envia a cópia
            '====================================================
            If chkCopia.Value = 1 Then
            ChkCopiaEnviada.Value = 1
            End If
           
            Else
            Permitido = False
            USMsgBox "Não foi possivel enviar o email, por favor tente novamente.", vbCritical, "CAPRIND v5.0"
                Exit Sub
            End If
        End If
 End If
 
        If Permitido = True Then
            USMsgBox ("E-mail enviado com sucesso."), vbInformation, "CAPRIND v5.0"
            '==================================
            Modulo = "Compras/Pedido"
            Evento = "Enviar e-mail"
            ID_documento = .txtIDPedido
            Documento = "Nº pedido: " & .txtPedido
            Documento1 = ""
            ProcGravaEvento
            '==================================
            TextoCampo = ", Email_enviado = 'True'"
            .Chk_email_enviado.Value = 1
        Else
            TextoCampo = ""
            .Chk_email_enviado.Value = 0
        End If
        Conexao.Execute "Update Compras_pedido Set " & Familiatext & TextoCampo & " where IDpedido = " & .txtIDPedido
    End With
End If


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

Private Sub Form_Resize()
On Error GoTo tratar_erro

FrmEnviarEmail.chkCopia.Value = 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
