VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmEstoque_ImportarNFe_item 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estoque - Recebimento - Importar XML - Cadastrar dados do recebimento"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   60
      TabIndex        =   8
      Top             =   990
      Width           =   10215
      Begin VB.ComboBox Cmb_codigo_ref 
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
         Left            =   1890
         Sorted          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Codigo de referência."
         Top             =   390
         Width           =   1890
      End
      Begin VB.TextBox txtUN 
         Alignment       =   2  'Center
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
         Height          =   335
         Left            =   9645
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   390
         Width           =   390
      End
      Begin VB.TextBox txtDescricao 
         Alignment       =   2  'Center
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
         Height          =   335
         Left            =   3810
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   390
         Width           =   5805
      End
      Begin VB.TextBox txtCodInterno 
         Alignment       =   2  'Center
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
         Height          =   335
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   390
         Width           =   1695
      End
      Begin VB.ComboBox cmbLocal_armaz 
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
         Left            =   3435
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Local de armazenamento."
         Top             =   1050
         Width           =   5310
      End
      Begin VB.TextBox txtCorrida 
         Alignment       =   2  'Center
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
         Height          =   335
         Left            =   180
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Número da corrida."
         Top             =   1050
         Width           =   1635
      End
      Begin VB.TextBox txtCertificado 
         Alignment       =   2  'Center
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
         Height          =   335
         Left            =   1837
         MaxLength       =   50
         TabIndex        =   5
         ToolTipText     =   "Número do certificado."
         Top             =   1050
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker Txt_data_recebimento 
         Height          =   315
         Left            =   8760
         TabIndex        =   7
         ToolTipText     =   "Data do recebimento."
         Top             =   1050
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   488308737
         CurrentDate     =   39057
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Código interno"
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
         Left            =   472
         TabIndex        =   17
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Código de ref."
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
         Left            =   2318
         TabIndex        =   16
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
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
         Left            =   6367
         TabIndex        =   15
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Un."
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
         Left            =   9713
         TabIndex        =   14
         Top             =   180
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Data receb."
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
         Left            =   8970
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Local armazenamento"
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
         Left            =   5310
         TabIndex        =   12
         Top             =   840
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Certificado"
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
         Left            =   2235
         TabIndex        =   11
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Corrida"
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
         Left            =   735
         TabIndex        =   10
         Top             =   840
         Width           =   525
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
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
      ButtonKey5      =   "5"
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   2550
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmEstoque_ImportarNFe_item.frx":0000
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmEstoque_ImportarNFe_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcGravar()
On Error GoTo tratar_erro

Certificado = txtCertificado
Corrida = txtcorrida
DesenhoProduto = cmbLocal_armaz
Data_Prog = Txt_data_recebimento
Permitido = True
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
    
Select Case KeyCode
    Case vbKeyF3: ProcGravar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEsc: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 10215, 4, True
Certificado = ""
Corrida = ""
DesenhoProduto = ""
Data_Prog = ""
Txt_data_recebimento = Date
Proccarregalocarm
procCarregarProduto

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

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcGravar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Proccarregalocarm()
On Error GoTo tratar_erro

cmbLocal_armaz.Clear
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select ELC.Descricao from Estoque_Localarmazenamento_criar ELC INNER JOIN Estoque_Localarmazenamento EL ON ELC.ID = EL.idemb_locarm where EL.codinterno = '" & TBCompras_Lista!Desenho & "' and ELC.Descricao is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
    Do While TBAliquota.EOF = False
        If IsNull(TBAliquota!Descricao) = False Then cmbLocal_armaz.AddItem TBAliquota!Descricao
        TBAliquota.MoveNext
    Loop
    
    'Se for serviço, mão de obra ou remessa cria um local de armz padrão
    DesenhoProduto = ""
    Set TBCodigoDesc = CreateObject("adodb.recordset")
    TBCodigoDesc.Open "Select Remessa, Tipo, ID_CFOP from Compras_pedido_lista where IDlista = " & TBCompras_Lista!IDlista, Conexao, adOpenKeyset, adLockOptimistic
    If TBCodigoDesc.EOF = False Then
        If TBCodigoDesc!Tipo = "S" Then
            DesenhoProduto = "SERVIÇOS"
        ElseIf TBCodigoDesc!Remessa = True Then
                DesenhoProduto = "RETORNO DE MERCADORIA"
            ElseIf IsNull(TBCodigoDesc!ID_CFOP) = False And TBCodigoDesc!ID_CFOP <> "" And TBCodigoDesc!ID_CFOP <> "0" Then
                    Set TBCFOP = CreateObject("adodb.recordset")
                    TBCFOP.Open "Select IDCountCfop from tbl_NaturezaOperacao where IDCountCfop = " & TBCodigoDesc!ID_CFOP & " and MaoObra = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBCFOP.EOF = False Then
                        DesenhoProduto = "INDUSTRIALIZAÇÃO"
                    End If
                    TBCFOP.Close
        End If
    End If
    TBCodigoDesc.Close
    If DesenhoProduto <> "" Then cmbLocal_armaz.AddItem DesenhoProduto
Else
    ProcCarregaComboLA cmbLocal_armaz, False, False
End If
TBAliquota.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procCarregarProduto()
On Error GoTo tratar_erro

txtCodinterno = ""
Cmb_codigo_ref.Clear
txtdescricao = ""
txtUN = ""
Set TBAliquota = CreateObject("adodb.recordset")
TBAliquota.Open "Select Codproduto, desenho, Descricao, unidade from projproduto where desenho = '" & TBCompras_Lista!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAliquota.EOF = False Then
    txtCodinterno = TBCompras_Lista!Desenho
    txtdescricao = IIf(IsNull(TBAliquota!Descricao), "", TBAliquota!Descricao)
    txtUN = IIf(IsNull(TBAliquota!Unidade), "", TBAliquota!Unidade)
    ProcCarregaComboCodRef Cmb_codigo_ref, "P.codproduto = " & TBAliquota!Codproduto, 0, "", False, True
End If
TBAliquota.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
