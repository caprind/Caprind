VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCompras_pedido_cancelar 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compras - Pedido - Status"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4620
   ClipControls    =   0   'False
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
   ForeColor       =   &H8000000D&
   Icon            =   "frmCompras_pedido_cancelar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   2910
      Top             =   180
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmCompras_pedido_cancelar.frx":0442
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   4545
      _ExtentX        =   8017
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
      ButtonKey3      =   "6"
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
      ButtonKey4      =   "7"
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
      ButtonKey5      =   "8"
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   825
      Left            =   60
      TabIndex        =   3
      Top             =   1020
      Visible         =   0   'False
      Width           =   4545
      Begin VB.TextBox txtResponsavel 
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
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   60
         MouseIcon       =   "frmCompras_pedido_cancelar.frx":223E
         MousePointer    =   99  'Custom
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Nome do fornecedor."
         Top             =   375
         Width           =   3135
      End
      Begin VB.TextBox txtData 
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
         Height          =   315
         Left            =   3330
         MaxLength       =   60
         MouseIcon       =   "frmCompras_pedido_cancelar.frx":2548
         MousePointer    =   99  'Custom
         TabIndex        =   4
         ToolTipText     =   "Cidade."
         Top             =   375
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3675
         TabIndex        =   7
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1290
         TabIndex        =   6
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   60
      TabIndex        =   1
      Top             =   1020
      Width           =   4545
      Begin VB.TextBox txtmotivo 
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
         Height          =   1635
         Left            =   210
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         ToolTipText     =   "Motivo."
         Top             =   390
         Width           =   4125
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo do cancelamento"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1402
         TabIndex        =   8
         Top             =   180
         Width           =   1740
      End
   End
End
Attribute VB_Name = "frmCompras_pedido_cancelar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 4545, 5, True
If Compras_Pedido = True Then
    With frmCompras_Pedido
        If .Compras_pedido_Prod = False And .Compras_pedido_serv = False Then
            If .txtStatus = "CANCELADO" Then
                procCancelado
                Set TBPedido = CreateObject("adodb.recordset")
                TBPedido.Open "Select * from compras_pedido where idpedido = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
                If TBPedido.EOF = False Then
                    txtResponsavel = IIf(IsNull(TBPedido!Resp_cancelado), "", TBPedido!Resp_cancelado)
                    txtData = IIf(IsNull(TBPedido!Data_cancelado), "", Format(TBPedido!Data_cancelado, "dd/mm/yy"))
                    txtMotivo = IIf(IsNull(TBPedido!Motivo_cancelado), "", TBPedido!Motivo_cancelado)
                End If
                TBPedido.Close
            End If
        ElseIf .Compras_pedido_Prod = True Then
                If .txtstatus_item = "CANCELADO" Then
                    procCancelado
                    Set TBPedido = CreateObject("adodb.recordset")
                    TBPedido.Open "Select * from compras_pedido_lista where idlista = " & .TXTIDLista, Conexao, adOpenKeyset, adLockOptimistic
                    If TBPedido.EOF = False Then
                        txtResponsavel = IIf(IsNull(TBPedido!Resp_cancelado), "", TBPedido!Resp_cancelado)
                        txtData = IIf(IsNull(TBPedido!Data_cancelado), "", Format(TBPedido!Data_cancelado, "dd/mm/yy"))
                        txtMotivo = IIf(IsNull(TBPedido!Motivo_cancelado), "", TBPedido!Motivo_cancelado)
                    End If
                    TBPedido.Close
                End If
            Else
                If .txtStatus_serv = "CANCELADO" Then
                    procCancelado
                    Set TBPedido = CreateObject("adodb.recordset")
                    TBPedido.Open "Select * from compras_pedido_lista where idlista = " & .txtIDLista_serv, Conexao, adOpenKeyset, adLockOptimistic
                    If TBPedido.EOF = False Then
                        txtResponsavel = IIf(IsNull(TBPedido!Resp_cancelado), "", TBPedido!Resp_cancelado)
                        txtData = IIf(IsNull(TBPedido!Data_cancelado), "", Format(TBPedido!Data_cancelado, "dd/mm/yy"))
                        txtMotivo = IIf(IsNull(TBPedido!Motivo_cancelado), "", TBPedido!Motivo_cancelado)
                    End If
                    TBPedido.Close
                End If
        End If
    End With
ElseIf Plano_centro_de_custo = True Then
        With Frm_centro_de_custo
            Caption = "Custos - Centro de custo - Status"
            Label2.Caption = "Motivo do bloqueio"
            Label2.Left = 1597
            contador = 0
            IDAntigo = 0
            
            For InitFor = 1 To .Lista.ListItems.Count
                If .Lista.ListItems.Item(InitFor).Checked = True Then
                    IDAntigo = .Lista.ListItems(InitFor)
                    contador = contador + 1
                End If
            Next InitFor
                    
            If contador = 1 Then
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select DtBloq, MotivoBloq, RespBloq from Usuarios_setor where ID = " & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    If IsNull(TBContas!DtBloq) = False Then procCancelado
                    txtData = IIf(IsNull(TBContas!DtBloq), "", Format(TBContas!DtBloq, "dd/mm/yy"))
                    txtResponsavel = IIf(IsNull(TBContas!RespBloq), "", TBContas!RespBloq)
                    txtMotivo = IIf(IsNull(TBContas!MotivoBloq), "", TBContas!MotivoBloq)
                End If
                TBContas.Close
            End If
        End With
    Else
        With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
            If Vendas_PI = True Then Caption = "Vendas - Pedido interno - Status" Else Caption = "Vendas - Proposta comercial - Status"
            
            contador = 0
            IDAntigo = 0
            
            If Sit_REG = 1 Then
                For InitFor = 1 To .Lista.ListItems.Count
                    If .Lista.ListItems.Item(InitFor).Checked = True Then
                        IDAntigo = .Lista.ListItems(InitFor)
                        contador = contador + 1
                    End If
                Next InitFor
            ElseIf Sit_REG = 2 Then
                For InitFor = 1 To .Listprod.ListItems.Count
                    If .Listprod.ListItems.Item(InitFor).Checked = True Then
                        IDAntigo = .Listprod.ListItems(InitFor)
                        contador = contador + 1
                    End If
                Next InitFor
            Else
                For InitFor = 1 To .ListaServicos.ListItems.Count
                    If .ListaServicos.ListItems.Item(InitFor).Checked = True Then
                        IDAntigo = .ListaServicos.ListItems(InitFor)
                        contador = contador + 1
                    End If
                Next InitFor
            End If
                    
            If contador = 1 Then
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select DtCancelado, MotivoCancelado, RespCancelado from " & IIf(Sit_REG = 1, "vendas_proposta where cotacao = ", "vendas_carteira where Codigo = ") & IDAntigo, Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    If IsNull(TBContas!DtCancelado) = False Then procCancelado
                    txtData = IIf(IsNull(TBContas!DtCancelado), "", Format(TBContas!DtCancelado, "dd/mm/yy"))
                    txtResponsavel = IIf(IsNull(TBContas!RespCancelado), "", TBContas!RespCancelado)
                    txtMotivo = IIf(IsNull(TBContas!MotivoCancelado), "", TBContas!MotivoCancelado)
                End If
                TBContas.Close
            End If
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

Sub ProcSalvar_SemLista()
On Error GoTo tratar_erro

With frmCompras_Pedido
    Acao = "salvar"
    If .Compras_pedido_Prod = False And .Compras_pedido_serv = False Then
        If txtMotivo.Text = "" And .txtStatus <> "CANCELADO" Then
            NomeCampo = "o motivo"
            ProcVerificaAcao
            If txtMotivo.Enabled = True Then txtMotivo.SetFocus
            Exit Sub
        End If
        
        Set TBPedido = CreateObject("adodb.recordset")
        TBPedido.Open "Select * from compras_pedido where IDpedido = " & frmCompras_Pedido.txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
        If TBPedido.EOF = False Then
            USMsgBox ("Status alterado com sucesso."), vbInformation, "CAPRIND v5.0"
            Evento = "Alterar status"
            If TBPedido!idcotacao = 0 Then
Pedido:
                If .txtStatus = "CANCELADO" Then
                    Conexao.Execute "Update compras_pedido_lista Set Status_Item = 'AGUARDANDO APROVAÇÃO', Resp_cancelado = null, Data_cancelado = null, motivo_cancelado = null  where IDpedido = " & IDlista
                    Conexao.Execute "Update compras_pedido Set Status_pedido = 'AGUARDANDO APROVAÇÃO', Resp_cancelado = null, Data_cancelado = null, motivo_cancelado = null where IDpedido = " & IDlista
                    .txtStatus = "AGUARDANDO APROVAÇÃO"
                Else
                    Conexao.Execute "Update compras_pedido_lista Set Status_Item = 'CANCELADO', Resp_cancelado = '" & pubUsuario & "', Data_cancelado = '" & Date & "', motivo_cancelado = '" & txtMotivo & "'  where IDpedido = " & IDlista
                    Conexao.Execute "Update compras_pedido Set Status_pedido = 'CANCELADO', Resp_cancelado = '" & pubUsuario & "', Data_cancelado = '" & Date & "', motivo_cancelado = '" & txtMotivo & "'  where IDpedido = " & IDlista
                    .txtStatus = "CANCELADO"
                End If
            Else
                Set TBPedido = CreateObject("adodb.recordset")
                TBPedido.Open "Select * from compras_pedido_lista where IDpedido = " & frmCompras_Pedido.txtIDPedido, Conexao, adOpenKeyset, adLockOptimistic
                If TBPedido.EOF = False Then
                    ProcCotacao
                Else
                    GoTo Pedido
                End If
            End If
        End If
        TBPedido.Close
        Documento1 = ""
        .ProcAtualizalistapedido (IIf(ReturnNumbersOnly(Left(.lblPaginas(3).Caption, Len(.lblPaginas(3).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(3).Caption, Len(.lblPaginas(3).Caption) - 5))))
    ElseIf .Compras_pedido_Prod = True Then
            If .txtstatus_item = "CANCELADO" And .txtResponsavel_aprovacao <> "" And .txtResponsavel_aprovacao <> pubUsuario Then
                USMsgBox ("Somente o usuário que aprovou o pedido pode alterar o status deste produto."), vbExclamation, "CAPRIND v5.0"
                Exit Sub
            End If
            If txtMotivo.Text = "" And .txtstatus_item <> "CANCELADO" Then
                NomeCampo = "o motivo"
                ProcVerificaAcao
                If txtMotivo.Enabled = True Then txtMotivo.SetFocus
                Exit Sub
            End If
            
            USMsgBox ("Status do produto alterado com sucesso."), vbInformation, "CAPRIND v5.0"
            Evento = "Alterar status do produto"
            If .txtstatus_item = "CANCELADO" Then
                If .txtResponsavel_aprovacao = "" Then
                    .txtstatus_item = .txtStatus
                    TextoFiltro = "AGUARDANDO APROVAÇÃO"
                Else
                    If .txtStatus = "COMPRADO" Then
                        .txtstatus_item = "COMPRADO"
                        TextoFiltro = "N_RECEBIDO"
                    Else
                        .txtstatus_item = "APROVADO"
                        TextoFiltro = "APROVADO"
                    End If
                End If
                Conexao.Execute "Update compras_pedido_lista Set Status_Item = '" & TextoFiltro & "', Resp_cancelado = Null, Data_cancelado = Null, motivo_cancelado = Null where IDlista = " & .TXTIDLista
            Else
                Conexao.Execute "Update compras_pedido_lista Set Status_Item = 'CANCELADO', Resp_cancelado = '" & pubUsuario & "', Data_cancelado = '" & Date & "', motivo_cancelado = '" & txtMotivo & "'  where IDlista = " & .TXTIDLista
                .txtstatus_item = "CANCELADO"
            End If
            FunAtualizaStatusPC frmCompras_Pedido.txtIDPedido
            Documento1 = "Cód. interno: " & .txtNomenclatura
            .ProcLimpaCamposItem False
            .ProcAtualizalista
            .ProcAtualizalistapedido (IIf(ReturnNumbersOnly(Left(.lblPaginas(3).Caption, Len(.lblPaginas(3).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(3).Caption, Len(.lblPaginas(3).Caption) - 5))))
        Else
            If .txtStatus_serv = "CANCELADO" And .txtResponsavel_aprovacao <> "" And txtResponsavel_aprovacao <> pubUsuario Then
                USMsgBox ("Somente o usuário que aprovou o pedido pode alterar o status deste serviço."), vbExclamation, "CAPRIND v5.0"
                Exit Sub
            End If
            If txtMotivo.Text = "" And .txtStatus_serv <> "CANCELADO" Then
                NomeCampo = "o motivo"
                ProcVerificaAcao
                If txtMotivo.Enabled = True Then txtMotivo.SetFocus
                Exit Sub
            End If
            
            USMsgBox ("Status do serviço alterado com sucesso."), vbInformation, "CAPRIND v5.0"
            Evento = "Alterar status do serviço"
            If .txtStatus_serv = "CANCELADO" Then
                If .txtResponsavel_aprovacao = "" Then
                    .txtStatus_serv = .txtStatus
                    TextoFiltro = "AGUARDANDO APROVAÇÃO"
                Else
                    .txtStatus_serv = "COMPRADO"
                    TextoFiltro = "N_RECEBIDO"
                End If
                Conexao.Execute "Update compras_pedido_lista Set Status_Item = '" & TextoFiltro & "', Resp_cancelado = Null, Data_cancelado = Null, motivo_cancelado = Null where IDlista = " & .txtIDLista_serv
            Else
                Conexao.Execute "Update compras_pedido_lista Set Status_Item = 'CANCELADO', Resp_cancelado = '" & pubUsuario & "', Data_cancelado = '" & Date & "', motivo_cancelado = '" & txtMotivo & "'  where IDlista = " & .txtIDLista_serv
                .txtStatus_serv = "CANCELADO"
            End If
            FunAtualizaStatusPC frmCompras_Pedido.txtIDPedido
            Documento1 = "Cód. interno: " & .txtCodigo
            .ProcLimpaCamposServ False
            .ProcAtualizalistaServ
            .ProcAtualizalistapedido (IIf(ReturnNumbersOnly(Left(.lblPaginas(3).Caption, Len(.lblPaginas(3).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(3).Caption, Len(.lblPaginas(3).Caption) - 5))))
    End If
    '==================================
    Modulo = "Compras/Pedido"
    ID_documento = IDlista
    Documento = "Nº pedido: " & .txtPedido.Text
    ProcGravaEvento
    '==================================
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcSalvar_ComLista(ListaCancelar_PI As ListView)
On Error GoTo tratar_erro

Evento = "Alterar status"
If Plano_centro_de_custo = False Then
    With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
        If Sit_REG = 1 Then TextoFiltro = "vendas_proposta where Cotacao" Else TextoFiltro = "vendas_carteira where codigo"
        For InitFor = 1 To ListaCancelar_PI.ListItems.Count
            If ListaCancelar_PI.ListItems.Item(InitFor).Checked = True Then
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select * from " & TextoFiltro & " = " & ListaCancelar_PI.ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    If IsNull(TBContas!DtCancelado) = False Then
                        TBContas!DtCancelado = Null
                        TBContas!RespCancelado = Null
                        TBContas!MotivoCancelado = Null
                        
                        'Atualiza o status dos itens
                        If Sit_REG = 1 Then
                            Conexao.Execute "Update vendas_carteira Set Liberacao = '" & IIf(Vendas_PI = True, "VENDIDA", "ABERTA EM ANALISE") & "', DtCancelado = Null, RespCancelado = Null, MotivoCancelado = Null where Liberacao = 'CANCELADO' and Cotacao = " & ListaCancelar_PI.ListItems.Item(InitFor)
                            Conexao.Execute "Update vendas_carteira Set Liberacao = 'FATURADO PARCIAL', DtCancelado = Null, RespCancelado = Null, MotivoCancelado = Null where Liberacao = 'FATURADO' and DtCancelado IS NOT NULL and Cotacao = " & ListaCancelar_PI.ListItems.Item(InitFor)
                        Else
                            If TBContas!Liberacao = "FATURADO" Then TBContas!Liberacao = "FATURADO PARCIAL" Else TBContas!Liberacao = IIf(Vendas_PI = True, "VENDIDA", "ABERTA EM ANALISE")
                        End If
                    Else
                        TBContas!DtCancelado = Date
                        TBContas!RespCancelado = pubUsuario
                        TBContas!MotivoCancelado = IIf(txtMotivo = "", Null, txtMotivo)
        
                        'Atualiza o status dos itens
                        If Sit_REG = 1 Then
                            Conexao.Execute "Update vendas_carteira Set Liberacao = 'CANCELADO', DtCancelado = '" & Date & "', RespCancelado = '" & pubUsuario & "', MotivoCancelado = '" & txtMotivo & "' where Liberacao = '" & IIf(Vendas_PI = True, "VENDIDA", "ABERTA EM ANALISE") & "' and Cotacao = " & ListaCancelar_PI.ListItems.Item(InitFor)
                            Conexao.Execute "Update vendas_carteira Set Liberacao = 'FATURADO', DtCancelado = '" & Date & "', RespCancelado = '" & pubUsuario & "', MotivoCancelado = '" & txtMotivo & "' where Liberacao = 'FATURADO PARCIAL' and Cotacao = " & ListaCancelar_PI.ListItems.Item(InitFor)
                        Else
                            If TBContas!Liberacao = "FATURADO PARCIAL" Then TBContas!Liberacao = "FATURADO" Else TBContas!Liberacao = "CANCELADO"
                        End If
                    End If
                    TBContas.Update
                    FunAtualizaStatusPropPI (TBContas!Cotacao)
                    
                    '==================================
                    Modulo = IIf(Vendas_PI = True, "Vendas/Pedido interno", "Vendas/Proposta comercial")
                    ID_documento = TBContas!Cotacao
                    If Sit_REG = 1 Then
                        Documento = IIf(Vendas_PI = True, "Nº pedido: ", "Nº proposta: ") & TBContas!Ncotacao & " - Rev.: " & TBContas!Revisao
                        Documento1 = ""
                    Else
                        Documento = IIf(Vendas_PI = True, "Nº pedido: ", "Nº proposta: ") & .txtCotacao & " - Rev.: " & .txtrevisao
                        Documento1 = "Cód. interno: " & IIf(IsNull(TBContas!Desenho), "", TBContas!Desenho)
                    End If
                    ProcGravaEvento
                    '==================================
                End If
                TBContas.Close
            End If
        Next InitFor
        
        If Sit_REG = 1 Then
            .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
            Set TBAbrir = CreateObject("adodb.recordset")
            StrSql = "Select VP.*, CL.CPF_CNPJ as CNPJ_CPF, CL.CEP as CEP, CL.RG_IE from vendas_proposta VP inner join Clientes CL on VP.IDcliente = CL.IDCliente where cotacao ="
            TBAbrir.Open StrSql & .txtID, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                .ProcPuxaDados
                .ProcPuxaTotais
            End If
            TBAbrir.Close
        ElseIf Sit_REG = 2 Then
            .ProcAtualizalistaProdutos (IIf(ReturnNumbersOnly(Left(.lblPaginas1.Caption, Len(.lblPaginas1.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas1.Caption, Len(.lblPaginas1.Caption) - 5))))
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select * from vendas_carteira where Codigo = " & .txtid_produto, Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = False Then .ProcPuxaDadosLista
            TBProduto.Close
        Else
            .ProcAtualizalistaServicos (IIf(ReturnNumbersOnly(Left(.lblPaginas2.Caption, Len(.lblPaginas2.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas2.Caption, Len(.lblPaginas2.Caption) - 5))))
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select * from vendas_carteira where Codigo = " & .txtid_servico, Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = False Then .ProcpuxadadoslistaServicos
            TBProduto.Close
        End If
    End With
Else
    With Frm_centro_de_custo
        For InitFor = 1 To ListaCancelar_PI.ListItems.Count
            If ListaCancelar_PI.ListItems.Item(InitFor).Checked = True Then
                Set TBContas = CreateObject("adodb.recordset")
                TBContas.Open "Select ID, CODIGO, Setor, DtBloq, RespBloq, MotivoBloq from Usuarios_Setor where ID = " & ListaCancelar_PI.ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBContas.EOF = False Then
                    If IsNull(TBContas!DtBloq) = False Then
                        TBContas!DtBloq = Null
                        TBContas!RespBloq = Null
                        TBContas!MotivoBloq = Null
                    Else
                        TBContas!DtBloq = Date
                        TBContas!RespBloq = pubUsuario
                        TBContas!MotivoBloq = IIf(txtMotivo = "", Null, txtMotivo)
                    End If
                    TBContas.Update
                    
                    '==================================
                    Modulo = "Custos/Centro de custo"
                    ID_documento = TBContas!ID
                    Documento = "Código: " & TBContas!CODIGO & " - Descrição: " & TBContas!Setor
                    Documento1 = ""
                    ProcGravaEvento
                    '==================================
                End If
                TBContas.Close
            End If
        Next InitFor
        
        .ProcCarregaLista
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select * from Usuarios_Setor where ID = " & .txtID, Conexao, adOpenKeyset, adLockOptimistic
        If TBLISTA.EOF = False Then
            .ProcCarregaDados
        End If
        TBLISTA.Close
    End With
End If
USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
Unload Me

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

Sub ProcCotacao()
On Error GoTo tratar_erro

IDConta = 0 'ID da cotação
With frmCompras_Pedido
    Set TBCompras_Lista = CreateObject("adodb.recordset")
    TBCompras_Lista.Open "Select * from compras_pedido_lista where idpedido = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Lista.EOF = False Then
        Do While TBCompras_Lista.EOF = False
            If .txtStatus = "CANCELADO" Then
                Evento = "Alterar status"
            
                Set TBCotacao = CreateObject("adodb.recordset")
                TBCotacao.Open "Select * from compras_cotacao where id_cotacao = " & IIf(IsNull(TBCompras_Lista!ID_cotacao), 0, TBCompras_Lista!ID_cotacao) & " and statuscotacao = 'APROVADA'", Conexao, adOpenKeyset, adLockOptimistic
                If TBCotacao.EOF = False Then
                    'Veirifica se possui novo pedido para cotação
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from compras_pedido where idcotacao = " & IIf(IsNull(TBCompras_Lista!ID_cotacao), 0, TBCompras_Lista!ID_cotacao) & " and IDpedido <> " & IIf(IsNull(TBCompras_Lista!IDpedido), 0, TBCompras_Lista!IDpedido) & " and Status_pedido <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        USMsgBox ("Não é possivel alterar o status, pois a cotação vinculada já possui outro pedido de compra."), vbExclamation, "CAPRIND v5.0"
                        Exit Sub
                    End If
                    TBFI.Close
                    
                    'Verificar se existe fornecedores aprovados na cotação
                    Set TBFornecedor = CreateObject("adodb.recordset")
                    TBFornecedor.Open "Select cotacao_item.IDitemlista, Cotacao_fornecedor.* FROM Cotacao_item INNER JOIN Cotacao_fornecedor ON Cotacao_item.ID = Cotacao_fornecedor.IDitem where Cotacao_fornecedor.IDcot = " & IIf(IsNull(TBCompras_Lista!ID_cotacao), 0, TBCompras_Lista!ID_cotacao) & " and Cotacao_fornecedor.aprovadoforn = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFornecedor.EOF = True Then
                        USMsgBox ("Não é possivel alterar o status, pois a cotação vinculada não possui fornecedor aprovado."), vbExclamation, "CAPRIND v5.0"
                        Exit Sub
                    End If
                    TBFornecedor.Close
                End If
                IDConta = IIf(IsNull(TBCompras_Lista!ID_cotacao), 0, TBCompras_Lista!ID_cotacao)
                
                Conexao.Execute "Update Compras_Cotacao Set Statuscotacao = 'APROVADA', dataaprovada = '" & .txtData & "' where id_cotacao = " & IIf(IsNull(TBCompras_Lista!ID_cotacao), 0, TBCompras_Lista!ID_cotacao)
                Conexao.Execute "UPDATE Compras_pedido_lista_custo Set IDpedido = " & .txtIDPedido & "  WHERE ID_requisicao = " & TBCompras_Lista!ID_Requisicao
                Conexao.Execute "Update compras_pedido Set Status_pedido = 'AGUARDANDO APROVAÇÃO', Resp_cancelado = null, Data_cancelado = null, motivo_cancelado = null where IDpedido = " & IDlista
                TBCompras_Lista.Delete
            Else
                Evento = "Cancelar"
                Conexao.Execute "Update compras_pedido Set Status_pedido = 'CANCELADO', Resp_cancelado = '" & pubUsuario & "', Data_cancelado = '" & Date & "', motivo_cancelado = '" & txtMotivo & "'  where IDpedido = " & IDlista
                If TBCompras_Lista!ID_cotacao <> 0 Then
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from compras_pedido_lista ", Conexao, adOpenKeyset, adLockOptimistic
                    TBFI.AddNew
                    TBFI!ID_cotacao = IIf(IsNull(TBCompras_Lista!ID_cotacao), 0, TBCompras_Lista!ID_cotacao)
                    TBFI!IDpedido = IIf(IsNull(TBCompras_Lista!IDpedido), 0, TBCompras_Lista!IDpedido)
                    TBFI!Status_Item = "CANCELADO"
                    TBFI!Codproduto = TBCompras_Lista!Codproduto
                    TBFI!Desenho = IIf(IsNull(TBCompras_Lista!Desenho), "", TBCompras_Lista!Desenho)
                    TBFI!Descricao = IIf(IsNull(TBCompras_Lista!Descricao), "", TBCompras_Lista!Descricao)
                    TBFI!Descricao_comercial = IIf(IsNull(TBCompras_Lista!Descricao_comercial), "", TBCompras_Lista!Descricao_comercial)
                    TBFI!detalheitem = TBCompras_Lista!detalheitem
                    TBFI!quant_req = IIf(IsNull(TBCompras_Lista!quant_req), 0, TBCompras_Lista!quant_req)
                    TBFI!Quant_Comp = TBCompras_Lista!Quant_Comp
                    TBFI!Desconto = TBCompras_Lista!Desconto
                    TBFI!ValorDesconto = TBCompras_Lista!ValorDesconto
                    TBFI!preco_unitario_desconto = TBCompras_Lista!preco_unitario_desconto
                    TBFI!preco_unitario = TBCompras_Lista!preco_unitario
                    TBFI!preco_total = TBCompras_Lista!preco_total
                    TBFI!IPI = TBCompras_Lista!IPI
                    TBFI!Familia = IIf(IsNull(TBCompras_Lista!Familia), "", TBCompras_Lista!Familia)
                    TBFI!Un = IIf(IsNull(TBCompras_Lista!Un), "", TBCompras_Lista!Un)
                    TBFI!Unidade_com = IIf(IsNull(TBCompras_Lista!Unidade_com), "", TBCompras_Lista!Unidade_com)
                    TBFI!ICMS = TBCompras_Lista!ICMS
                    TBFI!vlrICMS = TBCompras_Lista!vlrICMS
                    TBFI!VlrIPI = TBCompras_Lista!VlrIPI
                    TBFI!Prazo = TBCompras_Lista!Prazo
                    TBFI!Obs_pedido = TBCompras_Lista!Obs_pedido
                    TBFI!Remessa = TBCompras_Lista!Remessa
                    TBFI!Tipo = IIf(IsNull(TBCompras_Lista!Tipo), "", TBCompras_Lista!Tipo)
                    TBFI!ISSQN = TBCompras_Lista!ISSQN
                    TBFI!VlrISSQN = TBCompras_Lista!VlrISSQN
                    TBFI!Ordem = IIf(IsNull(TBCompras_Lista!Ordem), 0, TBCompras_Lista!Ordem)
                    TBFI!OS = IIf(IsNull(TBCompras_Lista!OS), 0, TBCompras_Lista!OS)
                    TBFI!Resp_cancelado = pubUsuario
                    TBFI!Data_cancelado = Date
                    TBFI!Motivo_cancelado = txtMotivo
                    TBFI!Prioridade = TBCompras_Lista!Prioridade
                    TBFI.Update
                    TBFI.Close
                End If
                Conexao.Execute "Update Compras_Cotacao Set Statuscotacao = 'LIBERADA', dataaprovada = null where id_cotacao = " & IIf(IsNull(TBCompras_Lista!ID_cotacao), 0, TBCompras_Lista!ID_cotacao)
                Conexao.Execute "UPDATE Compras_pedido_lista_custo Set IDpedido = 0 WHERE ID_requisicao = " & TBCompras_Lista!ID_Requisicao
            End If
            TBCompras_Lista.MoveNext
        Loop
        If .txtStatus <> "CANCELADO" Then
            Conexao.Execute "Update compras_pedido_lista Set Status_Item = 'COTANDO', IDPedido = 0, Resp_cancelado = Null, Data_cancelado = Null, motivo_cancelado = Null where IDpedido = " & IDlista & " and Status_item <> 'CANCELADO'"
            .txtStatus = "CANCELADO"
        Else
            Conexao.Execute "UPDATE Compras_pedido_lista Set IDpedido = " & .txtIDPedido & ", Status_Item = 'AGUARDANDO APROVAÇÃO' WHERE ID_cotacao = " & IDConta & " and IDpedido = 0"
            .txtStatus = "AGUARDANDO APROVAÇÃO"
        End If
    End If
    TBCompras_Lista.Close
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procCancelado()
On Error GoTo tratar_erro

Frame4.Visible = True
Frame2.Top = 1830
Frame2.Enabled = False
Height = 4425

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcSalvar()
On Error GoTo tratar_erro

If Compras_Pedido = True Then
    ProcSalvar_SemLista
ElseIf Plano_centro_de_custo = True Then
        ProcSalvar_ComLista Frm_centro_de_custo.Lista
    Else
        With IIf(Vendas_PI = True, frmVendas_PI, frmVendas_proposta)
            If Sit_REG = 1 Then
                ProcSalvar_ComLista .Lista
            ElseIf Sit_REG = 2 Then
                    ProcSalvar_ComLista .Listprod
                Else
                    ProcSalvar_ComLista .ListaServicos
            End If
        End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
