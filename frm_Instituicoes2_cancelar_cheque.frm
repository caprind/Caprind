VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frm_Instituicoes2_cancelar_cheque 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Financeiro - Instituições - Cancelar cheque"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4230
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
   Icon            =   "frm_Instituicoes2_cancelar_cheque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Motivo do cancelamento do cheque"
      Height          =   2115
      Left            =   55
      TabIndex        =   1
      Top             =   990
      Width           =   4125
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
         Height          =   1725
         Left            =   210
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         ToolTipText     =   "Motivo do cancelamento do cheque."
         Top             =   270
         Width           =   3705
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   3060
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frm_Instituicoes2_cancelar_cheque.frx":0442
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   4125
      _ExtentX        =   7276
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
      ButtonCaption1  =   "Cancelar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Cancelar (F4)"
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
      ButtonWidth1    =   50
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
      ButtonLeft2     =   54
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
      ButtonLeft3     =   58
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
      ButtonLeft4     =   96
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
      ButtonLeft5     =   124
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
   End
End
Attribute VB_Name = "frm_Instituicoes2_cancelar_cheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF4: ProcCancelar
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

ProcCarregaToolBar1 Me, 4125, 5, True

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

Sub ProcCancelar()
On Error GoTo tratar_erro

If txtMotivo.Text = "" Then
    USMsgBox ("Informe o motivo antes de cancelar este(s) cheque(s)."), vbExclamation, "CAPRIND v5.0"
    txtMotivo.SetFocus
    Exit Sub
End If
Permitido = False
With frm_Instituicoes.Lst_cheque
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente cancelar este(s) cheque(es)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from tbl_Instituicoes WHERE ID = " & frm_Instituicoes.txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                '==================================
                Modulo = "Financeiro/Instituições"
                Evento = "Cancelar cheque emitido"
                ID_documento = .ListItems(InitFor)
                Documento = "Cheque nº: " & .ListItems(InitFor).ListSubItems(2) & " - Instituição bancária: " & TBFI!Txt_descricao
                Documento1 = ""
                ProcGravaEvento
                '==================================
            
                Set TBFIltro = CreateObject("adodb.recordset")
                TBFIltro.Open "Select * from tbl_ContasPagar where idintconta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                If TBFIltro.EOF = False Then
                    If IsNull(TBFIltro!tituloref) = True Or TBFIltro!tituloref = "" Then tituloref = 0 Else tituloref = TBFIltro!tituloref
                    
                    'Verifica se a conta paga parcial já está liquidada
                    Set TBContas = CreateObject("adodb.recordset")
                    TBContas.Open "Select * from tbl_contaspagar where idintconta = " & tituloref & " and parcial = 'True' and tituloref <> '" & TBFIltro!IDintconta & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBContas.EOF = False Then
                        ProcCriaNovaConta
                        ProcCriaChequeCancelado
                        Set TBCorretiva = CreateObject("adodb.recordset")
                        TBCorretiva.Open "Select * from tbl_contaspagar where idintconta = " & TBFIltro!tituloref, Conexao, adOpenKeyset, adLockOptimistic
                        If TBCorretiva.EOF = False Then
                            ValorParcial = TBFIltro!ValorPago
                            Pendente = TBCorretiva!dbl_valorpagto
                            TBCorretiva!dbl_valorpagto = (Pendente + ValorParcial)
                            
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select * from tbl_contaspagar where tituloref = '" & TBFIltro!tituloref & "' and idintconta <> " & TBFIltro!tituloref & " and idintconta <> " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                If TBCorretiva!Bloqueado = False Then TBCorretiva!status = "TÍTULO PAGO PARCIAL"
                            Else
                                If TBCorretiva!Bloqueado = False Then TBCorretiva!status = "TÍTULO EM ABERTO"
                                TBCorretiva!Parcial = False
                                TBCorretiva!pagoparcial = 0
                                TBCorretiva!ValorPendente = 0
                                TBCorretiva!tituloref = ""
                                TBCorretiva!valorprincipal = 0
                            End If
                            TBAbrir.Close
                            
                            'Fluxo de Caixa
                            Cheque = "Cheque n. " & .ListItems(InitFor).ListSubItems(2)
                            Set TBFluxo = CreateObject("adodb.recordset")
                            TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where Operacao = 'Débito' and Instituicao = '" & TBFI!Txt_descricao & "' and ID_empresa = " & TBFI!ID_empresa & " and Descricao = '" & Cheque & "'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBFluxo.EOF = False Then
                                TBFluxo!valor = TBFluxo!valor - .ListItems(InitFor).ListSubItems(4)
                                TBFluxo.Update
                                If TBFluxo!valor = 0 Then TBFluxo.Delete
                            End If
                            TBFluxo.Close
                            Set TBFluxo = CreateObject("adodb.recordset")
                            TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                            If TBFluxo.EOF = True Then TBFluxo.AddNew
                            TBFluxo!Operacao = "À Debitar"
                            TBFluxo!Data = TBCorretiva!dt_Pagamento
                            TBFluxo!valor = TBCorretiva!dbl_valorpagto
                            TBFluxo!Descricao = TBCorretiva!Txt_fornecedor
                            TBFluxo!status = "N"
                            TBFluxo!int_NotaFiscal = TBCorretiva!txt_ndocumento
                            TBCorretiva!IDFluxo = TBFluxo!IDFluxo
                            TBFluxo!Instituicao = Null
                            TBFluxo!Hora = Null
                            TBFluxo!Cheque = 0
                            TBFluxo!Bloqueado = False
                            TBFluxo!ID_empresa = TBCorretiva!ID_empresa
                            TBFluxo.Update
                            TBFluxo.Close
                        End If
                        TBCorretiva.Update
                        TBCorretiva.Close
                        
                        Set TBFamilia = CreateObject("adodb.recordset")
                        TBFamilia.Open "select * from familia_financeiro where idconta = " & tituloref & " and tipoconta = 'P' order by ID_PC", Conexao, adOpenKeyset, adLockOptimistic
                        If TBFamilia.EOF = False Then
                            Do While TBFamilia.EOF = False
                                Set TBCiclo = CreateObject("adodb.recordset")
                                TBCiclo.Open "Select * from familia_financeiro where IDConta = " & .ListItems(InitFor) & " and ID_PC = " & TBFamilia!ID_PC & " and tipoconta = 'P'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBCiclo.EOF = False Then
                                    TBFamilia!valor = TBFamilia!valor + ValorParcial
                                    TBFamilia.Update
                                    TBCiclo.Delete
                                End If
                                TBCiclo.Close
                                TBFamilia.MoveNext
                            Loop
                        End If
                        TBFamilia.Close
                        
                        Set TBCorretiva = CreateObject("adodb.recordset")
                        TBCorretiva.Open "Select * from tbl_ContasPagar where idintconta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBCorretiva.EOF = False Then
                            'Fluxo de Caixa
                            Conexao.Execute "DELETE from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo)
                            
                            TBCorretiva.Delete
                        End If
                        TBCorretiva.Close
                    Else
                        ProcCriaNovaConta
                        ProcCriaChequeCancelado
                        Set TBCorretiva = CreateObject("adodb.recordset")
                        TBCorretiva.Open "Select * from tbl_contaspagar where idintconta = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
                        If TBCorretiva.EOF = False Then
                            status = TBCorretiva!status
                            
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select * from tbl_contaspagar where tituloref = '" & tituloref & "'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                If TBCorretiva!Bloqueado = False Then TBCorretiva!status = "TÍTULO PAGO PARCIAL"
                            Else
                                If TBCorretiva!Bloqueado = False Then TBCorretiva!status = "TÍTULO EM ABERTO"
                                TBCorretiva!Parcial = False
                                TBCorretiva!pagoparcial = 0
                                TBCorretiva!ValorPendente = 0
                                TBCorretiva!tituloref = ""
                                TBCorretiva!valorprincipal = 0
                            End If
                            TBAbrir.Close
                                           
                            'Fluxo de Caixa
                            If status <> "DEPÓSITO EM CHEQUE" Then
                                Cheque = "Cheque n. " & .ListItems(InitFor).ListSubItems(2)
                                Set TBFluxo = CreateObject("adodb.recordset")
                                TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where Operacao = 'Débito' and Instituicao = '" & TBFI!Txt_descricao & "' and ID_empresa = " & TBFI!ID_empresa & " and Descricao = '" & Cheque & "'", Conexao, adOpenKeyset, adLockOptimistic
                                If TBFluxo.EOF = False Then
                                    TBFluxo!valor = TBFluxo!valor - .ListItems(InitFor).ListSubItems(4)
                                    TBFluxo.Update
                                    If TBFluxo!valor = 0 Then TBFluxo.Delete
                                End If
                                TBFluxo.Close
                                Set TBFluxo = CreateObject("adodb.recordset")
                                TBFluxo.Open "Select * from tbl_Fluxo_de_caixa where IDFluxo = " & IIf(IsNull(TBCorretiva!IDFluxo), 0, TBCorretiva!IDFluxo), Conexao, adOpenKeyset, adLockOptimistic
                                If TBFluxo.EOF = True Then TBFluxo.AddNew
                                TBFluxo!Operacao = "À Debitar"
                                TBFluxo!Data = TBCorretiva!dt_Pagamento
                                TBFluxo!valor = TBCorretiva!dbl_valorpagto
                                TBFluxo!Descricao = TBCorretiva!Txt_fornecedor
                                TBFluxo!status = "N"
                                TBFluxo!int_NotaFiscal = TBCorretiva!txt_ndocumento
                                TBCorretiva!IDFluxo = TBFluxo!IDFluxo
                                TBFluxo!Instituicao = Null
                                TBFluxo!Hora = Null
                                TBFluxo!Cheque = 0
                                TBFluxo!Bloqueado = False
                                TBFluxo!ID_empresa = TBCorretiva!ID_empresa
                                TBFluxo.Update
                                TBFluxo.Close
                            End If
                                        
                            TBCorretiva!Logsit = "N"
                            TBCorretiva!DataBaixa = Null
                            TBCorretiva!Bom_para = Null
                            TBCorretiva!ValorPago = 0
                            TBCorretiva!NDoctoBaixa = ""
                            TBCorretiva!Banco = ""
                            TBCorretiva!Obs = ""
                            TBCorretiva!Favorecido = ""
                            TBCorretiva!Obscheque = ""
                            TBCorretiva!Dias_atraso = 0
                            TBCorretiva!Juros = 0
                            TBCorretiva!Juros_valor = 0
                            TBCorretiva!Multa = 0
                            TBCorretiva!Multa_valor = 0
                            TBCorretiva!Desconto = 0
                            TBCorretiva!Desconto_valor = 0
                            TBCorretiva.Update
                            
                            Conexao.Execute "DELETE from familia_financeiro where IDconta = " & TBCorretiva!IDintconta & " and Pago_recebido = 'True' and tipoconta = 'P' and Deposito_transf = 'False'"
                            Conexao.Execute "Update familia_financeiro Set Pago_recebido = 'False' where idconta = " & TBCorretiva!IDintconta & " and tipoconta = 'P'"
                            
                            If status = "DEPÓSITO EM CHEQUE" Then TBCorretiva.Delete
                        
                        End If
                        TBCorretiva.Close
                    End If
                    TBContas.Close
                End If
                TBFIltro.Close
                
                'Exclui cheque da tabela de depósitos, transferencias e saque
                Conexao.Execute "DELETE from tbl_instituicoes_transf where id_banco_rem = '" & TBFI!ID & "' and FormaBaixa = 'CHEQUE' and NDoctoBaixa = '" & .ListItems(InitFor).ListSubItems(2) & "'"
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) cheque(s) antes de cancelar."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Cheque(s) cancelado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    frm_Instituicoes.ProcCarregaListaCheque
    frm_Instituicoes.Frame7.Enabled = False
    Unload Me
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriaNovaConta()
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_ContasPagar", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
TBGravar!int_codforn = TBFIltro!int_codforn
TBGravar!Logsit = Null
TBGravar!Txt_fornecedor = TBFIltro!Txt_fornecedor
TBGravar!FormaBaixa = TBFIltro!FormaBaixa
TBGravar!DataBaixa = TBFIltro!DataBaixa
TBGravar!Bom_para = TBFIltro!Bom_para
TBGravar!ValorPago = TBFIltro!ValorPago
TBGravar!NDoctoBaixa = TBFIltro!NDoctoBaixa
TBGravar!Banco = TBFIltro!Banco
TBGravar!Obs = TBFIltro!Obs
TBGravar!Favorecido = TBFIltro!Favorecido
TBGravar!Obscheque = TBFIltro!Obscheque
TBGravar!impresso = TBFIltro!impresso
TBGravar!status = "CHEQUE CANCELADO"
TBGravar!ID_empresa = TBFIltro!ID_empresa
TBGravar.Update
IDlista = TBGravar!IDintconta
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriaChequeCancelado()
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Cheques_Cancelados", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
TBGravar!Data_cancelamento = Date
TBGravar!Responsavel = pubUsuario
TBGravar!ID_conta = IDlista
TBGravar!motivo = txtMotivo
TBGravar.Update
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcCancelar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
