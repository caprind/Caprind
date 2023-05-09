VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form frm_Instituicoes2_compensar_cheque 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compensar cheque"
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   2040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   55
      TabIndex        =   1
      Top             =   -60
      Width           =   1935
      Begin VB.CommandButton Cmd_salvar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1410
         Picture         =   "frm_Instituicoes2_compensar_cheque.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salvar (F3)"
         Top             =   375
         Width           =   315
      End
      Begin MSComCtl2.DTPicker Txt_data 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         ToolTipText     =   "Data da compensação."
         Top             =   375
         Width           =   1245
         _ExtentX        =   2196
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
         Format          =   183304193
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
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
         Index           =   4
         Left            =   630
         TabIndex        =   2
         Top             =   180
         Width           =   345
      End
   End
End
Attribute VB_Name = "frm_Instituicoes2_compensar_cheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_salvar_Click()
On Error GoTo tratar_erro

Cheque = ""
Cheque1 = ""
If frm_Instituicoes.Cheques_Emitidos = True Then
    Permitido = False
    With frm_Instituicoes.Lst_cheque
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                If Permitido = False Then
                    If USMsgBox("Deseja realmente compensar este(s) cheque(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
                End If
1:
                Permitido = True
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select ID, ID_empresa, Txt_descricao from tbl_Instituicoes WHERE ID = " & frm_Instituicoes.txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    Cheque = "Cheque n. " & .ListItems(InitFor).ListSubItems(2)
                    If Cheque <> Cheque1 Then
                        '==================================
                        Modulo = "Financeiro/Instituições"
                        Evento = "Compensar cheque emitido"
                        ID_documento = .ListItems(InitFor)
                        Documento = "Cheque nº: " & .ListItems(InitFor).ListSubItems(2) & " - Instituição bancária: " & TBFI!Txt_descricao
                        Documento1 = ""
                        ProcGravaEvento
                        '==================================
                       
                        Conexao.Execute "Update tbl_Fluxo_de_caixa Set Data = '" & Txt_data & "', Hora = '" & Now & "', Bloqueado = 'False' where Operacao = 'Débito' and Instituicao = '" & TBFI!Txt_descricao & "' and ID_empresa = " & TBFI!ID_empresa & " and Descricao = '" & Cheque & "'"
                        Conexao.Execute "Update tbl_ContasPagar Set Data_movimentacao = '" & Txt_data & "' where Banco = '" & TBFI!Txt_descricao & "' and ID_empresa = " & TBFI!ID_empresa & " and NDoctoBaixa = '" & Cheque & "'"
                        
                        Set TBContas = CreateObject("adodb.recordset")
                        TBContas.Open "Select NDoctoBaixa from tbl_ContasPagar where IdIntConta = " & .ListItems(InitFor) & " and Status = 'DEPÓSITO EM CHEQUE'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBContas.EOF = False Then
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select IDFluxo, IDFluxo_rec from tbl_instituicoes_transf where NDoctoBaixa = '" & TBContas!NDoctoBaixa & "' and id_banco_rem = " & TBFI!ID & " and FormaBaixa = 'CHEQUE'", Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                'Corrige saldo do banco recebedor e compensa o cheque
                                Set TBFluxo = CreateObject("adodb.recordset")
                                TBFluxo.Open "Select ID_empresa, Instituicao, Valor from tbl_Fluxo_de_caixa where IDFluxo = " & TBAbrir!IDFluxo_Rec, Conexao, adOpenKeyset, adLockOptimistic
                                If TBFluxo.EOF = False Then
                                    Conexao.Execute "Update tbl_Fluxo_de_caixa Set Data = '" & Txt_data & "', Hora = '" & Now & "', Bloqueado = 'False' where Operacao = 'Crédito' and Instituicao = '" & TBFluxo!Instituicao & "' and ID_empresa = " & TBFluxo!ID_empresa & " and Descricao = '" & Cheque & "'"
                                    Set TBProduto = CreateObject("adodb.recordset")
                                    TBProduto.Open "Select Saldo from tbl_instituicoes where txt_descricao = '" & TBFluxo!Instituicao & "' and ID_empresa = " & TBFluxo!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                                    If TBProduto.EOF = False Then
                                        TBProduto!Saldo = Format(TBProduto!Saldo + TBFluxo!valor, "###,##0.00")
                                        TBProduto.Update
                                    End If
                                End If
                                'Corrige saldo do banco remetente
                                Set TBFluxo = CreateObject("adodb.recordset")
                                TBFluxo.Open "Select ID_empresa, Instituicao, Valor from tbl_Fluxo_de_caixa where IDFluxo = " & TBAbrir!IDFluxo, Conexao, adOpenKeyset, adLockOptimistic
                                If TBFluxo.EOF = False Then
                                    Set TBProduto = CreateObject("adodb.recordset")
                                    TBProduto.Open "Select Saldo from tbl_instituicoes where txt_descricao = '" & TBFluxo!Instituicao & "' and ID_empresa = " & TBFluxo!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                                    If TBProduto.EOF = False Then
                                        TBProduto!Saldo = Format(TBProduto!Saldo - TBFluxo!valor, "###,##0.00")
                                        frm_Instituicoes.txtSaldo = Format(TBProduto!Saldo, "###,##0.00")
                                        TBProduto.Update
                                    End If
                                    TBProduto.Close
                                End If
                                TBFluxo.Close
                            End If
                            TBAbrir.Close
                        Else
                            Set TBGravar = CreateObject("adodb.recordset")
                            TBGravar.Open "Select Saldo from tbl_instituicoes where txt_Descricao = '" & TBFI!Txt_descricao & "' and ID_empresa = " & TBFI!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                            If TBGravar.EOF = False Then
                                Cheque = "Cheque n. " & .ListItems(InitFor).ListSubItems(2)
                                Set TBFluxo = CreateObject("adodb.recordset")
                                TBFluxo.Open "Select Valor from tbl_Fluxo_de_caixa where Operacao = 'Débito' and Instituicao = '" & TBFI!Txt_descricao & "' and Descricao = '" & Cheque & "' and ID_empresa = " & TBFI!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                                If TBFluxo.EOF = False Then
                                    TBGravar!Saldo = Format(TBGravar!Saldo - TBFluxo!valor, "###,##0.00")
                                    frm_Instituicoes.txtSaldo = Format(TBGravar!Saldo, "###,##0.00")
                                End If
                                TBFluxo.Close
                                TBGravar.Update
                            End If
                            TBGravar.Close
                        End If
                        TBContas.Close
                    End If
                    Cheque1 = "Cheque n. " & .ListItems(InitFor).ListSubItems(2)
                End If
                TBFI.Close
            End If
        Next InitFor
    End With
    If Permitido = False Then
        USMsgBox ("Informe os cheque(s) antes de compensar."), vbExclamation, "CAPRIND v5.0"
    Else
        USMsgBox ("Cheque(s) compensados com sucesso."), vbInformation, "CAPRIND v5.0"
        frm_Instituicoes.ProcCarregaListaCheque
    End If
Else
    Permitido = False
    With frm_Instituicoes.Lista_cheque
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                If Permitido = False Then
                    If USMsgBox("Deseja realmente compensar este(s) cheque(s)?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 2 Else Exit Sub
                End If
2:
                Permitido = True
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select ID, ID_empresa, Txt_descricao from tbl_Instituicoes WHERE ID = " & frm_Instituicoes.txtCodBanco, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = False Then
                    Cheque = "Cheque n. " & .ListItems(InitFor).ListSubItems(2)
                    If Cheque <> Cheque1 Then
                        '==================================
                        Modulo = "Financeiro/Instituições"
                        Evento = "Compensar cheque recebido"
                        ID_documento = .ListItems(InitFor)
                        Documento = "Cheque nº: " & .ListItems(InitFor).ListSubItems(2) & " - Instituição bancária: " & TBFI!Txt_descricao
                        Documento1 = ""
                        ProcGravaEvento
                        '==================================
                        
                        Conexao.Execute "Update tbl_Fluxo_de_caixa Set Data = '" & Txt_data & "', Hora = '" & Now & "', Bloqueado = 'False' where Operacao = 'Crédito' and Instituicao = '" & TBFI!Txt_descricao & "' and Descricao = '" & Cheque & "'"
                        Conexao.Execute "Update tbl_contas_receber Set Data_movimentacao = '" & Txt_data & "' where Banco = '" & TBFI!Txt_descricao & "' and ID_empresa = " & TBFI!ID_empresa & " and NDoctoBaixa = '" & Cheque & "'"
                        
                        Set TBGravar = CreateObject("adodb.recordset")
                        TBGravar.Open "Select Saldo from tbl_instituicoes where txt_Descricao = '" & TBFI!Txt_descricao & "' and ID_empresa = " & TBFI!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                        If TBGravar.EOF = False Then
                            Cheque = "Cheque n. " & .ListItems(InitFor).ListSubItems(2)
                            Set TBFluxo = CreateObject("adodb.recordset")
                            TBFluxo.Open "Select Valor from tbl_Fluxo_de_caixa where Operacao = 'Crédito' and Instituicao = '" & TBFI!Txt_descricao & "' and Descricao = '" & Cheque & "' and ID_empresa = " & TBFI!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
                            If TBFluxo.EOF = False Then
                                TBGravar!Saldo = Format(TBGravar!Saldo + TBFluxo!valor, "###,##0.00")
                                frm_Instituicoes.txtSaldo = Format(TBGravar!Saldo, "###,##0.00")
                            End If
                            TBFluxo.Close
                            TBGravar.Update
                        End If
                        TBGravar.Close
                    End If
                    Cheque1 = "Cheque n. " & .ListItems(InitFor).ListSubItems(2)
                End If
                TBFI.Close
            End If
        Next InitFor
    End With
    If Permitido = False Then
        USMsgBox ("Informe os cheque(s) antes de compensar."), vbExclamation, "CAPRIND v5.0"
    Else
        USMsgBox ("Cheque(s) compensados com sucesso."), vbInformation, "CAPRIND v5.0"
        frm_Instituicoes.ProcCarregaListaCheque
    End If
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: Cmd_salvar_Click
    'Case vbKeyF1: Ajuda
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Txt_data = Date

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
