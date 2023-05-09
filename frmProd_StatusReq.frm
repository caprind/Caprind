VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmprod_StatusReq 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gerenciamento de ordem - Status"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3690
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00000000&
      Height          =   1965
      Left            =   55
      TabIndex        =   4
      Top             =   0
      Width           =   3585
      Begin DrawSuite2022.USButton Cmd_cancelado 
         Height          =   360
         Left            =   180
         TabIndex        =   2
         Top             =   1050
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   635
         Caption         =   "Alterar para CANCELADO"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   1
      End
      Begin DrawSuite2022.USButton Cmd_mov_estoque 
         Height          =   360
         Left            =   180
         TabIndex        =   3
         Top             =   1470
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   635
         Caption         =   "Alterar conf. movimentação do estoque"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   1
      End
      Begin DrawSuite2022.USButton Cmd_sim 
         Height          =   360
         Left            =   180
         TabIndex        =   0
         Top             =   180
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   635
         Caption         =   "Alterar para SIM"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   1
      End
      Begin DrawSuite2022.USButton Cmd_nao 
         Height          =   360
         Left            =   180
         TabIndex        =   1
         Top             =   600
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   635
         Caption         =   "Alterar para NÃO"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         PicSizeH        =   48
         PicSizeW        =   48
         Theme           =   1
      End
   End
End
Attribute VB_Name = "frmprod_StatusReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_sim_Click()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With frmprod.ListaRequisicao
    If .ListItems.Count = 0 Then Exit Sub
    If USMsgBox("Deseja realmente alterar o status deste material " & .SelectedItem.SubItems(1) & " para SIM?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        If .SelectedItem.ListSubItems(8) = "SIM" Then
            USMsgBox "Não é permitido alterar o status do material, pois o mesmo já foi baixado integralmente.", vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
        Conexao.Execute "Update Producaomaterial Set Saida = 'SIM' where Idmateriaprima = " & .SelectedItem
        Evento = "Alterar status do material requisitado para sim"
        ProcSalvarEvento .SelectedItem, .SelectedItem.ListSubItems(1)
    End If
End With
Unload Me
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_nao_Click()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With frmprod.ListaRequisicao
    If .ListItems.Count = 0 Then Exit Sub
    If USMsgBox("Deseja realmente alterar o status deste material " & .SelectedItem.SubItems(1) & " para NÃO?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        If .SelectedItem.ListSubItems(8) = "NÃO" Then
            USMsgBox "Não é permitido alterar o status do material, pois o mesmo não foi baixado.", vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
        Conexao.Execute "Update Producaomaterial Set Saida = 'NÃO' where Idmateriaprima = " & .SelectedItem
        Evento = "Alterar status do material requisitado para não"
        ProcSalvarEvento .SelectedItem, .SelectedItem.ListSubItems(1)
    End If
End With
Unload Me
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_cancelado_Click()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With frmprod.ListaRequisicao
    If .ListItems.Count = 0 Then Exit Sub
    If USMsgBox("Deseja realmente alterar o status deste material " & .SelectedItem.SubItems(1) & " para CANCELADO?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        If .SelectedItem.ListSubItems(8) = "SIM" Then
            USMsgBox "Não é permitido alterar o status do material, pois o mesmo já foi baixado integralmente.", vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
        If .SelectedItem.ListSubItems(8) = "CANCEL." Then
            USMsgBox "Não é permitido alterar o status do material, pois o mesmo já foi cancelado.", vbExclamation, "CAPRIND v5.0"
            Exit Sub
        End If
        Conexao.Execute "Update Producaomaterial Set Saida = 'CANCEL.' where Idmateriaprima = " & .SelectedItem
        Evento = "Alterar status do material requisitado para cancelado"
        ProcSalvarEvento .SelectedItem, .SelectedItem.ListSubItems(1)
    End If
End With
Unload Me
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_mov_estoque_Click()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With frmprod.ListaRequisicao
    If .ListItems.Count = 0 Then Exit Sub
    If USMsgBox("Deseja realmente alterar o status deste material " & .SelectedItem.SubItems(1) & "?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select SUM(Saida) as quantidade from estoque_movimentacao where oe = '" & frmprod.txtof.Text & "' and desenho = '" & .SelectedItem.ListSubItems(1) & "' and documento = '" & frmprod.txtof & "' and (operacao = 'SAIDA_ORDEM' or operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Qtd = IIf(IsNull(TBAbrir!quantidade), 0, TBAbrir!quantidade)
            Qtde = .SelectedItem.SubItems(3)
            If Qtd = 0 Then
                SttusTexto = "NÃO"
            ElseIf Qtd >= Qtde Then
                    SttusTexto = "SIM"
                Else
                    SttusTexto = "PARCIAL"
            End If
        End If
        Conexao.Execute "UPDATE producaomaterial set Saida = '" & SttusTexto & "' where idmateriaprima = " & .SelectedItem
        '==================================
        Evento = "Alterar status do material requisitado"
        ProcSalvarEvento .SelectedItem, .SelectedItem.ListSubItems(1)
    End If
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
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarEvento(ID_documento As Long, Codinterno As String)
On Error GoTo tratar_erro

USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
With frmprod
    '==================================
    Modulo = "PCP/Gerenciamento de ordem"
    ID_documento = ID_documento
    Documento = "Ordem: " & .txtof.Text & " - Cód. interno: " & .txtdesenho
    Documento1 = "Cód. interno: " & Codinterno
    ProcGravaEvento
    '==================================
    .ProcCarregaListaRequisicao
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
