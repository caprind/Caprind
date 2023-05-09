VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmprod_Status 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Gerenciamento de ordem | Alterar status"
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   2460
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   820
      DibPicture      =   "frmprod_Status.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmprod_Status.frx":9AAD
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   1665
      Left            =   450
      TabIndex        =   0
      Top             =   630
      Width           =   4005
      Begin DrawSuite2022.USButton cmdMRP 
         Height          =   420
         Left            =   390
         TabIndex        =   1
         Top             =   240
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   741
         Caption         =   "Marcar\desmarcar como MRP gerado"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         PicSizeH        =   48
         PicSizeW        =   48
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USButton cmdDesmarcarExpedir 
         Height          =   420
         Left            =   390
         TabIndex        =   2
         Top             =   1080
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   741
         Caption         =   "Desmarcar como expedido"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         PicSizeH        =   48
         PicSizeW        =   48
         ShowFocusRect   =   0   'False
      End
      Begin DrawSuite2022.USButton cmdExpedir 
         Height          =   420
         Left            =   390
         TabIndex        =   3
         Top             =   660
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   741
         Caption         =   "Marcar como expedido"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         PicSizeH        =   48
         PicSizeW        =   48
         ShowFocusRect   =   0   'False
      End
   End
End
Attribute VB_Name = "frmprod_Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDesmarcarExpedir_Click()
On Error GoTo tratar_erro
    
With frmprod
    If USMsgBox("Deseja realmente cancelar a expedição deste produto?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Conexao.Execute "Update Vendas_carteira Set qtdeexpedida = 0 where Codigo = " & .listaitens.SelectedItem & " and liberacao <> 'CANCELADO' and liberacao <> 'PERDIDO P/ PRAZO' and liberacao <> 'PERDIDO P/ PREÇO' and liberacao <> 'PORTAL ELETRONICO' and Qtdeexpedida <> 0"
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "PCP/Gerenciamento de ordem"
        Evento = "Cancelar expedição"
        ID_documento = .listaitens.SelectedItem
        Documento = "Ped. interno: " & .listaitens.SelectedItem.ListSubItems(6) & " - Rev.: " & .listaitens.SelectedItem.ListSubItems(7) & " - Cód. interno: " & .listaitens.SelectedItem.ListSubItems(3)
        Documento1 = ""
        ProcGravaEvento
        '==================================
        .ProcAtualizalista_carteira
    End If
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdExpedir_Click()
On Error GoTo tratar_erro

With frmprod
    If USMsgBox("Deseja realmente marcar este produto como expedido?", vbYesNo, "CAPRIND v5.0") = vbYes Then
Mensagem1:
            QuantSolicitado = 0
            Expedida = InputBox("Favor informar a quantidade expedida.")
            If IsNumeric(Expedida) = True Then
                QuantSolicitado = Expedida
            Else
                If Expedida = "" Then Exit Sub
                USMsgBox ("Só é permitido número neste campo."), vbInformation, "CAPRIND v5.0"
                GoTo Mensagem1
            End If
            If QuantSolicitado <> 0 Then
                Set TBExecucao = CreateObject("adodb.recordset")
                TBExecucao.Open "Select * from vendas_carteira where Codigo = " & .listaitens.SelectedItem & " and liberacao <> 'CANCELADO' and liberacao <> 'PERDIDO P/ PRAZO' and liberacao <> 'PERDIDO P/ PREÇO' and liberacao <> 'PORTAL ELETRONICO'", Conexao, adOpenKeyset, adLockOptimistic
                If TBExecucao.EOF = False Then
                    TBExecucao!qtdeexpedida = Format(QuantSolicitado, "###,##0.00")
                    TBExecucao.Update
                    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
                    '==================================
                    Modulo = "PCP/Gerenciamento de ordem"
                    Evento = "Mudar p/ expedido"
                    ID_documento = .listaitens.SelectedItem
                    Documento = "Ped. interno: " & .listaitens.SelectedItem.ListSubItems(6) & " - Rev.: " & .listaitens.SelectedItem.ListSubItems(7) & " - Cód. interno: " & .listaitens.SelectedItem.ListSubItems(3)
                    Documento1 = ""
                    ProcGravaEvento
                    '==================================
                    .ProcAtualizalista_carteira
                Else
                    USMsgBox ("Não é permitido marcar este produto como expedido devido o status."), vbInformation, "CAPRIND v5.0"
                End If
                TBExecucao.Close
            End If
    End If
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdMRP_Click()
On Error GoTo tratar_erro

With frmprod
    If .listaitens.SelectedItem.SubItems(3) = "NÃO" Then
        If USMsgBox("Deseja marcar o produto como MRP gerado?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            Set TBCarteira = CreateObject("adodb.recordset")
            TBCarteira.Open "Select * from vendas_carteira where Codigo = " & .listaitens.SelectedItem & " and OE = 'False' order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
            If TBCarteira.EOF = False Then
                TBCarteira!OE = True
                TBCarteira.Update
                USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
                '==================================
                Modulo = "PCP/Gerenciamento de ordem"
                Evento = "Marcar MRP gerado"
                ID_documento = .listaitens.SelectedItem
                Documento = "Ped. interno: " & .listaitens.SelectedItem.ListSubItems(18) & " - Rev.: " & .listaitens.SelectedItem.ListSubItems(19) & " - Cód. interno: " & .listaitens.SelectedItem.ListSubItems(4)
                Documento1 = ""
                ProcGravaEvento
                '==================================
                .ProcAtualizalista_carteira
            Else
                USMsgBox ("Este produto já está com o MRP gerado."), vbInformation, "CAPRIND v5.0"
            End If
            TBCarteira.Close
        End If
    Else
        If USMsgBox("Deseja marcar o produto como MRP não gerado?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            Set TBCarteira = CreateObject("adodb.recordset")
            TBCarteira.Open "Select * from vendas_carteira where Codigo = " & .listaitens.SelectedItem & " order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
            If TBCarteira.EOF = False Then
                If TBCarteira!OE = False Then
                    TBCarteira!OE = True
                Else
                    TBCarteira!OE = False
                End If
            End If
            
                TBCarteira.Update
                USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
                '==================================
                Modulo = "PCP/Gerenciamento de ordem"
                Evento = "Marcar MRP não gerado"
                ID_documento = .listaitens.SelectedItem
                Documento = "Ped. interno: " & .listaitens.SelectedItem.ListSubItems(18) & " - Rev.: " & .listaitens.SelectedItem.ListSubItems(19) & " - Cód. interno: " & .listaitens.SelectedItem.ListSubItems(4)
                Documento1 = ""
                ProcGravaEvento
                '==================================
                .ProcAtualizalista_carteira
            TBCarteira.Close
        End If
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
