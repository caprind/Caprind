VERSION 5.00
Begin VB.Form frmQualidadePPAP_FMEA_Localizar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Qualidade - PPAP - FMEA - Localizar"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8865
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   2415
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   30
      TabIndex        =   10
      Top             =   0
      Width           =   8805
      Begin VB.CommandButton cmdFiltrar 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   180
         MouseIcon       =   "frmQualidadePPAP_FMEA_Localizar.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmQualidadePPAP_FMEA_Localizar.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Filtrar (F2)"
         Top             =   180
         Width           =   630
      End
      Begin VB.CommandButton imgSair 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7980
         MouseIcon       =   "frmQualidadePPAP_FMEA_Localizar.frx":08D1
         MousePointer    =   99  'Custom
         Picture         =   "frmQualidadePPAP_FMEA_Localizar.frx":0A23
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Sair (Esc)"
         Top             =   180
         Width           =   630
      End
      Begin VB.CommandButton cmdAjuda 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7350
         MouseIcon       =   "frmQualidadePPAP_FMEA_Localizar.frx":11F6
         MousePointer    =   99  'Custom
         Picture         =   "frmQualidadePPAP_FMEA_Localizar.frx":1348
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Ajuda (F1)"
         Top             =   180
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   30
      TabIndex        =   0
      Top             =   870
      Width           =   8805
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   4620
         TabIndex        =   4
         Top             =   210
         Width           =   3975
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1470
            MouseIcon       =   "frmQualidadePPAP_FMEA_Localizar.frx":17EA
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   180
            MouseIcon       =   "frmQualidadePPAP_FMEA_Localizar.frx":193C
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2760
            MouseIcon       =   "frmQualidadePPAP_FMEA_Localizar.frx":1A8E
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmQualidadePPAP_FMEA_Localizar.frx":1BE0
         Left            =   180
         List            =   "frmQualidadePPAP_FMEA_Localizar.frx":1BF9
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4365
      End
      Begin VB.TextBox txtTexto 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         MaxLength       =   255
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   8415
      End
      Begin VB.ComboBox cmbfamilia 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   180
         MousePointer    =   99  'Custom
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Familia."
         Top             =   1050
         Width           =   8415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
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
         Left            =   3660
         TabIndex        =   9
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Filtrar por"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1942
         TabIndex        =   8
         Top             =   180
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmQualidadePPAP_FMEA_Localizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Família" Then
    cmbfamilia.Visible = True
    txtTexto.Visible = False
Else
    cmbfamilia.Visible = False
    txtTexto.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdFiltrar_Click()
On Error GoTo tratar_erro

With frmQualidadePPAP_FMEA
    .Lista.ListItems.Clear
    If txtTexto.Text <> "" Or cmbfamilia.Visible = True Then
        If cmbfiltrarpor = "Código interno" Then
            If Optinicio.Value = True Then .SQL_FMEA = "Select * from QualidadePPAP_FMEA where CodInterno like '" & txtTexto.Text & "%' order by ID desc"
            If Optmeio.Value = True Then .SQL_FMEA = "Select * from QualidadePPAP_FMEA where codinterno like '%" & txtTexto.Text & "%' order by id desc"
            If Optfim.Value = True Then .SQL_FMEA = "Select * from QualidadePPAP_FMEA where codinterno like '%" & txtTexto.Text & "' order by id desc"
        End If
        
        If cmbfiltrarpor = "Código de referência" Then
            Set TBFamilia = CreateObject("adodb.recordset")
            If Optinicio.Value = True Then TBFamilia.Open "Select codproduto from item_aplicacoes where n_referencia like '" & txtTexto.Text & "%' order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
            If Optmeio.Value = True Then TBFamilia.Open "Select codproduto from item_aplicacoes where n_referencia like '%" & txtTexto.Text & "%' order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
            If Optfim.Value = True Then TBFamilia.Open "Select codproduto from item_aplicacoes where n_referencia like '%" & txtTexto.Text & "' order by codproduto", Conexao, adOpenKeyset, adLockOptimistic
            If TBFamilia.EOF = False Then
                Codproduto = 0
                Do While TBFamilia.EOF = False
                    If Codproduto <> TBFamilia!Codproduto Then
                        .SQL_FMEA = "Select * from QualidadePPAP_FMEA where IDProduto = " & TBFamilia!Codproduto & " order by id desc"
                        .ProcCarregaLista
                    End If
                    Codproduto = TBFamilia!Codproduto
                    TBFamilia.MoveNext
                Loop
            End If
            TBFamilia.Close
            Unload Me
            Exit Sub
        End If
        
        If cmbfiltrarpor = "Descrição" Then
            If Optinicio.Value = True Then .SQL_FMEA = "Select * from QualidadePPAP_FMEA where descricao like '" & txtTexto.Text & "%' order by id desc"
            If Optmeio.Value = True Then .SQL_FMEA = "Select * from QualidadePPAP_FMEA where descricao like '%" & txtTexto.Text & "%' order by id desc"
            If Optfim.Value = True Then .SQL_FMEA = "Select * from QualidadePPAP_FMEA where descricao like '%" & txtTexto.Text & "' order by id desc"
        End If
        
        If cmbfiltrarpor = "Cliente" Then
            If Optinicio.Value = True Then .SQL_FMEA = "Select * FROM QualidadePPAP_FMEA where cliente like '" & txtTexto.Text & "%' order by id desc"
            If Optmeio.Value = True Then .SQL_FMEA = "Select * from QualidadePPAP_FMEA where cliente like '%" & txtTexto.Text & "%' order by id desc"
            If Optfim.Value = True Then .SQL_FMEA = "Select * from QualidadePPAP_FMEA where cliente like '%" & txtTexto.Text & "' order by id desc"
        End If
        
        If cmbfiltrarpor = "Fornecedor" Then
            If Optinicio.Value = True Then .SQL_FMEA = "Select * FROM QualidadePPAP_FMEA where fornecedor like '" & txtTexto.Text & "%' order by id desc"
            If Optmeio.Value = True Then .SQL_FMEA = "Select * from QualidadePPAP_FMEA where fornecedor like '%" & txtTexto.Text & "%' order by id desc"
            If Optfim.Value = True Then .SQL_FMEA = "Select * from QualidadePPAP_FMEA where fornecedor like '%" & txtTexto.Text & "' order by id desc"
        End If
        
        If cmbfiltrarpor = "Número FMEA" Then
            If Optinicio.Value = True Then .SQL_FMEA = "Select * FROM QualidadePPAP_FMEA where FMEA like '" & txtTexto.Text & "%' order by id desc"
            If Optmeio.Value = True Then .SQL_FMEA = "Select * from QualidadePPAP_FMEA where FMEA like '%" & txtTexto.Text & "%' order by id desc"
            If Optfim.Value = True Then .SQL_FMEA = "Select * from QualidadePPAP_FMEA where FMEA like '%" & txtTexto.Text & "' order by id desc"
        End If
        
        If cmbfiltrarpor = "Família" Then
            Set TBFamilia = CreateObject("adodb.recordset")
            TBFamilia.Open "Select * from projproduto where Classe = '" & cmbfamilia & "' order by classe", Conexao, adOpenKeyset, adLockOptimistic
            If TBFamilia.EOF = False Then
                Do While TBFamilia.EOF = False
                    .SQL_FMEA = "Select * from QualidadePPAP_FMEA where IDProduto = " & TBFamilia!Codproduto & " order by id desc"
                    .ProcCarregaLista
                    TBFamilia.MoveNext
                Loop
            End If
            TBFamilia.Close
            Unload Me
            Exit Sub
        End If
    Else
        .SQL_FMEA = "Select * from QualidadePPAP_FMEA order by id desc"
    End If
    .ProcCarregaLista
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
    Case vbKeyF2: cmdFiltrar_Click
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

cmbfiltrarpor = "Número FMEA"
cmbfamilia.Clear
Set TBFamilia = CreateObject("adodb.recordset")
TBFamilia.Open "Select * from projfamilia order by familia", Conexao, adOpenKeyset, adLockOptimistic
If TBFamilia.EOF = False Then
    Do While TBFamilia.EOF = False
        cmbfamilia.AddItem TBFamilia!Familia
        TBFamilia.MoveNext
    Loop
End If
TBFamilia.Close
Optinicio.Value = True
txtTexto.Visible = True
cmbfamilia.Visible = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub imgSair_Click()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If txtTexto <> "" Then cmbfamilia.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
