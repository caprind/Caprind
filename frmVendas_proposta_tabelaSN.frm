VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_proposta_tabelaSN 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vendas - Proposta comercial - Tabela do simples nacional"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7365
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmVendas_proposta_tabelaSN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   855
      Left            =   30
      TabIndex        =   2
      Top             =   -30
      Width           =   6645
      Begin VB.ComboBox Cmb_tipo_TBSN 
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
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmVendas_proposta_tabelaSN.frx":1042
         Left            =   180
         List            =   "frmVendas_proposta_tabelaSN.frx":1058
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Tabela do simples nacional."
         Top             =   390
         Width           =   6285
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Alinhar à Direita
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparente
         Caption         =   "Tabela do simples nacional"
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
         Left            =   2370
         TabIndex        =   3
         Top             =   180
         Width           =   1890
      End
   End
   Begin DrawSuite2022.USButton Cmd_OK 
      Height          =   765
      Left            =   6720
      TabIndex        =   1
      Top             =   60
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   1349
      BorderColor     =   8421504
      BorderColorDisabled=   0
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColor2  =   14737632
      GradientColor3  =   12632256
      GradientColor4  =   12632256
      PicSizeH        =   48
      PicSizeW        =   48
      Theme           =   1
   End
End
Attribute VB_Name = "frmVendas_proposta_tabelaSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_OK_Click()
On Error GoTo tratar_erro

Acao = "sair"
If Cmb_tipo_TBSN = "" Then
    NomeCampo = "a tabela"
    ProcVerificaAcao
    Cmb_tipo_TBSN.SetFocus
    Exit Sub
End If
If Vendas_Proposta = True Then
    With frmVendas_proposta
        If Left(Cmb_tipo_TBSN, 1) = "T" Then
            Select Case Mid(Cmb_tipo_TBSN, 8, 3)
                Case "I -": .TabelaSN_Proposta = 1
                Case "II ": .TabelaSN_Proposta = 2
                Case "III": .TabelaSN_Proposta = 3
                Case "IV ": .TabelaSN_Proposta = 4
            End Select
        Else
            .TabelaSN_Proposta = 6
        End If
    End With
ElseIf Vendas_PI = True Then
        With frmVendas_PI
            If Left(Cmb_tipo_TBSN, 1) = "T" Then
                Select Case Mid(Cmb_tipo_TBSN, 8, 3)
                    Case "I -": .TabelaSN_PI = 1
                    Case "II ": .TabelaSN_PI = 2
                    Case "III": .TabelaSN_PI = 3
                    Case "IV ": .TabelaSN_PI = 4
                End Select
            Else
                .TabelaSN_PI = 6
            End If
        End With
    Else
        With frmVendas_programacao
            If Left(Cmb_tipo_TBSN, 1) = "T" Then
                Select Case Mid(Cmb_tipo_TBSN, 8, 3)
                    Case "I -": .TabelaSN_Prog = 1
                    Case "II ": .TabelaSN_Prog = 2
                    Case "III": .TabelaSN_Prog = 3
                    Case "IV ": .TabelaSN_Prog = 4
                End Select
            Else
                .TabelaSN_Prog = 6
            End If
        End With
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

If Vendas_PI = True Then
    Caption = Replace(Caption, "Proposta comercial", "Pedido interno")
    IDlista = frmVendas_PI.Cmb_empresa.ItemData(frmVendas_PI.Cmb_empresa.ListIndex)
ElseIf Vendas_Programacao = True Then
        Caption = Replace(Caption, "Proposta comercial", "Programação")
        IDlista = frmVendas_programacao.Cmb_empresa.ItemData(frmVendas_programacao.Cmb_empresa.ListIndex)
    Else
        IDlista = frmVendas_proposta.Cmb_empresa.ItemData(frmVendas_proposta.Cmb_empresa.ListIndex)
End If
With Cmb_tipo_TBSN
    .Clear
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Tabela FROM Impostos_TabelaDAS where ID_empresa = " & IDlista & " and Ativado = 1 group by Tabela", Conexao, adOpenKeyset, adLockOptimistic
    Do While TBAbrir.EOF = False
        Select Case TBAbrir!Tabela
            Case 1: .AddItem "Tabela I - Partilha do Simples Nacional – Comércio"
            Case 2: .AddItem "Tabela II - Partilha do Simples Nacional - Indústria"
            Case 3: .AddItem "Tabela III - Partilha do Simples Nacional - Serviços e Locação de Bens Móveis"
            Case 4: .AddItem "Tabela IV - Partilha do Simples Nacional - Serviços"
            Case 5: .AddItem "Tabela V - Partilha do Simples Nacional - Partilha do Simples Nacional - Receitas decorrentes da prestação de serviços relacionados no § 5º-I do art. 18 da LC 123/2016"
        End Select
        TBAbrir.MoveNext
    Loop
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
