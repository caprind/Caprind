VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmContas_pagas_menuimpressao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contas pagas"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2445
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
   ScaleHeight     =   1740
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Menu relatórios"
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
      Height          =   1695
      Left            =   55
      TabIndex        =   3
      Top             =   0
      Width           =   2355
      Begin DrawSuite2022.USButton Cmd_detalhado 
         Height          =   360
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   635
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Detalhado"
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
      Begin DrawSuite2022.USButton Cmd_resumido 
         Height          =   360
         Left            =   180
         TabIndex        =   1
         Top             =   750
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   635
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Resumido"
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
      Begin DrawSuite2022.USButton Cmd_recibo 
         Height          =   360
         Left            =   180
         TabIndex        =   2
         Top             =   1200
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   635
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Recibo"
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
End
Attribute VB_Name = "frmContas_pagas_menuimpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_detalhado_Click()
On Error GoTo tratar_erro

With frmContas_Pagas
    If .Filtro_Contas_Pagas_PC = False Then NomeRel = "Contas_pagas.rpt" Else NomeRel = "Contas_pagas_conta contabil.rpt"
    ProcVerificaContasSelRel .lst_ContasPagas, IIf(.Cmb_opcao_lista = "Relatório", True, False)
    If Familiatext <> "" Then
        ProcImprimirRel "(" & Familiatext & ") and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""
    Else
        ProcImprimirRel .FormulaRel_Contas_Pagas & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_recibo_Click()
On Error GoTo tratar_erro

NomeRelAntigo = NomeRel
NomeRel = "Contas_pagas_recibo.rpt"
With frmContas_Pagas
    ProcVerificaContasSelRel .lst_ContasPagas, IIf(.Cmb_opcao_lista = "Relatório", True, False)
    If Familiatext <> "" Then
        ProcImprimirRel Familiatext, ""
    Else
        ProcImprimirRel .FormulaRel_Contas_Pagas, ""
    End If
End With
NomeRel = NomeRelAntigo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_resumido_Click()
On Error GoTo tratar_erro

NomeRelAntigo = NomeRel
NomeRel = "Contas_pagas_resumido.rpt"
With frmContas_Pagas
    ProcVerificaContasSelRel .lst_ContasPagas, IIf(.Cmb_opcao_lista = "Relatório", True, False)
    If Familiatext <> "" Then
        ProcImprimirRel "(" & Familiatext & ") and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""
    Else
        ProcImprimirRel .FormulaRel_Contas_Pagas & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""
    End If
End With
NomeRel = NomeRelAntigo

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

Private Sub Form_Load()
On Error GoTo tratar_erro

If Imprimir = False Then
    Cmd_detalhado.Enabled = False
    Cmd_resumido.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
