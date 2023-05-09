VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmContas_pagar_menuimpressao 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menu relatórios"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   1965
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
   ScaleHeight     =   1260
   ScaleWidth      =   1965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
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
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   55
      TabIndex        =   2
      Top             =   0
      Width           =   1875
      Begin DrawSuite2022.USButton Cmd_detalhado 
         Height          =   360
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
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
         Top             =   690
         Width           =   1515
         _ExtentX        =   2672
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
   End
End
Attribute VB_Name = "frmContas_pagar_menuimpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_detalhado_Click()
On Error GoTo tratar_erro

With frmContas_Pagar
    ProcVerificaContasSelRel .lst_contas, IIf(.Cmb_opcao_lista = "Relatório", True, False)
    If Familiatext <> "" Then
        ProcImprimirRel "(" & Familiatext & ") and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""
    Else
        ProcImprimirRel .FormulaRel_Contas_Pagar & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_resumido_Click()
On Error GoTo tratar_erro

NomeRelAntigo = NomeRel
NomeRel = "Contas_pagar_resumido.rpt"
With frmContas_Pagar
    ProcVerificaContasSelRel .lst_contas, IIf(.Cmb_opcao_lista = "Relatório", True, False)
    If Familiatext <> "" Then
        ProcImprimirRel "(" & Familiatext & ") and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""
    Else
        ProcImprimirRel .FormulaRel_Contas_Pagar & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""
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
