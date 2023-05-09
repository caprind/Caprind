VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmQualidade_Relatorios_NC_Menu_Impressao 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Não conformidade"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3045
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
   ScaleHeight     =   1290
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Frame Frame2 
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
      Height          =   1245
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   2955
      Begin DrawSuite2022.USButton Cmd_padrao 
         Height          =   360
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   635
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Padrão"
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
      Begin DrawSuite2022.USButton Cmd_grafico 
         Height          =   360
         Left            =   180
         TabIndex        =   1
         Top             =   750
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   635
         BorderColor     =   8421504
         BorderColorDisabled=   0
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Gráfico"
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
Attribute VB_Name = "frmQualidade_Relatorios_NC_Menu_Impressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_grafico_Click()
On Error GoTo tratar_erro

Unload Me
If Sit_REG = 1 Then
    With frmQualidade_Relatorios_NC
        If .Opt_individual.Value = True Then
            If .optDetalhado.Value = True Then
                NomeRel = "CQ_nc_individual_detalhado grafico.rpt"
            Else
                USMsgBox ("Não existe relatório disponível para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
                Exit Sub
            End If
        Else
            NomeRel = "CQ_nc_comparativo_resumido grafico.rpt"
        End If
    End With
Else
    With frmPCP_Relatorios_NC
        If .Opt_individual.Value = True Then
            If .optDetalhado.Value = True Then
                NomeRel = "PCP_nc_individual_detalhado grafico.rpt"
            Else
                USMsgBox ("Não existe relatório disponível para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
                Exit Sub
            End If
        Else
            NomeRel = "PCP_nc_comparativo_resumido grafico.rpt"
        End If
    End With
End If
ProcImprimirRelGrafico "{Producao_Relatorios.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios.Modulo} = '" & Formulario & "' and {Producao_Relatorios_Total.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_padrao_Click()
On Error GoTo tratar_erro

Unload Me
If Sit_REG = 1 Then
    With frmQualidade_Relatorios_NC
        If .Opt_individual.Value = True Then
            If .optDetalhado.Value = True Then NomeRel = "CQ_nc_individual_detalhado.rpt" Else NomeRel = "CQ_nc_resumido.rpt"
            'ProcImprimirRel
        Else
            NomeRel = "CQ_nc_resumido.rpt"
            'TabelaRel = 1
            'OrdenarRel = 0
            'If .cmbfiltrarpor = "Ordem" Then CampoRel = 2 Else CampoRel = 14
            'ProcImprimirRelOrdenado
        End If
    End With
Else
    With frmPCP_Relatorios_NC
        If .Opt_individual.Value = True Then
            If .optDetalhado.Value = True Then NomeRel = "PCP_nc_individual_detalhado.rpt" Else NomeRel = "PCP_nc_resumido.rpt"
        Else
            NomeRel = "PCP_nc_resumido.rpt"
        End If
    End With
End If
ProcImprimirRel "{Producao_Relatorios.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios.Modulo} = '" & Formulario & "' and {Producao_Relatorios_Total.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""

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
