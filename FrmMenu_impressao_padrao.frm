VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form FrmMenu_impressao_padrao 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "CAPRIND v5.0 | Menu relatórios"
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4185
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   714
      DibPicture      =   "FrmMenu_impressao_padrao.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "FrmMenu_impressao_padrao.frx":108C4
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
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
      Height          =   2025
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   3495
      Begin DrawSuite2022.USButton cmdNormal 
         Height          =   750
         Left            =   660
         TabIndex        =   0
         Top             =   270
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   1323
         DibPicture      =   "FrmMenu_impressao_padrao.frx":10BDE
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
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
         ForeColor       =   16777215
         ForeColorDown   =   16777215
         ForeColorOver   =   16777215
         GradientColor1  =   5263559
         GradientColor2  =   5263559
         GradientColor3  =   5263559
         GradientColor4  =   5263559
         GradientColorDisabled1=   13160660
         GradientColorDisabled2=   13160660
         GradientColorDisabled3=   13160660
         GradientColorDisabled4=   13160660
         GradientColorDown1=   4013465
         GradientColorDown2=   4013465
         GradientColorDown3=   4013465
         GradientColorDown4=   4013465
         GradientColorOver1=   4408288
         GradientColorOver2=   4408288
         GradientColorOver3=   4408288
         GradientColorOver4=   4408288
         PicAlign        =   7
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
      Begin DrawSuite2022.USButton cmdGrafico 
         Height          =   780
         Left            =   660
         TabIndex        =   1
         Top             =   1080
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   1376
         DibPicture      =   "FrmMenu_impressao_padrao.frx":214A2
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
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
         ForeColor       =   16777215
         ForeColorDown   =   16777215
         ForeColorOver   =   16777215
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         PicAlign        =   7
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   3
      End
   End
End
Attribute VB_Name = "FrmMenu_impressao_padrao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGrafico_Click()
On Error GoTo tratar_erro

If Vendas_Relatorio_Historico = True Then
    With frmVendas_Relatorios_Historico
        If .Opt_individual.Value = True Then
            If .optDetalhado.Value = True Then
                NomeRel = "Vendas_historico_individual_detalhado grafico.rpt"
            Else
                USMsgBox ("Não existe relatório disponível para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
                Exit Sub
            End If
        Else
            NomeRel = "Vendas_historico_comparativo_resumido grafico.rpt"
        End If
    End With
ElseIf Vendas_Relatorio_IndiceAtraso = True Then
        With frmVendas_Relatorios_Indice_Atraso
            If .Opt_individual.Value = True Then
                If .optDetalhado.Value = True Then
                    NomeRel = "Indice_atraso_individual_detalhado grafico.rpt"
                Else
                    USMsgBox ("Não existe relatório disponível para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
                    Exit Sub
                End If
            Else
                NomeRel = "Indice_atraso_comparativo_resumido grafico.rpt"
            End If
        End With
    ElseIf Compras_Relatorio_IndiceAtraso = True Then
            With frmCompras_Relatorios_Indice_Atraso
                If .Opt_individual.Value = True Then
                    If .optDetalhado.Value = True Then
                        NomeRel = "Indice_atraso_individual_detalhado grafico.rpt"
                    Else
                        USMsgBox ("Não existe relatório disponível para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
                        Exit Sub
                    End If
                Else
                    NomeRel = "Indice_atraso_comparativo_resumido grafico.rpt"
                End If
            End With
        ElseIf Vendas_Relatorio_Comissao = True Then
'                With frmVendas_comissao
''                    If .Opt_individual.Value = True Then
'                        If .optDetalhado.Value = True Then
'                            NomeRel = "Vendas_comissao_individual_detalhado grafico.rpt"
'                        Else
'                            USMsgBox ("Não existe relatório disponível para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
'                            Exit Sub
'                        End If
'                    Else
'                        NomeRel = "Vendas_comissao_comparativo_resumido grafico.rpt"
'                    End If
'                End With
            ElseIf PCP_relatorios_indice_atraso = True Then
                    With frmRelatorios_indice_atraso
                        If .Opt_individual.Value = True Then
                            If .optDetalhado.Value = True Then
                                NomeRel = "PCP_Indice_atraso_individual_detalhado grafico.rpt"
                            Else
                                USMsgBox ("Não existe relatório disponível para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
                                Exit Sub
                            End If
                        Else
                            NomeRel = "PCP_Indice_atraso_comparativo_resumido grafico.rpt"
                        End If
                    End With
                ElseIf Manutencao_Relatorio_Historico = True Then
                        With frmManutencao_relatorios
                            If .Opt_individual.Value = True Then
                                If .optDetalhado.Value = True Then
                                    USMsgBox ("Não existe relatório disponível para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
                                    Exit Sub
                                End If
                            Else
                                NomeRel = "Manutencao_relatorio_comparativo_resumido grafico.rpt"
                            End If
                        End With
End If
ProcImprimirRelGrafico "{Producao_Relatorios.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios.Modulo} = '" & Formulario & "' and {Producao_Relatorios_Total.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdNormal_Click()
On Error GoTo tratar_erro

If Vendas_Relatorio_Historico = True Then
    If frmVendas_Relatorios_Historico.optDetalhado.Value = True Then NomeRel = "Vendas_historico_individual_detalhado.rpt" Else NomeRel = "Vendas_historico_resumido.rpt"
ElseIf Vendas_Relatorio_IndiceAtraso = True Then
        With frmVendas_Relatorios_Indice_Atraso
            If .Opt_individual.Value = True Then
                If .optDetalhado.Value = True Then NomeRel = "Indice_atraso_individual_detalhado.rpt" Else NomeRel = "Indice_atraso_individual_resumido.rpt"
            Else
                NomeRel = "Indice_atraso_comparativo_resumido.rpt"
            End If
        End With
    ElseIf Compras_Relatorio_IndiceAtraso = True Then
            With frmCompras_Relatorios_Indice_Atraso
                If .Opt_individual.Value = True Then
                    If .optDetalhado.Value = True Then NomeRel = "Compras_Indice_atraso_individual_detalhado.rpt" Else NomeRel = "Compras_Indice_atraso_individual_resumido.rpt"
                Else
                    NomeRel = "Compras_Indice_atraso_comparativo_resumido.rpt"
                End If
            End With
        ElseIf Vendas_Relatorio_Comissao = True Then
                With frmVendas_comissao
'                    If .Opt_individual.Value = True Then
                        If .optDetalhado.Value = True Then NomeRel = "Vendas_comissao_individual_detalhado.rpt" Else NomeRel = "Vendas_comissao_individual_resumido.rpt"
'                    Else
'                        NomeRel = "Vendas_comissao_comparativo_resumido.rpt"
'                    End If
                End With
            ElseIf PCP_relatorios_indice_atraso = True Then
                    With frmRelatorios_indice_atraso
                        If .Opt_individual.Value = True Then
                            If .optDetalhado.Value = True Then NomeRel = "PCP_Indice_atraso_individual_detalhado.rpt" Else NomeRel = "PCP_Indice_atraso_individual_resumido.rpt"
                        Else
                            NomeRel = "PCP_Indice_atraso_comparativo_resumido.rpt"
                        End If
                    End With
                ElseIf Manutencao_Relatorio_Historico = True Then
                        With frmManutencao_relatorios
                            If .Opt_individual.Value = True Then
                                If .optDetalhado.Value = True Then NomeRel = "Manutencao_relatorio_individual_detalhado.rpt" Else NomeRel = "Manutencao_relatorio_individual_resumido.rpt"
                            Else
                                NomeRel = "Manutencao_relatorio_comparativo_resumido.rpt"
                            End If
                        End With
            
End If
ProcImprimirRel "{Producao_Relatorios.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios.Modulo} = '" & Formulario & "' and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""

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

