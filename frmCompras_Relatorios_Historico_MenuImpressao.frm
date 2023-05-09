VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCompras_Relatorios_Historico_MenuImpressao 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Menu relatórios"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   3555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
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
   ScaleHeight     =   3975
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   820
      DibPicture      =   "frmCompras_Relatorios_Historico_MenuImpressao.frx":0000
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmCompras_Relatorios_Historico_MenuImpressao.frx":1C95
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   3570
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   714
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
      Height          =   2565
      Left            =   300
      TabIndex        =   2
      Top             =   690
      Width           =   2955
      Begin DrawSuite2022.USButton cmdNormal 
         Height          =   930
         Left            =   270
         TabIndex        =   0
         Top             =   300
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   1640
         DibPicture      =   "frmCompras_Relatorios_Historico_MenuImpressao.frx":1FAF
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
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         GradientColor1  =   1154291
         GradientColor2  =   1154291
         GradientColor3  =   1154291
         GradientColor4  =   1154291
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   8438015
         GradientColorOver2=   8438015
         GradientColorOver3=   8438015
         GradientColorOver4=   8438015
         GradientColorDown1=   16576
         GradientColorDown2=   16576
         GradientColorDown3=   16576
         GradientColorDown4=   16576
         PicAlign        =   8
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   5
      End
      Begin DrawSuite2022.USButton cmdGrafico 
         Height          =   930
         Left            =   270
         TabIndex        =   1
         Top             =   1410
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   1640
         DibPicture      =   "frmCompras_Relatorios_Historico_MenuImpressao.frx":40D5
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
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         PicAlign        =   8
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   3
      End
   End
End
Attribute VB_Name = "frmCompras_Relatorios_Historico_MenuImpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdGrafico_Click()
On Error GoTo tratar_erro

If frmCompras_Relatorios_Historico.Opt_individual.Value = True Then
    If frmCompras_Relatorios_Historico.optDetalhado.Value = True Then
        NomeRel = "Compras_historico_individual_detalhado grafico.rpt"
    Else
        USMsgBox ("Não existe relatório disponível para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
Else
    NomeRel = "Compras_historico_comparativo_resumido grafico.rpt"
End If
ProcImprimirRelGrafico "{Producao_Relatorios.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios.Modulo} = '" & Formulario & "' and {Producao_Relatorios_Total.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdNormal_Click()
On Error GoTo tratar_erro

If frmCompras_Relatorios_Historico2.optDetalhado.Value = True Then
    NomeRel = "Compras_historico_individual_detalhado.rpt"
    'Debug.print FormulaRelatorio
    ProcImprimirRel FormulaRelatorio, ""
Else
    NomeRel = "Compras_historico_resumido.rpt"
    ProcImprimirRel "{Producao_Relatorios.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios.Modulo} = '" & Formulario & "' and {Producao_Relatorios_Total.Responsavel}= '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""
End If


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
