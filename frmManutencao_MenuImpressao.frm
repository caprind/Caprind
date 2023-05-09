VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmManutencao_MenuImpressao 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'None
   Caption         =   "Relatório manutenção"
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManutencao_MenuImpressao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   714
      CaptionDelimiter=   ""
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
   End
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
      Height          =   2385
      Left            =   120
      TabIndex        =   2
      Top             =   510
      Width           =   4185
      Begin DrawSuite2022.USButton Cmd_padrao 
         Height          =   750
         Left            =   210
         TabIndex        =   0
         Top             =   1260
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1323
         Caption         =   "Ficha de manutenção"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
      Begin DrawSuite2022.USButton Cmd_cronograma 
         Height          =   750
         Left            =   210
         TabIndex        =   1
         Top             =   270
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1323
         Caption         =   "Cronograma de manutenção"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
Attribute VB_Name = "frmManutencao_MenuImpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_padrao_Click()
On Error GoTo tratar_erro

With frmManutencao
    If .txtId = 0 Then
        USMsgBox ("Informe a manutenção antes de visualizar impressão."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    NomeRel = "Manutencao.rpt"
    If .txttipo <> "Solicitação" Then
    ProcImprimirRel "{Manutencao.CodMan} = '" & .txtCodigo & "'", ""
    Else
    ProcImprimirRel "{Manutencao.CodSol} = '" & .txtCodigo & "'", ""
    End If
    
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_cronograma_Click()
On Error GoTo tratar_erro

With frmManutencao
.FormulaRel_Manutencao = ""
.FormulaRelSubReport_Manutencao = ""

    If .Lista.ListItems.Count = 0 Then Exit Sub
    NomeRel = "Manutencao_cronograma.rpt"
    ProcImprimirRel .FormulaRel_Manutencao, .FormulaRelSubReport_Manutencao
    'Debug.print .FormulaRel_Manutencao
    'Debug.print .FormulaRelSubReport_Manutencao
End With



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
