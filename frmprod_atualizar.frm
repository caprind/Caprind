VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmprod_atualizar 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Gerenciamento de ordem | Atualizar"
   ClientHeight    =   8220
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5160
   ControlBox      =   0   'False
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
   ScaleHeight     =   8220
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin DrawSuite2022.USButton btnAtualizar 
      Height          =   765
      Left            =   360
      TabIndex        =   27
      Top             =   6930
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   1349
      DibPicture      =   "frmprod_atualizar.frx":0000
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Atualizar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Theme           =   4
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   26
      Top             =   7815
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   741
      DibPicture      =   "frmprod_atualizar.frx":62E4
      CaptionDelimiter=   "|"
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
      Icon            =   "frmprod_atualizar.frx":10491
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin VB.CheckBox Chk_filtrar_backup 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar do backup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   390
      TabIndex        =   0
      Top             =   570
      Width           =   1545
   End
   Begin VB.Frame Frame2 
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
      Height          =   4995
      Left            =   360
      TabIndex        =   20
      Top             =   1560
      Width           =   4425
      Begin VB.CheckBox Chk18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "18 - Quantidade não conforme na qualidade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   4710
         Width           =   4095
      End
      Begin VB.CheckBox Chk17 
         BackColor       =   &H00E0E0E0&
         Caption         =   "17 - Valor das ordens no estoque"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   4410
         Width           =   4095
      End
      Begin VB.CheckBox Chk16 
         BackColor       =   &H00E0E0E0&
         Caption         =   "16 - Valor de terceiros nas ordens"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   4140
         Width           =   4095
      End
      Begin VB.CheckBox Chk15 
         BackColor       =   &H00E0E0E0&
         Caption         =   "15 - Valor dos materiais nas ordens"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   3876
         Width           =   4095
      End
      Begin VB.CheckBox Chk14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "14 - Número das ordens no estoque"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   3612
         Width           =   4095
      End
      Begin VB.CheckBox Chk13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "13 - Código do cliente nas ordens"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   3348
         Width           =   4095
      End
      Begin VB.CheckBox Chk12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "12 - Quantidade não conforme nas ordens"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   3084
         Width           =   4095
      End
      Begin VB.CheckBox Chk11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "11 - Dados das ordens"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   2820
         Width           =   4095
      End
      Begin VB.CheckBox Chk8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "08 - Resultados por operador, máq. e turno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   2028
         Width           =   4095
      End
      Begin VB.CheckBox Chk9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "09 - Quantidade expedida nas ordens e ped."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   2292
         Width           =   4095
      End
      Begin VB.CheckBox Chk10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "10 - Dimensão total do material nas ordens"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   2556
         Width           =   4095
      End
      Begin VB.CheckBox Chk3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "03 - OS's"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   708
         Width           =   4095
      End
      Begin VB.CheckBox Chk6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "06 - Evento dos apontamentos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   1500
         Width           =   4095
      End
      Begin VB.CheckBox Chk5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "05 - Turno dos apontamentos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   1236
         Width           =   4095
      End
      Begin VB.CheckBox Chk7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "07 - Quantidade de dias dos apontamentos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   1764
         Width           =   4095
      End
      Begin VB.CheckBox Chk1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "01 - Número das OS's nos apontamentos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Width           =   4095
      End
      Begin VB.CheckBox Chk2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "02 - Status dos materiais"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   444
         Width           =   4095
      End
      Begin VB.CheckBox Chk4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "04 - Hora dos apontamentos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   972
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   360
      TabIndex        =   21
      Top             =   900
      Width           =   4425
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   3030
         TabIndex        =   19
         ToolTipText     =   "Data final."
         Top             =   270
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   130809857
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   1260
         TabIndex        =   18
         ToolTipText     =   "Data inicio."
         Top             =   270
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   130809857
         CurrentDate     =   39057
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparente
         Caption         =   "Até :"
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
         Left            =   2595
         TabIndex        =   23
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparente
         Caption         =   "Dt. emissão:"
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
         Left            =   270
         TabIndex        =   22
         Top             =   270
         Width           =   900
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   6660
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor2      =   0
      SearchText      =   "Atualizando..."
      Value           =   0
   End
End
Attribute VB_Name = "frmprod_atualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcAtualizar()
On Error GoTo tratar_erro

If Chk1.Value = 0 And Chk2.Value = 0 And Chk3.Value = 0 And Chk4.Value = 0 And Chk5.Value = 0 And Chk6.Value = 0 And chk7.Value = 0 And chk8.Value = 0 And chk9.Value = 0 And chk10.Value = 0 And chk11.Value = 0 And chk12.Value = 0 And chk13.Value = 0 And Chk14.Value = 0 And Chk15.Value = 0 And Chk16.Value = 0 And Chk17.Value = 0 And Chk18.Value = 0 Then
    USMsgBox ("Informe uma das opções antes de atualizar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
If Chk_filtrar_backup.Value = 1 Then
    NomeTabelaAp = "ProducaoFases_Backup"
    NomeTabelaApTotalizacao = "ProducaoFases_Totalizacao_Backup"
Else
    NomeTabelaAp = "ProducaoFases"
    NomeTabelaApTotalizacao = "ProducaoFases_Totalizacao"
End If

frmprod.ProcAtualizacao

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnAtualizar_Click()
On Error GoTo tratar_erro

'If USMsgBox("Deseja realmente realizar as atualizações selecionadas?", vbYesNo, "CAPRIND v5.0") = vbYes Then
 ProcAtualizar
'End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcAtualizar
    Case vbKeyEscape: ProcSair
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

msk_fltInicio.Value = Date
msk_fltFim.Value = Date

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


