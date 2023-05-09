VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmFaturamento_Prod_serv_copiar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Faturamento | Gerar cópias NFe"
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFaturamento_Prod_serv_copiar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   2865
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   767
      DibPicture      =   "frmFaturamento_Prod_serv_copiar.frx":000C
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
      ForeColor       =   16777215
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
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
      Height          =   2025
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   3780
      Begin VB.TextBox Txt_n_copias 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   2130
         TabIndex        =   1
         ToolTipText     =   "Número de cópias."
         Top             =   390
         Width           =   795
      End
      Begin VB.CheckBox chk_fixar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fixar dia da emissão"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1050
         TabIndex        =   2
         Top             =   870
         Width           =   1755
      End
      Begin MSComCtl2.DTPicker Txt_inicio_emissao 
         Height          =   315
         Left            =   810
         TabIndex        =   0
         ToolTipText     =   "Data de início do pagamento."
         Top             =   390
         Width           =   1185
         _ExtentX        =   2090
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
         Format          =   488898561
         CurrentDate     =   39057
      End
      Begin DrawSuite2022.USButton cmdCopiar 
         Height          =   615
         Left            =   660
         TabIndex        =   7
         Top             =   1200
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   1085
         DibPicture      =   "frmFaturamento_Prod_serv_copiar.frx":1B6F
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Gerar cópia(s) NFe"
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
         PicSize         =   5
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Início emissão"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   900
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "N° cópias"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2205
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Top             =   180
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmFaturamento_Prod_serv_copiar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCopiar_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente copiar a(s) nota(s) selecionada(s)?", vbYesNo, "CAPRIND v4.0") = vbYes Then
ProcCopiar
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF3: ProcCopiar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

If Formulario = "Estoque/Ordem de faturamento" Then Caption = "Ordem de fat. - Copiar" Else Caption = "Nota fiscal - Copiar"
Txt_inicio_emissao.Value = Date
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCopiar()
On Error GoTo tratar_erro

valor = IIf(Txt_n_copias = "", 0, Txt_n_copias)
If valor <= 0 Then
    USMsgBox ("Informe o número de cópias antes de copiar."), vbExclamation, "CAPRIND v5.0"
    Txt_n_copias.SetFocus
    Exit Sub
End If
With frmFaturamento_Prod_Serv.ListaNota
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            PAGTO = Format(Txt_inicio_emissao.Value, "dd/mm/yyyy")
            
            For InitFor1 = 1 To valor
                If chk_fixar.Value = 1 Then
                    DT = PAGTO
                    DiaX = Day(DT)
                    MesX = Month(DT)
                    AnoX = Year(DT)
                Else
                    DT = PAGTO
                    DiaX = Day(DT)
                    MesX = Month(DT)
                    AnoX = Year(DT)
                    If Weekday(DT) = vbSunday Then
                        Dataini = DT
                        Dataini = Dataini + 1
                        DT = Dataini
                    End If
                    If Weekday(DT) = vbSaturday Then
                        Dataini = DT
                        Dataini = Dataini + 2
                        DT = Dataini
                    End If
                End If
                
                Dataini = DT
                frmFaturamento_Prod_Serv.ProcCopiarNF .ListItems(InitFor), Dataini
                
                MesX = MesX + 1
                If MesX > 12 Then
                    AnoX = AnoX + 1
                    MesX = 1
                End If
                If DiaX = 29 And MesX = 2 Then DiaX = 28
                If DiaX = 30 And MesX = 2 Then DiaX = 28
                If DiaX = 31 And MesX = 2 Then DiaX = 28
                If DiaX = 31 And MesX = 4 Then DiaX = 30
                If DiaX = 31 And MesX = 6 Then DiaX = 30
                If DiaX = 31 And MesX = 9 Then DiaX = 30
                If DiaX = 31 And MesX = 11 Then DiaX = 30
                PAGTO = Format(DiaX, "00") & "/" & Format(MesX, "00") & "/" & Format(AnoX, "0000")
            Next InitFor1
        End If
    Next InitFor
End With

If Formulario = "Estoque/Ordem de faturamento" Then MsgTexto = "Ordem(ns) de faturamento" Else MsgTexto = "Nota(s) fiscal(ais)"
USMsgBox (MsgTexto & " copiada(s) com sucesso."), vbInformation, "CAPRIND v5.0"
With frmFaturamento_Prod_Serv
    .ProcCarregaDadosNota IIf(.txtID = "", 0, .txtID)
    .ProcCarregaListaNota (1)
    .Novo_Nota = False
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosCCRealizado(ID_CC As Long)
On Error GoTo tratar_erro

TBFamilia!ID_financeiro = TBProduto!IDintconta
TBFamilia!Data = DT
TBFamilia!Responsavel = pubUsuario
TBFamilia!ID_empresa = TBCiclo!ID_empresa
TBFamilia!Operacao = "Débito"
TBFamilia!ID_CC = ID_CC
TBFamilia!valor = TBCiclo!valor
TBFamilia!Percentual = TBCiclo!Percentual
TBFamilia!Bloqueado = False

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

Private Sub Txt_inicio_emissao_Change()
On Error GoTo tratar_erro

If Left(Txt_inicio_emissao.Value, 2) = "31" Or Left(Txt_inicio_emissao.Value, 5) = "30/01" Or Left(Txt_inicio_emissao.Value, 5) = "31/01" Then
    chk_fixar.Enabled = False
Else
    chk_fixar.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_n_copias_Change()
On Error GoTo tratar_erro

If Txt_n_copias.Text <> "" Then
    VerifNumero = Txt_n_copias.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_n_copias.Text = ""
        Txt_n_copias.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
