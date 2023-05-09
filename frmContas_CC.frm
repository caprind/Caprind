VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContas_CC 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Centro de custo"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8655
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTotalCentro 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   5580
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "0,00"
      ToolTipText     =   "Valor total centro de custo."
      Top             =   6420
      Width           =   1515
   End
   Begin VB.TextBox txtSaldoCentro 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   7110
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "0,00"
      ToolTipText     =   "Saldo."
      Top             =   6420
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   1425
      Left            =   55
      TabIndex        =   12
      Top             =   990
      Width           =   8565
      Begin VB.TextBox txtPercentual 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Percentual."
         Top             =   960
         Width           =   1155
      End
      Begin VB.CheckBox chkValor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6090
         TabIndex        =   3
         Top             =   750
         Width           =   675
      End
      Begin VB.CheckBox chkPercentual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Percentual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7245
         TabIndex        =   5
         Top             =   750
         Width           =   1065
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   5670
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Valor."
         Top             =   960
         Width           =   1515
      End
      Begin VB.ComboBox Cmb_CC 
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
         Height          =   330
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Centro de custo."
         Top             =   960
         Width           =   5475
      End
      Begin VB.TextBox txtData 
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
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Data do cadastro."
         Top             =   375
         Width           =   1185
      End
      Begin VB.TextBox txtResponsavel 
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
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Responsável pelo cadastro."
         Top             =   375
         Width           =   6975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Centro de custo*"
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
         Left            =   2295
         TabIndex        =   15
         Top             =   750
         Width           =   1245
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
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
         Index           =   4
         Left            =   600
         TabIndex        =   14
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável"
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
         Index           =   9
         Left            =   4410
         TabIndex        =   13
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.TextBox txtID 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2400
      TabIndex        =   11
      Text            =   "0"
      Top             =   3240
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txtValor_total 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   4050
      TabIndex        =   8
      Text            =   "0,00"
      ToolTipText     =   "Valor total."
      Top             =   6420
      Width           =   1515
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3705
      Left            =   60
      TabIndex        =   7
      Top             =   2430
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   6535
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Código"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   7364
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Valor"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Percentual"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "ID_CC"
         Object.Width           =   0
      EndProperty
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   3780
      Top             =   120
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmContas_CC.frx":0000
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   16
      Top             =   0
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   1720
      ButtonCount     =   7
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Novo"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Novo (Insert)"
      ButtonKey1      =   "1"
      ButtonAlignment1=   2
      BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   33
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonCaption2  =   "Salvar"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Salvar (F3)"
      ButtonKey2      =   "2"
      ButtonAlignment2=   2
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   37
      ButtonTop2      =   2
      ButtonWidth2    =   38
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonCaption3  =   "Excluir"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Excluir (F4)"
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   77
      ButtonTop3      =   2
      ButtonWidth3    =   39
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonAlignment4=   2
      ButtonType4     =   1
      ButtonStyle4    =   -1
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState4    =   -1
      ButtonLeft4     =   118
      ButtonTop4      =   4
      ButtonWidth4    =   2
      ButtonHeight4   =   54
      ButtonCaption5  =   "Ajuda"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Ajuda (F1)"
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   122
      ButtonTop5      =   2
      ButtonWidth5    =   36
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonCaption6  =   "Sair"
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonToolTipText6=   "Sair (Esc)"
      ButtonKey6      =   "6"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft6     =   160
      ButtonTop6      =   2
      ButtonWidth6    =   26
      ButtonHeight6   =   21
      ButtonUseMaskColor6=   0   'False
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonKey7      =   "7"
      ButtonAlignment7=   2
      BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState7    =   5
      ButtonLeft7     =   188
      ButtonTop7      =   2
      ButtonWidth7    =   24
      ButtonHeight7   =   24
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   17
      Top             =   6450
      Width           =   3885
      _ExtentX        =   6853
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo"
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
      Left            =   7635
      TabIndex        =   20
      Top             =   6210
      Width           =   465
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total centro"
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
      Left            =   5820
      TabIndex        =   19
      Top             =   6210
      Width           =   1035
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vlr. total"
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
      Left            =   4447
      TabIndex        =   18
      Top             =   6210
      Width           =   720
   End
End
Attribute VB_Name = "frmContas_CC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Contas_CC As Boolean 'OK

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Faturamento = True Then
    If frmFaturamento_Prod_Serv.lst_Duplicata.SelectedItem.ListSubItems(4) = "SIM" Then
        USMsgBox ("Não é permitido excluir centro de custo, pois a duplicata já foi enviada para o financeiro."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) centro(s) de custo?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub
            End If
            Permitido = True
            
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from CC_realizado where id = " & .ListItems(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                Conexao.Execute "DELETE from CC_realizado where ID_origem = " & TBFI!ID & " and Operacao = 'Débito' and ID_financeiro is not null"
                
                TBFI.Delete
                
                '==================================
                Modulo = Formulario
                Evento = "Excluir"
                ID_documento = .ListItems(InitFor)
                Documento1 = "Código: " & .ListItems(InitFor).ListSubItems(1) & " - Descrição: " & .ListItems(InitFor).ListSubItems(2)
                ProcGravaEvento
                '==================================
            End If
            TBFI.Close
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) centro(s) de custo antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Centro(s) de custo excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista
    Novo_Contas_CC = False
    Frame1.Enabled = False
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro

If Faturamento = True Then
    If frmFaturamento_Prod_Serv.lst_Duplicata.SelectedItem.ListSubItems(4) = "SIM" Then
        USMsgBox ("Não é criar novo centro de custo, pois a duplicata já foi enviada para o financeiro."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
End If
ProcLimpaCampos
Novo_Contas_CC = True
Frame1.Enabled = True
Cmb_CC.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Contas_CC = True Then
    If USMsgBox("O centro de custo ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Contas_CC = True Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
End If
Novo_Contas_CC = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
txtData = Format(Date, "dd/mm/yy")
txtResponsavel = pubUsuario
Cmb_CC.ListIndex = -1
chkValor.Value = 0
txtValor = ""
chkPercentual.Value = 0
txtPercentual = ""
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Frame1.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
If Faturamento = True Then
    If frmFaturamento_Prod_Serv.lst_Duplicata.SelectedItem.ListSubItems(4) = "SIM" Then
        USMsgBox ("Não é permitido salvar o centro de custo, pois a duplicata já foi enviada para o financeiro."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
End If

Acao = "salvar"
If Cmb_CC = "" Then
    NomeCampo = "o centro de custo"
    ProcVerificaAcao
    Cmb_CC.SetFocus
    Exit Sub
End If
If chkPercentual.Value = 1 Then
    Valor1 = IIf(txtPercentual = "", 0, txtPercentual)
    If Valor1 = 0 Then
        NomeCampo = "o percentual"
        ProcVerificaAcao
        txtPercentual.SetFocus
        Exit Sub
    End If
End If
Valor1 = IIf(txtValor = "", 0, txtValor)
If Valor1 = 0 Then
    NomeCampo = "o valor"
    ProcVerificaAcao
    chkValor.Value = 1
    txtValor.SetFocus
    Exit Sub
End If

'Verifica se o valor do centro de custo já ultrapassou o valor do item
Qtde = 0
Qtd = IIf(txtValor = "", 0, txtValor)
qt = txtvalor_total
If Faturamento = True Then TextoFiltro = "CC_realizado.ID_duplicata = " Else TextoFiltro = "CC_realizado.ID_financeiro = "
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Sum(CC_realizado.Valor) as qtde from CC_realizado INNER JOIN Usuarios_setor ON CC_realizado.ID_CC = Usuarios_setor.ID where " & TextoFiltro & IDlista & " and CC_realizado.ID <> " & txtId & " and Usuarios_setor.Consolidacao = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Qtde = IIf(IsNull(TBAbrir!Qtde), 0, TBAbrir!Qtde)
End If
TBAbrir.Close
If Format((Qtde + Qtd), "###,##0.00") > qt Then
    USMsgBox ("Não é permitido salvar, pois o valor do centro de custo ultrapassou o valor total da conta."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

'Verifica se o centro de custo possui previsão orçamentária, se não tiver ele bloqueia
Permitido = False
Permitido1 = False
ID_CC = 0
Formulario = "Financeiro/Autorização de centro de custo sem previsão"
If Financeiro_Contas_Pagar = True Then
    With frmContas_Pagar
        ID_empresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
        TextoFiltro = "Select ID_PC from Familia_financeiro where IDConta = " & .txtidintconta & " and TipoConta = 'P'"
    End With
ElseIf Faturamento = True Then
        With frmFaturamento_Prod_Serv
            TextoFiltro = "Select ID_PC from projproduto P INNER JOIN tbl_Detalhes_Nota N on P.CodProduto = N.codproduto where N.ID_Nota = " & .txtId & " and P.ID_PC IS NOT NULL"
            IDempresa = .txtIDEmpresa
        End With
        Formulario = "Faturamento/Autorização de centro de custo sem previsão"
    Else
        With frmContas_Pagas
            ID_empresa = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
            TextoFiltro = "Select ID_PC from Familia_financeiro where IDConta = " & .txtidintconta & " and TipoConta = 'P'"
        End With
End If

Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select Codigo from Empresa where Codigo = " & ID_empresa & " and Bloc_CC_Previsao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    If Novo_Contas_CC = False Then ID_CC = Lista.SelectedItem.ListSubItems(5)

    If ID_CC <> Cmb_CC.ItemData(Cmb_CC.ListIndex) Then
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then
            If USMsgBox("A " & IIf(Faturamento = True, "nota fiscal", "conta") & " não possui conta contábil cadastrada, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbYes Then frmAprovar_CC.Show 1
        Else
            Do While TBProduto.EOF = False And Permitido1 = False
                Permitido = False
                Set TBCQ = CreateObject("adodb.recordset")
                TBCQ.Open "Select US.Id from Usuarios_setor US INNER JOIN Usuarios_setor_previsao USP on US.Id = USP.ID_CC where US.ID = " & Cmb_CC.ItemData(Cmb_CC.ListIndex) & " and USP.ID_PC = " & TBProduto!ID_PC, Conexao, adOpenKeyset, adLockOptimistic
                If TBCQ.EOF = True Then
                    If USMsgBox("O centro de custo não possui previsão orçamentária, deseja prosseguir assim mesmo?", vbYesNo, "CAPRIND v5.0") = vbYes Then frmAprovar_CC.Show 1
                Else
                    Permitido = True
                End If
                TBCQ.Close
                If Permitido = False Then
                    TBProduto.Close
                    Exit Sub
                End If
                TBProduto.MoveNext
            Loop
        End If
        TBProduto.Close
        If Permitido = False Then Exit Sub
    End If
End If
TBTempo.Close

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select * from CC_realizado where id = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = True Then TBFI.AddNew
ProcEnviaDados Cmb_CC.ItemData(Cmb_CC.ListIndex), 0
TBFI.Update
txtId = TBFI!ID

'Grava movimentação no centro consolidado
Set TBAfericao = CreateObject("adodb.recordset")
TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & Cmb_CC.ItemData(Cmb_CC.ListIndex), Conexao, adOpenKeyset, adLockOptimistic
If TBAfericao.EOF = False Then
    Do While TBAfericao.EOF = False
        Set TBFI = CreateObject("adodb.recordset")
        TBFI.Open "Select * from CC_realizado where ID_CC = " & TBAfericao!ID_CC & " and ID_origem = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBFI.EOF = True Then TBFI.AddNew
        ProcEnviaDados TBAfericao!ID_CC, txtId
        TBFI.Update
        TBFI.Close
        
        Set TBCiclo = CreateObject("adodb.recordset")
        TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
        If TBCiclo.EOF = False Then
            Do While TBCiclo.EOF = False
                Set TBFI = CreateObject("adodb.recordset")
                TBFI.Open "Select * from CC_realizado where ID_CC = " & TBCiclo!ID_CC & " and ID_origem = " & txtId, Conexao, adOpenKeyset, adLockOptimistic
                If TBFI.EOF = True Then TBFI.AddNew
                ProcEnviaDados TBCiclo!ID_CC, txtId
                TBFI.Update
                TBFI.Close
                
                TBCiclo.MoveNext
            Loop
        End If
        TBCiclo.Close
        
        TBAfericao.MoveNext
    Loop
End If
TBAfericao.Close

ProcCarregaLista
If Novo_Contas_CC = True Then
    USMsgBox ("Novo valor do centro de custo cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Novo"
Else
    USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
    Evento = "Alterar"
    If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
        Lista.SelectedItem = Lista.ListItems(CodigoLista)
        Lista.SetFocus
    End If
End If
'==================================
If Faturamento = True Then Modulo = "Faturamento/Nota fiscal/Terceiros" Else Modulo = Formulario
ID_documento = txtId
Documento1 = "Centro de custo: " & Cmb_CC
ProcGravaEvento
'==================================
Novo_Contas_CC = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkPercentual_Click()
On Error GoTo tratar_erro

If chkPercentual.Value = 1 Then
    chkValor.Value = 0
    With txtValor
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
    With txtPercentual
        .Text = ""
        valor = txtvalor_total
        Valor1 = txtSaldoCentro
        If Valor1 > 0 Then .Text = (Valor1 / valor) * 100
        .Locked = False
        .TabStop = True
        .SetFocus
    End With
Else
    With txtPercentual
        If chkValor.Value = 0 Then
            .Text = ""
            txtValor = ""
        End If
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkValor_Click()
On Error GoTo tratar_erro

If chkValor.Value = 1 Then
    chkPercentual.Value = 0
    With txtPercentual
        .Text = ""
        .Locked = True
        .TabStop = False
    End With
    With txtValor
        .Text = ""
        valor = txtSaldoCentro
        If valor > 0 Then .Text = txtSaldoCentro
        .Locked = False
        .TabStop = True
        .SetFocus
    End With
Else
    With txtValor
        If chkPercentual.Value = 0 Then
            .Text = ""
            txtPercentual = ""
        End If
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
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

ProcCarregaToolBar1 Me, 8565, 7, True
If Financeiro_Contas_Pagar = True Then
    Caption = "Financeiro - Contas a pagar - Centro de custo"
    Formulario = "Financeiro/Contas a pagar/Centro de custo"
    With frmContas_Pagar
        Id_Item = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
        IDlista = .txtidintconta
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select dbl_valorpagto from tbl_contaspagar where idintconta = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            txtvalor_total = IIf(IsNull(TBAbrir!dbl_valorpagto), 0, Format(TBAbrir!dbl_valorpagto, "###,##0.00"))
        End If
        TBAbrir.Close
        Documento = "Documento: " & .txtNDocumento
    End With
ElseIf Faturamento = True Then
        Caption = "Faturamento - Nota fiscal - Terceiros - Centro de custo"
        With frmFaturamento_Prod_Serv
            Id_Item = .txtIDEmpresa.Text
            IDlista = .Txt_ID_duplicata
            'txtValor_total = .txtValorDuplicata
            Documento = "N° nota: " & .txtNFiscal & " - Tipo: " & TipoNF & " - Série: " & .txtSerie
        End With
    Else
        Caption = "Financeiro - Contas pagas - Centro de custo"
        Formulario = "Financeiro/Contas pagas/Centro de custo"
        With frmContas_Pagas
            Id_Item = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
            IDlista = .txtidintconta
            txtvalor_total = .txt_ValorPago
            Documento = "Documento: " & .txtNDocumento
        End With
End If

ProcLimpaVariaveisPrincipais
ProcCarregaComboSetor Cmb_CC, "Setor IS NOT NULL and DtBloq IS NULL and (Consolidacao = 'False' or Consolidacao is null)", "", False, True, False, "", True, False
ProcCarregaLista

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CC_realizado.*, Usuarios_setor.Codigo, Usuarios_setor.Setor, Usuarios_setor.DtBloq, Usuarios_setor.ID as ID_CC_usuarios from CC_realizado INNER JOIN Usuarios_setor ON CC_realizado.ID_CC = Usuarios_setor.ID where CC_realizado.id = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    ProcLimpaCampos
    txtId = TBLISTA!ID
    txtData = Format(TBLISTA!Data, "dd/mm/yy")
    txtResponsavel = TBLISTA!Responsavel
    
    If IsNull(TBLISTA!CODIGO) = False And TBLISTA!CODIGO <> "" Then
        If IsNull(TBLISTA!DtBloq) = False Then
            Cmb_centro.AddItem TBLISTA!CODIGO & " - " & IIf(IsNull(TBLISTA!Setor), "", TBLISTA!Setor)
            Cmb_centro.ItemData(Cmb_centro.NewIndex) = TBLISTA!ID_CC_usuarios
        End If
        Cmb_CC = TBLISTA!CODIGO & " - " & IIf(IsNull(TBLISTA!Setor), "", TBLISTA!Setor)
    Else
        If IsNull(TBLISTA!DtBloq) = False Then
            Cmb_centro.AddItem IIf(IsNull(TBLISTA!Setor), "", TBLISTA!Setor)
            Cmb_centro.ItemData(Cmb_centro.NewIndex) = TBLISTA!ID_CC_usuarios
        End If
        Cmb_CC = IIf(IsNull(TBLISTA!Setor), "", TBLISTA!Setor)
    End If
    
    txtValor = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
    txtPercentual = IIf(IsNull(TBLISTA!Percentual), "", Format(TBLISTA!Percentual, "###,##0.00"))

    Novo_Contas_CC = False
    CodigoLista = Lista.SelectedItem.index
End If
TBLISTA.Close
Frame1.Enabled = True
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPercentual_Change()
On Error GoTo tratar_erro

If chkPercentual.Value = 1 Then
    txtValor = ""
    If txtPercentual <> "" Then
        VerifNumero = txtPercentual
        ProcVerificaNumero
        If VerifNumero = False Then
            txtPercentual = ""
            txtPercentual.SetFocus
            Exit Sub
        End If
        Qtde = txtPercentual
        Qtd = txtvalor_total
        qt = (Qtd * Qtde) / 100
        txtValor = Format(qt, "###,##0.00")
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPercentual_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtPercentual

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPercentual_LostFocus()
On Error GoTo tratar_erro

txtPercentual = Format(txtPercentual, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValor_Change()
On Error GoTo tratar_erro

If chkValor.Value = 1 Then
    txtPercentual = ""
    If txtValor <> "" Then
        VerifNumero = txtValor
        ProcVerificaNumero
        If VerifNumero = False Then
            txtValor = ""
            txtValor.SetFocus
            Exit Sub
        End If
        Qtde = txtValor
        Qtd = txtvalor_total
        If Qtd <> 0 Then qt = (Qtde * 100) / Qtd Else qt = 0
        txtPercentual = Format(qt, "###,##0.00")
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValor_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtValor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtValor_LostFocus()
On Error GoTo tratar_erro

txtValor = Format(txtValor, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcEnviaDados(ID_CC As Long, ID_origem As Long)
On Error GoTo tratar_erro

If Faturamento = True Then
    TBFI!id_duplicata = IDlista
    TBFI!ID_financeiro = 0
Else
    TBFI!ID_financeiro = IDlista
    TBFI!id_duplicata = 0
End If
If Financeiro_Contas_Pagar = True Then
    TBFI!Data = frmContas_Pagar.txtDtpagto
ElseIf Faturamento = True Then
        'TBFI!data = frmFaturamento_Prod_Serv.txt_Vencimento
    Else
        TBFI!Data = frmContas_Pagas.txtDataPagto
End If
If txtResponsavel <> "" Then TBFI!Responsavel = txtResponsavel Else TBFI!Responsavel = pubUsuario
TBFI!ID_empresa = Id_Item
TBFI!Operacao = "Débito"
TBFI!ID_CC = ID_CC
TBFI!valor = txtValor
TBFI!Percentual = txtPercentual
TBFI!Bloqueado = False
TBFI!ID_origem = ID_origem

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

ValorTotal = 0
Lista.ListItems.Clear
If Faturamento = True Then TextoFiltro = "CC_realizado.ID_duplicata = " Else TextoFiltro = "CC_realizado.ID_financeiro = "
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select CC_realizado.*, Usuarios_setor.Codigo, Usuarios_setor.Setor from CC_realizado INNER JOIN Usuarios_setor ON CC_realizado.ID_CC = Usuarios_setor.ID where " & TextoFiltro & IDlista & " and Usuarios_setor.Consolidacao = 'False' and (CC_realizado.ID_origem is null or CC_realizado.ID_origem = 0) order by Usuarios_setor.Codigo", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CODIGO), "", TBLISTA!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Setor), "", TBLISTA!Setor)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!valor), "", Format(TBLISTA!valor, "###,##0.00"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Percentual), "", Format(TBLISTA!Percentual, "###,##0.00"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!ID_CC), "", TBLISTA!ID_CC)
            ValorTotal = Format(ValorTotal + IIf(IsNull(TBLISTA!valor), 0, TBLISTA!valor), "###,##0.00")
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close
txtTotalCentro = Format(ValorTotal, "###,##0.00")
Qtd = IIf(txtvalor_total = "", 0, txtvalor_total)
txtSaldoCentro = Format(Qtd - ValorTotal, "###,##0.00")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcSalvar
    Case 3: ProcExcluir
    'Case 5: ProcAjuda
    Case 6: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
