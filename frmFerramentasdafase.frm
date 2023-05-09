VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFerramentasdafase 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Processos - Gerenciamento de processos - Utensílios da fase"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   1080
      ScreenWidth     =   2560
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10500
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.TextBox Txt_ID_acessorio 
      BackColor       =   &H80000014&
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
      Left            =   2730
      MouseIcon       =   "frmFerramentasdafase.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Text            =   "0"
      ToolTipText     =   "idferramenta."
      Top             =   6570
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Height          =   825
      Left            =   55
      TabIndex        =   14
      Top             =   1830
      Width           =   15195
      Begin VB.TextBox Txt_codinterno 
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
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   390
         Width           =   2055
      End
      Begin VB.CommandButton cmdProduto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2265
         Picture         =   "frmFerramentasdafase.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Localizar ferramenta."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox Txt_descricao 
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
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   390
         Width           =   10755
      End
      Begin VB.TextBox Txt_quant 
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
         Left            =   13440
         TabIndex        =   9
         ToolTipText     =   "Quantidade."
         Top             =   390
         Width           =   1572
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade*"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   13761
         TabIndex        =   17
         Top             =   180
         Width           =   930
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno*"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   712
         TabIndex        =   16
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7702
         TabIndex        =   15
         Top             =   180
         Width           =   690
      End
   End
   Begin VB.TextBox txtid 
      BackColor       =   &H80000014&
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
      Left            =   2070
      MouseIcon       =   "frmFerramentasdafase.frx":040C
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Text            =   "0"
      ToolTipText     =   "idferramenta."
      Top             =   6570
      Visible         =   0   'False
      Width           =   645
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   7070
      Left            =   60
      TabIndex        =   10
      Top             =   2670
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   12462
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   19059
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Tipo"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Height          =   825
      Left            =   55
      TabIndex        =   11
      Top             =   990
      Width           =   15195
      Begin VB.TextBox txtVersao 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9288
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Versão da fase."
         Top             =   390
         Width           =   1015
      End
      Begin VB.TextBox txtrev 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2106
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "Revisão."
         Top             =   390
         Width           =   475
      End
      Begin VB.TextBox txtfase 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10314
         MaxLength       =   20
         TabIndex        =   4
         ToolTipText     =   "Fase."
         Top             =   390
         Width           =   1065
      End
      Begin VB.TextBox txtcodinterno 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         MaxLength       =   20
         TabIndex        =   0
         ToolTipText     =   "Código interno."
         Top             =   390
         Width           =   1905
      End
      Begin VB.TextBox txtmaquina 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11400
         MaxLength       =   20
         TabIndex        =   5
         ToolTipText     =   "Posto de trabalho."
         Top             =   390
         Width           =   3585
      End
      Begin VB.TextBox txtdescricao 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2592
         MaxLength       =   20
         TabIndex        =   2
         ToolTipText     =   "Descrição."
         Top             =   390
         Width           =   6675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Versão"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9543
         TabIndex        =   21
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posto de trabalho"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12555
         TabIndex        =   19
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fase"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   10674
         TabIndex        =   18
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5584
         TabIndex        =   13
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   682
         TabIndex        =   12
         Top             =   180
         Width           =   900
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   23
      Top             =   9750
      Width           =   15195
      _ExtentX        =   26802
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
      SearchText      =   ""
      Value           =   0
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   22
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
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
         Name            =   "Tahoma"
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
      ButtonUseMaskColor7=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   11640
         Top             =   195
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmFerramentasdafase.frx":0716
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmFerramentasdafase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_Processo_Ferramenta As Boolean 'OK

Private Sub cmdProduto_Click()
On Error GoTo tratar_erro

Vendas_Analise = False
Permitido = False
frmFerramentasdafase_localizar.Show 1
If Permitido = True Then Txt_quant.SetFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF2: cmdProduto_Click
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 7, True
Formulario = "Engenharia/Processos"
Direitos
With frmProcessos
    txtCodinterno.Text = .txtdesenho.Text
    txtRev.Text = .txtrevdesenho.Text
    txtVersao.Text = .cmbVersao
    txtFase.Text = .txtFase.Text
    txtmaquina.Text = .cmbMaquina.Text
    txtdescricao.Text = .txtProduto.Text
End With
ProcAtualizaFerramentas

ProcRemoveObjetosResize Me
       
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro

If Excluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Permitido = False
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If Permitido = False Then
                If USMsgBox("Deseja realmente excluir este(s) utensílio(s) da fase?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from ferramentas where idferramenta = " & .ListItems(InitFor)
            
            '==================================
            Modulo = "Engenharia/Processos/Utensílios da fase"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            With frmProcessos
                Documento = "Processo: " & .txtidprocesso & " - Rev.: " & .txtrevproc & " - Cód. interno: " & .txtdesenho & " - Rev.: " & .txtrevdesenho & " - Fase: " & .txtFase
            End With
            Documento1 = "Cód. interno: " & .ListItems(InitFor).SubItems(1)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) utensílio(s) da fase antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Utensílio(s) da fase excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcAtualizaFerramentas
    ProcLimpaCampos
End If
      
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovo()
On Error GoTo tratar_erro
  
If Incluir = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If FunVerificaRegistroValidado("Processos", "IDProcesso = " & frmProcessos.txtidprocesso, "processo", "utensílio da fase", "criar novo", True, True) = False Then Exit Sub
Frame4.Enabled = True
ProcLimpaCampos
Novo_Processo_Ferramenta = True
cmdProduto_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtId = 0
Txt_ID_acessorio = 0
Txt_codinterno = ""
Txt_descricao = ""
Txt_quant = ""
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro
  
If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & " você não tem acesso a este recurso."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
If Frame4.Enabled = False Then
    ProcVerificaSalvar
    Exit Sub
End If
Acao = "salvar"
If Txt_codinterno.Text = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    cmdProduto_Click
    Exit Sub
End If
Quant = IIf(Txt_quant = "", 0, Txt_quant)
If Quant <= 0 Then
    NomeCampo = "a quantidade"
    ProcVerificaAcao
    Txt_quant.SetFocus
    Exit Sub
End If
With frmProcessos
    Set TBFerramentas = CreateObject("adodb.recordset")
    TBFerramentas.Open "Select * from ferramentas where IDFerramenta = " & IIf(txtId = "", 0, txtId), Conexao, adOpenKeyset, adLockOptimistic
    If TBFerramentas.EOF = True Then
        TBFerramentas.AddNew
    Else
        If FunVerificaRegistroValidado("Processos", "IDProcesso = " & .txtidprocesso, "processo", "o utensílio da fase", "alterar", True, True) = False Then Exit Sub
    End If
    TBFerramentas!IDPROCESSO = .txtidprocesso.Text
    TBFerramentas!IDFase = .ListaFases.SelectedItem
    TBFerramentas!ID_acessorio = IIf(Txt_ID_acessorio = "", 0, Txt_ID_acessorio)
    TBFerramentas!Numero = Txt_codinterno.Text
    TBFerramentas!quantidade = Txt_quant.Text
    TBFerramentas.Update
    txtId.Text = TBFerramentas!IDFerramenta
    TBFerramentas.Close
    
    If Novo_Processo_Ferramenta = True Then
        USMsgBox ("Novo utensílio da fase cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Novo"
        ProcAtualizaFerramentas
    Else
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar"
        ProcAtualizaFerramentas
        If CodigoLista <> 0 And Lista.ListItems.Count <> 0 Then
            Lista.SelectedItem = Lista.ListItems(CodigoLista)
            Lista.SetFocus
        End If
    End If
    '==================================
    Modulo = "Engenharia/Processos/Utensílios da fase"
    ID_documento = txtId
    Documento = "Processo: " & .txtidprocesso & " - Rev.: " & .txtrevproc & " - Cód. interno: " & .txtdesenho & " - Rev.: " & .txtrevdesenho & " - Fase: " & .txtFase
    Documento1 = "Cód. interno: " & Txt_codinterno
    ProcGravaEvento
    '==================================
    Novo_Processo_Ferramenta = False
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then
                .ListItems.Item(InitFor).Checked = False
            Else
                If FunVerificaRegistroValidadoSemMsg("Processos", "IDprocesso = " & frmProcessos.txtidprocesso, True) = False Then
                    .ListItems.Item(InitFor).Checked = False
                    GoTo Proximo
                End If
                .ListItems.Item(InitFor).Checked = True
Proximo:
            End If
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True And .ListItems.Item(InitFor) = Item Then
            If FunVerificaRegistroValidado("Processos", "IDProcesso = " & frmProcessos.txtidprocesso, "processo", "utensílio da fase", "excluir este", True, True) = False Then .ListItems.Item(InitFor).Checked = False
        End If
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
  
If Lista.ListItems.Count = 0 Then Exit Sub
Set TBFerramentas = CreateObject("adodb.recordset")
TBFerramentas.Open "Select F.*, P.descricao from Ferramentas F INNER JOIN projproduto P ON F.Numero = P.Desenho where F.IDFerramenta = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBFerramentas.EOF = False Then
    ProcLimpaCampos
    txtId = TBFerramentas!IDFerramenta
    Txt_ID_acessorio = TBFerramentas!ID_acessorio
    Txt_codinterno.Text = IIf(IsNull(TBFerramentas!Numero), 0, TBFerramentas!Numero)
    Txt_descricao.Text = IIf(IsNull(TBFerramentas!Descricao), "", TBFerramentas!Descricao)
    Txt_quant.Text = IIf(IsNull(TBFerramentas!quantidade), "", Format(TBFerramentas!quantidade, "###,##0.0000"))
    CodigoLista = Lista.SelectedItem.index
    Frame4.Enabled = True
    Novo_Processo_Ferramenta = False
End If
TBFerramentas.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaFerramentas()
On Error GoTo tratar_erro
  
Lista.ListItems.Clear
Set TBFerramentas = CreateObject("adodb.recordset")
TBFerramentas.Open "Select F.*, P.descricao from Ferramentas F INNER JOIN projproduto P ON F.Numero = P.Desenho where F.IDProcesso = " & frmProcessos.txtidprocesso & " and F.idFase = " & frmProcessos.ListaFases.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBFerramentas.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBFerramentas.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBFerramentas.EOF = False
        With Lista.ListItems
            .Add , , TBFerramentas!IDFerramenta
            .Item(.Count).SubItems(1) = TBFerramentas!Numero
            .Item(.Count).SubItems(2) = TBFerramentas!Descricao
            .Item(.Count).SubItems(3) = Format(TBFerramentas!quantidade, "###,##0.0000")
            If IsNull(TBFerramentas!ID_acessorio) = False And TBFerramentas!ID_acessorio <> 0 Then .Item(.Count).SubItems(4) = "Acessório" Else .Item(.Count).SubItems(4) = "Ferramenta"
        End With
        TBFerramentas.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBFerramentas.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_quant_Change()
On Error GoTo tratar_erro

If Txt_quant.Text <> "" Then
    VerifNumero = Txt_quant.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_quant.Text = ""
        Txt_quant.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_quant_LostFocus()
On Error GoTo tratar_erro

Txt_quant.Text = Format(Txt_quant.Text, "###,##0.0000")

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

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_Processo_Ferramenta = True Then
    If USMsgBox("O utensílio ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_Processo_Ferramenta = True Then Exit Sub Else Unload Me
    End If
End If
Novo_Processo_Ferramenta = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
