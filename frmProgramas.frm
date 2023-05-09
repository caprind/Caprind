VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgramas 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Processos - Gerenciamento de processos - Programas da fase"
   ClientHeight    =   10065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15390
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
   ScaleHeight     =   10065
   ScaleWidth      =   15390
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
      FormHeightDT    =   10530
      FormWidthDT     =   15510
      FormScaleHeightDT=   10065
      FormScaleWidthDT=   15390
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.TextBox Txt_ID 
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
      Left            =   1620
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   6780
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4095
      Left            =   60
      TabIndex        =   10
      Top             =   5655
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483628
      BorderStyle     =   1
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Programa"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Ciclo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   21175
      EndProperty
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   14
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   1720
      ButtonCount     =   9
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
      ButtonCaption4  =   "Importar"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Importa programa de arquivo (F6)"
      ButtonKey4      =   "4"
      ButtonAlignment4=   2
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   118
      ButtonTop4      =   2
      ButtonWidth4    =   50
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Exportar"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Exporta programa para arquivo (F7)"
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
      ButtonLeft5     =   170
      ButtonTop5      =   2
      ButtonWidth5    =   50
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonAlignment6=   2
      ButtonType6     =   1
      ButtonStyle6    =   -1
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState6    =   -1
      ButtonLeft6     =   222
      ButtonTop6      =   4
      ButtonWidth6    =   2
      ButtonHeight6   =   54
      ButtonCaption7  =   "Ajuda"
      ButtonEnabled7  =   0   'False
      ButtonIconSize7 =   32
      ButtonToolTipText7=   "Ajuda (F1)"
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
      ButtonLeft7     =   226
      ButtonTop7      =   2
      ButtonWidth7    =   36
      ButtonHeight7   =   21
      ButtonUseMaskColor7=   0   'False
      ButtonCaption8  =   "Sair"
      ButtonEnabled8  =   0   'False
      ButtonIconSize8 =   32
      ButtonToolTipText8=   "Sair (Esc)"
      ButtonKey8      =   "8"
      ButtonAlignment8=   2
      BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft8     =   264
      ButtonTop8      =   2
      ButtonWidth8    =   26
      ButtonHeight8   =   21
      ButtonUseMaskColor8=   0   'False
      ButtonEnabled9  =   0   'False
      ButtonIconSize9 =   32
      ButtonKey9      =   "9"
      ButtonAlignment9=   2
      BeginProperty ButtonFont9 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState9    =   5
      ButtonLeft9     =   292
      ButtonTop9      =   2
      ButtonWidth9    =   24
      ButtonHeight9   =   24
      ButtonUseMaskColor9=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   11640
         Top             =   195
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmProgramas.frx":0000
         Count           =   1
      End
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
      Height          =   3825
      Left            =   60
      TabIndex        =   11
      Top             =   1830
      Width           =   15195
      Begin VB.TextBox txtPrograma 
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
         MaxLength       =   11
         TabIndex        =   6
         ToolTipText     =   "Número do programa."
         Top             =   390
         Width           =   1440
      End
      Begin VB.TextBox Txt_ciclo 
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
         Left            =   1630
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Ciclo."
         Top             =   390
         Width           =   945
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   9420
         Top             =   1590
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Filter          =   "*.nc"
         FontName        =   "Tahoma"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8820
         Top             =   1590
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProgramas.frx":4151
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtDescricao_programa 
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
         Left            =   2590
         MaxLength       =   100
         TabIndex        =   8
         ToolTipText     =   "Descrição."
         Top             =   390
         Width           =   12405
      End
      Begin RichTextLib.RichTextBox txtprogramacao 
         Height          =   2925
         Left            =   180
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Programa CNC."
         Top             =   780
         Width           =   14835
         _ExtentX        =   26167
         _ExtentY        =   5159
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmProgramas.frx":4211
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ciclo"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1935
         TabIndex        =   15
         Top             =   180
         Width           =   330
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Programa*"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   510
         TabIndex        =   13
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição*"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8402
         TabIndex        =   12
         Top             =   180
         Width           =   780
      End
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
      Left            =   60
      TabIndex        =   16
      Top             =   990
      Width           =   15195
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   682
         TabIndex        =   21
         Top             =   180
         Width           =   900
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5584
         TabIndex        =   20
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fase"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   10674
         TabIndex        =   19
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posto de trabalho"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12555
         TabIndex        =   18
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Versão"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9543
         TabIndex        =   17
         Top             =   180
         Width           =   495
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   22
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
End
Attribute VB_Name = "frmProgramas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_programa As Boolean 'OK

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_ID = 0
txtPrograma.Text = ""
Txt_ciclo = ""
txtDescricao_programa.Text = ""
txtprogramacao.Text = ""
CodigoLista = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExportar()
On Error GoTo tratar_erro

If txtPrograma = "" Then Exit Sub
If Novo_programa = True Then
    USMsgBox ("Salve o programa antes de exportar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
Filter = "(*.*) | *.*"
With CD1
    .filename = ""
    .Filter = Filter
    .InitDir = App.Path
    .DefaultExt = "*.*"
    CD1.ShowSave
    If Dir$(.filename) = "" Then
        Call FunExportaArquivo(.filename)
    Else
        If USMsgBox("Deseja substituir o arquivo existente ?", vbYesNo, "CAPRIND v5.0") = vbYes Then Call FunExportaArquivo(.filename)
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Exportação cancelada."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBProgramas = CreateObject("adodb.recordset")
TBProgramas.Open "Select * from Programas where IDProcesso = " & frmProcessos.txtidprocesso & " AND IDFase = " & frmProcessos.ListaFases.SelectedItem & " order by idprograma", Conexao, adOpenKeyset, adLockOptimistic
If TBProgramas.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBProgramas.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBProgramas.EOF = False
        With Lista.ListItems
            .Add , , TBProgramas!IDPrograma
            .Item(.Count).SubItems(1) = IIf(IsNull(TBProgramas!programa), "", TBProgramas!programa)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBProgramas!Ciclo), "", TBProgramas!Ciclo)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBProgramas!Descricao), "", TBProgramas!Descricao)
        End With
        TBProgramas.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBProgramas.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procImportar()
On Error GoTo tratar_erro

If txtPrograma = "" Then Exit Sub
If FunVerificaRegistroValidado("Processos", "IDProcesso = " & frmProcessos.txtidprocesso, "processo", "o programa da fase", "importar", True, True) = False Then Exit Sub
Filter = "(*.*) | *.*"
With CD1
    .filename = ""
    .Filter = Filter
    .filename = App.Path
    .DefaultExt = "*.*"
    .ShowOpen
    If Dir$(.filename) <> "" Then txtprogramacao.filename = .filename
End With

Exit Sub
tratar_erro:
    USMsgBox ("Importação cancelada com sucesso."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyInsert: ProcNovo
    Case vbKeyF3: ProcSalvar
    Case vbKeyF4: ProcExcluir
    Case vbKeyF6: procImportar
    Case vbKeyF7: ProcExportar
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 9, True
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
ProcCarregaLista

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
                If USMsgBox("Deseja realmente excluir este(s) programa(s) da fase?", vbYesNo, "CAPRIND v5.0") = vbYes Then GoTo 1 Else Exit Sub
            End If
1:
            Permitido = True
            Conexao.Execute "DELETE from Programas where IDPrograma = " & .ListItems(InitFor)
            
            '==================================
            Modulo = "Engenharia/Processos/Programas da fase"
            Evento = "Excluir"
            ID_documento = .ListItems(InitFor)
            With frmProcessos
                Documento = "Processo: " & .txtidprocesso & " - Rev.: " & .txtrevproc & " - Cód. interno: " & .txtdesenho & " - Rev.: " & .txtrevdesenho & " - Fase: " & .txtFase
            End With
            Documento1 = "Programa: " & .ListItems(InitFor).SubItems(1) & " - Ciclo: " & .ListItems(InitFor).SubItems(2)
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With
If Permitido = False Then
    USMsgBox ("Informe o(s) programa(s) da fase antes de excluir."), vbExclamation, "CAPRIND v5.0"
Else
    USMsgBox ("Programa(s) da fase excluído(s) com sucesso."), vbInformation, "CAPRIND v5.0"
    ProcLimpaCampos
    ProcCarregaLista
    Novo_programa = False
    Frame4.Enabled = False
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
With frmProcessos
    If FunVerificaRegistroValidado("Processos", "IDProcesso = " & .txtidprocesso, "processo", "programa da fase", "criar novo", True, True) = False Then Exit Sub
    ProcLimpaCampos
    Set TBProgramas = CreateObject("adodb.recordset")
    TBProgramas.Open "Select programa from programas order by programa desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBProgramas.EOF = False Then
        txtPrograma.Text = TBProgramas!programa + 1
    Else
        txtPrograma.Text = 1
    End If
    Set TBProgramas = CreateObject("adodb.recordset")
    TBProgramas.Open "Select Ciclo from programas where idprocesso = " & .txtidprocesso & " and idfase = " & .ListaFases.SelectedItem & " order by Ciclo desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBProgramas.EOF = False Then
        Ciclo = TBProgramas!Ciclo + 1
    Else
        Ciclo = 1
    End If
    TBProgramas.Close
    Txt_ciclo.Text = Ciclo
    Frame4.Enabled = True
    Novo_programa = True
    txtPrograma.SetFocus
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

If Novo_programa = True Then
    If USMsgBox("O programa ainda não foi salvo, deseja salvar antes de fechar o módulo?", vbYesNo) = vbYes Then
        ProcSalvar
        If Novo_programa = True Then Exit Sub Else Unload Me
    End If
End If
Novo_programa = False
Unload Me

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
If txtDescricao_programa = "" Then
    NomeCampo = "a descrição do programa"
    ProcVerificaAcao
    txtDescricao_programa.SetFocus
    Exit Sub
End If
If txtprogramacao.Text = "" Then
    USMsgBox ("Importe o programa antes de salvar."), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If

If Novo_programa = True Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "select IDPrograma from programas where Programa = " & txtPrograma, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        USMsgBox ("Número do programa já existente, favor alterar."), vbExclamation, "CAPRIND v5.0"
        txtPrograma.SetFocus
        Exit Sub
    End If
End If

With frmProcessos
    Set TBProgramas = CreateObject("adodb.recordset")
    TBProgramas.Open "Select * from Programas where IDPrograma = " & Txt_ID, Conexao, adOpenKeyset, adLockOptimistic
    If TBProgramas.EOF = True Then
        TBProgramas.AddNew
    Else
        If FunVerificaRegistroValidado("Processos", "IDProcesso = " & .txtidprocesso, "processo", "programa da fase", "alterar", True, True) = False Then Exit Sub
    End If
    TBProgramas!Data = Date
    TBProgramas!IDPROCESSO = .txtidprocesso
    TBProgramas!Desenho = txtCodinterno
    TBProgramas!maquina = txtmaquina
    TBProgramas!Ciclo = Txt_ciclo.Text
    TBProgramas!IDFase = .ListaFases.SelectedItem
    TBProgramas!programa = txtPrograma
    TBProgramas!Descricao = txtDescricao_programa.Text
    TBProgramas!Programacao = IIf(txtprogramacao.TextRTF = "", Null, txtprogramacao.TextRTF)
    TBProgramas.Update
    Txt_ID = TBProgramas!IDPrograma
    TBProgramas.Close
    If Novo_programa = True Then
        USMsgBox ("Novo programa da fase cadastrado com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Novo"
    Else
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        Evento = "Alterar"
        If CodigoLista <> 0 And ListaHistorico.ListItems.Count <> 0 Then
            ListaHistorico.SelectedItem = ListaHistorico.ListItems(CodigoLista)
            ListaHistorico.SetFocus
        End If
    End If
    '==================================
    Modulo = "Engenharia/Processos/Programas da fase"
    ID_documento = Txt_ID
    With frmProcessos
        Documento = "Processo: " & .txtidprocesso & " - Rev.: " & .txtrevproc & " - Cód. interno: " & .txtdesenho & " - Rev.: " & .txtrevdesenho & " - Fase: " & .txtFase
    End With
    Documento1 = "Programa: " & txtPrograma & " - Ciclo: " & Txt_ciclo
    ProcGravaEvento
    '==================================
    Novo_programa = False
    ProcCarregaLista
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
            If FunVerificaRegistroValidado("Processos", "IDProcesso = " & frmProcessos.txtidprocesso, "processo", "programa da fase", "excluir este", True, True) = False Then .ListItems.Item(InitFor).Checked = False
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
Set TBProgramas = CreateObject("adodb.recordset")
TBProgramas.Open "Select * from Programas where IDPrograma = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBProgramas.EOF = False Then
    Txt_ID = TBProgramas!IDPrograma
    txtPrograma = TBProgramas!programa
    Txt_ciclo.Text = TBProgramas!Ciclo
    txtDescricao_programa.Text = TBProgramas!Descricao
    txtprogramacao.TextRTF = IIf(IsNull(TBProgramas!Programacao), "", TBProgramas!Programacao)
    CodigoLista = Lista.SelectedItem.index
End If
TBProgramas.Close
Frame4.Enabled = True
Novo_programa = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunExportaArquivo(Arquivo As String)
On Error GoTo tratar_erro

txtprogramacao.SaveFile Arquivo, 0
USMsgBox "Arquivo exportado com sucesso em: " & Chr(13) & Chr(13) & Arquivo & Chr(13) & Chr(13), vbInformation, "CAPRIND v5.0"

Exit Function
tratar_erro:
    USMsgBox ("Este arquivo não é valido, favor selecionar um arquivo de texto."), vbInformation, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcSalvar
    Case 3: ProcExcluir
    Case 4: procImportar
    Case 5: ProcExportar
    'Case 7: ProcAjuda
    Case 8: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
