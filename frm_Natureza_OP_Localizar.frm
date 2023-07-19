VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frm_Natureza_OP_Localizar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Fiscal | CFOP | Localizar"
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   5835
   ClipControls    =   0   'False
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
   ScaleHeight     =   3795
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   13
      Top             =   3390
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   767
      DibPicture      =   "frm_Natureza_OP_Localizar.frx":0000
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
      Icon            =   "frm_Natureza_OP_Localizar.frx":3650
      ShowMaximizeButton=   0   'False
      ShowMinimizeButton=   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções para pesquisa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   270
      TabIndex        =   8
      Top             =   720
      Width           =   5115
      Begin MSMask.MaskEdBox txtTexto1 
         Height          =   255
         Left            =   3630
         TabIndex        =   0
         Top             =   570
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         BackColor       =   12640511
         MaxLength       =   5
         Mask            =   "#.###"
         PromptChar      =   "_"
      End
      Begin DrawSuite2022.USButton btnFiltrar 
         Height          =   405
         Left            =   3600
         TabIndex        =   3
         Top             =   1650
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   714
         DibPicture      =   "frm_Natureza_OP_Localizar.frx":396A
         Caption         =   "Filtrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         GradientColor1  =   5263559
         GradientColor2  =   5263559
         GradientColor3  =   5263559
         GradientColor4  =   5263559
         GradientColorDisabled1=   13160660
         GradientColorDisabled2=   13160660
         GradientColorDisabled3=   13160660
         GradientColorDisabled4=   13160660
         GradientColorOver1=   4408288
         GradientColorOver2=   4408288
         GradientColorOver3=   4408288
         GradientColorOver4=   4408288
         GradientColorDown1=   4013465
         GradientColorDown2=   4013465
         GradientColorDown3=   4013465
         GradientColorDown4=   4013465
         PicSize         =   1
         ShowFocusRect   =   0   'False
         Theme           =   4
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
         Height          =   510
         Left            =   300
         TabIndex        =   11
         Top             =   990
         Width           =   4485
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   7
            Top             =   180
            Width           =   705
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1350
            TabIndex        =   5
            Top             =   180
            Width           =   675
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   180
            Value           =   -1  'True
            Width           =   705
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2490
            TabIndex        =   6
            Top             =   180
            Width           =   585
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         ItemData        =   "frm_Natureza_OP_Localizar.frx":6FBA
         Left            =   330
         List            =   "frm_Natureza_OP_Localizar.frx":6FC7
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Opções para filtro."
         Top             =   540
         Width           =   2715
      End
      Begin VB.TextBox txtTexto 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.###"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
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
         Left            =   3060
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   540
         Width           =   1725
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Localizar por"
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
         Left            =   1237
         TabIndex        =   10
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Localizar"
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
         Left            =   3660
         TabIndex        =   9
         Top             =   330
         Width           =   630
      End
   End
End
Attribute VB_Name = "frm_Natureza_OP_Localizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcFiltrar()
On Error GoTo tratar_erro
Dim varValidacao As String

With frm_Natureza_OP

    If txtTexto <> "" Or txtTexto1.Text <> "_.___" Then
        Select Case cmbfiltrarpor
            Case "CFOP": TextoFiltro = "id_CFOP"
            varTexto = txtTexto1.Text
            
            Case "Natureza da operação": TextoFiltro = "txt_Descricao"
            varTexto = txtTexto.Text
            
            Case "ID": TextoFiltro = "IDCountCfop"
            varTexto = txtTexto.Text

        End Select
        
        
        If Optinicio.Value = True Then
            .StrSql_CFOP = "Select * from tbl_NaturezaOperacao where " & TextoFiltro & " like '" & varTexto & "%' order by id_CFOP, IDCountCfop"
            .FormulaRel_CFOP = "{tbl_NaturezaOperacao." & TextoFiltro & "} like '" & varTexto & "*'"
        End If
        If Optmeio.Value = True Then
            .StrSql_CFOP = "Select * from tbl_NaturezaOperacao where " & TextoFiltro & " like '%" & varTexto & "%' order by id_CFOP, IDCountCfop"
            .FormulaRel_CFOP = "{tbl_NaturezaOperacao." & TextoFiltro & "} like '*" & varTexto & "*'"
        End If
        If Optfim.Value = True Then
            .StrSql_CFOP = "Select * from tbl_NaturezaOperacao where " & TextoFiltro & " like '%" & varTexto & "' order by id_CFOP, IDCountCfop"
            .FormulaRel_CFOP = "{tbl_NaturezaOperacao." & TextoFiltro & "} like '*" & varTexto & "'"
        End If
        If optIgual.Value = True Then
            .StrSql_CFOP = "Select * from tbl_NaturezaOperacao where " & TextoFiltro & " = '" & varTexto & "' order by id_CFOP, IDCountCfop"
            .FormulaRel_CFOP = "{tbl_NaturezaOperacao." & TextoFiltro & "} = '" & varTexto & "'"
        End If
    Else

        .StrSql_CFOP = "Select * from tbl_NaturezaOperacao order by id_CFOP, IDCountCfop"
        .FormulaRel_CFOP = ""

    End If
    .ProcCarregaLista (1)
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnFiltrar_Click()
On Error GoTo tratar_erro

ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Change()
On Error GoTo tratar_erro

If cmbfiltrarpor.Text = "CFOP" Then
    txtTexto1.Visible = True
    txtTexto.Visible = True
Else
    txtTexto1.Visible = True
    txtTexto.Visible = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor.Text = "CFOP" Then
    txtTexto1.Visible = True
    txtTexto.Visible = True
Else
    txtTexto1.Visible = False
    txtTexto.Visible = True
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
    Case vbKeyF2: ProcFiltrar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

cmbfiltrarpor = "CFOP"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

