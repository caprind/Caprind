VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Orcamento_abrir 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Outros | Orçamentos - Localizar"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   8385
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
   ScaleHeight     =   7245
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2085
      Left            =   4410
      TabIndex        =   11
      Top             =   750
      Width           =   3585
      Begin VB.CheckBox chkPeriodo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Por período de emissão"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   18
         Top             =   0
         Width           =   2355
      End
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   2250
         TabIndex        =   12
         ToolTipText     =   "Data final."
         Top             =   420
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
         Format          =   198639617
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   570
         TabIndex        =   13
         ToolTipText     =   "Data inicio."
         Top             =   420
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
         Format          =   198639617
         CurrentDate     =   39057
      End
      Begin DrawSuite2022.USButton btnFiltrar 
         Height          =   915
         Left            =   210
         TabIndex        =   16
         Top             =   960
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   1614
         DibPicture      =   "frm_Orcamento_abrir.frx":0000
         Caption         =   "Filtrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         PicAlign        =   7
         PicSize         =   3
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Até:"
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
         Left            =   1890
         TabIndex        =   15
         Top             =   510
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   510
         Width           =   255
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   767
      DibPicture      =   "frm_Orcamento_abrir.frx":3650
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
      Icon            =   "frm_Orcamento_abrir.frx":6CA0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opões para filtro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   360
      TabIndex        =   7
      Top             =   750
      Width           =   4035
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   270
         TabIndex        =   9
         Top             =   750
         WhatsThisHelpID =   210
         Width           =   3525
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
            Left            =   1680
            TabIndex        =   5
            Top             =   180
            Width           =   555
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
            Left            =   180
            TabIndex        =   3
            Top             =   180
            Value           =   -1  'True
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
            Left            =   930
            TabIndex        =   4
            Top             =   180
            Width           =   675
         End
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
            Left            =   2490
            TabIndex        =   6
            Top             =   180
            Width           =   705
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
         ItemData        =   "frm_Orcamento_abrir.frx":6FBA
         Left            =   270
         List            =   "frm_Orcamento_abrir.frx":6FCA
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3555
      End
      Begin VB.TextBox txtTexto 
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
         Left            =   270
         MaxLength       =   255
         TabIndex        =   1
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1560
         Width           =   3525
      End
      Begin VB.ComboBox Cmb_status 
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
         Height          =   330
         ItemData        =   "frm_Orcamento_abrir.frx":7014
         Left            =   270
         List            =   "frm_Orcamento_abrir.frx":7016
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1560
         Width           =   3525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
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
         Left            =   1275
         TabIndex        =   8
         Top             =   1350
         Width           =   1470
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4065
      Left            =   150
      TabIndex        =   17
      Top             =   2970
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   7170
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
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "N"
         Text            =   "Orçamento"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Item"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Unidade"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Descrição"
         Object.Width           =   3881
      EndProperty
   End
End
Attribute VB_Name = "frm_Orcamento_abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

DataFiltro = "data Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"

CamposFiltro = " ID_Orcamento, Codigo, txtCodigoproduto, txtunidade, txtdescricao"
    
    If txtTexto.Visible = True Then
       Select Case cmbfiltrarpor
           Case "Número do orçamento": TextoFiltro = "Codigo"
           Case "Código interno": TextoFiltro = "txtCodigoproduto"
           Case "Código de referência": TextoFiltro = "txtreferencia"
           Case "Descrição": TextoFiltro = "txtdescricao"
       End Select
       If chkPeriodo.Value = 1 Then
            If Optinicio.Value = True Then StrSql = "Select " & CamposFiltro & " from Vendas_orcamento where " & TextoFiltro & " like '" & txtTexto.Text & "%' and " & DataFiltro & " order by Codigo desc, ID_orcamento desc"
            If Optmeio.Value = True Then StrSql = "Select " & CamposFiltro & " from Vendas_orcamento where " & TextoFiltro & " like '%" & txtTexto.Text & "%' and " & DataFiltro & " order by Codigo desc, ID_orcamento desc"
            If Optfim.Value = True Then StrSql = "Select " & CamposFiltro & " from Vendas_orcamento where " & TextoFiltro & " like '%" & txtTexto.Text & "' and " & DataFiltro & " order by Codigo desc, ID_orcamento desc"
            If optIgual.Value = True Then StrSql = "Select " & CamposFiltro & " from Vendas_orcamento where " & TextoFiltro & " = '" & txtTexto.Text & "' and " & DataFiltro & " order by Codigo desc, ID_orcamento desc"
       Else
            If Optinicio.Value = True Then StrSql = "Select " & CamposFiltro & " from Vendas_orcamento where " & TextoFiltro & " like '" & txtTexto.Text & "%' order by Codigo desc, ID_orcamento desc"
            If Optmeio.Value = True Then StrSql = "Select " & CamposFiltro & " from Vendas_orcamento where " & TextoFiltro & " like '%" & txtTexto.Text & "%' order by Codigo desc, ID_orcamento desc"
            If Optfim.Value = True Then StrSql = "Select " & CamposFiltro & " from Vendas_orcamento where " & TextoFiltro & " like '%" & txtTexto.Text & "' order by Codigo desc, ID_orcamento desc"
            If optIgual.Value = True Then StrSql = "Select " & CamposFiltro & " from Vendas_orcamento where " & TextoFiltro & " = '" & txtTexto.Text & "' order by Codigo desc, ID_orcamento desc"
       End If
        
    Else
        StrSql = "Select " & CamposFiltro & " from Vendas_orcamento where " & DataFiltro & " order by Codigo desc, ID_orcamento desc"
    End If

Lista.ListItems.Clear

Set TBProduto = CreateObject("adodb.recordset")
'Debug.print StrSql

TBProduto.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBProduto.EOF = False Then
    Contador = 0
    Do While TBProduto.EOF = False
        With Lista.ListItems
            .Add , , TBProduto!ID_orcamento
            .Item(.Count).SubItems(1) = IIf(IsNull(TBProduto!CODIGO), "", TBProduto!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBProduto!txtcodigoproduto), "", TBProduto!txtcodigoproduto)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBProduto!txtunidade), "", TBProduto!txtunidade)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBProduto!txtdescricao), "", TBProduto!txtdescricao)
        End With
        TBProduto.MoveNext
        Contador = Contador + 1
    Loop
End If


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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: ProcSair
    Case vbKeyF2: ProcFiltrar
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

cmbfiltrarpor = "Número do orçamento"
msk_fltInicio.Value = Date
msk_fltFim.Value = Date

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

Private Sub Lista_DblClick()
On Error GoTo tratar_erro
  
If Lista.ListItems.Count = 0 Then Exit Sub

IDlista = Lista.SelectedItem
Unload Me


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub
