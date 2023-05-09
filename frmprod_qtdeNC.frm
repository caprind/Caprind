VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmprod_qtdeNC 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informar quantidade por descrição da NC"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4875
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2385
      Left            =   2970
      TabIndex        =   19
      Top             =   -90
      Width           =   1845
      Begin VB.CommandButton Cmd_backspace 
         BackColor       =   &H00FFFFFF&
         Caption         =   "<-------"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   720
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmprod_qtdeNC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1800
         Width           =   990
      End
      Begin VB.CommandButton Number 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   0
         Left            =   180
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmprod_qtdeNC.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1800
         Width           =   480
      End
      Begin VB.CommandButton Number 
         BackColor       =   &H00FFFFFF&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   3
         Left            =   1230
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmprod_qtdeNC.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1260
         Width           =   480
      End
      Begin VB.CommandButton Number 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   2
         Left            =   720
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmprod_qtdeNC.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1260
         Width           =   480
      End
      Begin VB.CommandButton Number 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   180
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmprod_qtdeNC.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1260
         Width           =   480
      End
      Begin VB.CommandButton Number 
         BackColor       =   &H00FFFFFF&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   6
         Left            =   1230
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmprod_qtdeNC.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   750
         Width           =   480
      End
      Begin VB.CommandButton Number 
         BackColor       =   &H00FFFFFF&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   5
         Left            =   720
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmprod_qtdeNC.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   750
         Width           =   480
      End
      Begin VB.CommandButton Number 
         BackColor       =   &H00FFFFFF&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   4
         Left            =   180
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmprod_qtdeNC.frx":1546
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   750
         Width           =   480
      End
      Begin VB.CommandButton Number 
         BackColor       =   &H00FFFFFF&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   9
         Left            =   1230
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmprod_qtdeNC.frx":1850
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   480
      End
      Begin VB.CommandButton Number 
         BackColor       =   &H00FFFFFF&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   8
         Left            =   720
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmprod_qtdeNC.frx":1B5A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   480
      End
      Begin VB.CommandButton Number 
         BackColor       =   &H00FFFFFF&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   7
         Left            =   180
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmprod_qtdeNC.frx":1E64
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Número da OS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   30
      TabIndex        =   18
      Top             =   600
      Width           =   2865
      Begin VB.TextBox Txt_OS 
         Alignment       =   2  'Centralizar
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   460
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Número da OS."
         Top             =   240
         Width           =   2505
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Conforme"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   30
      TabIndex        =   17
      Top             =   1440
      Width           =   1432
      Begin VB.TextBox txtTOK 
         Alignment       =   2  'Centralizar
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   460
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade conforme."
         Top             =   270
         Width           =   1050
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Não conf."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1463
      TabIndex        =   16
      Top             =   1440
      Width           =   1432
      Begin VB.TextBox txtTNC 
         Alignment       =   2  'Centralizar
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   460
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade não conforme."
         Top             =   270
         Width           =   1050
      End
   End
   Begin DrawSuite2022.USButton Cmd_F3 
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Top             =   90
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   714
      BorderColor     =   14404026
      BorderColorDown =   11632444
      BorderColorOver =   11632444
      ButtonShape     =   1
      Caption         =   "F3 - GRAVAR"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColor2  =   16777215
      GradientColor3  =   16777215
      GradientColorDown2=   16246986
      GradientColorDown3=   15189380
      GradientColorDown4=   14596208
      GradientColorOver1=   16643560
      GradientColorOver2=   16576988
      GradientColorOver3=   16441780
      GradientColorOver4=   16178091
      PicAlign        =   8
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
   End
   Begin DrawSuite2022.USButton Cmd_esc 
      Height          =   405
      Left            =   1500
      TabIndex        =   4
      Top             =   90
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   714
      BorderColor     =   14404026
      BorderColorDown =   11632444
      BorderColorOver =   11632444
      ButtonShape     =   1
      Caption         =   "ESC - VOLTAR"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColor2  =   16777215
      GradientColor3  =   16777215
      GradientColorDown2=   16246986
      GradientColorDown3=   15189380
      GradientColorDown4=   14596208
      GradientColorOver1=   16643560
      GradientColorOver2=   16576988
      GradientColorOver3=   16441780
      GradientColorOver4=   16178091
      PicAlign        =   8
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5025
      Left            =   30
      TabIndex        =   20
      Top             =   2280
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   8864
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   14737632
      BackColorBkg    =   14737632
      GridColor       =   14737632
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaco
      BorderWidth     =   2
      Height          =   555
      Left            =   30
      Top             =   30
      Width           =   2865
   End
End
Attribute VB_Name = "frmprod_qtdeNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Linha As Integer 'OK
Public QtdeDescricaoNC As Integer 'OK

Private Sub Cmd_backspace_Click()
On Error GoTo tratar_erro

If Grid.TextMatrix(Grid.RowSel, 1) = "" Then Exit Sub
Grid.TextMatrix(Grid.RowSel, 1) = Left(Grid.TextMatrix(Grid.RowSel, 1), Len(Grid.TextMatrix(Grid.RowSel, 1)) - 1)
ProcCalculaTotalNC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_esc_Click()
On Error GoTo tratar_erro

Gravar = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmd_F3_Click()
On Error GoTo tratar_erro

ProcGravar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCalculaTotalNC()
On Error GoTo tratar_erro

Contador = 0
valor = 0
With Grid
    For InitFor = 1 To (.rows)
        If IsNumeric(Grid.TextMatrix(Contador, 1)) = True Then
            If Grid.TextMatrix(Contador, 1) <> "" Then valor = valor + Grid.TextMatrix(Contador, 1)
        End If
        Contador = Contador + 1
    Next InitFor
End With
txtTNC = valor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcGravar
    Case vbKeyEscape: Cmd_esc_Click
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcGravar()
On Error GoTo tratar_erro

If valor <> TNC Then
    USMsgBox ("A quantidade não conforme está diferente da informada no apontamento, favor verificar"), vbExclamation, "CAPRIND v5.0"
    Exit Sub
End If
ReDim ArrayQtdeDescNC(1 To QtdeDescricaoNC, 1 To 2)
With Grid
    For i = 1 To (.rows - 1)
        ArrayQtdeDescNC(i, 1) = .TextMatrix(i, 0)
        ArrayQtdeDescNC(i, 2) = .TextMatrix(i, 1)
    Next
End With
Gravar = True
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

With frmprod
    Txt_OS = .cmbAPOS
    txtTOK = TOK
End With
ProcCarregaListaDNC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaDNC()
On Error GoTo tratar_erro
Dim l As Long

With Grid
    .Row = 0
    .Col = 0
    
    .TextMatrix(0, 0) = "Descrição"
    .TextMatrix(0, 1) = "Qtde."
    .ColWidth(0) = 3500
    .ColWidth(1) = 900
    
    .CellFontBold = True
    
        
    .Col = 1
    .CellFontBold = True
    .ColAlignment(1) = 3
    
    QtdeDescricaoNC = 1
    l = 1
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select ID, Causa from CQ_NC_FABRICA_causa where DtValidacao IS NOT NULL order by Causa", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        Contador = 1
        QtdeDescricaoNC = TBLISTA.RecordCount
        .rows = TBLISTA.RecordCount + 1

        Do While TBLISTA.EOF = False
            .TextMatrix(l, 0) = IIf(IsNull(TBLISTA!Causa), "", TBLISTA!Causa)
            
            .Row = l
            .Col = 0
            .CellFontBold = True
            
            l = l + 1
            TBLISTA.MoveNext
        Loop
    End If
    TBLISTA.Close
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
On Error GoTo tratar_erro

If KeyAscii = 8 Then
    Cmd_backspace_Click
ElseIf KeyAscii = 13 Then
        ProcAvancaCelula
    Else
        Select Case KeyAscii
            Case 48: Number_Click (0)
            Case 49: Number_Click (1)
            Case 50: Number_Click (2)
            Case 51: Number_Click (3)
            Case 52: Number_Click (4)
            Case 53: Number_Click (5)
            Case 54: Number_Click (6)
            Case 55: Number_Click (7)
            Case 56: Number_Click (8)
            Case 57: Number_Click (9)
            Case 96: Number_Click (0)
            Case 97: Number_Click (1)
            Case 98: Number_Click (2)
            Case 99: Number_Click (3)
            Case 100: Number_Click (4)
            Case 101: Number_Click (5)
            Case 102: Number_Click (6)
            Case 103: Number_Click (7)
            Case 104: Number_Click (8)
            Case 105: Number_Click (9)
        End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcAvancaCelula()
 On Error GoTo tratar_erro
    
With Grid
    .HighLight = flexHighlightNever
    If .Col < .Cols - 1 Then
        .Col = .Col + 1
    Else
        If .Row < .rows - 1 Then
            .Row = .Row + 1  'Desce uma linha
            .Col = 1
        Else
            .Row = 1
            .Col = 1
        End If
    End If
    If .CellTop + .CellHeight > .Top + .Height Then
        .TopRow = .TopRow + 1
    End If
    .HighLight = flexHighlightAlways
End With
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
   
End Sub

Private Sub Number_Click(index As Integer)
On Error GoTo tratar_erro

Grid.TextMatrix(Grid.RowSel, 1) = Grid.TextMatrix(Grid.RowSel, 1) & Number(index).Caption
ProcCalculaTotalNC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
