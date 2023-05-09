VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_PI_CST 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   2265
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.Frame Frame1 
      Caption         =   "Escolha a CST do ICMS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   150
      TabIndex        =   1
      Top             =   540
      Width           =   1905
      Begin DrawSuite2022.USButton btnOK 
         Height          =   285
         Left            =   1260
         TabIndex        =   4
         Top             =   480
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "OK"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
      Begin VB.ComboBox Cmb_CST_ICMS 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmVendas_PI_CST.frx":0000
         Left            =   390
         List            =   "frmVendas_PI_CST.frx":00BB
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Situação tributária ICMS."
         Top             =   480
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparente
         Caption         =   "CST ICMS"
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
         Index           =   93
         Left            =   420
         TabIndex        =   3
         Top             =   270
         Width           =   705
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   767
      DibPicture      =   "frmVendas_PI_CST.frx":020B
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmVendas_PI_CST.frx":738B
      ShowClose       =   0   'False
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmVendas_PI_CST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()
On Error GoTo tratar_erro

TBCotacao!txt_CST = Cmb_CST_ICMS.Text
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
Cmb_CST_ICMS.Clear
'========================
' Busca CST do ICMS
'========================
    Set TBAliquota = CreateObject("adodb.recordset")
    TBAliquota.Open "Select DISTINCT CST_ICMS FROM tbl_NaturezaOperacao_CST where ID_CFOP = " & TBCotacao!ID_CFOP, Conexao, adOpenKeyset, adLockOptimistic
    If TBAliquota.EOF = False Then
        Do While TBAliquota.EOF = False
            If IsNull(TBAliquota!CST_ICMS) = False And TBAliquota!CST_ICMS <> "" Then Cmb_CST_ICMS.AddItem TBAliquota!CST_ICMS
            TBAliquota.MoveNext
        Loop
    End If
    TBAliquota.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
