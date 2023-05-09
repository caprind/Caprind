VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmProd_imp_filtro 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "PCP | Gerenciamento de ordem - Localizar"
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   9360
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
   Icon            =   "frmProd_imp_filtro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   820
      DibPicture      =   "frmProd_imp_filtro.frx":1042
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
      Icon            =   "frmProd_imp_filtro.frx":4692
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   28
      Top             =   5130
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   714
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
      Height          =   735
      Left            =   330
      TabIndex        =   26
      Top             =   2070
      Width           =   6225
      Begin VB.CheckBox chkEscopo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escopo"
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
         Left            =   180
         TabIndex        =   13
         Top             =   300
         Value           =   1  'Checked
         Width           =   825
      End
      Begin VB.CheckBox chkSemEscopo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fora do escopo"
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
         Left            =   1140
         TabIndex        =   14
         Top             =   300
         Value           =   1  'Checked
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Res. valid."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7860
      TabIndex        =   25
      Top             =   2070
      Width           =   1215
      Begin VB.ComboBox Cmb_validado_res 
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
         ItemData        =   "frmProd_imp_filtro.frx":56E4
         Left            =   180
         List            =   "frmProd_imp_filtro.frx":56F1
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Validado."
         Top             =   270
         Width           =   840
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Validado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   24
      Top             =   2070
      Width           =   1215
      Begin VB.ComboBox Cmb_validado 
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
         ItemData        =   "frmProd_imp_filtro.frx":5701
         Left            =   180
         List            =   "frmProd_imp_filtro.frx":570E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Validado."
         Top             =   270
         Width           =   840
      End
   End
   Begin VB.ComboBox Cmb_empresa 
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
      Height          =   330
      ItemData        =   "frmProd_imp_filtro.frx":571E
      Left            =   1380
      List            =   "frmProd_imp_filtro.frx":5720
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Empresa."
      Top             =   1710
      Width           =   7695
   End
   Begin VB.CheckBox ChkData2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dt. conclusão"
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
      Left            =   3630
      TabIndex        =   10
      Top             =   4590
      Width           =   1395
   End
   Begin VB.CheckBox ChkData1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Prazo entrega"
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
      Left            =   1935
      TabIndex        =   9
      Top             =   4590
      Width           =   1425
   End
   Begin VB.CheckBox ChkData 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dt. emissão"
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
      Left            =   450
      TabIndex        =   8
      Top             =   4590
      Width           =   1245
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   3660
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmProd_imp_filtro.frx":5722
      Count           =   1
   End
   Begin VB.Frame Frame4 
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
      Height          =   1515
      Left            =   270
      TabIndex        =   18
      Top             =   2820
      Width           =   8805
      Begin VB.Frame Frame5 
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
         Height          =   510
         Left            =   3810
         TabIndex        =   27
         Top             =   210
         Width           =   4785
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3930
            TabIndex        =   7
            Top             =   180
            Width           =   705
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1470
            TabIndex        =   5
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   4
            Top             =   180
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   6
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
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
         Height          =   330
         ItemData        =   "frmProd_imp_filtro.frx":7910
         Left            =   180
         List            =   "frmProd_imp_filtro.frx":793E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3525
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
         Left            =   180
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   8415
      End
      Begin VB.ComboBox cmbfamilia 
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
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Status."
         Top             =   1050
         Width           =   8415
      End
      Begin VB.Label Label3 
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
         Left            =   3645
         TabIndex        =   20
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
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
         Left            =   1522
         TabIndex        =   19
         Top             =   180
         Width           =   840
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   22
      Top             =   480
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1720
      ButtonCount     =   5
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Filtrar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Filtrar (F2)"
      ButtonKey1      =   "1"
      ButtonAlignment1=   2
      BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   42
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonAlignment2=   2
      ButtonType2     =   1
      ButtonStyle2    =   -1
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState2    =   -1
      ButtonLeft2     =   46
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   50
      ButtonTop3      =   2
      ButtonWidth3    =   41
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
      ButtonKey4      =   "4"
      ButtonAlignment4=   2
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   93
      ButtonTop4      =   2
      ButtonWidth4    =   30
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   5
      ButtonLeft5     =   125
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   270
      TabIndex        =   17
      Top             =   4350
      Width           =   8805
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   7200
         TabIndex        =   12
         ToolTipText     =   "Data final para pesquisa."
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   490012673
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   5430
         TabIndex        =   11
         ToolTipText     =   "Data início para pesquisa."
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   490012675
         CurrentDate     =   39057
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "à"
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
         Left            =   6945
         TabIndex        =   21
         Top             =   240
         Width           =   90
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa :"
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
      Left            =   480
      TabIndex        =   23
      Top             =   1800
      Width           =   825
   End
End
Attribute VB_Name = "frmProd_imp_filtro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pesquisa_ordem As String 'ok

Private Sub ChkData_Click()
On Error GoTo tratar_erro

If chkData.Value = 1 Then
    ChkData1.Value = 0
    ChkData2.Value = 0
    Frame3.Enabled = True
    msk_fltInicio.SetFocus
Else
    If ChkData1.Value = 0 And ChkData2.Value = 0 Then
        Frame3.Enabled = False
        msk_fltInicio.Value = Date
        msk_fltFim.Value = Date
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ChkData1_Click()
On Error GoTo tratar_erro

If ChkData1.Value = 1 Then
    chkData.Value = 0
    ChkData2.Value = 0
    Frame3.Enabled = True
    msk_fltInicio.SetFocus
Else
    If chkData.Value = 0 And ChkData2.Value = 0 Then
        Frame3.Enabled = False
        msk_fltInicio.Value = Date
        msk_fltFim.Value = Date
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ChkData2_Click()
On Error GoTo tratar_erro

If ChkData2.Value = 1 Then
    ChkData1.Value = 0
    chkData.Value = 0
    Frame3.Enabled = True
    msk_fltInicio.SetFocus
Else
    If chkData.Value = 0 And ChkData1.Value = 0 Then
        Frame3.Enabled = False
        msk_fltInicio.Value = Date
        msk_fltFim.Value = Date
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

If (cmbfiltrarpor = "Ordem" Or cmbfiltrarpor = "OS") And txtTexto <> "" Then
    VerifNumero = txtTexto
    ProcVerificaNumero
    If VerifNumero = False Then
        txtTexto = ""
        txtTexto.SetFocus
        Exit Sub
    End If
End If
With cmbfamilia
    If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Status" Or cmbfiltrarpor = "Tipo" Or cmbfiltrarpor = "Prioridade" Or cmbfiltrarpor = "Reposição" Then
        txtTexto.Visible = False
        .Visible = True
        .Clear
        .AddItem ""
        If cmbfiltrarpor = "Família" Then
            ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and Fabricacao = 'True'", True
        ElseIf cmbfiltrarpor = "Status" Then
                .AddItem "Aberta"
                .AddItem "Produzindo"
                .AddItem "Concluída"
                .AddItem "Cancelada"
                .AddItem "Aguardando"
                .AddItem "Entregue"
                .AddItem "Sem material"
            ElseIf cmbfiltrarpor = "Prioridade" Then
                    .AddItem "Urgente"
                    .AddItem "Normal"
                ElseIf cmbfiltrarpor = "Reposição" Then
                        .AddItem "Sim"
                        .AddItem "Não"
                    Else
                        .AddItem "Produto final"
                        .AddItem "Subconjunto"
                        .AddItem "Componente"
                        .AddItem "Serviço"
        End If
    Else
        txtTexto.Visible = True
        .Visible = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    'Case vbKeyF1: Ajuda
    Case vbKeyF2: ProcFiltrar
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 9450, 5, True
ProcCarregaComboEmpresa Cmb_empresa, False
cmbfiltrarpor = "Ordem"
txtTexto.Visible = True
msk_fltInicio.Value = Date
msk_fltFim.Value = Date

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
Acao = "filtrar"
'If chkEscopo.Value = 0 And chkSemEscopo.Value = 0 Then
'    NomeCampo = "uma das opções de escopo"
'    ProcVerificaAcao
'    Exit Sub
'End If

With frmprod
'    If chkEscopo.Value = 0 And chkSemEscopo.Value = 1 Then
'        FiltroEscopo = " and (P.Escopo = 'False' or P.Escopo IS NULL)"
'        FiltroEscopoRel = " and ({Producao.Escopo} = False or ISNULL({Producao.Escopo}) = True)"
'    ElseIf chkEscopo.Value = 1 And chkSemEscopo.Value = 0 Then
'            FiltroEscopo = " and P.Escopo = 'True'"
'            FiltroEscopoRel = " and {Producao.Escopo} = True"
'        Else
'            FiltroEscopo = ""
'            FiltroEscopoRel = ""
'    End If
    
    ValidFiltro = ""
    ValidFiltroRel = ""
    If Cmb_validado <> "" Then
        If Cmb_validado = "Sim" Then
            ValidFiltro = " and P.DtValidacao IS NOT NULL"
            ValidFiltroRel = " and NOT(ISNULL({Producao.DtValidacao}))"
        Else
            ValidFiltro = " and P.DtValidacao IS NULL"
            ValidFiltroRel = " and ISNULL({Producao.DtValidacao}) = True"
        End If
    End If
    ValidFiltroRes = ""
    ValidFiltroResRel = ""
    If Cmb_validado_res <> "" Then
        If Cmb_validado_res = "Sim" Then
            ValidFiltroRes = " and P.DtValidacao_custo IS NOT NULL"
            ValidFiltroResRel = " and NOT(ISNULL({Producao.DtValidacao_custo}))"
        Else
            ValidFiltroRes = " and P.DtValidacao_custo IS NULL"
            ValidFiltroResRel = " and ISNULL({Producao.DtValidacao_custo}) = True"
        End If
    End If
    
    DataFiltro = "P.desenho IS NOT NULL"
    DataFiltroRel = "{Producao.desenho} <> 'Null'"
    If chkData.Value = 1 Then
        DataFiltro = "(P.data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
        DataFiltroRel = "{Producao.data} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {Producao.data} <= Date(" & _
                                    Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
    End If
    If ChkData1.Value = 1 Then
        DataFiltro = "(P.PrazoEntrega) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
        DataFiltroRel = "{Producao.PrazoEntrega} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {Producao.PrazoEntrega} <= Date(" & _
                                    Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
    End If
    If ChkData2.Value = 1 Then
        DataFiltro = "(P.dataentrega) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
        DataFiltroRel = "{Producao.dataentrega} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {Producao.dataentrega} <= Date(" & _
                                    Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
    End If
    
    CamposFiltro = "P.Ordem, P.PrazoEntrega, P.Desenho, P.N_Referencia, P.Produto, P.dataentrega, P.DtValidacao, P.DtValidacao_custo, P.Impof, E.Empresa"
    INNERJOINTEXTO = "Select " & CamposFiltro & " from ((((((Producao P LEFT JOIN Ordemservico OS ON P.Ordem = OS.Ordem) LEFT JOIN projproduto PP ON PP.Desenho = P.Desenho) LEFT JOIN Producao_pedidos PPE ON PPE.Ordem = P.Ordem) LEFT JOIN Vendas_carteira VC ON VC.codigo = PPE.IDCarteira) LEFT JOIN Vendas_proposta VP ON VP.Cotacao = VC.Cotacao) LEFT JOIN Outros_SolicitacaoPCP OSP ON OSP.ID = VC.ID_solicitacao) LEFT JOIN Empresa E ON E.Codigo = P.ID_empresa"
    TextoFiltroPadrao = "P.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & DataFiltro & FiltroEscopo & ValidFiltro & ValidFiltroRes & " group by " & CamposFiltro & " order by P.Ordem desc"
    TextoFiltroPadraoRel = "{producao.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and " & DataFiltroRel & FiltroEscopoRel & ValidFiltroRes & ValidFiltroResRel
    
    If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
        If cmbfiltrarpor = "Ordem" Or cmbfiltrarpor = "OS" Then
            If cmbfiltrarpor = "Ordem" Then TextoFiltro = "P.Ordem" Else TextoFiltro = "OS.IDProducao"
            .StrSql_Ordem_Localizar = INNERJOINTEXTO & " where " & TextoFiltro & " = " & txtTexto & " and " & TextoFiltroPadrao
            .FormulaRel_Ordem = "{" & IIf(Left(TextoFiltro, 1) = "P", Replace(TextoFiltro, "P.", "Producao."), Replace(TextoFiltro, "OS.", "Ordemservico.")) & " } = " & txtTexto & " and " & TextoFiltroPadraoRel
        ElseIf cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Status" Then
                If cmbfiltrarpor = "Família" Then TextoFiltro = "PP.classe" Else TextoFiltro = "P.Status"
                .StrSql_Ordem_Localizar = INNERJOINTEXTO & " where " & TextoFiltro & " = '" & cmbfamilia & "' and " & TextoFiltroPadrao
                .FormulaRel_Ordem = "{" & IIf(Left(TextoFiltro, 2) = "PP", Replace(TextoFiltro, "PP.", "projproduto."), Replace(TextoFiltro, "P.", "Producao.")) & " } = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
            ElseIf cmbfiltrarpor = "Tipo" Then
                    Select Case cmbfamilia:
                        Case "Componente":
                            TextoFiltro = "P.Tipo = 'F'"
                            TextoFiltroRel = "Tipo} = 'F'"
                        Case "Subconjunto":
                            TextoFiltro = "P.Tipo = 'M'"
                            TextoFiltroRel = "Tipo} = 'M'"
                        Case "Produto final"
                            TextoFiltro = "P.Tipo = 'E'"
                            TextoFiltroRel = "Tipo} = 'E'"
                        Case "Serviço"
                            TextoFiltro = "P.Tipo = 'S'"
                            TextoFiltroRel = "Tipo} = 'S'"
                    End Select
                    .StrSql_Ordem_Localizar = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao
                    .FormulaRel_Ordem = "{Producao." & TextoFiltroRel & " and " & TextoFiltroPadraoRel
                ElseIf cmbfiltrarpor = "Prioridade" Then
                        If cmbfamilia = "Urgente" Then
                            TextoFiltro = "P.IMPREQ = 'True'"
                            TextoFiltroRel = "{Producao.IMPREQ} = True"
                        Else
                            TextoFiltro = "P.IMPREQ = 'False'"
                            TextoFiltroRel = "{Producao.IMPREQ} = False"
                        End If
                        .StrSql_Ordem_Localizar = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao
                        .FormulaRel_Ordem = TextoFiltroRel & " and " & TextoFiltroPadraoRel
                    ElseIf cmbfiltrarpor = "Reposição" Then
                            If cmbfamilia = "Sim" Then
                                TextoFiltro = "P.Reposicao = 'True'"
                                TextoFiltroRel = "{Producao.Reposicao} = True"
                            Else
                                TextoFiltro = "P.Reposicao = 'False'"
                                TextoFiltroRel = "{Producao.Reposicao} = False"
                            End If
                            .StrSql_Ordem_Localizar = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao
                            .FormulaRel_Ordem = TextoFiltroRel & " and " & TextoFiltroPadraoRel
                        Else
                            Select Case cmbfiltrarpor
                                Case "Código interno": TextoFiltro = "P.desenho"
                                Case "Código de referência": TextoFiltro = "P.N_Referencia"
                                Case "Descrição": TextoFiltro = "P.Produto"
                                Case "Cliente": TextoFiltro = "P.Cliente"
                                Case "Pedido interno": TextoFiltro = "VP.Ncotacao"
                                Case "Pedido do cliente": TextoFiltro = "VC.PCcliente"
                                Case "Solicitação de produção": TextoFiltro = "OSP.Requisicaotexto"
                            End Select
                            Select Case Left(TextoFiltro, 2)
                                Case "P.": TextoFiltroRel = Replace(TextoFiltro, "P.", "Producao.")
                                Case "VP": TextoFiltroRel = Replace(TextoFiltro, "VP.", "Vendas_proposta.")
                                Case "VC": TextoFiltroRel = Replace(TextoFiltro, "VC.", "vendas_carteira.")
                                Case "OS": TextoFiltroRel = Replace(TextoFiltro, "OS.", "Outros_SolicitacaoPCP.")
                            End Select
                            .StrSql_Ordem_Localizar = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
                            .FormulaRel_Ordem = "{" & TextoFiltroRel & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
        End If
    Else
        .StrSql_Ordem_Localizar = INNERJOINTEXTO & " where " & TextoFiltroPadrao
        .FormulaRel_Ordem = TextoFiltroPadraoRel
    End If
    .atualiza_lista_ordens (1)
End With
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

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If txtTexto <> "" Then
    If cmbfiltrarpor = "Ordem" Or cmbfiltrarpor = "OS" Then
        VerifNumero = txtTexto
        ProcVerificaNumero
        If VerifNumero = False Then
            txtTexto = ""
            txtTexto.SetFocus
            Exit Sub
        End If
    End If
    cmbfamilia.ListIndex = -1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


