VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_PI_lista 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "CAPRIND v5.0 | Vendas | Localizar documento"
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   8640
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
   ScaleHeight     =   3705
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   873
      DibPicture      =   "frmVendas_PI_lista.frx":0000
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
      Icon            =   "frmVendas_PI_lista.frx":3650
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empresa"
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
      Left            =   120
      TabIndex        =   27
      Top             =   630
      Width           =   4545
      Begin VB.ComboBox Cmb_empresa 
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
         ItemData        =   "frmVendas_PI_lista.frx":396A
         Left            =   180
         List            =   "frmVendas_PI_lista.frx":396C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   270
         Width           =   4185
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Status"
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
      Left            =   4680
      TabIndex        =   26
      Top             =   630
      Width           =   2595
      Begin VB.ComboBox cmbStatus 
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
         ItemData        =   "frmVendas_PI_lista.frx":396E
         Left            =   180
         List            =   "frmVendas_PI_lista.frx":3970
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Status."
         Top             =   270
         Width           =   2220
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
      Left            =   7290
      TabIndex        =   25
      Top             =   630
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
         ItemData        =   "frmVendas_PI_lista.frx":3972
         Left            =   180
         List            =   "frmVendas_PI_lista.frx":397F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Validado."
         Top             =   270
         Width           =   840
      End
   End
   Begin VB.CheckBox Chk_retorno 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Retorno"
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
      Left            =   2370
      TabIndex        =   13
      Top             =   3150
      Width           =   1005
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   4110
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmVendas_PI_lista.frx":398F
      Count           =   1
   End
   Begin VB.CheckBox Chk_emissao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Emissão"
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
      Left            =   330
      TabIndex        =   11
      Top             =   3150
      Width           =   915
   End
   Begin VB.CheckBox Chk_venda 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Venda"
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
      Left            =   1350
      TabIndex        =   12
      Top             =   3150
      Width           =   795
   End
   Begin VB.Frame Frame1 
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
      Height          =   1515
      Left            =   120
      TabIndex        =   18
      Top             =   1350
      Width           =   8385
      Begin VB.CheckBox chkFantasia 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nome fantasia"
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
         Left            =   8460
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frase"
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
         Left            =   3420
         TabIndex        =   28
         Top             =   210
         Width           =   3105
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
            Left            =   1710
            TabIndex        =   9
            Top             =   210
            Width           =   585
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
            TabIndex        =   7
            Top             =   210
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
            TabIndex        =   8
            Top             =   210
            Width           =   645
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
            Left            =   2340
            TabIndex        =   10
            Top             =   210
            Width           =   705
         End
      End
      Begin VB.ComboBox Cmb_alteracao 
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
         ItemData        =   "frmVendas_PI_lista.frx":5B7D
         Left            =   6750
         List            =   "frmVendas_PI_lista.frx":5B8A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Opções para filtro."
         Top             =   1050
         Width           =   1455
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmVendas_PI_lista.frx":5B9A
         Left            =   180
         List            =   "frmVendas_PI_lista.frx":5BC2
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   3195
      End
      Begin VB.TextBox txtTexto 
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
         Left            =   180
         TabIndex        =   2
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Width           =   6555
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
         ToolTipText     =   "Texto para pesquisa."
         Top             =   1050
         Visible         =   0   'False
         Width           =   6555
      End
      Begin MSMask.MaskEdBox txtCpf 
         Height          =   315
         Left            =   180
         TabIndex        =   4
         ToolTipText     =   "Número do CPF."
         Top             =   1050
         Visible         =   0   'False
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtcnpj 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         ToolTipText     =   "Número do CNPJ."
         Top             =   1050
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.###.###/####-##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Com alteração"
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
         Left            =   6855
         TabIndex        =   24
         Top             =   840
         Width           =   1035
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
         Left            =   2722
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1350
         TabIndex        =   19
         Top             =   180
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   675
      Left            =   120
      TabIndex        =   21
      Top             =   2850
      Width           =   8385
      Begin DrawSuite2022.USButton btnFiltrar 
         Height          =   405
         Left            =   7050
         TabIndex        =   30
         Top             =   180
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   714
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
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
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   5700
         TabIndex        =   15
         ToolTipText     =   "Data final."
         Top             =   240
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
         Format          =   137822209
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   3810
         TabIndex        =   14
         ToolTipText     =   "Data inicio."
         Top             =   240
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
         Format          =   137822209
         CurrentDate     =   39057
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
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
         Height          =   285
         Left            =   5295
         TabIndex        =   23
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "De :"
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
         Height          =   285
         Left            =   3450
         TabIndex        =   22
         Top             =   270
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmVendas_PI_lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrSql_PI_Localizar As String 'OK

Private Sub btnFiltrar_Click()
On Error GoTo tratar_erro

ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_emissao_Click()
On Error GoTo tratar_erro

If Chk_emissao.Value = 1 Then
    Chk_venda.Value = 0
    Chk_retorno.Value = 0
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_retorno_Click()
On Error GoTo tratar_erro

If Chk_retorno.Value = 1 Then
    Chk_emissao.Value = 0
    Chk_venda.Value = 0
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_venda_Click()
On Error GoTo tratar_erro

If Chk_venda.Value = 1 Then
    Chk_emissao.Value = 0
    Chk_retorno.Value = 0
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkFantasia_Click()
On Error GoTo tratar_erro

If chkFantasia.Value = 1 Then
UseNomeFantasia = True
Else
UseNomeFantasia = False
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

If cmbfiltrarpor = "Família" Then
    txtTexto.Visible = False
    cmbfamilia.Visible = True
    txtcnpj.Visible = False
    txtCpf.Visible = False
 ElseIf cmbfiltrarpor = "CNPJ" Then
        txtTexto.Visible = False
        cmbfamilia.Visible = False
        txtcnpj.Visible = True
        txtCpf.Visible = False
    ElseIf cmbfiltrarpor = "CPF" Then
            txtTexto.Visible = False
            cmbfamilia.Visible = False
            txtcnpj.Visible = False
            txtCpf.Visible = True
        Else
            txtTexto.Visible = True
            cmbfamilia.Visible = False
            txtcnpj.Visible = False
            txtCpf.Visible = False
End If

If cmbfiltrarpor = "Nome Fantasia" Then
chkFantasia.Value = 1
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

If Vendas_PI = True Then
    TipoPadrao = "(VP.Tipo = 'PE' or VP.Tipo = 'PRPE')"
    TipoPadraoRel = "({vendas_proposta.Tipo} = 'PE' or {vendas_proposta.Tipo} = 'PRPE')"
    CampoFiltroValid = "DtValidacaoPI"
Else
    TipoPadrao = "(VP.Tipo = 'PR' or VP.tipo = 'PRPE')"
    TipoPadraoRel = "({vendas_proposta.Tipo} = 'PR' or {vendas_proposta.tipo} = 'PRPE')"
    CampoFiltroValid = "DtValidacao"
End If
StatusFiltro = ""
StatusFiltroRel = ""
If cmbstatus <> "" Then
    StatusFiltro = " and VP.Status = '" & cmbstatus & "'"
    StatusFiltroRel = " and {vendas_proposta.Status} = '" & cmbstatus & "'"
End If
ValidFiltro = ""
ValidFiltroRel = ""
If Cmb_validado <> "" Then
    If Cmb_validado = "Sim" Then
        ValidFiltro = " and VP." & CampoFiltroValid & " IS NOT NULL"
        ValidFiltroRel = " and NOT(ISNULL({vendas_proposta." & CampoFiltroValid & "}))"
    Else
        ValidFiltro = " and VP." & CampoFiltroValid & " IS NULL"
        ValidFiltroRel = " and ISNULL({vendas_proposta." & CampoFiltroValid & ") = True"
    End If
End If
TextoFiltroAlt = ""
TextoFiltroAltRel = ""
If Cmb_alteracao = "Sim" Then
    If Vendas_PI = True Then
        TextoFiltroAlt1 = " and VCA.Tipo = 'VPI'"
        TextoFiltroAlt1Rel = " and {vendas_carteira_alteracoes.Tipo} = 'VPI'"
    Else
        TextoFiltroAlt1 = " and VCA.Tipo = 'VPR'"
        TextoFiltroAlt1Rel = " and {vendas_carteira_alteracoes.Tipo} = 'VPR'"
    End If
    TextoFiltroAlt = " and VCA.ID IS NOT NULL" & TextoFiltroAlt1
    TextoFiltroAltRel = " and Not(IsNull({vendas_carteira_alteracoes.ID}))" & TextoFiltroAlt1Rel
ElseIf Cmb_alteracao = "Não" Then
        TextoFiltroAlt = " and VCA.ID IS NULL"
        TextoFiltroAltRel = " and ISNULL({vendas_carteira_alteracoes.ID})"
End If
DataFiltro = ""
DataFiltroRel = ""
If Chk_emissao.Value = 1 Or Chk_venda.Value = 1 Or Chk_retorno.Value = 1 Then
    If Chk_emissao.Value = 1 Then
        Data_Solicitacao = "VP.Data"
    ElseIf Chk_venda.Value = 1 Then
            Data_Solicitacao = "VP.Datavendas"
        Else
            Data_Solicitacao = "VC.Data_retorno"
    End If
    DataFiltro = " and " & Data_Solicitacao & " Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
    DataFiltroRel = "and {" & IIf(Left(Data_Solicitacao, 2) = "VP", Replace(Data_Solicitacao, "VP", "vendas_proposta"), Replace(Data_Solicitacao, "VC", "vendas_carteira")) & "} >= Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {" & IIf(Left(Data_Solicitacao, 2) = "VP", Replace(Data_Solicitacao, "VP", "vendas_proposta"), Replace(Data_Solicitacao, "VC", "vendas_carteira")) & "} <= Date(" & _
                                Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ")"
End If

CamposFiltro = "VP.IDcliente,VP.vTotalFrete, VP.Cotacao, VP.Data, VP.Ncotacao, VP.Revisao, VP.Cliente, VP.Status, VP.DtValidacao, VP.DtValidacaoPI, VP.ordenarproposta"
INNERJOINTEXTO = "Select " & CamposFiltro & " from ((vendas_proposta VP LEFT JOIN Clientes C ON C.IDcliente = VP.IDcliente) LEFT JOIN vendas_carteira VC ON VP.cotacao = VC.cotacao) LEFT JOIN vendas_carteira_alteracoes VCA ON VCA.ID_carteira = VC.Codigo"
TextoFiltroPadrao = TipoPadrao & " and VP.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltro & TextoFiltroAlt & StatusFiltro & ValidFiltro & " group by " & CamposFiltro & " order by VP.ordenarproposta desc, VP.cotacao desc"
TextoFiltroPadraoRel = TipoPadraoRel & " and {vendas_proposta.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & DataFiltroRel & TextoFiltroAltRel & StatusFiltroRel & ValidFiltroRel

If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Or txtcnpj.Visible = True And txtcnpj <> "__.___.___/____-__" Or txtCpf.Visible = True And txtCpf <> "___.___.___-__" Then
    If cmbfiltrarpor = "Família" Then
        If cmbfiltrarpor = "Família" Then TextoFiltro = "VC.Familia" Else TextoFiltro = "VP.Status"
        FiltroTexto = INNERJOINTEXTO & " where VC.Familia = '" & cmbfamilia & "' and " & TextoFiltroPadrao
        FiltroTextoRel = "{vendas_carteira.Familia} = '" & cmbfamilia & "' and " & TextoFiltroPadraoRel
    ElseIf cmbfiltrarpor = "CNPJ" Or cmbfiltrarpor = "CPF" Then
            If cmbfiltrarpor = "CNPJ" Then
                TextoFiltro = "C.CPF_CNPJ = '" & txtcnpj & "'"
                TextoFiltroRel = "{Clientes.CPF_CNPJ} = '" & txtcnpj & "'"
            Else
                TextoFiltro = "C.CPF_CNPJ = '" & txtCpf & "'"
                TextoFiltroRel = "{Clientes.CPF_CNPJ} = '" & txtCpf & "'"
            End If
            FiltroTexto = INNERJOINTEXTO & " where " & TextoFiltro & " and " & TextoFiltroPadrao
            FiltroTextoRel = TextoFiltroRel & " and " & TextoFiltroPadraoRel
        Else
            Select Case cmbfiltrarpor
                Case "Proposta": TextoFiltro = "VP.Ncotacao"
                Case "Pedido": TextoFiltro = "VP.Ncotacao"
                Case "Referência": TextoFiltro = "VP.Referente"
                Case "Descrição referência": TextoFiltro = "VP.REF"
                Case "Cliente": TextoFiltro = "VP.Cliente"
                Case "Nome Fantasia": TextoFiltro = "C.nomefantasia"
                Case "Código de referência": TextoFiltro = "VC.n_referencia"
                Case "Pedido cliente": TextoFiltro = "VC.PCCliente"
                Case "Código interno": TextoFiltro = "VC.Desenho"
                Case "Descrição": TextoFiltro = "VC.Descricao_tecnica"
                Case "Descrição comercial": TextoFiltro = "VC.Descricao"
            End Select
            FiltroTexto = INNERJOINTEXTO & " where " & TextoFiltro & FunVerifTipoFiltroIMF(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadrao
            FiltroTextoRel = "{" & IIf(Left(TextoFiltro, 2) = "VP", Replace(TextoFiltro, "VP", "vendas_proposta"), Replace(TextoFiltro, "VC", "vendas_carteira")) & "}" & FunVerifTipoFiltroIMFRel(Optinicio, Optmeio, Optfim, optIgual, txtTexto) & " and " & TextoFiltroPadraoRel
    End If
Else
    FiltroTexto = INNERJOINTEXTO & " where " & TextoFiltroPadrao
    FiltroTextoRel = TextoFiltroPadraoRel
End If
If Vendas_PI = True Then
    With frmVendas_PI
        .StrSql_PI_Localizar = FiltroTexto
        .StrSql_PI_LocalizarRel = FiltroTextoRel
        .ProcCarregaLista (1)
    End With
Else
    With frmVendas_proposta
        .StrSql_Proposta_Localizar = FiltroTexto
        .StrSql_Proposta_LocalizarRel = FiltroTextoRel
        .ProcCarregaLista (1)
    End With
End If
Unload Me

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
UseNomeFantasia = False
ProcCarregaComboEmpresa Cmb_empresa, False
If Vendas_PI = True Then
    Caption = "Administrativo - Vendas - Pedido interno - Localizar"
    With cmbfiltrarpor
        .AddItem "Pedido"
        .Text = "Pedido"
    End With
Else
    Caption = "Administrativo - Vendas - Proposta comercial - Localizar"
    With cmbfiltrarpor
        .AddItem "Proposta"
        .Text = "Proposta"
    End With
End If
With cmbstatus
    .AddItem ""
    If Vendas_Proposta = True Then .AddItem "ABERTA EM ANALISE"
    .AddItem "CANCELADA"
    .AddItem "FATURADA"
    .AddItem "FATURADA PARCIAL"
    .AddItem "PERDIDA P/ PRAZO"
    .AddItem "PERDIDA P/ PREÇO"
    .AddItem "PORTAL ELETRONICO"
    .AddItem "REVISADA"
    .AddItem "VENDIDA"
    .AddItem "VENDIDA PARCIAL"
End With
ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and vendas = 'True'", True
msk_fltInicio = Date
msk_fltFim = Date

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If txtTexto <> "" Then cmbfamilia.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
