VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmImpostosLPLR 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "CAPRIND v5.0 | Impostos"
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImpostosLPLR.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   4965
   StartUpPosition =   1  'Centralizar no Mestre
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aliquotas a serem aplicadas"
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
      Height          =   2205
      Left            =   2340
      TabIndex        =   1
      Top             =   570
      Width           =   2415
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   0
         Left            =   750
         Top             =   390
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "IRPJ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "IRPJ"
      End
      Begin VB.TextBox txtCSLL 
         Alignment       =   2  'Centralizar
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'Nenhum
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
         Height          =   280
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   585
      End
      Begin VB.TextBox txtCOFINS 
         Alignment       =   2  'Centralizar
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'Nenhum
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
         Height          =   280
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1065
         Width           =   585
      End
      Begin VB.TextBox txtPis 
         Alignment       =   2  'Centralizar
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'Nenhum
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
         Height          =   280
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1425
         Width           =   585
      End
      Begin VB.TextBox txtIPI 
         Alignment       =   2  'Centralizar
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'Nenhum
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
         Height          =   280
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   3420
         Width           =   585
      End
      Begin VB.TextBox txtICMS 
         Alignment       =   2  'Centralizar
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'Nenhum
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
         Height          =   280
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1770
         Width           =   585
      End
      Begin VB.TextBox txtIRPJ 
         Alignment       =   2  'Centralizar
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'Nenhum
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
         Height          =   280
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   585
      End
      Begin DrawSuite2022.USLabel USLabel2 
         Height          =   195
         Left            =   750
         Top             =   735
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "CSLL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "CSLL"
      End
      Begin DrawSuite2022.USLabel USLabel3 
         Height          =   195
         Left            =   750
         Top             =   1095
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "Cofins"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "Cofins"
      End
      Begin DrawSuite2022.USLabel USLabel4 
         Height          =   195
         Left            =   750
         Top             =   1440
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "PIS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "PIS"
      End
      Begin DrawSuite2022.USLabel USLabel6 
         Height          =   195
         Left            =   750
         Top             =   3450
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "IPI"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "IPI"
      End
      Begin DrawSuite2022.USLabel USLabel7 
         Height          =   195
         Left            =   750
         Top             =   1770
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "ICMS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         NoHTMLCaption   =   "ICMS"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   1
         Left            =   2010
         Top             =   420
         Width           =   165
         _ExtentX        =   291
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         NoHTMLCaption   =   "%"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   2
         Left            =   2010
         Top             =   775
         Width           =   165
         _ExtentX        =   291
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         NoHTMLCaption   =   "%"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   3
         Left            =   2010
         Top             =   1130
         Width           =   165
         _ExtentX        =   291
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         NoHTMLCaption   =   "%"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   4
         Left            =   2010
         Top             =   1485
         Width           =   165
         _ExtentX        =   291
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         NoHTMLCaption   =   "%"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   5
         Left            =   2010
         Top             =   3495
         Width           =   165
         _ExtentX        =   291
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         NoHTMLCaption   =   "%"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   6
         Left            =   2010
         Top             =   1830
         Width           =   165
         _ExtentX        =   291
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         NoHTMLCaption   =   "%"
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   2340
      TabIndex        =   11
      Top             =   2700
      Width           =   2415
      Begin VB.TextBox txtpTotal 
         Alignment       =   2  'Centralizar
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'Nenhum
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
         Height          =   280
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   270
         Width           =   585
      End
      Begin DrawSuite2022.USLabel USLabel5 
         Height          =   195
         Left            =   750
         Top             =   300
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "Total"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         NoHTMLCaption   =   "Total"
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   195
         Index           =   7
         Left            =   2010
         Top             =   330
         Width           =   165
         _ExtentX        =   291
         _ExtentY        =   344
         Autosize        =   0   'False
         Caption         =   "%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         NoHTMLCaption   =   "%"
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   10
      Top             =   3600
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   714
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Regime tributário"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   150
      TabIndex        =   8
      Top             =   570
      Width           =   2145
      Begin VB.TextBox txtRegime 
         Alignment       =   2  'Centralizar
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   270
         Width           =   1905
      End
      Begin DrawSuite2022.USAlphaImage USAlphaImage1 
         Height          =   1500
         Left            =   180
         Top             =   690
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   2646
         Image           =   "frmImpostosLPLR.frx":000C
         Props           =   5
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   714
      DibPicture      =   "frmImpostosLPLR.frx":60A2
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
      Icon            =   "frmImpostosLPLR.frx":101C5
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmImpostosLPLR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo tratar_erro
Dim IRPJ As Double
Dim IPI As Double
Dim PIS As Double
Dim Cofins As Double
Dim ICMS As Double
Dim CSLL As Double

'===========================================================
' Se não for regime simples nacional
'===========================================================
If frm_orcamento.txtidregime <> 1 Then
txtRegime = frm_orcamento.txtRegime.Text
'===========================================================
' Busca impostos inerentes a empresa IRPJ, CSLL, PIS, Cofins
'===========================================================
 Set TBAbrir = CreateObject("adodb.recordset")
  TBAbrir.Open "Select * FROM Impostos where ID_empresa = " & IDempresa, Conexao, adOpenKeyset, adLockOptimistic
  txtIRPJ = Format(TBAbrir!IRPJ_produtos, "###,##0.00")
  txtCSLL = Format(TBAbrir!CSLL_produtos, "###,##0.00")
  txtCOFINS = Format(TBAbrir!Cofins_produtos, "###,##0.00")
  txtPis = Format(TBAbrir!PIS_produtos, "###,##0.00")
 TBAbrir.Close
 
'===========================================================
' Busca impostos inerentes a NCM ICMS, e IPI
'===========================================================
 Set TBAbrir = CreateObject("adodb.recordset")
  If frm_orcamento.txtNCM.Text <> "" Then
   TBAbrir.Open "Select * FROM tbl_ClassificacaoFiscal where IDIntClasse = '" & frm_orcamento.txtNCM.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
'    txtIPI = Format(TBAbrir!dbl_IPI, "###,##0.00")
    txtICMS = Format(TBAbrir!dbl_ICMS_de, "###,##0.00")
    IRPJ = txtIRPJ.Text
'    IPI = txtIPI.Text
    PIS = txtPis.Text
    Cofins = txtCOFINS.Text
    ICMS = txtICMS.Text
    CSLL = txtCSLL.Text
    Total = IRPJ + ICMS + IPI + PIS + Cofins + CSLL
    Total = IRPJ + ICMS + PIS + Cofins + CSLL
    frm_orcamento.txtp7 = Format(Total, "###,##0.00")
    txtpTotal.Text = Format(Total, "###,##0.00")
  TBAbrir.Close
  End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
