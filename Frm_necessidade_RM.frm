VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form Frm_necessidade_RM 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estoque - Necessidade - Alterar requisição de material da ordem"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12180
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   23
      Top             =   0
      Width           =   12045
      _ExtentX        =   21246
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
      ButtonCaption1  =   "Salvar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Salvar (F3)"
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
      ButtonWidth1    =   38
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
      ButtonLeft2     =   42
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   46
      ButtonTop3      =   2
      ButtonWidth3    =   36
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   84
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
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
      ButtonState5    =   5
      ButtonLeft5     =   112
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   10860
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "Frm_necessidade_RM.frx":0000
         Count           =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
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
      Top             =   1020
      Width           =   12045
      Begin VB.TextBox Txt_ID 
         Alignment       =   2  'Center
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
         Left            =   3513
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   375
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox txtUn_com 
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
         Left            =   8821
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Unidade comercial."
         Top             =   375
         Width           =   630
      End
      Begin VB.TextBox txtQtde_PC 
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
         Left            =   10665
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade de peças."
         Top             =   375
         Width           =   1185
      End
      Begin VB.TextBox txtOrdem 
         Alignment       =   2  'Center
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Ordem."
         Top             =   375
         Width           =   1490
      End
      Begin VB.TextBox txtUN 
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
         Left            =   8177
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Unidade de estoque."
         Top             =   375
         Width           =   630
      End
      Begin VB.TextBox txtDesenho 
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
         Left            =   1684
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   375
         Width           =   1815
      End
      Begin VB.TextBox txtDescricao 
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
         Left            =   3513
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   375
         Width           =   4650
      End
      Begin VB.TextBox txtQtde 
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
         Left            =   9465
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade."
         Top             =   375
         Width           =   1185
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Un. com."
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
         Left            =   8814
         TabIndex        =   26
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade PÇ"
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
         Left            =   10717
         TabIndex        =   25
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem"
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
         Left            =   180
         TabIndex        =   24
         Top             =   180
         Width           =   1490
      End
      Begin VB.Label Label34 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Un. est."
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
         Left            =   8200
         TabIndex        =   21
         Top             =   180
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
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
         Left            =   5493
         TabIndex        =   17
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno"
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
         Left            =   2137
         TabIndex        =   16
         Top             =   180
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
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
         Left            =   9637
         TabIndex        =   15
         Top             =   180
         Width           =   840
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   55
      TabIndex        =   18
      Top             =   1860
      Width           =   12045
      Begin VB.TextBox txtUN_com_Similar 
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
         Left            =   8805
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         ToolTipText     =   "Unidade comercial."
         Top             =   375
         Width           =   630
      End
      Begin VB.TextBox txtUn_Similar 
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
         Left            =   8160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         ToolTipText     =   "Unidade de estoque."
         Top             =   375
         Width           =   630
      End
      Begin VB.TextBox txtQtde_PC_Similar 
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
         Left            =   10650
         MaxLength       =   50
         TabIndex        =   13
         ToolTipText     =   "Quantidade de peças."
         Top             =   375
         Width           =   1185
      End
      Begin VB.TextBox txtQtde_Similar 
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
         Left            =   9450
         MaxLength       =   50
         TabIndex        =   12
         ToolTipText     =   "Quantidade."
         Top             =   375
         Width           =   1185
      End
      Begin VB.TextBox txtDesc_Similar 
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
         Left            =   2700
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   375
         Width           =   5445
      End
      Begin VB.TextBox txtDesenho_Similar 
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
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   375
         Width           =   2085
      End
      Begin VB.CommandButton cmdDesenho 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2280
         Picture         =   "Frm_necessidade_RM.frx":1E03
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Localizar produtos."
         Top             =   375
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno"
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
         Left            =   772
         TabIndex        =   29
         Top             =   180
         Width           =   900
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade PÇ"
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
         Left            =   10702
         TabIndex        =   28
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Un. com."
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
         Left            =   8798
         TabIndex        =   27
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Un. est."
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
         Left            =   8183
         TabIndex        =   22
         Top             =   180
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
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
         Left            =   9622
         TabIndex        =   20
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
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
         Left            =   5077
         TabIndex        =   19
         Top             =   180
         Width           =   690
      End
   End
End
Attribute VB_Name = "Frm_necessidade_RM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDesenho_Click()
On Error GoTo tratar_erro

PCP_AlterarRM = False
Frm_necessidade_Item.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvar()
On Error GoTo tratar_erro

Acao = "salvar"
If txtDesenho_Similar = "" Then
    NomeCampo = "o código interno"
    ProcVerificaAcao
    cmdDesenho_Click
    Exit Sub
End If
If txtQtde_Similar.Locked = False Then
    Total = IIf(txtQtde_Similar = "", 0, txtQtde_Similar)
    If Total <= 0 Then
        NomeCampo = "a quantidade"
        ProcVerificaAcao
        txtQtde_Similar.SetFocus
        Exit Sub
    End If
Else
    Total = IIf(txtQtde_PC_Similar = "", 0, txtQtde_PC_Similar)
    If Total <= 0 Then
        NomeCampo = "a quantidade de peças"
        ProcVerificaAcao
        txtQtde_PC_Similar.SetFocus
        Exit Sub
    End If
End If
With Frm_necessidade
    If FunAlterarProdSimiliarOrdem(.Cmb_empresa.ItemData(.Cmb_empresa.ListIndex), txtDesenho_Similar, IIf(.Opt_vendas.Value = False, txtOrdem, Txt_ID), txtdesenho, IIf(txtQtde_Similar = "", 0, txtQtde_Similar), IIf(txtQtde_PC_Similar = "", 0, txtQtde_PC_Similar), .Opt_vendas.Value) = True Then
        If .Opt_vendas.Value = True Then
            MsgTexto = " do pedido"
            MsgTexto1 = "Ped. interno: "
        Else
            MsgTexto = " da ordem"
            MsgTexto1 = "Ordem: "
        End If
        USMsgBox ("Requisição de material " & MsgTexto & " alterada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = Formulario
        Evento = "Alterar requisição de material " & MsgTexto
        ID_documento = Txt_ID
        Documento = MsgTexto1 & txtOrdem & " - Cód. interno de: " & txtdesenho
        Documento1 = MsgTexto1 & txtOrdem & " - Cód. interno para: " & txtDesenho_Similar
        ProcGravaEvento
        '==================================
        Unload Me
        
        .ProcLimpaCampos
        .ProcCarregaLista
    Else
        USMsgBox ("Não foi alterado a requisição do material " & MsgTexto & ", pois não está ativo o recurso para produtos similares no cadastro da empresa."), vbExclamation, "CAPRIND v5.0"
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
    Case vbKeyF3: ProcSalvar
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 12045, 5, True

If Compras_Necessidade = True Then
    Caption = "Compras - Necessidade - Alterar requisição de material da ordem"
    Formulario = "Compras/Necessidade"
ElseIf PCP_Necessidade = True Then
        Caption = "PCP - Necessidade - Alterar requisição de material da ordem"
        Formulario = "PCP/Necessidade"
    Else
        Caption = "Estoque - Necessidade - Alterar requisição de material da ordem"
        Formulario = "Estoque/Necessidade"
End If
With Frm_necessidade
    If .Opt_vendas.Value = True Then
        Label11.Caption = "Ped. int."
        txtOrdem.ToolTipText = "Pedido interno"
    End If
    If .SSTab1.Tab = 0 Then
        Txt_ID = .Lista_necessidade.SelectedItem
        If .Opt_vendas.Value = True Then txtOrdem = .Lista_necessidade.SelectedItem.SubItems(4) Else txtOrdem = .Lista_necessidade.SelectedItem.SubItems(5)
        txtdesenho = .Lista.SelectedItem.SubItems(1)
        txtQtde = .Lista_necessidade.SelectedItem.SubItems(1)
        txtQtde_PC = .Lista_necessidade.SelectedItem.SubItems(3)
    Else
        Txt_ID = .Lista_detalhado.SelectedItem
        If .Opt_vendas.Value = True Then txtOrdem = .Lista_detalhado.SelectedItem.SubItems(6) Else txtOrdem = .Lista_detalhado.SelectedItem.SubItems(7)
        txtdesenho = .Lista_detalhado.SelectedItem.SubItems(1)
        txtQtde = .Lista_detalhado.SelectedItem.SubItems(3)
        txtQtde_PC = .Lista_detalhado.SelectedItem.SubItems(5)
    End If
End With

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "select Unidade, unidade_com, descricao from projproduto where desenho = '" & txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    txtUN = IIf(IsNull(TBItem!Unidade), "", TBItem!Unidade)
    txtUn_com = IIf(IsNull(TBItem!Unidade_com), "", TBItem!Unidade_com)
    txtdescricao = IIf(IsNull(TBItem!Descricao), "", TBItem!Descricao)
End If
TBItem.Close

If FunVerifMovimentacaoEstPC(Frm_necessidade.Cmb_empresa.ItemData(Frm_necessidade.Cmb_empresa.ListIndex)) = True And txtQtde_PC > 0 Then
    With txtQtde_Similar
        .Locked = True
        .TabStop = False
    End With
    With txtQtde_PC_Similar
        .Locked = False
        .TabStop = True
    End With
Else
    With txtQtde_Similar
        .Locked = False
        .TabStop = True
    End With
    With txtQtde_PC_Similar
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_PC_Similar_Change()
On Error GoTo tratar_erro

If txtQtde_PC_Similar.Text <> "" Then
    VerifNumero = txtQtde_PC_Similar.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde_PC_Similar.Text = ""
        txtQtde_PC_Similar.SetFocus
        Exit Sub
    End If
    If txtQtde_PC_Similar.Locked = False Then ProcCalculaQtdesSimilar txtDesenho_Similar, txtQtde_Similar, txtQtde_PC_Similar
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Similar_Change()
On Error GoTo tratar_erro

If txtQtde_Similar.Text <> "" Then
    VerifNumero = txtQtde_Similar.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtde_Similar.Text = ""
        txtQtde_Similar.SetFocus
        Exit Sub
    End If
    If txtQtde_Similar.Locked = False Then ProcCalculaQtdesSimilar txtDesenho_Similar, txtQtde_Similar, txtQtde_PC_Similar
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtde_Similar_LostFocus()
On Error GoTo tratar_erro

txtQtde_Similar = Format(txtQtde_Similar, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtDesenho_Similar = ""
txtDesc_Similar = ""
txtUn_Similar = ""
txtUN_com_Similar = ""
txtQtde_Similar = ""
txtQtde_PC_Similar = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
