VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCompras_pedido_liberar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Compras Pedido | Liberar alteração"
   ClientHeight    =   7155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   5055
   StartUpPosition =   1  'CenterOwner
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   6750
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   714
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   5715
      Left            =   540
      TabIndex        =   1
      Top             =   630
      Width           =   4005
      Begin VB.TextBox txtJustificativa 
         Height          =   2415
         Left            =   420
         MaxLength       =   240
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtSenha 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   450
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   3660
         Width           =   3075
      End
      Begin DrawSuite2022.USButton btnLiberar 
         Height          =   1125
         Left            =   390
         TabIndex        =   2
         Top             =   4200
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   1984
         DibPicture      =   "frmCompras_pedido_liberar.frx":0000
         Caption         =   "Liberar alteração"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         PicAlign        =   7
         PicSize         =   4
         PicSizeH        =   48
         PicSizeW        =   48
         ShowFocusRect   =   0   'False
         Theme           =   3
      End
      Begin DrawSuite2022.USLabel USLabel1 
         Height          =   240
         Left            =   630
         TabIndex        =   7
         Top             =   3420
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   423
         Caption         =   "Informe a senha para liberação"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   64
         NoHTMLCaption   =   "Informe a senha para liberação"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Informe o motivo da alteração."
         Height          =   240
         Left            =   630
         TabIndex        =   6
         Top             =   450
         Width           =   2655
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   714
      DibPicture      =   "frmCompras_pedido_liberar.frx":7E58
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
      Icon            =   "frmCompras_pedido_liberar.frx":FCB0
   End
End
Attribute VB_Name = "frmCompras_pedido_liberar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnLiberar_Click()
On Error GoTo tratar_erro
Liberado = False
LiberarData = False
If txtJustificativa = "" Then Exit Sub

If txtSenha.Text = "LIBdata2020" And txtJustificativa.Text <> "" Then
    If USMsgBox("Deseja realmente liberar a modificação do prazo do item no pedido de compras?", vbYesNo, "CAPRIND v5.01") = vbYes Then
        LiberarData = True
        StrSql = "Insert into Compras_Pedido_Lista_Alteracoes (IDLista,Responsavel,Data,Justificativa) VALUES ('" & IDlista & "','" & pubUsuario & "','" & Date & "','" & txtJustificativa & "')"
        'Debug.print StrSql
        
        Conexao.Execute StrSql
        USMsgBox "Modificação do prazo do item liberada com sucesso", vbInformation, "CAPRIND v5.0"
        Unload Me
    Else
        LiberarData = False
        Unload Me
    End If
End If

If txtSenha.Text = "LIBdata2021" And txtJustificativa.Text <> "" Then
    If USMsgBox("Deseja realmente liberar a modificação dos dados do item no pedido de compras?", vbYesNo, "CAPRIND v5.01") = vbYes Then
        Liberado = True
        frmCompras_Pedido.ProcDesbloqueiaCamposItem
        frmCompras_Pedido.ProcDesbloqueiaCamposServ
        StrSql = "Insert into Compras_Pedido_Lista_Alteracoes (IDLista,Responsavel,Data,Justificativa) VALUES ('" & IDlista & "','" & pubUsuario & "','" & Date & "','" & txtJustificativa & "')"
        Conexao.Execute StrSql
        USMsgBox "Modificação dos dados do item liberada com sucesso", vbInformation, "CAPRIND v5.0"
        Unload Me
    Else
        Liberado = False
        Unload Me
    End If
End If



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
