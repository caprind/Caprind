VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_Transporte_Tipo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Transportadora"
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   3270
   StartUpPosition =   1  'Centralizar no Mestre
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Transportadora tipo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   300
      TabIndex        =   1
      Top             =   600
      Width           =   2625
      Begin DrawSuite2022.USButton btnCliente 
         Height          =   585
         Left            =   390
         TabIndex        =   2
         ToolTipText     =   "Transporte feito pelo cliente"
         Top             =   390
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1032
         DibPicture      =   "frmVendas_Transporte_Tipo.frx":0000
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Cliente"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         PicAlign        =   7
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
      Begin DrawSuite2022.USButton btnEmpresa 
         Height          =   585
         Left            =   390
         TabIndex        =   3
         ToolTipText     =   "Transporte feito pela empresa"
         Top             =   1020
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1032
         DibPicture      =   "frmVendas_Transporte_Tipo.frx":8558
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Empresa"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         PicAlign        =   7
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
      Begin DrawSuite2022.USButton btnFornecedor 
         Height          =   585
         Left            =   390
         TabIndex        =   4
         ToolTipText     =   "Transporte feito por fornecedor"
         Top             =   1650
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1032
         DibPicture      =   "frmVendas_Transporte_Tipo.frx":103B0
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Fornecedor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         PicAlign        =   7
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   4
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   714
      DibPicture      =   "frmVendas_Transporte_Tipo.frx":15E15
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmVendas_Transporte_Tipo.frx":1B87A
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmVendas_Transporte_Tipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCliente_Click()
On Error GoTo tratar_erro

If Vendas_Proposta = True Then
If Transporte1 = True Then
 frmVendas_proposta.txtTipoTransp(0) = "Cliente"
 frmVendas_proposta.txtTransportadora.Text = frmVendas_proposta.txtcliente.Text
 frmVendas_proposta.txtidTransportadora.Text = frmVendas_proposta.txtIDCliente
 Else
 frmVendas_proposta.txtTipoTransp(1) = "Cliente"
 frmVendas_proposta.txtRedespacho.Text = frmVendas_proposta.txtcliente.Text
 End If
End If

If Vendas_PI = True Then
 If Transporte1 = True Then
 frmVendas_PI.txtTipoTransp(0).Text = "Cliente"
 frmVendas_PI.txtTransportadora.Text = frmVendas_PI.txtcliente.Text
 frmVendas_PI.txtidTransportadora.Text = frmVendas_PI.txtIDCliente
 Else
 frmVendas_PI.txtTipoTransp(1).Text = "Cliente"
 frmVendas_PI.txtRedespacho.Text = frmVendas_PI.txtcliente.Text
 End If
End If

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub btnEmpresa_Click()
On Error GoTo tratar_erro

If Vendas_Proposta = True Then
If Transporte1 = True Then
frmVendas_proposta.txtTipoTransp(0).Text = "Empresa"
frmVendas_proposta.txtTransportadora.Text = frmVendas_proposta.Cmb_empresa.Text
frmVendas_proposta.txtidTransportadora.Text = frmVendas_proposta.Cmb_empresa.ItemData(frmVendas_proposta.Cmb_empresa.ListIndex)
Else
frmVendas_proposta.txtTipoTransp(1).Text = "Empresa"
frmVendas_proposta.txtRedespacho.Text = frmVendas_proposta.Cmb_empresa.Text
End If
End If

If Vendas_PI = True Then
If Transporte1 = True Then
frmVendas_PI.txtTipoTransp(0).Text = "Empresa"
frmVendas_PI.txtTransportadora.Text = frmVendas_PI.Cmb_empresa.Text
frmVendas_PI.txtidTransportadora.Text = frmVendas_PI.Cmb_empresa.ItemData(frmVendas_PI.Cmb_empresa.ListIndex)
Else
frmVendas_PI.txtTipoTransp(1).Text = "Empresa"
frmVendas_PI.txtRedespacho.Text = frmVendas_PI.Cmb_empresa.Text
End If
End If

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnFornecedor_Click()
On Error GoTo tratar_erro

If Vendas_Proposta = True Then
If Transporte1 = True Then
frmVendas_proposta.txtTipoTransp(0).Text = "Fornecedor"
Else
frmVendas_proposta.txtTipoTransp(1).Text = "Fornecedor"
End If
End If

If Vendas_PI = True Then
If Transporte1 = True Then
frmVendas_PI.txtTipoTransp(0).Text = "Fornecedor"
Else
frmVendas_PI.txtTipoTransp(1).Text = "Fornecedor"
End If
End If


FrmVendas_LocalizarTransporte.Show 1

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
