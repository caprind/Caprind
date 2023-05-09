VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmEstoque_CorrigirLote 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Estoque | Corrigir movimentação"
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   9570
   StartUpPosition =   1  'Centralizar no Mestre
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados da RE"
      Height          =   1845
      Left            =   150
      TabIndex        =   9
      Top             =   570
      Width           =   9195
      Begin VB.TextBox txtSaldoRE 
         Alignment       =   2  'Centralizar
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8220
         TabIndex        =   30
         Top             =   630
         Width           =   825
      End
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Centralizar
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
         Height          =   345
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   630
         Width           =   4125
      End
      Begin VB.TextBox txtun 
         Alignment       =   2  'Centralizar
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
         Height          =   345
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1290
         Width           =   405
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Centralizar
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
         Height          =   345
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1290
         Width           =   1695
      End
      Begin VB.TextBox txtDescricao 
         Alignment       =   2  'Centralizar
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
         Height          =   345
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1290
         Width           =   6645
      End
      Begin VB.TextBox txtRE 
         Alignment       =   2  'Centralizar
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
         Height          =   345
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   630
         Width           =   1065
      End
      Begin VB.TextBox txtLote 
         Alignment       =   2  'Centralizar
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
         Height          =   345
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   630
         Width           =   1065
      End
      Begin MSComCtl2.DTPicker DataRE 
         Height          =   345
         Left            =   2400
         TabIndex        =   14
         Top             =   630
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         _Version        =   393216
         Format          =   137035777
         CurrentDate     =   43825
      End
      Begin DrawSuite2022.USButton btnBuscaData 
         Height          =   315
         Left            =   3750
         TabIndex        =   32
         ToolTipText     =   "Buscar data do documento de entrada"
         Top             =   630
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         DibPicture      =   "FrmEstoque_Corrigir_lote.frx":0000
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   ""
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
         PicAlign        =   7
         PicSize         =   1
         Theme           =   4
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Centralizar
         BackStyle       =   0  'Transparente
         Caption         =   "Saldo"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   8370
         TabIndex        =   31
         Top             =   390
         Width           =   435
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Centralizar
         BackStyle       =   0  'Transparente
         Caption         =   "Un"
         Height          =   255
         Left            =   1890
         TabIndex        =   29
         Top             =   1050
         Width           =   585
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Centralizar
         BackStyle       =   0  'Transparente
         Caption         =   "Status"
         Height          =   255
         Left            =   5340
         TabIndex        =   22
         Top             =   390
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Centralizar
         BackStyle       =   0  'Transparente
         Caption         =   "Código"
         Height          =   255
         Left            =   330
         TabIndex        =   19
         Top             =   1050
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Centralizar
         BackStyle       =   0  'Transparente
         Caption         =   "Descrição"
         Height          =   255
         Left            =   4905
         TabIndex        =   18
         Top             =   1050
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Centralizar
         BackStyle       =   0  'Transparente
         Caption         =   "Data"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2550
         TabIndex        =   15
         Top             =   390
         Width           =   945
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Centralizar
         BackStyle       =   0  'Transparente
         Caption         =   "RE"
         Height          =   255
         Left            =   300
         TabIndex        =   13
         Top             =   390
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Centralizar
         BackStyle       =   0  'Transparente
         Caption         =   "Lote"
         Height          =   255
         Left            =   1380
         TabIndex        =   11
         Top             =   390
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dados da movimentação"
      Height          =   1155
      Left            =   150
      TabIndex        =   1
      Top             =   2430
      Width           =   9195
      Begin VB.TextBox txtOperacao 
         Alignment       =   2  'Centralizar
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
         Height          =   345
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   630
         Width           =   3285
      End
      Begin VB.TextBox txtDocumento 
         Alignment       =   2  'Centralizar
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
         Height          =   345
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   630
         Width           =   1005
      End
      Begin VB.TextBox txtSaida 
         Alignment       =   2  'Centralizar
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7500
         TabIndex        =   23
         Top             =   630
         Width           =   675
      End
      Begin DrawSuite2022.USButton btnCorrigir 
         Height          =   705
         Left            =   8280
         TabIndex        =   5
         Top             =   270
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1244
         DibPicture      =   "FrmEstoque_Corrigir_lote.frx":3650
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Corrigir"
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
         PicAlign        =   7
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         Theme           =   4
      End
      Begin VB.TextBox txtID 
         Alignment       =   2  'Centralizar
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
         Height          =   345
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   630
         Width           =   735
      End
      Begin VB.TextBox txtEntrada 
         Alignment       =   2  'Centralizar
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6720
         TabIndex        =   3
         Top             =   630
         Width           =   765
      End
      Begin MSComCtl2.DTPicker Data 
         Height          =   345
         Left            =   5310
         TabIndex        =   2
         Top             =   630
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         _Version        =   393216
         Format          =   136970241
         CurrentDate     =   43825
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Centralizar
         BackStyle       =   0  'Transparente
         Caption         =   "Operação"
         Height          =   255
         Left            =   3330
         TabIndex        =   28
         Top             =   390
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Centralizar
         BackStyle       =   0  'Transparente
         Caption         =   "Documento"
         Height          =   255
         Left            =   1020
         TabIndex        =   27
         Top             =   390
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Centralizar
         BackStyle       =   0  'Transparente
         Caption         =   "Entrada"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6810
         TabIndex        =   24
         Top             =   390
         Width           =   645
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Centralizar
         BackStyle       =   0  'Transparente
         Caption         =   "Saida"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7620
         TabIndex        =   8
         Top             =   390
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Centralizar
         BackStyle       =   0  'Transparente
         Caption         =   "Data"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5535
         TabIndex        =   7
         Top             =   390
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Centralizar
         BackStyle       =   0  'Transparente
         Caption         =   "ID"
         Height          =   255
         Left            =   330
         TabIndex        =   6
         Top             =   390
         Width           =   585
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   767
      DibPicture      =   "FrmEstoque_Corrigir_lote.frx":D0FD
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "FrmEstoque_Corrigir_lote.frx":E951
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmEstoque_CorrigirLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBuscaData_Click()
On Error GoTo tratar_erro

Select Case txtStatus.Text
Case "ENTRADA_ORDEM":

   Set TBAbrir = CreateObject("adodb.recordset")
   TBAbrir.Open "SELECT dataentrega FROM producao WHERE ordem = " & txtLote.Text, Conexao, adOpenKeyset, adLockOptimistic
   If TBAbrir.EOF = False Then
   DataRE.Value = TBAbrir!DataEntrega
   End If
   TBAbrir.Close

Case "ENTRADA_DEVOLUÇÃO":

Case "ENTRADA_NOTA_FISCAL":

Case "CONSIGNAÇÃO RECEBIDA":

End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnCorrigir_Click()
On Error GoTo tratar_erro

If RE <> 0 And IDlista <> 0 Then
If USMsgBox("Deseja realmente atualizar os dados da RE e da movimentação?", vbYesNo, "CAPRIND v5.0") = vbYes Then

Conexao.Execute "Update Estoque_Controle set Data = '" & DataRE.Value & "', Estoque_venda = " & Replace(txtSaldoRE.Text, ",", ".") & ", Estoque_Real = " & Replace(txtSaldoRE.Text, ",", ".") & ", Qtde = " & Replace(txtSaldoRE.Text, ",", ".") & ", Qtde_fisica = " & Replace(txtSaldoRE.Text, ",", ".") & " Where Idestoque = " & RE & ""
Conexao.Execute "Update Estoque_Movimentacao set Data = '" & DataRE.Value & "', Documento = '" & txtLote.Text & "' Where Idestoque = " & RE & " and Operacao = '" & txtStatus.Text & "'"
Conexao.Execute "Update Estoque_Movimentacao set Data = '" & Data.Value & "', Saida = " & Replace(txtSaida.Text, ",", ".") & ", Entrada = " & Replace(txtEntrada.Text, ",", ".") & " Where Idoperacao = " & IDlista & ""

USMsgBox "Dados atualizados com sucesso!", vbInformation, "CAPRIND v5.0"
frmestoque_Movimentacao.ProcFiltrar
frmestoque_Movimentacao.ProcCarregaGridMV

Unload Me
End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
