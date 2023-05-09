VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_Cliente_RelUF 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Vendas | Clientes | Menu impressão"
   ClientHeight    =   3510
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5550
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
   Icon            =   "frmVendas_Cliente_RElUF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USButton btnDetalhado 
      Height          =   855
      Left            =   420
      TabIndex        =   2
      Top             =   900
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   1508
      DibPicture      =   "frmVendas_Cliente_RElUF.frx":030A
      Caption         =   "Lista detalhada"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      Theme           =   3
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   820
      DibPicture      =   "frmVendas_Cliente_RElUF.frx":9DB7
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
      Icon            =   "frmVendas_Cliente_RElUF.frx":BEDD
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   3105
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USButton btnResumida 
      Height          =   855
      Left            =   420
      TabIndex        =   3
      Top             =   1860
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   1508
      DibPicture      =   "frmVendas_Cliente_RElUF.frx":C1F7
      Caption         =   "Lista resumida"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      Theme           =   4
   End
   Begin DrawSuite2022.USButton btnEmail 
      Height          =   855
      Left            =   2850
      TabIndex        =   4
      Top             =   900
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   1508
      DibPicture      =   "frmVendas_Cliente_RElUF.frx":163A4
      Caption         =   "Lista email"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   1154291
      BorderColorDisabled=   13160660
      BorderColorDown =   16576
      BorderColorOver =   8438015
      GradientColor1  =   1154291
      GradientColor2  =   1154291
      GradientColor3  =   1154291
      GradientColor4  =   1154291
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorOver1=   8438015
      GradientColorOver2=   8438015
      GradientColorOver3=   8438015
      GradientColorOver4=   8438015
      GradientColorDown1=   16576
      GradientColorDown2=   16576
      GradientColorDown3=   16576
      GradientColorDown4=   16576
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      Theme           =   5
   End
   Begin DrawSuite2022.USButton btnEtiqueta 
      Height          =   855
      Left            =   2850
      TabIndex        =   5
      Top             =   1860
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   1508
      DibPicture      =   "frmVendas_Cliente_RElUF.frx":1FE51
      Caption         =   "Lista etiquetas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   0
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   8421504
      GradientColor1  =   0
      GradientColor2  =   0
      GradientColor3  =   0
      GradientColor4  =   0
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorOver1=   8421504
      GradientColorOver2=   8421504
      GradientColorOver3=   8421504
      GradientColorOver4=   8421504
      GradientColorDown1=   4210752
      GradientColorDown2=   4210752
      GradientColorDown3=   4210752
      GradientColorDown4=   4210752
      PicSize         =   3
      PicSizeH        =   32
      PicSizeW        =   32
      ShowFocusRect   =   0   'False
      Theme           =   6
   End
End
Attribute VB_Name = "frmVendas_Cliente_RelUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcImprimir()
On Error GoTo tratar_erro
  
With frmVendas_cliente
    If Opt_imprimir_rel.Value = True Then
        NomeRel = "Clientes.rpt"
        ProcImprimirRel .FormulaRel_Cliente, ""
    ElseIf Opt_imprimir_lista.Value = True Then
            NomeRel = "Clientes_lista_email.rpt"
            ProcImprimirRel .FormulaRel_Cliente, ""
        ElseIf Opt_imprimir_rel_res.Value = True Then
                NomeRel = "Clientes_resumido.rpt"
                'Debug.print .FormulaRel_Cliente
                ProcImprimirRel .FormulaRel_Cliente, ""
            Else
                If USMsgBox("Deseja excluir os registros antigos antes de emitir as etiquetas.", vbYesNo) = vbYes Then
                    Conexao.Execute "DELETE from etiqueta where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' and Tipo = 'C'"
                End If
                frmVendas_Cliente_RElUF_etiqueta.Show 1
    End If
End With
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnDetalhado_Click()
On Error GoTo tratar_erro

With frmVendas_cliente
    NomeRel = "Clientes.rpt"
    ProcImprimirRel .FormulaRel_Cliente, ""
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnEmail_Click()
On Error GoTo tratar_erro

With frmVendas_cliente

    NomeRel = "Clientes_lista_email.rpt"
    ProcImprimirRel .FormulaRel_Cliente, ""

End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub btnEtiqueta_Click()
On Error GoTo tratar_erro

    If USMsgBox("Deseja excluir os registros antigos antes de emitir as etiquetas.", vbYesNo) = vbYes Then
        Conexao.Execute "DELETE from etiqueta where Modulo = '" & Formulario & "' and Responsavel = '" & pubUsuario & "' and Tipo = 'C'"
    End If
    frmVendas_Cliente_RElUF_etiqueta.Show 1


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnResumida_Click()
On Error GoTo tratar_erro

With frmVendas_cliente
    NomeRel = "Clientes_resumido.rpt"
    ProcImprimirRel .FormulaRel_Cliente, ""
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    'case vbkeyF1: cmdAjuda
    Case vbKeyF5: ProcImprimir
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

