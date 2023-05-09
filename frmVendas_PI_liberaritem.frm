VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_PI_liberaritem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Liberar produto/serviço p/ faturamento"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1395
      Left            =   60
      TabIndex        =   9
      Top             =   990
      Width           =   8925
      Begin VB.TextBox txtPedidoInterno 
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
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Número do pedido interno."
         Top             =   375
         Width           =   855
      End
      Begin VB.TextBox txtQuantidade 
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
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade."
         Top             =   375
         Width           =   1215
      End
      Begin VB.TextBox Txt_un_com 
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
         Left            =   3060
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Unidade comercial."
         Top             =   375
         Width           =   765
      End
      Begin VB.TextBox txtLiberada 
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
         Left            =   5070
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade liberada."
         Top             =   375
         Width           =   1215
      End
      Begin VB.TextBox txtRevisao 
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
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Revisão do pedido."
         Top             =   375
         Width           =   525
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   955
         Width           =   8565
      End
      Begin VB.TextBox txtCodigoInterno 
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
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   375
         Width           =   1455
      End
      Begin VB.TextBox txtLiberar 
         Alignment       =   1  'Right Justify
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
         Left            =   7530
         TabIndex        =   7
         ToolTipText     =   "Quantidade à liberar."
         Top             =   375
         Width           =   1215
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
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
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Saldo."
         Top             =   375
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   3120
         TabIndex        =   19
         ToolTipText     =   "Quantidade à liberada."
         Top             =   180
         Width           =   645
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
         Height          =   195
         Left            =   4027
         TabIndex        =   18
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. liberada"
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
         Left            =   5145
         TabIndex        =   17
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label Label6 
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
         Height          =   195
         Left            =   4117
         TabIndex        =   16
         ToolTipText     =   "Quantidade à liberada."
         Top             =   765
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código interno"
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
         Left            =   1792
         TabIndex        =   15
         ToolTipText     =   "Quantidade à liberada."
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rev."
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
         ToolTipText     =   "Quantidade à liberada."
         Top             =   180
         Width           =   345
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido"
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
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. liberar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7530
         TabIndex        =   11
         ToolTipText     =   "Quantidade à liberada."
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6675
         TabIndex        =   10
         Top             =   180
         Width           =   465
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   12
      Top             =   0
      Width           =   8925
      _ExtentX        =   15743
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
      ButtonCaption1  =   "Liberar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Liberar (F3)"
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
      ButtonWidth1    =   41
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
      ButtonLeft2     =   45
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
      ButtonLeft3     =   49
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
      ButtonLeft4     =   87
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonKey5      =   "5"
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
      ButtonLeft5     =   115
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   2550
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_PI_liberaritem.frx":0000
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmVendas_PI_liberaritem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcLiberar()
On Error GoTo tratar_erro

Acao = LCase(USToolBar1.ButtonCaption(1))
If Compras_Pedido = True Then
    NomeCampo = "a Quantidade a comprar"
ElseIf Compras_Cotacao = True Then
        NomeCampo = "a quantidade a cotar"
    ElseIf Faturamento = True Then
            NomeCampo = "a quantidade a relacionar"
        ElseIf Sit_REG = 1 Or Sit_REG = 2 Then
                NomeCampo = "a quantidade a liberar"
            ElseIf Sit_REG = 3 Then
                    NomeCampo = "a quantidade a faturar"
                ElseIf Sit_REG = 4 Then
                        NomeCampo = "a quantidade expedir"
                    Else
                        NomeCampo = "a quantidade liberar"
End If
valor = IIf(txtLiberar = "", 0, txtLiberar)
If valor = 0 Then
    ProcVerificaAcao
    txtLiberar.SetFocus
    Exit Sub
End If
Qtd = Len(NomeCampo)
Texto = Right(NomeCampo, Qtd - 2)
If valor > qt Then
    USMsgBox ("A " & Texto & " não pode ser maior que " & Format(qt, "###,##0.000") & "."), vbExclamation, "CAPRIND v5.0"
    txtLiberar.SetFocus
    Exit Sub
End If

Qtd = txtQuantidade
Valor1 = txtLiberada.Text
If Compras_Pedido = True Then
    With frmCompras_Pedido
        If Sit_Data = 1 Then
            If Sit_REG = 1 Then .ProcNovo_Necessidade .Opt_vendas Else .ProcNovo_Necessidade frmCompras_ListaProduto.Opt_vendas
        Else
            If (valor + Valor1) < Qtd Then
                qt = Qtd - (valor + Valor1)
                .ProcNovo_Solicitacao False
            End If
            If (valor + Valor1) >= Qtd Then .ProcAlterar_Solicitacao False
        End If
    End With
ElseIf Compras_Cotacao = True Then
        With frmcompras_reqcot
            If Sit_REG = 1 Then .ProcNovo_Necessidade .Opt_vendas Else .ProcNovo_Necessidade frmCompras_reqcot_abrir.Opt_vendas
        End With
    ElseIf Faturamento = True Then
            Valor2 = txtSaldo
            If valor > Valor2 Then
                USMsgBox ("A quantidade a relacionar não pode ser maior que o saldo."), vbExclamation, "CAPRIND v5.0"
                txtLiberar.SetFocus
                Exit Sub
            End If
            ValorNC = valor
        Else
            Set TBVendas = CreateObject("adodb.recordset")
            TBVendas.Open "Select VP.NCotacao, VP.Revisao, VC.* from vendas_carteira VC INNER JOIN vendas_proposta VP ON VC.Cotacao = VP.Cotacao where VC.Codigo = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
            If TBVendas.EOF = False Then
                If Sit_REG = 1 Or Sit_REG = 2 Then
                    TBVendas!Data_lib_fat = Date
                    TBVendas!Responsavel_lib_fat = pubUsuario
                    If (valor + Valor1) < Qtd Then TBVendas!Liberacao = "FATURAR PARCIAL" Else TBVendas!Liberacao = "FATURAR"
                    TBVendas!qtdeliberada = TBVendas!qtdeliberada + valor
                    TBVendas.Update
                    If Vendas_PI = True Then
                        USMsgBox ("Produto liberado para faturamento com sucesso."), vbInformation, "CAPRIND v5.0"
                        Evento = "Liberar produto p/ faturamento"
                    Else
                        If Sit_REG = 1 Or Sit_REG = 2 Then
                            If TBVendas!Tipo = "P" Then
                                TextoTipo = "Produto"
                                TextoTipo1 = "Produto"
                            Else
                                TextoTipo = "Serviço"
                                TextoTipo1 = "serviço"
                            End If
                            USMsgBox (TextoTipo & " liberado para faturamento com sucesso."), vbInformation, "CAPRIND v5.0"
                            Evento = "Liberar " & TextoTipo1 & " p/ faturamento"
                        Else
                            USMsgBox ("Programação liberada para faturamento com sucesso."), vbInformation, "CAPRIND v5.0"
                            Evento = "Liberar programação p/ faturamento"
                        End If
                    End If
                ElseIf Sit_REG = 3 Then
                        If (valor + Valor1) < Qtd Then TBVendas!Liberacao = "FATURADO PARCIAL" Else TBVendas!Liberacao = "FATURADO"
                        TBVendas!qtdeliberada = TBVendas!qtdeliberada + valor
                        TBVendas!QtdeFaturada = TBVendas!QtdeFaturada + valor
                        TBVendas!DataFaturamento = Date
                        TBVendas.Update
                        
                        If IsNull(TBVendas!ID_programacao) = False And TBVendas!ID_programacao <> "0" Then
                            'Programação
                            Set TBProgramas = CreateObject("adodb.recordset")
                            TBProgramas.Open "Select * from Vendas_programacao where ID_prog = " & TBVendas!ID_programacao, Conexao, adOpenKeyset, adLockOptimistic
                            If TBProgramas.EOF = False Then
                                TBProgramas!QtdeFaturada = Format(TBVendas!QtdeFaturada, "###,##0.00")
                                If TBProgramas!QtdeFaturada >= TBProgramas!quantidade Then
                                    TBProgramas!Status_prog = "FATURADO"
                                    TBProgramas!Ordenar = 4
                                Else
                                    TBProgramas!Status_prog = "PARCIAL"
                                    TBProgramas!Ordenar = 1
                                End If
                                TBProgramas.Update
                            
                                Set TBItem = CreateObject("adodb.recordset")
                                TBItem.Open "Select * from vendas_programa_item where ID_item = " & TBProgramas!Id_Item, Conexao, adOpenKeyset, adLockOptimistic
                                If TBItem.EOF = False Then
                                    Set TBAbrir = CreateObject("adodb.recordset")
                                    TBAbrir.Open "Select * from vendas_programacao where id_item = " & TBItem!Id_Item & " and status_prog <> 'PREVISÃO FUTURA'", Conexao, adOpenKeyset, adLockOptimistic
                                    If TBAbrir.EOF = True Then
                                        TBItem!Status_Item = "PREVISÃO FUTURA"
                                    Else
                                        Set TBAbrir = CreateObject("adodb.recordset")
                                        TBAbrir.Open "Select * from vendas_programacao where id_item = " & TBItem!Id_Item & " and status_prog <> 'ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
                                        If TBAbrir.EOF = True Then
                                            TBItem!Status_Item = "ABERTO"
                                        Else
                                            Set TBAbrir = CreateObject("adodb.recordset")
                                            TBAbrir.Open "Select * from vendas_programacao where id_item = " & TBItem!Id_Item & " and status_prog <> 'FATURADO'", Conexao, adOpenKeyset, adLockOptimistic
                                            If TBAbrir.EOF = True Then
                                                TBItem!Status_Item = "FATURADO"
                                            Else
                                                TBItem!Status_Item = "PARCIAL"
                                            End If
                                        End If
                                    End If
                                    TBAbrir.Close
                                    TBItem.Update
                                End If
                            End If
                            TBProgramas.Close
                            
                            'Programa
                            Set TBItem = CreateObject("adodb.recordset")
                            TBItem.Open "Select vendas_programa.ID, vendas_programa.Status from (vendas_programa INNER JOIN vendas_proposta ON vendas_programa.ID = vendas_proposta.ID_programa) INNER JOIN vendas_carteira ON vendas_carteira.Cotacao = vendas_proposta.Cotacao where vendas_carteira.Codigo = " & TBVendas!CODIGO, Conexao, adOpenKeyset, adLockOptimistic
                            If TBItem.EOF = False Then
                                Do While TBItem.EOF = False
                                    Set TBAbrir = CreateObject("adodb.recordset")
                                    TBAbrir.Open "Select * from vendas_programa_item where id = " & TBItem!ID & " and Status_Item <> 'PREVISÃO FUTURA'", Conexao, adOpenKeyset, adLockOptimistic
                                    If TBAbrir.EOF = True Then
                                        TBItem!status = "PREVISÃO FUTURA"
                                    Else
                                        Set TBAbrir = CreateObject("adodb.recordset")
                                        TBAbrir.Open "Select * from vendas_programa_item where id = " & TBItem!ID & " and Status_Item <> 'ABERTO'", Conexao, adOpenKeyset, adLockOptimistic
                                        If TBAbrir.EOF = True Then
                                            TBItem!status = "ABERTO"
                                        Else
                                            Set TBAbrir = CreateObject("adodb.recordset")
                                            TBAbrir.Open "Select * from vendas_programa_item where id = " & TBItem!ID & " and Status_Item <> 'FATURADO'", Conexao, adOpenKeyset, adLockOptimistic
                                            If TBAbrir.EOF = True Then
                                                TBItem!status = "FATURADO"
                                            Else
                                                TBItem!status = "PARCIAL"
                                            End If
                                        End If
                                    End If
                                    TBAbrir.Close
                                    TBItem.Update
                                    TBItem.MoveNext
                                Loop
                            End If
                            TBItem.Close
                        End If
                        
                        FunAtualizaStatusPropPI TBVendas!Cotacao
                        
                        If TBVendas!Tipo = "P" Then
                            TextoTipo = "Produto"
                            TextoTipo1 = "Produto"
                        Else
                            TextoTipo = "Serviço"
                            TextoTipo1 = "serviço"
                        End If
                        USMsgBox (TextoTipo & " faturado com sucesso."), vbInformation, "CAPRIND v5.0"
                        Evento = "Faturar " & TextoTipo1
                    Else
                        TBVendas!qtdeexpedida = TBVendas!qtdeexpedida + valor
                        TBVendas.Update
                        If TBVendas!Tipo = "P" Then
                            TextoTipo = "Produto"
                            TextoTipo1 = "Produto"
                        Else
                            TextoTipo = "Serviço"
                            TextoTipo1 = "serviço"
                        End If
                        USMsgBox (TextoTipo & " expedido com sucesso."), vbInformation, "CAPRIND v5.0"
                        Evento = "Expedir " & TextoTipo1
                End If
                '==================================
                Modulo = Formulario
                ID_documento = TBVendas!CODIGO
                Documento = "Nº pedido: " & TBVendas!Ncotacao & " - Rev.: " & TBVendas!Revisao
                Documento1 = "Cód. interno: " & TBVendas!Desenho
                ProcGravaEvento
                '==================================
        End If
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
    Case vbKeyF3: ProcLiberar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 8895, 4, True
If Vendas_PI = True Then
    Formulario = "Vendas/Pedido interno"
ElseIf Compras_Pedido = True Then
        Formulario = "Compras/Pedido"
        If Sit_Data = 1 Then
            Label3.Caption = "Pedido de compra"
            txtPedidoInterno.ToolTipText = "Número do pedido de compra"
        Else
            Label3.Caption = "Solicitação"
            txtPedidoInterno.ToolTipText = "Número da solicitação"
        End If
        Label3.Width = 1395
        txtPedidoInterno.Width = 1395
        Label4.Visible = False
        txtrevisao.Visible = False
        Caption = "Liberar produto/serviço p/ comprar"
        With USToolBar1
            .ButtonCaption(1) = "Comprar"
            .ButtonToolTipText(1) = "Comprar (F3)"
        End With
        Label7.Caption = "Qtde. comprada"
        txtLiberada.ToolTipText = "Quantidade comprada."
        Label2.Caption = "Qtde. comprar"
        txtLiberar.ToolTipText = "Quantidade à comprar."
    ElseIf Compras_Cotacao = True Then
            Formulario = "Compras/Cotação"
            Label3.Caption = "Cotação"
            Label3.Width = 1395
            txtPedidoInterno.Width = 1395
            txtPedidoInterno.ToolTipText = "Número da cotação"
            Label4.Visible = False
            txtrevisao.Visible = False
            Caption = "Liberar produto/serviço p/ cotar"
            With USToolBar1
                .ButtonCaption(1) = "Cotar"
                .ButtonToolTipText(1) = "Cotar (F3)"
            End With
            Label7.Caption = "Qtde. cotada"
            txtLiberada.ToolTipText = "Quantidade cotada."
            Label2.Caption = "Qtde. cotar"
            txtLiberar.ToolTipText = "Quantidade à cotar."
        ElseIf Faturamento = True Then
                Label3.Caption = "Nota"
                txtPedidoInterno.ToolTipText = "Número da nota"
                Label4.Caption = "Série"
                txtrevisao.ToolTipText = "Série"
                Caption = "Relacionar nota fiscal"
                With USToolBar1
                    .ButtonCaption(1) = "Relacionar"
                    .ButtonToolTipText(1) = "Relacionar (F3)"
                End With
                Label7.Caption = "Qt. relacionada"
                txtLiberada.ToolTipText = "Quantidade relacionada."
                Label2.Caption = "Qt. relacionar"
                txtLiberar.ToolTipText = "Quantidade à relacionar."
            ElseIf Sit_REG = 1 Or Sit_REG = 2 Then
                        If Sit_REG = 1 Then Formulario = "PCP/Gerenciamento de ordem" Else Formulario = "Vendas/Follow up"
                        Caption = "Liberar produto/serviço p/ faturamento"
                    ElseIf Sit_REG = 3 Then
                            Formulario = "Vendas/Follow up"
                            Caption = "Faturar produto/serviço"
                            With USToolBar1
                                .ButtonCaption(1) = "Faturar"
                                .ButtonToolTipText(1) = "Faturar (F3)"
                            End With
                            Label7.Caption = "Qtde. faturada"
                            txtLiberada.ToolTipText = "Quantidade faturada."
                            Label2.Caption = "Qtde. faturar"
                            txtLiberar.ToolTipText = "Quantidade à faturar."
                        ElseIf Sit_REG = 4 Then
                                Formulario = "Vendas/Follow up"
                                Caption = "Expedir produto/serviço"
                                With USToolBar1
                                    .ButtonCaption(1) = "Expedir"
                                    .ButtonToolTipText(1) = "Expedir (F3)"
                                End With
                                Label7.Caption = "Qtde. expedida"
                                txtLiberada.ToolTipText = "Quantidade expedida."
                                Label2.Caption = "Qtde. expedir"
                                txtLiberar.ToolTipText = "Quantidade à expedir."
                            Else
                                Formulario = "Vendas/Programação"
                                Caption = "Liberar programação p/ faturamento"
End If

Set TBCarteira = CreateObject("adodb.recordset")
If Compras_Pedido = True Then
    If Sit_Data = 1 Then
        TBCarteira.Open "Select Desenho, Descricao, Unidade_com from projproduto where codproduto = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
        If TBCarteira.EOF = False Then
            txtPedidoInterno.Text = frmCompras_Pedido.txtPedido
            txtCodigoInterno.Text = IIf(IsNull(TBCarteira!Desenho), "", TBCarteira!Desenho)
            txtdescricao.Text = IIf(IsNull(TBCarteira!Descricao), "", TBCarteira!Descricao)
            Txt_un_com.Text = IIf(IsNull(TBCarteira!Unidade_com), "", TBCarteira!Unidade_com)
            txtQuantidade.Text = IIf(IsNull(Qtde), 0, Format(Qtde, "###,##0.0000"))
            txtLiberada = "0,0000"
            txtSaldo.Text = "0,0000"
            txtLiberar.Text = IIf(IsNull(Qtde), 0, Format(Qtde, "###,##0.0000"))
        End If
    Else
        TBCarteira.Open "Select CPL.Desenho, CPL.Descricao, CPL.Unidade_com, CPL.quant_req, CR.Requisicaotexto from Compras_pedido_lista CPL INNER JOIN compras_requisicao CR on CPL.ID_requisicao = CR.ID_requisicao where CPL.idlista = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
        If TBCarteira.EOF = False Then
            txtPedidoInterno.Text = IIf(IsNull(TBCarteira!Requisicaotexto), "", TBCarteira!Requisicaotexto)
            txtCodigoInterno.Text = IIf(IsNull(TBCarteira!Desenho), "", TBCarteira!Desenho)
            txtdescricao.Text = IIf(IsNull(TBCarteira!Descricao), "", TBCarteira!Descricao)
            Txt_un_com.Text = IIf(IsNull(TBCarteira!Unidade_com), "", TBCarteira!Unidade_com)
            txtQuantidade.Text = IIf(IsNull(TBCarteira!quant_req), 0, Format(TBCarteira!quant_req, "###,##0.0000"))
            txtLiberada = "0,0000"
            txtSaldo.Text = "0,0000"
            txtLiberar.Text = IIf(IsNull(TBCarteira!quant_req), 0, Format(TBCarteira!quant_req, "###,##0.0000"))
        End If
    End If
ElseIf Compras_Cotacao = True Then
        TBCarteira.Open "Select Desenho, Descricao, Unidade_com from projproduto where codproduto = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
        If TBCarteira.EOF = False Then
            txtPedidoInterno.Text = frmcompras_reqcot.txtidcotacao
            txtCodigoInterno.Text = IIf(IsNull(TBCarteira!Desenho), "", TBCarteira!Desenho)
            txtdescricao.Text = IIf(IsNull(TBCarteira!Descricao), "", TBCarteira!Descricao)
            Txt_un_com.Text = IIf(IsNull(TBCarteira!Unidade_com), "", TBCarteira!Unidade_com)
            txtQuantidade.Text = IIf(IsNull(Qtde), 0, Format(Qtde, "###,##0.0000"))
            txtLiberada = "0,0000"
            txtSaldo.Text = "0,0000"
            txtLiberar.Text = IIf(IsNull(Qtde), 0, Format(Qtde, "###,##0.0000"))
        End If
    ElseIf Faturamento = True Then
            TBCarteira.Open "Select NF.int_NotaFiscal, NF.Serie, NFP.* from tbl_Detalhes_Nota NFP INNER JOIN tbl_Dados_Nota_Fiscal NF ON NF.ID = NFP.ID_nota where NFP.Int_codigo = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
            If TBCarteira.EOF = False Then
                OF = IIf(IsNull(TBCarteira!int_NotaFiscal), 0, TBCarteira!int_NotaFiscal)
                txtPedidoInterno.Text = OF
                txtrevisao = IIf(IsNull(TBCarteira!Serie), "", TBCarteira!Serie)
                txtCodigoInterno.Text = IIf(IsNull(TBCarteira!int_Cod_Produto), "", TBCarteira!int_Cod_Produto)
                txtdescricao.Text = IIf(IsNull(TBCarteira!Txt_descricao), "", TBCarteira!Txt_descricao)
                Txt_un_com.Text = IIf(IsNull(TBCarteira!Unidade_com), "", TBCarteira!Unidade_com)
                txtQuantidade.Text = IIf(IsNull(TBCarteira!int_Qtd), 0, Format(TBCarteira!int_Qtd, "###,##0.0000"))
                txtLiberada = IIf(IsNull(TBCarteira!int_Qtd), 0, Format(TBCarteira!int_Qtd - TBCarteira!Saldo, "###,##0.0000"))
                txtSaldo.Text = IIf(IsNull(TBCarteira!Saldo), 0, Format(TBCarteira!Saldo, "###,##0.0000"))
                txtLiberar.Text = IIf(IsNull(Qtde), 0, Format(Qtde, "###,##0.0000"))
            End If
        Else
            TBCarteira.Open "Select VP.NCotacao, VP.Revisao, VC.* from vendas_carteira VC INNER JOIN vendas_proposta VP ON VC.Cotacao = VP.Cotacao where VC.Codigo = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
            If TBCarteira.EOF = False Then
                txtPedidoInterno.Text = IIf(IsNull(TBCarteira!Ncotacao), "", TBCarteira!Ncotacao)
                txtrevisao.Text = IIf(IsNull(TBCarteira!Revisao), "", TBCarteira!Revisao)
                txtCodigoInterno.Text = IIf(IsNull(TBCarteira!Desenho), "", TBCarteira!Desenho)
                txtdescricao.Text = IIf(IsNull(TBCarteira!Descricao), "", TBCarteira!Descricao)
                Txt_un_com.Text = IIf(IsNull(TBCarteira!Unidade_com), "", TBCarteira!Unidade_com)
                txtQuantidade.Text = IIf(IsNull(TBCarteira!quantidade), 0, Format(TBCarteira!quantidade, "###,##0.0000"))
                If Sit_REG = 1 Or Sit_REG = 2 Then
                    txtLiberada = IIf(IsNull(TBCarteira!qtdeliberada), 0, Format(TBCarteira!qtdeliberada, "###,##0.0000"))
                    Saldo = (TBCarteira!quantidade) - (TBCarteira!qtdeliberada)
                ElseIf Sit_REG = 3 Then
                        txtLiberada = IIf(IsNull(TBCarteira!QtdeFaturada), 0, Format(TBCarteira!QtdeFaturada, "###,##0.0000"))
                        Saldo = (TBCarteira!quantidade) - (TBCarteira!QtdeFaturada)
                    Else
                        txtLiberada = IIf(IsNull(TBCarteira!qtdeexpedida), 0, Format(TBCarteira!qtdeexpedida, "###,##0.0000"))
                        Saldo = (TBCarteira!quantidade) - (TBCarteira!qtdeexpedida)
                End If
                txtSaldo.Text = Format(IIf(Saldo < 0, 0, Saldo), "###,##0.0000")
                txtLiberar.Text = Format(IIf(Saldo < 0, 0, Saldo), "###,##0.0000")
            End If
End If
TBCarteira.Close
qt = txtLiberar 'Carrega na variavel para depois verificar se a quantidade digitada é menor

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Permitido2 = False
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLiberar_Change()
On Error GoTo tratar_erro

If txtLiberar <> "" Then
    VerifNumero = txtLiberar
    ProcVerificaNumero
    If VerifNumero = False Then
        txtLiberar = ""
        txtLiberar.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLiberar_LostFocus()
On Error GoTo tratar_erro

txtLiberar = Format(txtLiberar, "###,##0.0000")
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtLiberar_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtLiberar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcLiberar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
