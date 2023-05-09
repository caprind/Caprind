VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmProd_Resultados_Material 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "PCP | Gerenciamento de ordem - Resultados da ordem detalhado - Material"
   ClientHeight    =   8010
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   13320
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmProd_Resultados_Material.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   13320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   13320
      _ExtentX        =   23495
      _ExtentY        =   767
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
      Icon            =   "FrmProd_Resultados_Material.frx":000C
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   7605
      Width           =   13320
      _ExtentX        =   23495
      _ExtentY        =   714
   End
   Begin VB.TextBox TxtTotal 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   11130
      TabIndex        =   3
      Top             =   6960
      Width           =   1965
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5385
      Left            =   210
      TabIndex        =   0
      Top             =   1530
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   9499
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Un."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "D"
         Text            =   "Dt. retirada"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Lote"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Corrida"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Certificado"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Valor unit."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Valor ICMS"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Tag             =   "N"
         Text            =   "Valor total"
         Object.Width           =   1940
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7350
      Width           =   13305
      _ExtentX        =   23469
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor2      =   0
      SearchText      =   ""
      Value           =   0
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   30
      TabIndex        =   2
      Top             =   450
      Width           =   13245
      _ExtentX        =   23363
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
      ButtonCaption1  =   "Relatório"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Relatório (F5)"
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
      ButtonWidth1    =   51
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonAlignment2=   2
      ButtonType2     =   1
      ButtonStyle2    =   -1
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState2    =   -1
      ButtonLeft2     =   55
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonUseMaskColor2=   0   'False
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
      ButtonLeft3     =   59
      ButtonTop3      =   2
      ButtonWidth3    =   36
      ButtonHeight3   =   21
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
      ButtonLeft4     =   97
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
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
      ButtonLeft5     =   125
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   6180
         Top             =   240
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "FrmProd_Resultados_Material.frx":0028
         Count           =   1
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Valor total de materiais utilizados :"
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
      Height          =   225
      Left            =   8490
      TabIndex        =   4
      Top             =   7020
      Width           =   3735
   End
End
Attribute VB_Name = "FrmProd_Resultados_Material"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    Case vbKeyF5: ProcImprimir
    Case vbKeyEscape: Unload Me
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 13250, 5, True
ValorTotal = 0

Lista.ListItems.Clear
If OF = 0 Then Exit Sub
Set TBMaterial = CreateObject("adodb.recordset")

StrSql = "Select PM.Descricao,PM.codigo, PM.pesounidade, PM.quantidade,PM.requisitado, PM.unidade, EM.IDestoque, EM.Data, EM.Saida, EM.VlrUnit,Sum(em.Saida*em.vlrunit) as VlrTotal, EC.LOTE, EC.Corrida, EC.Certificado from (Producaomaterial PM INNER JOIN estoque_movimentacao EM ON PM.CODIGO = EM.Desenho) INNER JOIN estoque_controle EC ON EC.IDestoque = EM.IDEstoque where PM.ORDEM = " & OF & " and EM.oe = '" & OF & "' and EM.documento = '" & OF & "' and (EM.operacao = 'SAIDA_ORDEM' or EM.operacao = 'SAIDA_ORDEM_PARCIAL') group by IDmateriaprima, PM.Descricao, PM.Tipo, PM.Pesometro, PM.codigo, PM.pesounidade, PM.dimensao,PM.ordem,PM.quantidade,PM.requisitado, PM.unidade,PM.pesototal,PM.saida,PM.DimensaoTotal,PM.Valor_saida_estoque,PM.versao,PM.Total_pc,PM.posicao,PM.ID_carteira,PM.Percentual_perda,PM.ID_partnumber_fabricante,PM.Obs,EM.IdEstoque, EM.Data, EM.Saida,EM.VlrUnit,EM.Lote,EC.Lote,EC.corrida,EC.CERTIFICADO order by PM.Codigo"
'StrSql = "Select PM.Descricao,PM.codigo, PM.pesounidade, PM.quantidade,PM.requisitado, PM.unidade,EM.IdEstoque, EM.Data, EM.Saida,EM.VlrUnit,EM.Lote,EC.Lote,EC.corrida,EC.CERTIFICADO, EM.IDestoque, EM.Data, EM.Saida, EM.VlrUnit,Sum(em.Saida*em.vlrunit) as VlrTotal, EC.LOTE, EC.Corrida, EC.Certificado from (Producaomaterial PM INNER JOIN estoque_movimentacao EM ON PM.CODIGO = EM.Desenho) INNER JOIN estoque_controle EC ON EC.IDestoque = EM.IDEstoque where PM.ORDEM = 48359 and EM.documento = '48359' and (EM.operacao = 'SAIDA_ORDEM' or EM.operacao = 'SAIDA_ORDEM_PARCIAL') group by IDmateriaprima, PM.Descricao, PM.Tipo, PM.Pesometro, PM.codigo, PM.pesounidade, PM.dimensao,PM.ordem,PM.quantidade,PM.requisitado, PM.unidade,PM.pesototal,PM.saida,PM.DimensaoTotal,PM.Valor_saida_estoque,PM.versao,PM.Total_pc,PM.posicao,PM.ID_carteira,PM.Percentual_perda,PM.ID_partnumber_fabricante,PM.Obs,EM.IdEstoque, EM.Data, EM.Saida,EM.VlrUnit,EM.Lote,EC.Lote,EC.corrida,EC.CERTIFICADO order by PM.Codigo"
'Debug.print StrSql

TBMaterial.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBMaterial.EOF = False Then
    NomeRel = "Pcp_resultados da ordem_material.rpt"
    Familiatext = "{Producaomaterial.Ordem} = " & OF & " and {Estoque_movimentacao.oe} = '" & OF & "' and {Estoque_movimentacao.documento} = '" & OF & "' and ({Estoque_movimentacao.operacao} = 'SAIDA_ORDEM' or {Estoque_movimentacao.operacao} = 'SAIDA_ORDEM_PARCIAL')"
    
    PBLista.Min = 0
    PBLista.Max = TBMaterial.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBMaterial.EOF = False
        With Lista.ListItems
            .Add , , TBMaterial!IDEstoque
            .Item(.Count).SubItems(1) = IIf(IsNull(TBMaterial!CODIGO), "", TBMaterial!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBMaterial!Descricao), "", TBMaterial!Descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBMaterial!Unidade), "", TBMaterial!Unidade)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBMaterial!Data), "", Format(TBMaterial!Data, "dd/mm/yy"))
            .Item(.Count).SubItems(5) = IIf(IsNull(TBMaterial!LOTE), "", TBMaterial!LOTE)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBMaterial!Corrida), "", TBMaterial!Corrida)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBMaterial!Certificado), "", TBMaterial!Certificado)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBMaterial!Saida), "", Format(TBMaterial!Saida, "###,##0.0000"))
            .Item(.Count).SubItems(9) = IIf(IsNull(TBMaterial!VlrUnit), "", Format(TBMaterial!VlrUnit, "###,##0.00"))
            .Item(.Count).SubItems(11) = IIf(IsNull(TBMaterial!vlrTotal), "", Format(TBMaterial!vlrTotal, "###,##0.00"))
            ValorTotal = ValorTotal + TBMaterial!vlrTotal
        End With
        TBMaterial.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
Else
    Set TBCompras_Lista = CreateObject("adodb.recordset")
    TBCompras_Lista.Open "Select NFP.*, E.Real, P.ID_empresa, P.Pedido, CST_ICMS.Valor_ICMS, CST_ICMS.Valor_ICMS_ST, CST_ICMS.Valor_ICMS_SN from ((((tbl_Detalhes_Nota_pedidos NFPP INNER JOIN tbl_Detalhes_Nota NFP ON NFP.Int_codigo = NFPP.ID_prod_NF) INNER JOIN Compras_pedido_lista PP ON NFPP.ID_carteira = PP.IDlista and NFPP.Codinterno = PP.Desenho) INNER JOIN tbl_Detalhes_Nota_CST_ICMS CST_ICMS ON CST_ICMS.ID_item = NFP.Int_codigo) INNER JOIN Compras_pedido P ON P.IDpedido = PP.IDpedido) INNER JOIN Empresa E ON E.Codigo = P.ID_empresa where NFP.Ordem = '" & OF & "' and PP.remessa = 'False' and (PP.OS is null or PP.OS = 0) order by NFP.int_Cod_Produto", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Lista.EOF = False Then
        NomeRel = "Pcp_resultados da ordem_material_NF.rpt"
        Familiatext = "{tbl_Detalhes_Nota.Ordem} = '" & OF & "' and {Compras_pedido_lista.remessa} = False and (ISNULL({Compras_pedido_lista.OS}) or {Compras_pedido_lista.OS} = 0)"
        
        PBLista.Min = 0
        PBLista.Max = TBCompras_Lista.RecordCount
        PBLista.Value = 1
        Contador = 0
        Do While TBCompras_Lista.EOF = False
            With Lista.ListItems
                .Add , , TBCompras_Lista!Int_codigo
                .Item(.Count).SubItems(1) = IIf(IsNull(TBCompras_Lista!int_Cod_Produto), "", TBCompras_Lista!int_Cod_Produto)
                .Item(.Count).SubItems(2) = IIf(IsNull(TBCompras_Lista!Txt_descricao), "", TBCompras_Lista!Txt_descricao)
                .Item(.Count).SubItems(3) = IIf(IsNull(TBCompras_Lista!txt_Unid), "", TBCompras_Lista!txt_Unid)
                .Item(.Count).SubItems(5) = IIf(IsNull(TBCompras_Lista!Pedido), "", TBCompras_Lista!Pedido)
                .Item(.Count).SubItems(8) = IIf(IsNull(TBCompras_Lista!int_Qtd), "", Format(TBCompras_Lista!int_Qtd, "###,##0.0000"))
                .Item(.Count).SubItems(9) = IIf(IsNull(TBCompras_Lista!dbl_ValorUnitario), "", Format(TBCompras_Lista!dbl_ValorUnitario, "###,##0.0000000000"))
                
                If IsNull(TBCompras_Lista!Valor_ICMS) = False And TBCompras_Lista!Valor_ICMS <> 0 Then
                    ValorICMS = TBCompras_Lista!Valor_ICMS
                ElseIf IsNull(TBCompras_Lista!Valor_ICMS_ST) = False And TBCompras_Lista!Valor_ICMS_ST <> 0 Then
                        ValorICMS = TBCompras_Lista!Valor_ICMS_ST
                    ElseIf IsNull(TBCompras_Lista!Valor_ICMS_SN) = False And TBCompras_Lista!Valor_ICMS_SN <> 0 Then
                            ValorICMS = TBCompras_Lista!Valor_ICMS_SN
                End If
                .Item(.Count).SubItems(10) = Format(ValorICMS, "###,##0.00")
                
                If TBCompras_Lista!Real = True Then
                    .Item(.Count).SubItems(11) = Format((IIf(IsNull(TBCompras_Lista!dbl_ValorTotal), 0, TBCompras_Lista!dbl_ValorTotal) - ValorICMS) - (IIf(IsNull(TBCompras_Lista!Total_PIS_prod), 0, TBCompras_Lista!Total_PIS_prod) + IIf(IsNull(TBCompras_Lista!Total_Cofins_prod), 0, TBCompras_Lista!Total_Cofins_prod)), "###,##0.00")
                Else
                    .Item(.Count).SubItems(11) = Format(IIf(IsNull(TBCompras_Lista!dbl_ValorTotal), 0, TBCompras_Lista!dbl_ValorTotal) - ValorICMS, "###,##0.00")
                End If
            End With
            Contador = Contador + 1
            PBLista.Value = Contador
            TBCompras_Lista.MoveNext
        Loop
        
    End If
    TBCompras_Lista.Close
End If
TBMaterial.Close
txtTotal.Text = Format(ValorTotal, "###,##0.00")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcImprimir
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub
ProcImprimirRel Familiatext, ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
