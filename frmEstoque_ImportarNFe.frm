VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEstoque_ImportarNFe 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estoque - Recebimento - Importar XML"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11490
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   1425
      Left            =   9450
      TabIndex        =   16
      Top             =   990
      Width           =   1935
      Begin VB.TextBox txtNotaFiscal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         MaxLength       =   9
         TabIndex        =   6
         ToolTipText     =   "Número da nota fiscal."
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox txtSerie 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1260
         MaxLength       =   3
         TabIndex        =   7
         ToolTipText     =   "Série."
         Top             =   960
         Width           =   495
      End
      Begin MSComCtl2.DTPicker txtDataemissao 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         ToolTipText     =   "Data de emissão da nota fiscal."
         Top             =   390
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   180813825
         CurrentDate     =   39057
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "N° nota fiscal   Série"
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
         Left            =   240
         TabIndex        =   18
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Data emissão"
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
         Left            =   487
         TabIndex        =   17
         Top             =   180
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   1425
      Left            =   55
      TabIndex        =   11
      Top             =   990
      Width           =   9375
      Begin VB.TextBox txtID_empresa 
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
         Left            =   210
         Locked          =   -1  'True
         MaxLength       =   255
         MouseIcon       =   "frmEstoque_ImportarNFe.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox txtfornecedor 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Fornecedor."
         Top             =   390
         Width           =   6045
      End
      Begin VB.TextBox txtPedido 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         MaxLength       =   50
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Programação de compra."
         Top             =   390
         Width           =   1815
      End
      Begin VB.CommandButton cmdPedido 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2010
         Picture         =   "frmEstoque_ImportarNFe.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Filtrar."
         Top             =   390
         Width           =   315
      End
      Begin VB.TextBox txtEmpresa 
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
         MaxLength       =   255
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Empresa."
         Top             =   960
         Width           =   8985
      End
      Begin VB.TextBox txtIDforn 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Código do fornecedor."
         Top             =   390
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor"
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
         Left            =   5730
         TabIndex        =   15
         Top             =   180
         Width           =   825
      End
      Begin VB.Label lblPedido 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido de compra"
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
         Left            =   450
         TabIndex        =   14
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Left            =   4365
         TabIndex        =   13
         Top             =   750
         Width           =   735
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         Left            =   2700
         TabIndex        =   12
         Top             =   180
         Width           =   165
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   3150
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmEstoque_ImportarNFe.frx":0725
      Count           =   1
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   11325
      _ExtentX        =   19976
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
      ButtonCaption1  =   "Importar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Importar XML (F3)"
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
      ButtonWidth1    =   50
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
      ButtonLeft2     =   54
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
      ButtonLeft3     =   58
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
      ButtonLeft4     =   96
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
      ButtonLeft5     =   124
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   4860
      Width           =   11325
      _ExtentX        =   19976
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
   Begin MSComctlLib.ListView Lista 
      Height          =   2415
      Left            =   60
      TabIndex        =   8
      Top             =   2430
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   11879
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Un."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   2646
      EndProperty
   End
End
Attribute VB_Name = "frmEstoque_ImportarNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub procImportar()
On Error GoTo tratar_erro

ID_nota = 0
Acao = "importar XML"
If txtPedido = "" Then
    NomeCampo = "o pedido de compra"
    ProcVerificaAcao
    txtPedido.SetFocus
    Exit Sub
End If
If txtnotafiscal = "" Then
    NomeCampo = "a nota fiscal"
    ProcVerificaAcao
    txtnotafiscal.SetFocus
    Exit Sub
End If
If txtSerie = "" Then
    NomeCampo = "a série"
    ProcVerificaAcao
    txtSerie.SetFocus
    Exit Sub
End If

'Verifica se a nota fiscal esta validada
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select ID from tbl_dados_nota_fiscal where ID_empresa = " & txtID_empresa & " and int_NotaFiscal = '" & txtnotafiscal & "' and Serie = '" & txtSerie & "' and Id_Int_Cliente = " & txtIDforn & " and int_TipoNota = 2 and DtValidacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    USMsgBox ("Não é possivel importar este número de nota fiscal, pois a mesma está validada."), vbExclamation, "CAPRIND v5.0"
    txtnotafiscal.SetFocus
    Exit Sub
End If

Permitido = False
Permitido1 = False
Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from ENT001 where EntNtfNumero = " & Format(txtnotafiscal) & " and EntNtfSerie = '" & txtSerie & "'", Conexao_NFe, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from ENT001DET where ENTNTFNumero = " & TBFIltro!ENTNTFNumero, Conexao_NFe, adOpenKeyset, adLockOptimistic
    Do While TBProduto.EOF = False
        Set TBCompras_Lista = CreateObject("adodb.recordset")
        TBCompras_Lista.Open "Select CPL.IDlista, CPL.IDpedido, CPL.desenho, CPL.Descricao, CPL.Quant_Comp, CPL.Quant_comp_PC from Compras_pedido CP INNER JOIN Compras_pedido_lista CPL ON CP.IDPedido = CPL.IDpedido where CP.pedido = '" & txtPedido & "' and CPL.desenho = '" & TBProduto!EntcProd & "' and (Status_item = 'N_RECEBIDO' or Status_item = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
        If TBCompras_Lista.EOF = True Then
            Set TBCompras_Lista = CreateObject("adodb.recordset")
            TBCompras_Lista.Open "Select CPL.IDlista, CPL.IDpedido, CPL.desenho, CPL.Descricao, CPL.Quant_Comp, CPL.Quant_comp_PC from (Compras_pedido CP INNER JOIN Compras_pedido_lista CPL ON CP.IDPedido = CPL.IDpedido) INNER JOIN Item_aplicacoes IP ON IP.Codproduto = CPL.Codproduto where CP.pedido = '" & txtPedido & "' and IP.N_referencia = '" & TBProduto!EntcProd & "' and (CPL.Status_item = 'N_RECEBIDO' or CPL.Status_item = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
            If TBCompras_Lista.EOF = False Then GoTo Continua
        Else
Continua:
            frmEstoque_ImportarNFe_item.Show 1
            If Permitido = False Then Exit Sub
            
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from estoque_controle", Conexao, adOpenKeyset, adLockOptimistic
            TBEstoque.AddNew
            
            IDFase = 0
            IDPlano = 0
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select Codproduto, Desenho, descricaotecnica, Unidade, Unidade_com, classe, ID_PC from projproduto where desenho = '" & TBCompras_Lista!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                TBEstoque!Desenho = TBItem!Desenho
                TBEstoque!Descricao = TBItem!descricaotecnica
                TBEstoque!Un = TBItem!Unidade
                TBEstoque!Un_com = TBItem!Unidade_com
                
                TBEstoque!Classe = TBItem!Classe
                IDFase = TBItem!Codproduto
                IDPlano = IIf(IsNull(TBItem!ID_PC), 0, TBItem!ID_PC)
                Familiatext = IIf(IsNull(TBItem!Classe), "", TBItem!Classe)
            End If
            TBItem.Close
            
            TBEstoque!LOTE = txtPedido
            TBEstoque!Data = Data_Prog
            TBEstoque!Responsavel = pubUsuario
            TBEstoque!Certificado = Certificado
            TBEstoque!Corrida = Corrida
            TBEstoque!local_armaz = DesenhoProduto
            TBEstoque!Fornecedor = TBFIltro!EntxNome
            TBEstoque!ID_empresa = txtID_empresa
            
            Permitido = False
            Set TBCFOP = CreateObject("adodb.recordset")
            TBCFOP.Open "Select * from projproduto where desenho = '" & TBCompras_Lista!Desenho & "' and Estoque = 'True'", Conexao, adOpenKeyset, adLockOptimistic
            If TBCFOP.EOF = False Then Permitido = True
            
            'Verifica se o produto do pedido é remessa
            Set TBCFOP = CreateObject("adodb.recordset")
            TBCFOP.Open "Select CPL.* from Compras_pedido_lista CPL LEFT JOIN projproduto P ON P.Desenho = CPL.Desenho where CPL.Idlista = " & TBCompras_Lista!IDlista & " and CPL.Remessa = 'True' and P.Subtipoitem <> 4", Conexao, adOpenKeyset, adLockOptimistic
            If TBCFOP.EOF = False Then Permitido = False
            
            'Verifica se o produto do pedido é mão de obra
            Set TBCFOP = CreateObject("adodb.recordset")
            TBCFOP.Open "Select CPL.* from (Compras_pedido_lista CPL LEFT JOIN projproduto P ON P.Desenho = CPL.Desenho) LEFT JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = CPL.ID_CFOP where CPL.Idlista = " & TBCompras_Lista!IDlista & " and CFOP.MaoObra = 'True' and P.Subtipoitem <> 4", Conexao, adOpenKeyset, adLockOptimistic
            If TBCFOP.EOF = False Then Permitido = False
            
            If Permitido = True Then
                TBEstoque!estoque_venda = IIf(IsNull(TBProduto!EntqCOM), 0, TBProduto!EntqCOM)
                TBEstoque!estoque_real = IIf(IsNull(TBProduto!EntqCOM), 0, TBProduto!EntqCOM)
                Valor_Cofins_Prod = FunCalculaQtdePC(TBCompras_Lista!Desenho, IIf(IsNull(TBProduto!EntqCOM), 0, TBProduto!EntqCOM), True, TBEstoque!Un_com)
                TBEstoque!estoque_real_PC = Valor_Cofins_Prod
            End If
            
            Qtd = IIf(IsNull(TBProduto!EntqCOM), 0, TBProduto!EntqCOM)
            TBEstoque!Qtde = Qtd
            TBEstoque!status = "ENTRADA_NOTA_FISCAL"
    
            TBEstoque.Update
            IDEstoque = TBEstoque!IDEstoque
            TBEstoque.Close
    
            ValorTotal = IIf(IsNull(TBProduto!EntvUnCom), 0, TBProduto!EntvUnCom)
            Valor3 = IIf(IsNull(TBProduto!EntvProd), 0, TBProduto!EntvProd)
            
            NovoValor = Replace(ValorTotal, ",", ".")
            Conexao.Execute "Update estoque_controle Set valor_unitario = " & NovoValor & " where IDestoque = " & IDEstoque
            If Permitido = True Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Valor_total from estoque_controle where IDestoque = " & IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    TBAbrir!Valor_total = Format((Qtd * ValorTotal) + IIf(IsNull(TBAbrir!Valor_total), 0, TBAbrir!Valor_total), "###.##0.00")
                    TBAbrir.Update
                End If
                TBAbrir.Close
            End If
            
            quantestoque = 0
            'Grava movimentação na tabela estoque_controle_recebimento
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from estoque_controle_recebimento where idpedido = " & TBCompras_Lista!IDpedido & " and idlista = " & TBCompras_Lista!IDlista & " and certificado = '" & Certificado & "' and corrida = '" & Corrida & "' and nota_fiscal = '" & txtnotafiscal & "' and Serie = '" & txtSerie & "' and local_armaz = '" & DesenhoProduto & "' and Programacao = 'False' and id_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = True Then TBEstoque.AddNew
            TBEstoque!Data_recebimento = Data_Prog
            TBEstoque!IDpedido = TBCompras_Lista!IDpedido
            TBEstoque!IDlista = TBCompras_Lista!IDlista
            TBEstoque!Desenho = TBCompras_Lista!Desenho
            TBEstoque!Certificado = Certificado
            TBEstoque!Corrida = Corrida
            TBEstoque!local_armaz = DesenhoProduto
            TBEstoque!Nota_fiscal = txtnotafiscal
            TBEstoque!Serie = txtSerie
            If txtnotafiscal <> "" Then TBEstoque!Data_emissao = txtDataemissao.Value Else TBEstoque!Data_emissao = Null
            TBEstoque!Responsavel = pubUsuario
            
            If Valor_Cofins_Prod = 0 Then
                quantestoque = Format(TBCompras_Lista!Quant_Comp, "###.##0.000")
                quantnovo = Format(TBProduto!EntqCOM, "###.##0.000")
            Else
                quantestoque = Format(IIf(IsNull(TBCompras_Lista!Quant_Comp_PC), Valor_Cofins_Prod, TBCompras_Lista!Quant_Comp_PC), "###.##0.000")
                quantnovo = Format(Valor_Cofins_Prod, "###.##0.000")
            End If
            If quantnovo < quantestoque Then TBEstoque!Parcial = True Else TBEstoque!Parcial = False
            TBEstoque!Programacao = False
            TBEstoque!ID_empresa = txtID_empresa
            TBEstoque.Update
            IDEstoque_recebimento = TBEstoque!ID
            TBEstoque.Close
            
            'Soma quantidade da movimentação
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Recebido, Recebido_PC from estoque_controle_recebimento where idpedido = " & TBCompras_Lista!IDpedido & " and idlista = " & TBCompras_Lista!IDlista & " and certificado = '" & Certificado & "' and corrida = '" & Corrida & "' and nota_fiscal = '" & txtnotafiscal & "' and Serie = '" & txtSerie & "' and local_armaz = '" & DesenhoProduto & "' and Programacao = 'False' and id_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                TBAbrir!Recebido = IIf(IsNull(TBAbrir!Recebido), 0, TBAbrir!Recebido) + TBProduto!EntqCOM
                TBAbrir!Recebido_PC = IIf(IsNull(TBAbrir!Recebido_PC), 0, TBAbrir!Recebido_PC) + Valor_Cofins_Prod
                TBAbrir.Update
            End If
            TBAbrir.Close
            
            quantestoque = 0
            quantnovo = 0
            'Grava movimentação na tabela estoque_movimentacao
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
            TBEstoque.AddNew
            TBEstoque!Destino = "Interno"
            TBEstoque!Terceiros = False
            TBEstoque!Operacao = "ENTRADA_NOTA_FISCAL"
            TBEstoque!IDEstoque = IDEstoque
            TBEstoque!Documento = txtnotafiscal
            TBEstoque!DtEmissao = Data_Prog
            TBEstoque!LOTE = txtPedido
            TBEstoque!Responsavel = pubUsuario
            TBEstoque!Data = Data_Prog
            TBEstoque!Descricao = TBCompras_Lista!Descricao
            TBEstoque!Desenho = TBCompras_Lista!Desenho
            TBEstoque!estoque_venda = Format(TBProduto!EntqCOM, "###.##0.000")
            TBEstoque!Entrada = Format(TBProduto!EntqCOM, "###.##0.000")
            TBEstoque!Entrada_PC = Format(Valor_Cofins_Prod, "###.##0.000")
            TBEstoque!Familia = Familiatext
            quantestoque = TBProduto!EntqCOM
            
            'Atualiza valor do material no estoque
            TBEstoque!VlrUnit = Format(ValorTotal, "###.##0.00000")
            TBEstoque!vlrTotal = Format(quantestoque * ValorTotal, "###.##0.00")
            
            TBEstoque!IDEstoque_recebimento = IDEstoque_recebimento
            TBEstoque!idlista_recebimento = TBCompras_Lista!IDlista
            TBEstoque!Destino = "Interno"
            TBEstoque!Terceiros = False
            
            Set TBNivel1 = CreateObject("adodb.recordset")
            TBNivel1.Open "Select * from estoque_movimentacao where pedidocompra = '" & txtPedido & "' and desenho = '" & TBCompras_Lista!Desenho & "' and destino = 'Terceiros'", Conexao, adOpenKeyset, adLockOptimistic
            If TBNivel1.EOF = False Then
                Set TBNivel2 = CreateObject("adodb.recordset")
                TBNivel2.Open "Select sum(Saida) as quantidade from estoque_movimentacao where pedidocompra = '" & txtPedido & "' and desenho = '" & TBCompras_Lista!Desenho & "' and destino = 'Terceiros'", Conexao, adOpenKeyset, adLockOptimistic
                If TBNivel2.EOF = False Then
                    Valor1 = IIf(IsNull(TBNivel2!quantidade), 0, TBNivel2!quantidade)
                End If
                TBNivel2.Close
                Set TBNivel2 = CreateObject("adodb.recordset")
                TBNivel2.Open "Select sum(entrada) as quantidade from estoque_movimentacao where pedidocompra = '" & txtPedido & "' and desenho = '" & TBCompras_Lista!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBNivel2.EOF = False Then
                    Valor2 = IIf(IsNull(TBNivel2!quantidade), 0, TBNivel2!quantidade)
                End If
                TBNivel2.Close
                Valor2 = Valor2 + Qtd
                If Valor1 <= Valor2 Then
                    Conexao.Execute "UPDATE estoque_movimentacao set Terceiros = 'False' where pedidocompra = '" & txtPedido & "' and desenho = '" & TBCompras_Lista!Desenho & "' and destino = 'Terceiros'"
                End If
                TBEstoque!Pedidocompra = txtPedido
                TBEstoque!IDpedido = IDpedido
            Else
                TBEstoque!Pedidocompra = Null
                TBEstoque!IDpedido = Null
            End If
            TBNivel1.Close
            TBEstoque.Update
            
            'Centro de custo
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select CPLC.Valor, CPLC.ID_CC, CP.Data from Compras_pedido_lista_custo CPLC INNER JOIN Compras_pedido CP ON CPLC.IDPedido = CP.IDPedido where CPLC.IDLista = " & TBCompras_Lista!IDlista & " and CP.id_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    Valor3 = TBAbrir!valor
                    qt = TBCompras_Lista!Quant_Comp
                    Qtde = TBEstoque!Entrada
                    valor = Format((Valor3 / qt) * Qtde, "###,##0.00")
                    
                    'Verifica se tem CC amarrado no produto, se for diferente do informado no pedido cria débito e crédito no CC do produto
                    Set TBExecucao = CreateObject("adodb.recordset")
                    TBExecucao.Open "Select ID_CC from projproduto where codproduto = " & IDFase & " and ID_CC is not null", Conexao, adOpenKeyset, adLockOptimistic
                    If TBExecucao.EOF = False Then
                        If TBExecucao!ID_CC <> "" Then
                            If TBAbrir!ID_CC <> TBExecucao!ID_CC Then
                                ProcSalvarCCRealizado TBAbrir!Data, txtID_empresa, "Débito", TBExecucao!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, TBCompras_Lista!IDlista, valor, True, False
                                
                                'Grava movimentação no centro consolidado
                                ProcSalvarRealCCConsolidado TBExecucao!ID_CC, "Débito", False, False, False
                                
                                ProcSalvarCCRealizado TBAbrir!Data, txtID_empresa, "Crédito", TBExecucao!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, TBCompras_Lista!IDlista, valor, True, False
                                
                                'Grava movimentação no centro consolidado
                                ProcSalvarRealCCConsolidado TBExecucao!ID_CC, "Crédito", True, True, False
                            End If
                        End If
                    End If
                    TBExecucao.Close
                    
                    ProcSalvarCCRealizado TBAbrir!Data, txtID_empresa, "Débito", TBAbrir!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, TBCompras_Lista!IDlista, valor, False, False
                    
                    'Grava movimentação no centro consolidado
                    ProcSalvarRealCCConsolidado TBAbrir!ID_CC, "Débito", False, False, False
                    
                    TBAbrir.MoveNext
                Loop
            Else
                'Verifica se tem CC amarrado no produto e cria um débito no CC do produto
                Set TBExecucao = CreateObject("adodb.recordset")
                TBExecucao.Open "Select ID_CC from projproduto where codproduto = " & IDFase & " and ID_CC is not null", Conexao, adOpenKeyset, adLockOptimistic
                If TBExecucao.EOF = False Then
                    qt = TBCompras_Lista!Quant_Comp
                    Qtde = TBEstoque!Entrada
                    valor = Format((Valor3 / qt) * Qtde, "###,##0.00")
                    
                    ProcSalvarCCRealizado Date, txtID_empresa, "Débito", TBExecucao!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, TBCompras_Lista!IDlista, valor, False, False
                    
                    'Grava movimentação no centro consolidado
                    ProcSalvarRealCCConsolidado TBExecucao!ID_CC, "Débito", False, False, False
                End If
                TBExecucao.Close
            End If
            TBAbrir.Close
            TBEstoque.Close
            
            'Atualiza o status do pedido, status e qtde recebida da programação
            ProcAlteraStatus_pedido TBCompras_Lista!IDlista, IIf(IsNull(TBCompras_Lista!Quant_Comp), 0, TBCompras_Lista!Quant_Comp), IIf(IsNull(TBCompras_Lista!Quant_Comp_PC), 0, TBCompras_Lista!Quant_Comp_PC)
            
            '==================================
            Modulo = "Estoque/Recebimento"
            Evento = "Receber"
            ID_documento = IDEstoque
            Documento = "Cód. interno: " & TBCompras_Lista!Desenho & " - Nº lote: " & txtPedido & " - Nº corrida: " & Corrida & " - Nº certificado: " & Certificado & " - Local armaz.: " & DesenhoProduto
            Documento1 = ""
            ProcGravaEvento
            '==================================
            Permitido1 = True
        End If
        TBCompras_Lista.Close
        TBProduto.MoveNext
    Loop
    TBProduto.Close
    
    If Permitido1 = True Then
        'Cria a nota fiscal
        procCriarNF
        ProcCarregaLista
        USMsgBox ("Nota fiscal importada com sucesso."), vbInformation, "CAPRIND v5.0"
    Else
        USMsgBox ("Não foi possivel importar a nota fiscal, pois o(s) produto(s) do pedido não foi(ram) localizado(s) na nota fiscal importada pelo GNFe."), vbInformation, "CAPRIND v5.0"
    End If
Else
    USMsgBox ("Não foi econtrado nenhuma nota fiscal importada com este número " & txtnotafiscal & " no GNFe."), vbInformation, "CAPRIND v5.0"
End If
'TBFIltro.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdpedido_Click()
On Error GoTo tratar_erro

IDpedido = 0
txtIDforn = ""
txtFornecedor = ""
txtEmpresa = ""
Lista.ListItems.Clear
ProcCarregaPedido

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaPedido()
On Error GoTo tratar_erro

If txtPedido <> "" Then
    Set TBCompras_Pedido = CreateObject("adodb.recordset")
    TBCompras_Pedido.Open "Select ID_empresa, Data, IDpedido, idfornecedor, Fornecedor from compras_pedido where pedido = '" & txtPedido & "' and (Status_pedido = 'ABERTO' or Status_pedido = 'PARCIAL' or Status_pedido = 'ENCERRADO')", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras_Pedido.EOF = False Then
        IDpedido = TBCompras_Pedido!IDpedido
        txtIDforn = TBCompras_Pedido!IDFornecedor
        txtFornecedor = IIf(IsNull(TBCompras_Pedido!Fornecedor) = False, TBCompras_Pedido!Fornecedor, "")
        Set TBExecucao = CreateObject("adodb.recordset")
        TBExecucao.Open "Select CODIGO, Empresa from Empresa where codigo = " & TBCompras_Pedido!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
        If TBExecucao.EOF = False Then
            txtID_empresa = IIf(IsNull(TBExecucao!CODIGO), "", TBExecucao!CODIGO)
            txtEmpresa = IIf(IsNull(TBExecucao!Empresa), "", TBExecucao!Empresa)
        End If
        TBExecucao.Close
    End If
    TBCompras_Pedido.Close
    ProcCarregaLista
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: procImportar
    'Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 11325, 5, True
ProcLimpaVariaveisPrincipais
Formulario = "Estoque/Recebimento/Importar nota de terceiros"
Direitos
txtDataemissao.Value = Date

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procCriarNF()
On Error GoTo tratar_erro

'Cria a nota fiscal
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_Dados_Nota_Fiscal where ID_empresa = " & txtID_empresa & " and int_NotaFiscal = '" & txtnotafiscal & "' and Serie = '" & txtSerie & "' and Id_Int_Cliente = " & txtIDforn & " and int_TipoNota = 2 and TipoNF = 'M1'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then
    TBGravar.AddNew
    TBGravar!TabelaSN = 0
    TBGravar!pedido_interno = False
    TBGravar!DtValidacaoOF = Now
    TBGravar!RespValidacaoOF = pubUsuario
    TBGravar!int_NotaFiscal = txtnotafiscal
    TBGravar!Serie = txtSerie
    TBGravar!int_TipoNota = "2"
    TBGravar!TipoNF = "M1"
    TBGravar!dt_Saida_Entrada = Format(Date, "dd/mm/yyyy")
    TBGravar!txt_Hora_Saida = Format(Now, "hh:mm:ss")

    TBGravar!txt_Razao_Nome = txtFornecedor
    Set TBFornecedor = CreateObject("adodb.recordset")
    TBFornecedor.Open "Select CF.*, CP.ID_empresa FROM Compras_fornecedores CF INNER JOIN Compras_pedido CP ON CF.IDCliente = CP.idfornecedor where CP.Pedido = '" & txtPedido & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFornecedor.EOF = False Then
        TBGravar!ID_empresa = TBFornecedor!ID_empresa
        TBGravar!Id_Int_Cliente = TBFornecedor!IDCliente
        TBGravar!txt_Endereco = IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
        TBGravar!Numero = IIf(IsNull(TBFornecedor!Numero), "", TBFornecedor!Numero)
        TBGravar!txt_Bairro = IIf(IsNull(TBFornecedor!Bairro), "", TBFornecedor!Bairro)
        TBGravar!Txt_CEP = IIf(IsNull(TBFornecedor!CEP), "", TBFornecedor!CEP)
        TBGravar!txt_Municipio = IIf(IsNull(TBFornecedor!Cidade), "", TBFornecedor!Cidade)
        TBGravar!txt_Fone_Fax = IIf(IsNull(TBFornecedor!Telefones), "", TBFornecedor!Telefones)
        
        TBGravar!txt_UF = IIf(IsNull(TBFornecedor!Estado), "", TBFornecedor!Estado)
        UF = IIf(IsNull(TBFornecedor!Estado), "", TBFornecedor!Estado)
        
        If TBFornecedor!idTipoEmpresa = 1 Then TBGravar!txt_CNPJ_CPF = IIf(IsNull(TBFornecedor!CPF_CNPJ), "", TBFornecedor!CPF_CNPJ)
        TBGravar!txt_IE_Cliente = IIf(IsNull(TBFornecedor!RG_IE), "", TBFornecedor!RG_IE)
        If TBFornecedor!Pessoa = "JURÍDICA" Then
            TBGravar!txt_tipocliente = "J"
        Else
            TBGravar!txt_tipocliente = "F"
        End If
    End If
    TBGravar!dt_DataEmissao = txtDataemissao
    TBGravar!txt_Hora_Saida = Format(Now, "hh:mm:ss")
    TBGravar!Int_status = "1"
    TBGravar!Aplicacao = "T"
    TBGravar.Update
    IDlista = TBGravar!ID
Else
    IDlista = TBGravar!ID
    
    'Verifica se a NF já foi validada e não permite alteração
    If IsNull(TBGravar!DtValidacao) = False Then
        USMsgBox ("Esta nota fiscal não será alterada, pois a mesma já foi validada."), vbInformation, "CAPRIND v5.0"
        TBGravar.Close
        Exit Sub
    End If
End If
TBGravar.Close

'Puxar chave de acesso
Set TBHistProc = CreateObject("adodb.recordset")
TBHistProc.Open "Select EntChaAcesso, EntindPag, EntFinNFe from ENT001 where EntNtfNumero = " & Format(txtnotafiscal) & " and EntNtfSerie = '" & txtSerie & "'", Conexao_NFe, adOpenKeyset, adLockOptimistic
If TBHistProc.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * From tbl_Dados_Nota_Fiscal_NFe where ID_nota = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        TBAbrir.AddNew
        TBAbrir!ID_nota = IDlista
        TBAbrir!status = Null
    End If
    TBAbrir!Chave_acesso = TBHistProc!EntChaAcesso
    TBAbrir!Forma_pagamento = TBHistProc!EntindPag
    TBAbrir!Finalidade_emissao = TBHistProc!EntFinNFe
    TBAbrir.Update
End If
TBHistProc.Close

'Cria ou altera os produtos
Desenho = ""
OrdemTexto = ""
valor = 0
ValorTotal = 0
OF = 0
NovoValor = ""
DesenhoProduto = ""
Set TBFornecedor = CreateObject("adodb.recordset")
TBFornecedor.Open "Select CF.* FROM Compras_fornecedores CF INNER JOIN Compras_pedido CP ON CF.IDCliente = CP.idfornecedor where CP.Pedido = '" & txtPedido & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFornecedor.EOF = False Then
    Set TBReceber = CreateObject("adodb.recordset")
    TBReceber.Open "Select ECR.*, CP.pedido from Estoque_Controle_recebimento ECR INNER JOIN Compras_pedido CP on ECR.idpedido = CP.idpedido where CP.idfornecedor = " & TBFornecedor!IDCliente & " and ECR.nota_fiscal = '" & txtnotafiscal & "' and ECR.Serie = '" & txtSerie & "' and ECR.Programacao = 'False' and ECR.id_empresa = " & txtID_empresa & " order by ECR.Desenho", Conexao, adOpenKeyset, adLockOptimistic
    If TBReceber.EOF = False Then
        Do While TBReceber.EOF = False
            Set TBPedido = CreateObject("adodb.recordset")
            TBPedido.Open "Select * from compras_pedido_lista where idlista = " & TBReceber!IDlista, Conexao, adOpenKeyset, adLockOptimistic
            If TBPedido.EOF = False Then
                If Desenho <> TBReceber!Desenho Or Desenho = TBReceber!Desenho And valor <> IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto) Or OrdemTexto <> IIf(IsNull(TBPedido!Ordem), 0, TBPedido!Ordem) Then
                    'Carrega valor unitario, primeiro verifica se o desenho bate na nota importada, se não verifica o cod de referencia e se ainda não achar puxa do pedido
                    Set TBHistProc = CreateObject("adodb.recordset")
                    TBHistProc.Open "Select EntvUnCom, EntcProd from ENT001DET where ENTNTFNumero = " & Format(txtnotafiscal) & " and EntNtfSerie = " & txtSerie & " and EntcProd = '" & TBReceber!Desenho & "'", Conexao_NFe, adOpenKeyset, adLockOptimistic
                    If TBHistProc.EOF = True Then
                        Set TBCorretiva = CreateObject("adodb.recordset")
                        TBCorretiva.Open "Select N_referencia from item_aplicacoes I INNER JOIN Projproduto P on I.codproduto = P.Codproduto where P.desenho = '" & TBReceber!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
                        Do While TBCorretiva.EOF = False
                            Set TBHistProc = CreateObject("adodb.recordset")
                            TBHistProc.Open "Select EntvUnCom, EntcProd from ENT001DET where ENTNTFNumero = " & Format(txtnotafiscal) & " and EntNtfSerie = " & txtSerie & " and EntcProd = '" & TBCorretiva!N_referencia & "'", Conexao_NFe, adOpenKeyset, adLockOptimistic
                            If TBHistProc.EOF = False Then
                                TBCorretiva.Close
                                GoTo Continua
                            End If
                            TBCorretiva.MoveNext
                        Loop
                        TBCorretiva.Close
                    End If
Continua:
                    If TBHistProc.EOF = False Then
                        DesenhoProduto = IIf(IsNull(TBHistProc!EntcProd), "", TBHistProc!EntcProd)
                        ValorTotal = IIf(IsNull(TBHistProc!EntvUnCom), 0, Format(TBHistProc!EntvUnCom, "###,##0.0000"))
                    Else
                        DesenhoProduto = IIf(IsNull(TBPedido!Desenho), "", TBPedido!Desenho)
                        ValorTotal = IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto)
                    End If
                    TBHistProc.Close
                    OF = IIf(IsNull(TBPedido!Ordem), 0, TBPedido!Ordem)
                    GoTo Prosseguir
                Else
                    GoTo Proximo
                End If
            End If
            TBPedido.Close
Prosseguir:
            If OF = 0 Then TextoFiltro = "(Ordem = 0 or Ordem is null)" Else TextoFiltro = "Ordem = '" & OF & "'"
            
            qt = 0
            NovoValor = Replace(ValorTotal, ",", ".")
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select Sum(ECR.recebido) as qt from compras_pedido_lista CPL INNER JOIN estoque_controle_recebimento ECR on CPL.idlista = ECR.idlista where CPL.preco_unitario_desconto = " & NovoValor & " and " & TextoFiltro & " and ECR.nota_fiscal = '" & txtnotafiscal & "' and ECR.Serie = '" & txtSerie & "' and ECR.Desenho = '" & TBReceber!Desenho & "' and ECR.id_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                qt = Format(IIf(IsNull(TBFI!qt), 0, TBFI!qt), "###,##0.0000")
            End If
            TBFI.Close
                        
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from tbl_Detalhes_Nota where int_Cod_Produto = '" & TBReceber!Desenho & "' and id_nota = " & IDlista & " and dbl_valorunitario = " & IIf(NovoValor = "", 0, NovoValor) & " and " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = True Then TBAbrir.AddNew
            TBAbrir!Tipo = "P"
            TBAbrir!int_Cod_Produto = TBReceber!Desenho
            TBAbrir!int_Qtd = qt
            TBAbrir!Saldo = qt
            TBAbrir!int_NotaFiscal = txtnotafiscal
            TBAbrir!ID_nota = IDlista
            
            Set TBPedido = CreateObject("adodb.recordset")
            TBPedido.Open "Select * from compras_pedido_lista where idpedido = " & TBReceber!IDpedido & " and IDLista = " & TBReceber!IDlista, Conexao, adOpenKeyset, adLockOptimistic
            If TBPedido.EOF = False Then
                TBAbrir!Txt_descricao = IIf(IsNull(TBPedido!Descricao), "", TBPedido!Descricao)
                TBAbrir!Codproduto = IIf(IsNull(TBPedido!Codproduto), "", TBPedido!Codproduto)
                'IDlista = IIf(IsNull(TBPedido!IDlista), "", TBPedido!IDlista)
                TBAbrir!txt_Unid = IIf(IsNull(TBPedido!Un), "", TBPedido!Un)
                TBAbrir!Unidade_com = IIf(IsNull(TBPedido!Unidade_com), "", TBPedido!Unidade_com)
                TBAbrir!Familia = IIf(IsNull(TBPedido!Familia), "", TBPedido!Familia)
                TBAbrir!N_referencia = IIf(IsNull(TBPedido!N_referencia), "", TBPedido!N_referencia)
                TBAbrir!Ordem = TBPedido!Ordem
                If TBPedido!Remessa = True Then TBAbrir!retorno = True
                        
                ProcCriaCFOP_NCM
                If IsNull(TBPedido!ID_CFOP) = False And TBPedido!ID_CFOP <> "" And IsNull(TBAbrir!ID_CFOP) = True Then TBAbrir!ID_CFOP = TBPedido!ID_CFOP
                If IsNull(TBPedido!ID_CF) = False And TBPedido!ID_CF <> "" And IsNull(TBAbrir!ID_CF) = True Then TBAbrir!ID_CF = TBPedido!ID_CF
                If IsNull(TBPedido!CST) = False And TBPedido!CST <> "" And IsNull(TBAbrir!txt_CST) = True Then TBAbrir!txt_CST = TBPedido!CST
            End If
            TBPedido.Close
            
            Set TBItem = CreateObject("adodb.recordset")
            TBItem.Open "Select ID_CFOP, ID_CF from projproduto where desenho = '" & TBReceber!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBItem.EOF = False Then
                If IsNull(TBAbrir!ID_CFOP) = True Or TBAbrir!ID_CFOP = "" Then TBAbrir!ID_CFOP = IIf(IsNull(TBItem!ID_CFOP), 0, TBItem!ID_CFOP)
                If IsNull(TBAbrir!ID_CF) = True Or TBAbrir!ID_CF = "" Then TBAbrir!ID_CF = IIf(IsNull(TBItem!ID_CF), 0, TBItem!ID_CF)
            End If
            
            If IsNull(TBAbrir!ID_CFOP) = False And TBAbrir!ID_CFOP <> "" Then
                Set TBItem = CreateObject("adodb.recordset")
                TBItem.Open "Select * from tbl_NaturezaOperacao_CST where ID_CFOP = " & TBAbrir!ID_CFOP, Conexao, adOpenKeyset, adLockOptimistic
                If TBItem.EOF = False Then
                    If TBItem.RecordCount = 1 Then
                        If IsNull(TBAbrir!txt_CST) = True Or TBAbrir!txt_CST = "" Then TBAbrir!txt_CST = TBItem!CST_ICMS
                        TBAbrir!CST_IPI = TBItem!CST_IPI
                        TBAbrir!CST_PIS = TBItem!CST_PIS
                        TBAbrir!CST_Cofins = TBItem!CST_Cofins
                    End If
                End If
            End If
            TBItem.Close
            
            Set TBPI_Lista_produto = CreateObject("adodb.recordset")
            TBPI_Lista_produto.Open "Select compras_pedido_lista.* from compras_pedido_lista inner join compras_pedido on compras_pedido_lista.idpedido = compras_pedido.idpedido where compras_pedido_lista.idlista = " & TBReceber!IDlista & " and compras_pedido.idpedido = " & TBReceber!IDpedido, Conexao, adOpenKeyset, adLockOptimistic
            If TBPI_Lista_produto.EOF = False Then
            
                'Carrega valor unitario, primeiro verifica se o desenho bate na nota importada, se não verifica o cod de referencia e se ainda não achar puxa do pedido
                Set TBHistProc = CreateObject("adodb.recordset")
                TBHistProc.Open "Select EntvUnCom, EntnItem, EntvProd, EntvFrete, EntvOutro_item, EntvSeg from ENT001DET where ENTNTFNumero = " & Format(txtnotafiscal) & " and EntNtfSerie = " & txtSerie & " and EntcProd = '" & DesenhoProduto & "'", Conexao_NFe, adOpenKeyset, adLockOptimistic
                If TBHistProc.EOF = False Then
                    TBAbrir!dbl_ValorUnitario = TBHistProc!EntvUnCom
                    
                    Set TBCodigoDesc = CreateObject("adodb.recordset")
                    TBCodigoDesc.Open "Select EntpICMS from ENT001DETICMSNORMALST where EntNtfNumero = " & Format(txtnotafiscal) & " and EntNtfSerie = " & txtSerie & " and EntnItem = " & TBHistProc!EntnItem, Conexao_NFe, adOpenKeyset, adLockOptimistic
                    If TBCodigoDesc.EOF = False Then
                        TBAbrir!int_ICMS = IIf(IsNull(TBCodigoDesc!EntpICMS), 0, TBCodigoDesc!EntpICMS)
                    End If

                    Set TBCodigoDesc = CreateObject("adodb.recordset")
                    TBCodigoDesc.Open "Select EntpIPI, EntvIPI from ENT001DETIPI where EntNtfNumero = " & Format(txtnotafiscal) & " and EntNtfSerie = " & txtSerie & " and EntnItem = " & TBHistProc!EntnItem, Conexao_NFe, adOpenKeyset, adLockOptimistic
                    If TBCodigoDesc.EOF = False Then
                        TBAbrir!int_IPI = IIf(IsNull(TBCodigoDesc!EntpIPI), 0, TBCodigoDesc!EntpIPI)
                        TBAbrir!dbl_valoripi = Format(IIf(IsNull(TBCodigoDesc!EntvIPI), 0, TBCodigoDesc!EntvIPI), "###,##0.00")
                    End If
                    TBCodigoDesc.Close
                    
                    TBAbrir!dbl_ValorTotal = Format(IIf(IsNull(TBHistProc!EntvProd), 0, TBHistProc!EntvProd), "###,##0.00")
                    TBAbrir!Valor_frete = Format(IIf(IsNull(TBHistProc!EntvFrete), 0, TBHistProc!EntvFrete), "###,##0.00")
                    TBAbrir!Valor_seguro = Format(IIf(IsNull(TBHistProc!EntvSeg), 0, TBHistProc!EntvSeg), "###,##0.00")
                    TBAbrir!Valor_acessorias = Format(IIf(IsNull(TBHistProc!EntvOutro_item), 0, TBHistProc!EntvOutro_item), "###,##0.00")
                Else
                    TBAbrir!dbl_ValorUnitario = TBPI_Lista_produto!preco_unitario_desconto
                    TBAbrir!int_ICMS = IIf(IsNull(TBPI_Lista_produto!ICMS), 0, TBPI_Lista_produto!ICMS)
                    TBAbrir!int_IPI = IIf(IsNull(TBPI_Lista_produto!IPI), 0, TBPI_Lista_produto!IPI)
                    TBAbrir!dbl_valoripi = Format(((IIf(IsNull(TBPI_Lista_produto!preco_unitario_desconto), 0, TBPI_Lista_produto!preco_unitario_desconto) * qt) * IIf(IsNull(TBPI_Lista_produto!IPI), 0, TBPI_Lista_produto!IPI)) / 100, "###,##0.00")
                    TBAbrir!dbl_ValorTotal = Format(IIf(IsNull(TBPI_Lista_produto!preco_unitario_desconto), 0, TBPI_Lista_produto!preco_unitario_desconto) * qt, "###,##0.00")
                    TBAbrir!Valor_frete = Format(IIf(IsNull(TBPI_Lista_produto!Frete), 0, TBPI_Lista_produto!Frete), "###,##0.00")
                    TBAbrir!Valor_seguro = Format(IIf(IsNull(TBPI_Lista_produto!Seguro), 0, TBPI_Lista_produto!Seguro), "###,##0.00")
                    TBAbrir!Valor_acessorias = Format(IIf(IsNull(TBPI_Lista_produto!Acessorias), 0, TBPI_Lista_produto!Acessorias), "###,##0.00")
                    TBAbrir!Tem_IPI_frete = TBPI_Lista_produto!Frete_IPI
                End If
                TBHistProc.Close
                If IsNull(TBPI_Lista_produto!OS) = False And TBPI_Lista_produto!OS <> "" Then ProcAtualizaCTTEROrdem TBPI_Lista_produto!OS
            End If
            TBPI_Lista_produto.Close
            TBAbrir.Update

            ProcAtualizaCST

            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * FROM tbl_Detalhes_Nota_pedidos where ID_nota = " & IDlista & " and ID_prod_NF = " & TBAbrir!Int_codigo & " and ID_carteira = " & IDlista & " and Codinterno = '" & TBAbrir!int_Cod_Produto & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = True Then TBGravar.AddNew
            TBGravar!ID_nota = IDlista
            TBGravar!ID_prod_NF = TBAbrir!Int_codigo
            TBGravar!ID_carteira = IDlista
            TBGravar!Codinterno = TBAbrir!int_Cod_Produto
            TBGravar!quantidade = TBAbrir!int_Qtd
            TBGravar.Update
            TBGravar.Close
            
            TBAbrir.Close
Proximo:
            Set TBPedido = CreateObject("adodb.recordset")
            TBPedido.Open "Select * from tbl_proposta_nota where id_nota = " & IDlista & " and proposta = '" & TBReceber!Pedido & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBPedido.EOF = True Then
                TBPedido.AddNew
                TBPedido!Proposta = TBReceber!Pedido
                TBPedido!NF = txtnotafiscal
                TBPedido!ID_nota = IDlista
                TBPedido.Update
            End If

            Set TBPedido = CreateObject("adodb.recordset")
            TBPedido.Open "Select * from compras_pedido_lista where idlista = " & TBReceber!IDlista, Conexao, adOpenKeyset, adLockOptimistic
            If TBPedido.EOF = False Then
                valor = IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto)
                OrdemTexto = IIf(IsNull(TBPedido!Ordem), "", TBPedido!Ordem)
            End If
            TBPedido.Close
            Desenho = TBReceber!Desenho
            TBReceber.MoveNext
        Loop
    End If
End If
TBFornecedor.Close

ProcGravarTotaisNota

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAlteraStatus_pedido(IDlista_AlteraStatus As Long, qtde_AlteraStatus As Double, qtde_PC_AlteraStatus As Double)
On Error GoTo tratar_erro

Set TBCompras = CreateObject("adodb.recordset")
TBCompras.Open "Select * from compras_pedido where pedido = '" & txtPedido.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBCompras.EOF = False Then
    IDpedido = TBCompras!IDpedido

    quantnovo = 0
    Valor_Cofins_Prod = 0
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select Sum(Recebido) as quantnovo, Sum(ISNULL(Recebido_PC, 0)) as Valor_Cofins_Prod from estoque_controle_recebimento where idpedido = " & IDpedido & " and idlista = " & IDlista_AlteraStatus & " and Programacao = 'False' and id_empresa = " & txtID_empresa, Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
        quantnovo = IIf(IsNull(TBEstoque!quantnovo), 0, TBEstoque!quantnovo)
        Valor_Cofins_Prod = IIf(IsNull(TBEstoque!Valor_Cofins_Prod), 0, TBEstoque!Valor_Cofins_Prod)
    End If
    TBEstoque.Close
    
    If Valor_Cofins_Prod > 0 Then
        quantnovo = Valor_Cofins_Prod
        quantestoque = qtde_PC_AlteraStatus
    Else
        quantestoque = qtde_AlteraStatus
    End If
    
    If quantnovo < quantestoque Then
        If USMsgBox("Este produto/serviço será recebido parcialmente, deseja encerrar o mesmo no pedido de compra?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            Status_Item = "RECEBIDO"
        Else
            Status_Item = "PARCIAL"
        End If
    End If
    If quantnovo >= quantestoque Then Status_Item = "RECEBIDO"
    Conexao.Execute "Update compras_pedido_lista Set Status_item = '" & Status_Item & "' where idpedido = " & IDpedido & " and idlista = " & IDlista_AlteraStatus
    Conexao.Execute "Update Compras_Programacao set Compras_Programacao.Status_prog = '" & Status_Item & "' from Compras_Programacao INNER JOIN compras_pedido_lista ON Compras_Programacao.ID_prog = compras_pedido_lista.ID_programacao where compras_pedido_lista.idpedido = " & IDpedido & " and compras_pedido_lista.idlista = " & IDlista_AlteraStatus
    
    'Verifica status do item
    Set TBCompras = CreateObject("adodb.recordset")
    TBCompras.Open "Select * from compras_pedido_lista where idpedido = " & IDpedido & " and status_item <> 'RECEBIDO' and status_item <> 'CANCELADO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCompras.EOF = True Then
        Status_pedido = "ENCERRADO"
    Else
        Status_pedido = "PARCIAL"
    End If
    TBCompras.Close
    
    'Pedido de compra
    Conexao.Execute "Update compras_pedido Set Status_pedido = '" & Status_pedido & "' where pedido = '" & txtPedido.Text & "'"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select IDlista, Desenho, Descricao, Un, Quant_Comp, Status_Item from compras_pedido_lista where IDpedido = " & IDpedido & " order by Desenho", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!IDlista
            .Item(.Count).SubItems(1) = TBLISTA!Desenho
            .Item(.Count).SubItems(2) = TBLISTA!Descricao
            .Item(.Count).SubItems(3) = TBLISTA!Un
            .Item(.Count).SubItems(4) = Format(TBLISTA!Quant_Comp, "###,##0.0000")
            .Item(.Count).SubItems(5) = IIf(TBLISTA!Status_Item = "N_RECEBIDO", "NÃO RECEBIDO", TBLISTA!Status_Item)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNotaFiscal_Change()
On Error GoTo tratar_erro
    
If txtnotafiscal.Text <> "" Then
    VerifNumero = txtnotafiscal.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtnotafiscal.Text = ""
        txtnotafiscal.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNotaFiscal_LostFocus()
On Error GoTo tratar_erro

If txtnotafiscal <> "" Then txtnotafiscal = FunTamanhoTextoZeroEsq(ReturnNumbersOnly(txtnotafiscal), 9)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtpedido_Change()
On Error GoTo tratar_erro

IDpedido = 0
txtIDforn = ""
txtFornecedor = ""
txtEmpresa = ""
Lista.ListItems.Clear

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: procImportar
    'Case 3: ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarCCRealizado(Data1 As Date, ID_empresa As Integer, Operacao As String, ID_CC As Long, Cod_produto As Long, ID_plano_contas As Long, ID_estoque As Long, ID_lista As Long, valor As Double, CC_produto As Boolean, Bloqueado As Boolean)
On Error GoTo tratar_erro

NovoValor = Replace(valor, ",", ".")
ProcINSERTINTO "CC_realizado", "Data, Responsavel, ID_empresa, Operacao, ID_CC, Cod_produto, ID_PC, ID_estoque, ID_lista, Valor, Bloqueado", "'" & Data & "', '" & pubUsuario & "', " & ID_empresa & ", '" & Operacao & "', " & ID_CC & ", " & Cod_produto & ", " & ID_plano_contas & ", " & IIf(ID_estoque = 0, "NULL", ID_estoque) & ", " & ID_lista & ", " & NovoValor & ", " & IIf(Bloqueado = True, 1, 0) & ""

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select ID from CC_realizado where ID_estoque = " & ID_estoque, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    If CC_produto = True And Operacao = "Crédito" Then Conexao.Execute "Update CC_realizado Set ID_ref_debito = " & TBGravar!ID - 1 & " where ID = " & TBGravar!ID
End If
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSalvarRealCCConsolidado(ID_CC As Long, Operacao As String, Credito As Boolean, CC_produto As Boolean, Bloqueado As Boolean)
On Error GoTo tratar_erro

Set TBAfericao = CreateObject("adodb.recordset")
TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & ID_CC, Conexao, adOpenKeyset, adLockOptimistic
If TBAfericao.EOF = False Then
    Do While TBAfericao.EOF = False
        ProcSalvarCCRealizado Date, txtID_empresa, Operacao, TBAfericao!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, TBCompras_Lista!IDlista, valor, CC_produto, Bloqueado
        
        Set TBCiclo = CreateObject("adodb.recordset")
        TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
        If TBCiclo.EOF = False Then
            Do While TBCiclo.EOF = False
                ProcSalvarCCRealizado Date, txtID_empresa, Operacao, TBCiclo!ID_CC, IDFase, IDPlano, TBEstoque!IDoperacao, TBCompras_Lista!IDlista, valor, CC_produto, Bloqueado
                TBCiclo.MoveNext
            Loop
        End If
        TBCiclo.Close
        
        TBAfericao.MoveNext
    Loop
End If
TBAfericao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCriaCFOP_NCM()
On Error GoTo tratar_erro

Set TBHistProc = CreateObject("adodb.recordset")
TBHistProc.Open "Select EntvProd, EntCFOP, EntNCM, EntnItem from ENT001DET where ENTNTFNumero = " & Format(txtnotafiscal) & " and EntNtfSerie = " & txtSerie & " and EntcProd = '" & DesenhoProduto & "'", Conexao_NFe, adOpenKeyset, adLockOptimistic
If TBHistProc.EOF = False Then
    TextoCFOP = "1." & Right(TBHistProc!EntCFOP, 3)
    TextoNCM = Left(TBHistProc!EntNCM, 4) & "." & Right(Left(TBHistProc!EntNCM, 6), 2) & "." & Right(TBHistProc!EntNCM, 2)
    
    'Puxa CST e aliquota de ICMS
    TemICMS_CFOP = "NÃO"
    CSTICMS = ""
    AliquotaICMS = 0
    Set TBCodigoDesc = CreateObject("adodb.recordset")
    TBCodigoDesc.Open "Select EntCST, EntpICMS from ENT001DETICMSNORMALST where EntNtfNumero = " & Format(txtnotafiscal) & " and EntNtfSerie = " & txtSerie & " and EntnItem = " & TBHistProc!EntnItem, Conexao_NFe, adOpenKeyset, adLockOptimistic
    If TBCodigoDesc.EOF = False Then
        TemICMS_CFOP = "SIM"
        AliquotaICMS = Replace(TBCodigoDesc!EntpICMS, ",", ".")
        If Len(TBCodigoDesc!EntCST) = 1 Then
            CSTICMS = "00" & TBCodigoDesc!EntCST
        ElseIf Len(TBCodigoDesc!EntCST) = 2 Then
            CSTICMS = "0" & TBCodigoDesc!EntCST
        Else
            CSTICMS = TBCodigoDesc!EntCST
        End If
    End If
    If CSTICMS <> "" Then TBAbrir!txt_CST = CSTICMS 'Salva CST de ICMS

    'Puxa aliquota de IPI
    TemIPI_CFOP = "NÃO"
    AliquotaIPI = 0
    Set TBCodigoDesc = CreateObject("adodb.recordset")
    TBCodigoDesc.Open "Select EntpIPI, EntCST_IPI from ENT001DETIPI where EntNtfNumero = " & txtnotafiscal & " and EntNtfSerie = " & txtSerie & " and EntnItem = " & TBHistProc!EntnItem, Conexao_NFe, adOpenKeyset, adLockOptimistic
    If TBCodigoDesc.EOF = False Then
        TemIPI_CFOP = "SIM"
        AliquotaIPI = Replace(TBCodigoDesc!EntpIPI, ",", ".")
    End If

    'Puxa aliquota de PIS
    TemPIS_CFOP = "False"
    AliquotaPIS = 0
    Set TBCodigoDesc = CreateObject("adodb.recordset")
    TBCodigoDesc.Open "Select EntpPIS, EntCST_pis from ENT001DETPIS where EntNtfNumero = " & Format(txtnotafiscal) & " and EntNtfSerie = " & txtSerie & " and EntnItem = " & TBHistProc!EntnItem, Conexao_NFe, adOpenKeyset, adLockOptimistic
    If TBCodigoDesc.EOF = False Then
        TemPIS_CFOP = "True"
        AliqutaCFOP = Replace(TBCodigoDesc!EntpPIS, ",", ".")
    End If
    
    'Puxa aliquota de Cofins
    TemCOFINS_CFOP = "False"
    AliquotaCofins = 0
    Set TBCodigoDesc = CreateObject("adodb.recordset")
    TBCodigoDesc.Open "Select EntpCOFINS, EntCST_cofins from ENT001DETCOFINS where EntNtfNumero = " & Format(txtnotafiscal) & " and EntNtfSerie = " & txtSerie & " and EntnItem = " & TBHistProc!EntnItem, Conexao_NFe, adOpenKeyset, adLockOptimistic
    If TBCodigoDesc.EOF = False Then
        TemCOFINS_CFOP = "True"
        AliquotaCofins = Replace(TBCodigoDesc!EntpCofins, ",", ".")
    End If
    
    'Puxa CFOP
    Set TBCodigoDesc = CreateObject("adodb.recordset")
    TBCodigoDesc.Open "Select N.IDCountCfop from tbl_NaturezaOperacao N INNER JOIN tbl_NaturezaOperacao_CST C ON N.IDCountCfop = C.ID_CFOP where N.ID_CFOP = '" & TextoCFOP & "' and N.txt_ICMS = '" & TemICMS_CFOP & "' and N.txt_IPI = '" & TemIPI_CFOP & "' and N.TemPIS = '" & TemPIS_CFOP & "' and N.TemCofins = '" & TemCOFINS_CFOP & "' and C.CST_ICMS = '" & CSTICMS & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCodigoDesc.EOF = True Then
        Set TBCodigoDesc = CreateObject("adodb.recordset")
        TBCodigoDesc.Open "Select IDCountCfop from tbl_NaturezaOperacao where ID_CFOP = '" & TextoCFOP & "'", Conexao, adOpenKeyset, adLockOptimistic
    End If
    If TBCodigoDesc.EOF = False Then
        TBAbrir!ID_CFOP = TBCodigoDesc!IDCountCfop 'Salva CFOP
    End If
    
    'Puxa CST de IPI, PIS e Cofins
    Set TBCodigoDesc = CreateObject("adodb.recordset")
    TBCodigoDesc.Open "Select C.CST_IPI, C.CST_PIS, C.CST_Cofins from tbl_NaturezaOperacao N INNER JOIN tbl_NaturezaOperacao_CST C ON N.IDCountCfop = C.ID_CFOP where N.ID_CFOP = '" & TextoCFOP & "' and C.CST_ICMS = '" & CSTICMS & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCodigoDesc.EOF = False Then
        TBAbrir!CST_IPI = TBCodigoDesc!CST_IPI 'Salva CST de IPI
        TBAbrir!CST_PIS = TBCodigoDesc!CST_PIS 'Salva CST de PIS
        TBAbrir!CST_Cofins = TBCodigoDesc!CST_Cofins 'Salva CST de Cofins
    End If
    
    'Verifica região do fornecedor
    Set TBCodigoDesc = CreateObject("adodb.recordset")
    TBCodigoDesc.Open "Select regiao from regioes where UF = '" & TBFornecedor!Estado & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCodigoDesc.EOF = False Then
        Select Case TBCodigoDesc!regiao
            Case "SS": TextoFiltroRegiao = "dbl_ICMS_SS"
            Case "NN": TextoFiltroRegiao = "dbl_ICMS_NN"
            Case "CO": TextoFiltroRegiao = "dbl_ICMS_CO"
            Case "DE": TextoFiltroRegiao = "dbl_ICMS_DE"
        End Select
    End If
    
    'Puxa NCM
    Set TBCodigoDesc = CreateObject("adodb.recordset")
    TBCodigoDesc.Open "Select IDclass from tbl_ClassificacaoFiscal where IDIntClasse = '" & TextoNCM & "' and " & TextoFiltroRegiao & " = " & AliquotaICMS & " and dbl_IPI = " & AliquotaIPI & " and PIS = " & AliquotaPIS & " and Cofins = " & AliquotaCofins, Conexao, adOpenKeyset, adLockOptimistic
    If TBCodigoDesc.EOF = True Then
        Set TBCodigoDesc = CreateObject("adodb.recordset")
        TBCodigoDesc.Open "Select IDclass from tbl_ClassificacaoFiscal where IDIntClasse = '" & TextoNCM & "'", Conexao, adOpenKeyset, adLockOptimistic
    End If
    If TBCodigoDesc.EOF = False Then
        TBAbrir!ID_CF = TBCodigoDesc!Idclass 'Salva NCM
    End If
    TBCodigoDesc.Close
End If
TBHistProc.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcGravarTotaisNota()
On Error GoTo tratar_erro

TotalProduto = 0 'Valor total produtos
TotalServicos = 0 'Valor total serviços
qtdeliberar = 0 'Qtde. produtos
qtdeliberada = 0 'Qtde de serviços
ValorPorc = 0 'VlrMPA
VlrICMS_suframa = 0 'Vlr Suframa
TotalIPI = 0 'Total IPI
BASECALCULO = 0 'Base calculo ICMS
Total_ICMS = 0 'Vlr total ICMS
TotalBCICMSCST = 0 'Base calculo ICMS ST
TotalICMSCST = 0 'Vlr total ICMS ST
Frete = 0 'Frete
Seguro = 0 'Seguro
Acessorias = 0 'Outras despesas
QuantsolicitadoN2 = 0 'Desconto
Valor_PIS_Prod = 0 'Pis produto
Valor_Cofins_Prod = 0 'Cofins produto
Valor_CSLL_Prod = 0 'CSLL produto
Valor_IRPJ_Prod = 0 'IRPJ produto
Valor_PIS_Serv = 0 'PIS serv
Valor_Cofins_Serv = 0 'Cofins serv
Valor_CSLL_Serv = 0 'CSLL serv
TotalISS = 0 'ISS
Valor_INSS_Serv = 0 'INSS
Valor_IRPJ_Serv = 0 'IRPJ
Valor_IRRF_Serv = 0 'IRRF
Valor_DAS = 0 'DAS
Valor_Retencao_PIS = 0 'Vlr retenção pis
Valor_Retencao_Cofins = 0 'Vlr retenção cofins
VlrTotalRetorno = 0
VlrTotalRetornoNF = 0
Valores = 0 'valor unitario
Vlr_total_aprox_tributos_prod = 0

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from tbl_Dados_Nota_Fiscal where ID = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_dadosadicionais where id_nota = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then
        TBAbrir.AddNew
        TBAbrir!mem_DadosAdicionais = Null
    End If
    TBAbrir!IDNotaFiscal = txtnotafiscal
    TBAbrir!ID_nota = IDlista
    TBAbrir.Update
    TBAbrir.Close
    
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select NFP.*, CST.Valor_BC, CST.Valor_ICMS, CST.Valor_BC_ST, CST.Valor_ICMS_ST from tbl_Detalhes_Nota NFP LEFT JOIN tbl_Detalhes_Nota_CST_ICMS CST ON CST.ID_item = NFP.Int_codigo where NFP.id_nota = " & IDlista & " and NFP.Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
    Do While TBProduto.EOF = False
        Valores = IIf(IsNull(TBProduto!dbl_ValorUnitario), 0, TBProduto!dbl_ValorUnitario) * TBProduto!int_Qtd
        
        If TBProduto!retorno = False Or (TBProduto!retorno = True And TBProduto!Soma_retorno_totalnf = True) Then
            If TBProduto!Tipo = "P" Then
                TotalProduto = TotalProduto + Format(IIf(IsNull(TBProduto!dbl_ValorUnitario), 0, TBProduto!dbl_ValorUnitario) * TBProduto!int_Qtd, "###,##0.00")
                qtdeliberar = qtdeliberar + IIf(IsNull(TBProduto!int_Qtd), 0, TBProduto!int_Qtd)
            Else
                TotalServicos = TotalServicos + Format(IIf(IsNull(TBProduto!dbl_ValorUnitario), 0, TBProduto!dbl_ValorUnitario) * TBProduto!int_Qtd, "###,##0.00")
                qtdeliberada = qtdeliberada + IIf(IsNull(TBProduto!int_Qtd), 0, TBProduto!int_Qtd)
            End If
            
            Frete = IIf(IsNull(TBProduto!Valor_frete), 0, TBProduto!Valor_frete)
            Seguro = IIf(IsNull(TBProduto!Valor_seguro), 0, TBProduto!Valor_seguro)
            Acessorias = IIf(IsNull(TBProduto!Valor_acessorias), 0, TBProduto!Valor_acessorias)
            QuantsolicitadoN2 = IIf(IsNull(TBProduto!Valor_desconto), 0, TBProduto!Valor_desconto) + IIf(IsNull(TBProduto!Valor_desconto_SUFRAMA), 0, TBProduto!Valor_desconto_SUFRAMA)
            
            Valor_total = Valores + IIf(IsNull(TBProduto!Valor_frete), 0, TBProduto!Valor_frete) + IIf(IsNull(TBProduto!Valor_seguro), 0, TBProduto!Valor_seguro) + IIf(IsNull(TBProduto!Valor_acessorias), 0, TBProduto!Valor_acessorias)
            ProcVerifImpostosEmpresa txtID_empresa, TBProduto!retorno, "", False, 0, False, IIf(IsNull(TBGravar!TabelaSN), 0, TBGravar!TabelaSN), 0
            'Impostos faturamento
            If DAS <> 0 Then
                NovoValor = Replace(DAS, ",", ".")
                Conexao.Execute "UPDATE tbl_Totais_Nota Set DAS = " & NovoValor & " where ID_nota = " & IDlista
                Valor_DAS = Valor_DAS + Format(((Valor_total * DAS) / 100), "###,##0.00")
            End If
            
            'Impostos
            TotalIPI = TotalIPI + IIf(IsNull(TBProduto!dbl_valoripi), 0, TBProduto!dbl_valoripi)
            BASECALCULO = BASECALCULO + IIf(IsNull(TBProduto!Valor_BC), 0, TBProduto!Valor_BC)
            Total_ICMS = Total_ICMS + IIf(IsNull(TBProduto!Valor_ICMS), 0, TBProduto!Valor_ICMS)
            TotalBCICMSCST = TotalBCICMSCST + IIf(IsNull(TBProduto!Valor_BC_ST), 0, TBProduto!Valor_BC_ST)
            TotalICMSCST = TotalICMSCST + IIf(IsNull(TBProduto!Valor_ICMS_ST), 0, TBProduto!Valor_ICMS_ST)
            
            'Impostos produtos
            If IsNull(TBProduto!Total_PIS_prod) = False Then Valor_PIS_Prod = Valor_PIS_Prod + TBProduto!Total_PIS_prod
            If IsNull(TBProduto!Total_Cofins_prod) = False Then Valor_Cofins_Prod = Valor_Cofins_Prod + TBProduto!Total_Cofins_prod
            If IsNull(TBProduto!Total_CSLL_prod) = False Then Valor_CSLL_Prod = Valor_CSLL_Prod + TBProduto!Total_CSLL_prod
            If IsNull(TBProduto!Total_IRPJ_prod) = False Then Valor_IRPJ_Prod = Valor_IRPJ_Prod + TBProduto!Total_IRPJ_prod
            
            'Impostos serviços
            If IsNull(TBProduto!Total_PIS_serv) = False Then Valor_PIS_Serv = Valor_PIS_Serv + TBProduto!Total_PIS_serv
            If IsNull(TBProduto!Total_Cofins_serv) = False Then Valor_Cofins_Serv = Valor_Cofins_Serv + TBProduto!Total_Cofins_serv
            If IsNull(TBProduto!Total_CSLL_serv) = False Then Valor_CSLL_Serv = Valor_CSLL_Serv + TBProduto!Total_CSLL_serv
            TotalISS = TotalISS + IIf(IsNull(TBProduto!VlrISS), 0, TBProduto!VlrISS)
            If IsNull(TBProduto!Total_INSS_serv) = False Then Valor_INSS_Serv = Valor_INSS_Serv + TBProduto!Total_INSS_serv
            If IsNull(TBProduto!Total_IRPJ_serv) = False Then Valor_IRPJ_Serv = Valor_IRPJ_Serv + TBProduto!Total_IRPJ_serv
            If IsNull(TBProduto!Total_IRRF_serv) = False Then Valor_IRRF_Serv = Valor_IRRF_Serv + TBProduto!Total_IRRF_serv
            
            'Soma retenção de PIS/Cofins
            Valor_Retencao_PIS = Valor_Retencao_PIS + IIf(IsNull(TBProduto!Valor_Retencao_PIS), 0, TBProduto!Valor_Retencao_PIS)
            Valor_Retencao_Cofins = Valor_Retencao_Cofins + IIf(IsNull(TBProduto!Valor_Retencao_Cofins), 0, TBProduto!Valor_Retencao_Cofins)
                        
            'Soma o ICMS do suframa
            VlrICMS_suframa = VlrICMS_suframa + IIf(IsNull(TBProduto!Valor_desconto_SUFRAMA), 0, TBProduto!Valor_desconto_SUFRAMA)
                
            'If NF_enviada = False And NFe_liberada = False Then ProcReacalculaICMSIPI
            'If TBProduto!Remessa = False Then qtdeliberar = qtdeliberar + IIf(IsNull(TBProduto!int_Qtd), 0, TBProduto!int_Qtd)
            ValorPorc = ValorPorc + IIf(IsNull(TBProduto!VlrMPA), 0, TBProduto!VlrMPA)
        Else
            VlrTotalRetorno = VlrTotalRetorno + Valores
            If TBProduto!Soma_retorno_totalnf = False Then VlrTotalRetornoNF = VlrTotalRetornoNF + Valores
        End If
        Vlr_total_aprox_tributos_prod = Vlr_total_aprox_tributos_prod + IIf(IsNull(TBProduto!Valor_aprox_tributos), 0, TBProduto!Valor_aprox_tributos)
        TBProduto.MoveNext
    Loop
    TBProduto.Close
    
    'Criar totalização da nota
    Set TBTotaisnota = CreateObject("adodb.recordset")
    TBTotaisnota.Open "Select * from tbl_totais_nota where id_Nota = " & IDlista, Conexao, adOpenKeyset, adLockOptimistic
    If TBTotaisnota.EOF = True Then TBTotaisnota.AddNew
    TBTotaisnota!int_NotaFiscal = txtnotafiscal
    TBTotaisnota!ID_nota = IDlista
    TBTotaisnota!dbl_Valor_Total_Produtos = Format(TotalProduto, "###,##0.00")
    TBTotaisnota!dbl_Valor_Total_Nota_Serv = Format(TotalServicos, "###,##0.00")
    TBTotaisnota!Qtde_total_serv = qtdeliberada
    TBTotaisnota!Qtde_total_prod = qtdeliberar
    TBTotaisnota!VlrMPA = ValorPorc
    TBTotaisnota!Valor_total_desconto_SUFRAMA = Format(VlrICMS_suframa, "###,##0.00")
    TBTotaisnota!dbl_Base_ICMS = Format(BASECALCULO, "###,##0.00")
    TBTotaisnota!dbl_Valor_ICMS = Format(Total_ICMS, "###,##0.00")
    TBTotaisnota!dbl_Base_ICMS_Subst = Format(TotalBCICMSCST, "###,##0.00")
    TBTotaisnota!dbl_Valor_ICMS_Subst = Format(TotalICMSCST, "###,##0.00")
    TBTotaisnota!dbl_Valor_Frete = Format(Frete, "###,##0.00")
    TBTotaisnota!dbl_Valor_Seguro = Format(Seguro, "###,##0.00")
    TBTotaisnota!dbl_Desp_Adicionais = Format(Acessorias, "###,##0.00")
    TBTotaisnota!dbl_Valor_Total_IPI = Format(TotalIPI, "###,##0.00")
    TBTotaisnota!Valor_total_desconto = Format(QuantsolicitadoN2, "###,##0.00")
    
    'Impostos produtos
    TBTotaisnota!Total_PIS_prod = Format(Valor_PIS_Prod, "###,##0.00")
    TBTotaisnota!Total_Cofins_prod = Format(Valor_Cofins_Prod, "###,##0.00")
    TBTotaisnota!Total_CSLL_prod = Format(Valor_CSLL_Prod, "###,##0.00")
    TBTotaisnota!Total_IRPJ_prod = Format(Valor_IRPJ_Prod, "###,##0.00")
    
    'Impostos serviços
    TBTotaisnota!Total_PIS_serv = Format(Valor_PIS_Serv, "###,##0.00")
    TBTotaisnota!Total_Cofins_serv = Format(Valor_Cofins_Serv, "###,##0.00")
    TBTotaisnota!Total_CSLL_serv = Format(Valor_CSLL_Serv, "###,##0.00")
    TBTotaisnota!dbl_valor_total_iss = Format(TotalISS, "###,##0.00")
    TBTotaisnota!Total_INSS_serv = Format(Valor_INSS_Serv, "###,##0.00")
    TBTotaisnota!Total_IRPJ_serv = Format(Valor_IRPJ_Serv, "###,##0.00")
    TBTotaisnota!Total_IRRF_serv = Format(Valor_IRRF_Serv, "###,##0.00")
    
    Valor_total = Format((TotalProduto + Frete + Seguro + Acessorias + TotalIPI + TotalICMSCST + TotalServicos) - QuantsolicitadoN2 - VlrICMS_suframa, "###,##0.00")
    
    'Impostos faturamento
    TBTotaisnota!Total_DAS = Format(Valor_DAS, "###,##0.00")
    
    'Retenção de PIS/Cofins
    TBTotaisnota!Total_retencao_PIS = Format(Valor_Retencao_PIS, "###,##0.00")
    TBTotaisnota!Total_retencao_Cofins = Format(Valor_Retencao_Cofins, "###,##0.00")
    
    'Crédito de ICMS
    TBTotaisnota!Total_Credito_ICMS = Total_ICMS
    
    If Valor_total <> 0 Then
        TBTotaisnota!dbl_Valor_Total_Nota = Format(Valor_total, "###,##0.00")
        
        'Valor total de ICMS simples nacional
        TBTotaisnota!Valor_total_ICMS_SN = 0
        
        'Atualiza serie e valor total da nf na tabela Estoque_controle_recebimento
        ValoresParcelas = Valor_total
        NovoValor = Replace(ValoresParcelas, ",", ".")
        Conexao.Execute "Update Estoque_controle_recebimento Set Serie = '" & txtSerie & "', Total_Nf = " & NovoValor & " where Nota_fiscal = '" & txtnotafiscal & "'"
    End If

    TBTotaisnota!Valor_total_Extenso = FunValorExtenso(TBTotaisnota!dbl_Valor_Total_Nota)
    TBTotaisnota!Total_retorno = Format(VlrTotalRetorno, "###,##0.00")
    TBTotaisnota!Valor_total_aprox_tributos = Format(Vlr_total_aprox_tributos_prod, "###,##0.00")
    
    TBTotaisnota.Update
    TBTotaisnota.Close
End If
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAtualizaCST()
On Error GoTo tratar_erro

CST_ICMS = False
CST_IPI = False
CST_PIS = False
CST_Cofins = False

'ICMS
If IsNull(TBAbrir!txt_CST) = False And TBAbrir!txt_CST <> "" Then
    InicioCST = Left(TBAbrir!txt_CST, 1)
    If Len(TBAbrir!txt_CST) = 4 Then FimCST = Right(TBAbrir!txt_CST, 3) Else FimCST = Right(TBAbrir!txt_CST, 2)
    
    CST_ICMS = True
    CST_Cofins = False
    CST_IPI = False
    CST_PIS = False
    
    Set TBCST = CreateObject("adodb.recordset")
    TBCST.Open "select * from tbl_Detalhes_Nota_CST_ICMS where id_item = " & TBAbrir!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
    ProcEnviadadosAtualizaCST
    TBCST.Close
End If

'IPI
If IsNull(TBAbrir!CST_IPI) = False And TBAbrir!CST_IPI <> "" Then
    FimCST = TBAbrir!CST_IPI
    CST_ICMS = False
    CST_Cofins = False
    CST_IPI = True
    CST_PIS = False
    
    Set TBCST = CreateObject("adodb.recordset")
    TBCST.Open "select * from tbl_Detalhes_Nota_CST_IPI where id_item = " & TBAbrir!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
    ProcEnviadadosAtualizaCST
    TBCST.Close
End If

'PIS
If IsNull(TBAbrir!CST_PIS) = False And TBAbrir!CST_PIS <> "" Then
    FimCST = TBAbrir!CST_PIS
    CST_ICMS = False
    CST_Cofins = False
    CST_IPI = False
    CST_PIS = True
    
    Set TBCST = CreateObject("adodb.recordset")
    TBCST.Open "select * from tbl_Detalhes_Nota_CST_PIS where id_item = " & TBAbrir!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
    ProcEnviadadosAtualizaCST
    TBCST.Close
End If

'Cofins
If IsNull(TBAbrir!CST_Cofins) = False And TBAbrir!CST_Cofins <> "" Then
    FimCST = TBAbrir!CST_Cofins
    CST_ICMS = False
    CST_Cofins = True
    CST_IPI = False
    CST_PIS = False
    
    Set TBCST = CreateObject("adodb.recordset")
    TBCST.Open "select * from tbl_Detalhes_Nota_CST_Cofins where id_item = " & TBAbrir!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
    ProcEnviadadosAtualizaCST
    TBCST.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviadadosAtualizaCST()
On Error GoTo tratar_erro

Set TBNivel15 = CreateObject("adodb.recordset")
TBNivel15.Open "Select Estado from Compras_fornecedores where IdCliente = " & txtIDforn, Conexao, adOpenKeyset, adLockOptimistic
If TBNivel15.EOF = False Then
    If TBCST.EOF = True Then TBCST.AddNew
    'ICMS
    If CST_ICMS = True Then
        TBCST!Id_Item = TBAbrir!Int_codigo
        TBCST!Origem_mercadoria = InicioCST
        TBCST!Tributacao_ICMS = FimCST
        If FimCST <> "40" And FimCST <> "41" And FimCST <> "50" And FimCST <> "60" And FimCST <> "101" And FimCST <> "102" And FimCST <> "103" And FimCST <> "300" And FimCST <> "400" And FimCST <> "500" Then
            If FimCST = "20" Or FimCST = "51" Or FimCST = "70" Or FimCST = "90" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "900" Then
                Set TBMaquinas = CreateObject("adodb.recordset")
                TBMaquinas.Open "Select * from regioes where uf = '" & TBNivel15!Estado & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBMaquinas.EOF = False Then
                    Set TBAfericao = CreateObject("adodb.recordset")
                    TBAfericao.Open "Select * from tbl_ClassificacaoFiscal where Idclass = " & IIf(IsNull(TBAbrir!ID_CF), 0, TBAbrir!ID_CF), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAfericao.EOF = False Then
                        Select Case TBMaquinas!regiao
                            Case "DE":
                                TBCST!Percentual_reducao_BC = TBAfericao!CTDE
                                If TBNivel15!Estado <> "MG" And FimCST <> "20" And FimCST <> "51" Then TBCST!Percentual_reducao_BC_ST = TBAfericao!CTDE
                            Case "SS":
                                TBCST!Percentual_reducao_BC = TBAfericao!CTSS
                                If TBNivel15!Estado <> "MG" And FimCST <> "20" And FimCST <> "51" Then TBCST!Percentual_reducao_BC_ST = TBAfericao!CTSS
                            Case "NN":
                                TBCST!Percentual_reducao_BC = TBAfericao!CTNN
                                If TBNivel15!Estado <> "MG" And FimCST <> "20" And FimCST <> "51" Then TBCST!Percentual_reducao_BC_ST = TBAfericao!CTNN
                            Case "CO":
                                TBCST!Percentual_reducao_BC = TBAfericao!CTCO
                                If TBNivel15!Estado <> "MG" And FimCST <> "20" And FimCST <> "51" Then TBCST!Percentual_reducao_BC_ST = TBAfericao!CTCO
                        End Select
                    End If
                    TBAfericao.Close
                End If
            End If
            
            If FimCST <> "201" And FimCST <> "202" And FimCST <> "203" Then
                ValorTotal = 0
                PV = (IIf(IsNull(TBAbrir!dbl_ValorTotal), 0, TBAbrir!dbl_ValorTotal) - IIf(IsNull(TBAbrir!Valor_desconto), 0, TBAbrir!Valor_desconto)) + IIf(IsNull(TBAbrir!Valor_frete), 0, TBAbrir!Valor_frete)
                
                IntICMS = IIf(IsNull(TBAbrir!int_ICMS), 0, TBAbrir!int_ICMS)
                Set TBCiclo = CreateObject("adodb.recordset")
                TBCiclo.Open "select * from tbl_NaturezaOperacao where IDCountCfop = " & IIf(IsNull(TBAbrir!ID_CFOP), 0, TBAbrir!ID_CFOP) & " and txt_ICMS = 'SIM'", Conexao, adOpenKeyset, adLockOptimistic
                If TBCiclo.EOF = False And IntICMS <> 0 Then
                    ProcCalculaBC txtID_empresa, TBCiclo!ID_CFOP, IIf(IsNull(TBAbrir!VlrMPA), 0, TBAbrir!VlrMPA), PV, TBAbrir!dbl_valoripi, TBCiclo!txt_Somar, TBCiclo!Somar_IPI_BC_ICMSST, TBCiclo!TemReducaoBC, IIf(TBAbrir!NaoArredondar = True, True, False), FimCST, "T", txtIDforn, txtFornecedor
                    If TBCiclo!txt_Somar = "SIM" Then TBCST!Valor_BC = Format(BC + IIf(IsNull(TBAbrir!Valor_seguro), 0, TBAbrir!Valor_seguro), "###,##0.00") Else TBCST!Valor_BC = Format(BC + IIf(IsNull(TBAbrir!Valor_seguro), 0, TBAbrir!Valor_seguro) + IIf(IsNull(TBAbrir!Valor_acessorias), 0, TBAbrir!Valor_acessorias), "###,##0.00")
                Else
                    TBCST!Valor_BC = 0
                End If
                TBCiclo.Close
                                
                'Calcula ICMS sem arredondar ou arredondando
                If IntICMS <> 0 Then
                    ValorTotal = TBCST!Valor_BC
                    TBCST!Valor_ICMS = Format((IntICMS * ValorTotal) / 100, "###,##0.00")
                Else
                    TBCST!Valor_ICMS = 0
                End If
            End If
        End If
        
        If FimCST = "101" Or FimCST = "201" Or FimCST = "900" Then
            TBCST!Valor_BC = 0
            TBCST!ICMS_SN = 0
            TBCST!Valor_ICMS_SN = 0
            
            Set TBCiclo = CreateObject("adodb.recordset")
            TBCiclo.Open "select * from tbl_NaturezaOperacao where IDCountCfop = " & IIf(IsNull(TBAbrir!ID_CFOP), 0, TBAbrir!ID_CFOP) & " and txt_ICMS = 'SIM'", Conexao, adOpenKeyset, adLockOptimistic
            If TBCiclo.EOF = False Then
                'If Faturamento_NF_Saida = True Then
                    'ProcValorImposto txtNotaFiscal, IIf(IsNull(TBAbrir!ID_CF), 0, TBAbrir!ID_CF), IIf(txtIDforn = "", 0, txtIDforn), txtfornecedor, TBNivel15!Estado, txtID_empresa, False, IIf(IsNull(TBAbrir!ID_CFOP), 0, TBAbrir!ID_CFOP)
               'Else
                    ProcValorImposto txtnotafiscal, IIf(IsNull(TBAbrir!ID_CF), 0, TBAbrir!ID_CF), IIf(txtIDforn = "", 0, txtIDforn), txtFornecedor, TBNivel15!Estado, txtID_empresa, True, IIf(IsNull(TBAbrir!ID_CFOP), 0, TBAbrir!ID_CFOP), 0
                'End If
                
                PV = (IIf(IsNull(TBAbrir!dbl_ValorTotal), 0, TBAbrir!dbl_ValorTotal) - IIf(IsNull(TBAbrir!Valor_desconto), 0, TBAbrir!Valor_desconto)) + IIf(IsNull(TBAbrir!Valor_frete), 0, TBAbrir!Valor_frete)
                IntICMS = IIf(IsNull(TBAbrir!ICMS_SN), 0, TBAbrir!ICMS_SN)
                ProcCalculaBC txtID_empresa, TBCiclo!ID_CFOP, IIf(IsNull(TBAbrir!VlrMPA), 0, TBAbrir!VlrMPA), PV, TBAbrir!dbl_valoripi, TBCiclo!txt_Somar, TBCiclo!Somar_IPI_BC_ICMSST, TBCiclo!TemReducaoBC, False, FimCST, "T", txtIDforn, txtFornecedor
                If TBCiclo!txt_Somar = "SIM" Then ValorTotal = Format(BC + IIf(IsNull(TBAbrir!Valor_seguro), 0, TBAbrir!Valor_seguro), "###,##0.00") Else ValorTotal = Format(BC + IIf(IsNull(TBAbrir!Valor_seguro), 0, TBAbrir!Valor_seguro) + IIf(IsNull(TBAbrir!Valor_acessorias), 0, TBAbrir!Valor_acessorias), "###,##0.00")
            End If
            TBCiclo.Close
            
            If IntICMS <> 0 Then
                TBCST!Valor_BC = ValorTotal
                TBCST!ICMS_SN = IntICMS
                TBCST!Valor_ICMS_SN = Format((ValorTotal * IntICMS) / 100, "###,##0.00")
            End If
        End If
    End If
    'IPI
    If CST_IPI = True Then
        TBCST!Id_Item = TBAbrir!Int_codigo
        TBCST!Codigo_situacaoTributaria = FimCST
        If FimCST = "00" Or FimCST = "49" Or FimCST = "50" Or FimCST = "99" Then TBCST!Valor_BC = IIf(TBAbrir!Tem_IPI_frete = True, IIf(IsNull(TBAbrir!dbl_ValorTotal), 0, TBAbrir!dbl_ValorTotal) + IIf(IsNull(TBAbrir!Valor_frete), 0, TBAbrir!Valor_frete), IIf(IsNull(TBAbrir!dbl_ValorTotal), 0, TBAbrir!dbl_ValorTotal))
    End If
    'PIS
    If CST_PIS = True Then
        TBCST!Id_Item = TBAbrir!Int_codigo
        TBCST!Codigo_situacaoTributaria = FimCST
        If FimCST = "01" Or FimCST = "03" Or FimCST = "49" Or FimCST = "98" Or FimCST = "99" Then TBCST!Valor_BC = IIf(IsNull(TBAbrir!dbl_ValorTotal), 0, TBAbrir!dbl_ValorTotal)
    End If
    'Cofins
    If CST_Cofins = True Then
        TBCST!Id_Item = TBAbrir!Int_codigo
        TBCST!Codigo_situacaoTributaria = FimCST
        If FimCST = "01" Or FimCST = "02" Or FimCST = "03" Or FimCST = "49" Or FimCST = "98" Or FimCST = "99" Then TBCST!Valor_BC = IIf(IsNull(TBAbrir!dbl_ValorTotal), 0, TBAbrir!dbl_ValorTotal)
    End If
    TBCST.Update
End If
TBNivel15.Close
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
