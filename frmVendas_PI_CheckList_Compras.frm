VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVendas_PI_CheckList_Compras 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Vendas | Pedido interno | Chek list de compras"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15390
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
   Icon            =   "frmVendas_PI_CheckList_Compras.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15390
   Begin FlexCell.Grid Grid1 
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   1020
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   12303
      BackColorActiveCellSel=   12648447
      BackColorBkg    =   16777215
      Cols            =   2
      DefaultFontSize =   6.75
      DisplayFocusRect=   0   'False
      DisplayRowIndex =   -1  'True
      ExtendLastCol   =   -1  'True
      GridColor       =   12632256
      ReadOnly        =   -1  'True
      Rows            =   2
      EnterKeyMoveTo  =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Filtrar"
      Height          =   435
      Left            =   13290
      TabIndex        =   2
      Top             =   1290
      Width           =   1425
   End
   Begin MSComctlLib.ListView Lista_comprados 
      Height          =   2040
      Left            =   0
      TabIndex        =   0
      Top             =   8010
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   3598
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Pedido"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fornecedor"
         Object.Width           =   21238
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "quantidade"
         Object.Width           =   2648
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Prazo entrega"
         Object.Width           =   2540
      EndProperty
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   26
      ScreenHeight    =   1080
      ScreenWidth     =   1920
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10500
      FormWidthDT     =   15510
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15390
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   1065
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   1879
      ButtonCount     =   4
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Filtrar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Filtrar (F2)"
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
      ButtonWidth1    =   36
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
      ButtonLeft2     =   40
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   60
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
      ButtonLeft3     =   44
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
      ButtonLeft4     =   82
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   12270
         Top             =   180
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmVendas_PI_CheckList_Compras.frx":000C
         Count           =   1
      End
   End
   Begin VB.Image imgFile 
      Height          =   240
      Left            =   270
      Picture         =   "frmVendas_PI_CheckList_Compras.frx":21F4
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgFolder 
      Height          =   240
      Left            =   0
      Picture         =   "frmVendas_PI_CheckList_Compras.frx":277E
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmVendas_PI_CheckList_Compras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Novo_Estrutura As Boolean 'OK
Dim StrSql_Engenharia_Estrutura As String 'OK
Public VersaoEstrutura As String 'OK

'GridEstrutura
Public m_Tree As New Node
Public m_Row As Long
Public m_Col As Long
Dim tempNode As Node
Dim intIndex, i As Integer
Dim CodRef As String, ValorCusto As String, DataValidacao As String, RespValidacao As String
Public IDProduto As Long, IDestrutura As Long

Sub ProcCarregaEstrutura()
On Error GoTo tratar_erro
Dim Aplicacao As String

'StrSql = "Select VC.Desenho, VC.Unidade, VC.Quantidade, VC.Descricao, PP.codproduto, isnull(EC.Saldo,0) as Saldo  from vendas_Carteira VC Inner Join ProjProduto PP on PP.desenho = VC.desenho Inner join Estoque_Controle_Saldo_Item EC on VC.desenho = EC.desenho where cotacao = '" & frmVendas_PI.txtID & "'"


Call m_Tree.Nodes.Clear
Grid1.rows = 1

m_Row = 1
m_Col = 1
Desenho = ""
Familiatext = ""
TotalPedidos = ""
TotalItens = ""
PrazoMinimo = ""
PrazoMaximo = ""

'Debug.print StrSql

Set TBLISTA = CreateObject("adodb.recordset")

TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    
    Contador1 = -1
    Contador = 0
    Do While Not TBLISTA.EOF
    Contador1 = Contador1 + 1
    arrNodes(Contador1).Level = 0
    Requisitado = TBLISTA!quantidade - TBLISTA!Saldo
    
    If TBLISTA!Vendas = True Then
    Aplicacao = "V"
    End If
    
    If TBLISTA!Compras = True Then
        Aplicacao = "C"
        StrSql = "select Count(IDPedido) as TotalPedidos, Desenho, MIN(prazo) AS Prazominimo, max(prazo) as Prazomaximo, sum(quant_comp) as Comprado from Compras_pedido_lista where Desenho = '" & TBLISTA!Desenho & "' And Status_Item = 'N_RECEBIDO' GROUP BY Desenho"
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            TotalPedidos = TBAbrir!TotalPedidos
            TotalItens = TBAbrir!Comprado
            PrazoMinimo = TBAbrir!PrazoMinimo
            PrazoMaximo = TBAbrir!PrazoMaximo
        End If
        TBAbrir.Close
    End If
    
    
    If TBLISTA!Producao = True Then
    Aplicacao = "P"
    End If
    
    
    arrNodes(Contador1).Text = TBLISTA!Desenho & vbTab & Aplicacao & vbTab & TBLISTA!Descricao & vbTab & TBLISTA!Unidade & vbTab & "1,00" & vbTab & Format(TBLISTA!quantidade, "###,##0.00") & vbTab & Format(TBLISTA!Saldo, "###,##0.00") & vbTab & Format(IIf(Requisitado > 0, Requisitado, 0), "###,##0.00") & vbTab & Format(IIf(TotalItens > 0, TotalItens, 0), "###,##0.00") & vbTab & PrazoMinimo & vbTab & PrazoaMximo
    
    Codproduto = TBLISTA!Codproduto
    Tipo = "A"
    
    TotalPedidos = ""
    TotalItens = ""
    PrazoMinimo = ""
    PrazoMaximo = ""


'===========================
' Carregar nivel 2
'===========================
StrSql = "Select PC.Desenho,PC.Descricao, PC.Unidade,PC.quantidade,PC.obs, PP.codProduto as CodPai, Isnull(EC.saldo,0)as Saldo, PP.Compras, PP.Vendas, PP.Producao  from projconjunto PC Inner join ProjProduto PP on PC.Desenho = PP.desenho left outer Join Estoque_Controle_Saldo_Item EC on PP.Desenho = EC.Desenho where PC.codproduto = " & Codproduto & " order by PC.Posicao, PC.Codigo"
                
                Codproduto = TBLISTA!Codproduto
                Set TBNivel2 = CreateObject("adodb.recordset")
                TBNivel2.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
'Debug.print StrSql

                If TBNivel2.EOF = False Then
                    Do While TBNivel2.EOF = False
                                Contador1 = Contador1 + 1
                                arrNodes(Contador1).Level = 1
                                Requisitado = (TBNivel2!quantidade * TBLISTA!quantidade) - TBNivel2!Saldo
                                Necessidade = (TBNivel2!quantidade * TBLISTA!quantidade)
                                
                                If TBLISTA!Vendas = True Then
                                Aplicacao = "V"
                                End If
                                
                                
                                If TBNivel2!Compras = True Then
                                Aplicacao = "C"
                                StrSql = "select Count(IDPedido) as TotalPedidos, Desenho, MIN(prazo) AS Prazominimo, max(prazo) as Prazomaximo, sum(quant_comp) as Comprado from Compras_pedido_lista where Desenho = '" & TBNivel2!Desenho & "' And Status_Item = 'N_RECEBIDO' GROUP BY Desenho"
                                Set TBAbrir = CreateObject("adodb.recordset")
                                TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                                If TBAbrir.EOF = False Then
                                    TotalPedidos = TBAbrir!TotalPedidos
                                    TotalItens = TBAbrir!Comprado
                                    PrazoMinimo = TBAbrir!PrazoMinimo
                                    PrazoMaximo = TBAbrir!PrazoMaximo
                                End If
                                TBAbrir.Close
                                
                                End If
                                
                                
                                If TBNivel2!Producao = True Then
                                Aplicacao = "P"
                                End If
                                
                                
                                arrNodes(Contador1).Text = TBNivel2!Desenho & vbTab & Aplicacao & vbTab & TBNivel2!Descricao & vbTab & TBNivel2!Unidade & vbTab & Format(TBNivel2!quantidade, "###,##0.00") & vbTab & Format(Necessidade, "###,##0.00") & vbTab & Format(TBNivel2!Saldo, "###,##0.00") & vbTab & Format(IIf(Requisitado > 0, Requisitado, 0), "###,##0.00") & vbTab & Format(IIf(TotalItens > 0, TotalItens, 0), "###,##0.00") & vbTab & PrazoMinimo & vbTab & PrazoMaximo & vbTab & TBNivel2!Obs
                                
                                
    TotalPedidos = ""
    TotalItens = ""
    PrazoMinimo = ""
    PrazoMaximo = ""
                                
'===========================
' Carregar nivel 3
'===========================
                                   Codproduto = TBNivel2!Codpai
                                    Set TBNivel3 = CreateObject("adodb.recordset")
                                    
                                    StrSql = "Select PC.Desenho,PC.Descricao, PC.Unidade,PC.quantidade,PC.obs, PP.codProduto as CodPai, Isnull(EC.saldo,0) as Saldo, PP.Compras, PP.Vendas, PP.Producao from projconjunto PC Inner join ProjProduto PP on PC.Desenho = PP.desenho left outer Join Estoque_Controle_Saldo_Item EC on PP.Desenho = EC.Desenho where PC.codproduto = " & Codproduto & " order by PC.Posicao, PC.Codigo"
'Debug.print StrSql

                                    TBNivel3.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                                    If TBNivel3.EOF = False Then
                                        Do While TBNivel3.EOF = False
                                                    Contador1 = Contador1 + 1
                                                    arrNodes(Contador1).Level = 2
                                                    Requisitado = (TBNivel3!quantidade * TBLISTA!quantidade) - TBNivel3!Saldo
                                                    Necessidade = (TBNivel3!quantidade * TBLISTA!quantidade)
                                                    
                                                    If TBLISTA!Vendas = True Then
                                                    Aplicacao = "V"
                                                    End If
                                                    
     
                                                    If TBNivel3!Compras = True Then
                                                    Aplicacao = "C"
                                                        StrSql = "select Count(IDPedido) as TotalPedidos, Desenho, MIN(prazo) AS Prazominimo, max(prazo) as Prazomaximo, sum(quant_comp) as Comprado from Compras_pedido_lista where Desenho = '" & TBNivel3!Desenho & "' And Status_Item = 'N_RECEBIDO' GROUP BY Desenho"
                                                        Set TBAbrir = CreateObject("adodb.recordset")
                                                        TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                                                        If TBAbrir.EOF = False Then
                                                            TotalPedidos = TBAbrir!TotalPedidos
                                                            TotalItens = TBAbrir!Comprado
                                                            PrazoMinimo = TBAbrir!PrazoMinimo
                                                            PrazoMaximo = TBAbrir!PrazoMaximo
                                                        End If
                                                        TBAbrir.Close
                                                    
                                                    End If
                                                    
                                                    If TBNivel3!Producao = True Then
                                                    Aplicacao = "P"
                                                    End If

     
                                                    arrNodes(Contador1).Text = TBNivel3!Desenho & vbTab & Aplicacao & vbTab & TBNivel3!Descricao & vbTab & TBNivel3!Unidade & vbTab & Format(TBNivel3!quantidade, "###,##0.00") & vbTab & Format(Necessidade, "###,##0.00") & vbTab & Format(TBNivel3!Saldo, "###,##0.00") & vbTab & Format(IIf(Requisitado > 0, Requisitado, 0), "###,##0.00") & vbTab & Format(IIf(TotalItens > 0, TotalItens, 0), "###,##0.00") & vbTab & PrazoMinimo & vbTab & PrazoMaximo & vbTab & TBNivel3!Obs

                                                   
    TotalPedidos = ""
    TotalItens = ""
    PrazoMinimo = ""
    PrazoMaximo = ""
                                                   
'===========================
' Carregar nivel 4
'===========================
                                                            Codproduto = TBNivel3!Codpai
                                                            Set TBNivel4 = CreateObject("adodb.recordset")
                                                            StrSql = "Select PC.Desenho,PC.Descricao, PC.Unidade,PC.quantidade,PC.obs, PP.codProduto as CodPai, Isnull(EC.saldo,0) as Saldo, PP.Compras, PP.Vendas, PP.Producao from projconjunto PC Inner join ProjProduto PP on PC.Desenho = PP.desenho left outer Join Estoque_Controle_Saldo_Item EC on PP.Desenho = EC.Desenho where PC.codproduto = " & Codproduto & " order by PC.Posicao, PC.Codigo"
                                                            TBNivel4.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                                                             If TBNivel4.EOF = False Then
                                                                 Do While TBNivel4.EOF = False
                                                                             Contador1 = Contador1 + 1
                                                                             arrNodes(Contador1).Level = 3
                                                                            Requisitado = (TBNivel4!quantidade * TBLISTA!quantidade) - TBNivel4!Saldo
                                                                            Necessidade = (TBNivel4!quantidade * TBLISTA!quantidade)
                                                                            QtTexto = Format(TBNivel4!quantidade, "0.0000")
                                                                            
                                                                            If TBLISTA!Vendas = True Then
                                                                            Aplicacao = "V"
                                                                            End If
                                                                                                                                                       
                                                                            If TBNivel4!Compras = True Then
                                                                            Aplicacao = "C"
                                                                                StrSql = "select Count(IDPedido) as TotalPedidos, Desenho, MIN(prazo) AS Prazominimo, max(prazo) as Prazomaximo, sum(quant_comp) as Comprado from Compras_pedido_lista where Desenho = '" & TBNivel4!Desenho & "' And Status_Item = 'N_RECEBIDO' GROUP BY Desenho"
                                                                                Set TBAbrir = CreateObject("adodb.recordset")
                                                                                TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
                                                                                If TBAbrir.EOF = False Then
                                                                                    TotalPedidos = TBAbrir!TotalPedidos
                                                                                    TotalItens = TBAbrir!Comprado
                                                                                    PrazoMinimo = TBAbrir!PrazoMinimo
                                                                                    PrazoMaximo = TBAbrir!PrazoMaximo
                                                                                End If
                                                                                TBAbrir.Close
                                                                            End If
                                                                            
                                                                            If TBNivel4!Producao = True Then
                                                                            Aplicacao = "P"
                                                                            End If
                                                                            arrNodes(Contador1).Text = TBNivel4!Desenho & vbTab & Aplicacacao & vbTab & TBNivel4!Descricao & vbTab & TBNivel4!Unidade & vbTab & Format(TBNivel4!quantidade, "###,##0.00") & vbTab & Format(Necessidade, "###,##0.00") & vbTab & Format(TBLISTA!Saldo, "###,##0.00") & vbTab & Format(IIf(Requisitado > 0, Requisitado, 0), "###,##0.00") & vbTab & Format(IIf(TotalItens > 0, TotalItens, 0), "###,##0.00") & vbTab & PrazoMinimo & vbTab & PrazoMaximo & vbTab & TBNivel4!Obs
    TotalPedidos = ""
    TotalItens = ""
    PrazoMinimo = ""
    PrazoMaximo = ""
                                                                     
                                                                     TBNivel4.MoveNext
                                                                 Loop
                                                             End If
                                                TBNivel3.MoveNext
                                            Loop
                                        End If
                        TBNivel2.MoveNext
                    Loop
                End If
        Contador = Contador + 1
        TBLISTA.MoveNext
    Loop

Call m_Tree.Nodes.Clear

 With Grid1
        
        .AutoRedraw = False
        .AllowUserPaste = cellTextOnly
        .ExtendLastCol = True
        .DrawMode = cellOwnerDraw
        .Cols = 13
        .rows = 1
        .Cell(0, 1).Text = "Cód. interno"
        .Cell(0, 2).Text = "Aplicação"
        .Cell(0, 3).Text = "Descrição"
        .Cell(0, 4).Text = "Un."
        .Cell(0, 5).Text = "Qtde"
        .Cell(0, 6).Text = "Requisitado"
        .Cell(0, 7).Text = "Em estoque"
        .Cell(0, 8).Text = "Necessidade"
        .Cell(0, 9).Text = "Comprado"
        .Cell(0, 10).Text = "Prazo minimo"
        .Cell(0, 11).Text = "Prazo máximo"
        
        .Cell(0, 12).Text = "Observações"
        
        .Column(0).Width = 30
        .Column(1).Width = 140
        .Column(1).Alignment = cellCenterCenter
        .Column(2).Width = 50
        .Column(2).Alignment = cellCenterCenter
        .Column(3).Width = 200
        .Column(3).Alignment = cellLeftCenter
        .Column(4).Width = 30
        .Column(4).Alignment = cellCenterCenter
        .Column(5).Width = 0
        .Column(5).Alignment = cellRightCenter
        .Column(6).Width = 55
        .Column(6).Alignment = cellRightCenter
        .Column(7).Width = 55
        .Column(7).Alignment = cellRightCenter
        .Column(8).Width = 55
        .Column(8).Alignment = cellRightCenter
        .Column(9).Width = 55
        .Column(9).Alignment = cellRightCenter
        .Column(10).Width = 70
        .Column(10).Alignment = cellCenterCenter
        .Column(11).Width = 70
        .Column(11).Alignment = cellCenterCenter
       
        .Column(12).Width = 100
        .Column(12).Alignment = cellCenterCenter

        .AutoRedraw = True
        .Refresh
  
       
        'First node
        Set tempNode = m_Tree.Nodes.Add("")
        .AddItem arrNodes(0).Text

        'Other nodes
        For intIndex = 1 To Contador1 'UBound(arrNodes)
            If arrNodes(intIndex).Level = arrNodes(intIndex - 1).Level Then
                Set tempNode = tempNode.Parent.Nodes.Add("")
            ElseIf arrNodes(intIndex).Level > arrNodes(intIndex - 1).Level Then
                Set tempNode = tempNode.Nodes.Add("")
            ElseIf arrNodes(intIndex).Level < arrNodes(intIndex - 1).Level Then
                For i = arrNodes(intIndex).Level To arrNodes(intIndex - 1).Level
                    Set tempNode = tempNode.Parent
                Next
                Set tempNode = tempNode.Nodes.Add("")
            End If
            .AddItem arrNodes(intIndex).Text
        Next

        .AutoRedraw = True
        .Refresh
    End With
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Grid1_Click()
On Error GoTo tratar_erro
Dim point As POINTAPI
Dim objCell As FlexCell.Cell
Dim intWidth As Integer

intWidth = 20

Call GetCursorPos(point)
Call ScreenToClient(Grid1.hWnd, point)
Set objCell = Grid1.HitTest(point.x, point.Y)

If Not objCell Is Nothing Then
    If objCell.Row >= m_Row And objCell.Col = m_Col Then
        Dim objNode As Node
        Set objNode = m_Tree.FindNode(objCell.Row - m_Row + 2)
        If Not objNode Is Nothing Then
            Dim i As Long, x As Long, Y As Long
            x = objCell.Left + 2 + (objNode.Level - 1) * intWidth
            Y = objCell.Top + (objCell.Height - 9) / 2
            If point.x >= x And point.x <= x + 9 And point.Y >= Y And point.Y <= Y + 9 Then
                If objNode.Expanded Then
                    objNode.Collapse
                    Grid1.AutoRedraw = False
                    For i = 1 To objNode.ChildrenCount
                        Grid1.RowHeight(objCell.Row + i) = 0
                    Next
                    Grid1.AutoRedraw = True
                    Grid1.Refresh
                Else
                    objNode.Expand
                    Grid1.AutoRedraw = False
                    For i = 1 To objNode.ChildrenCount
                        If objNode.FindNode(i + 1).Visible Then
                            Grid1.RowHeight(objCell.Row + i) = -1 'DefaultRowHeight
                        End If
                    Next
                    Grid1.AutoRedraw = True
                    Grid1.Refresh
                End If
            End If
        End If
    End If
End If

ProcBuscaPedidos (Grid1.ActiveCell.Row)


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBuscaPedidos(vRow As Integer)
On Error GoTo tratar_erro

Lista_comprados.ListItems.Clear

If Grid1.Cell(vRow, 2).Text <> "" And vRow > 0 Then
Desenho = Grid1.Cell(vRow, 1).Text
If Grid1.Cell(vRow, 2).Text = "C" Then
StrSql = "Select CPL.idlista, CP.Pedido, CP.Fornecedor, CPL.Quant_Comp, CPL.Prazo, CPL.Status_Item From compras_Pedido_Lista CPL Inner Join Compras_Pedido CP on CPL.IDPedido = CP.IDPedido Where CPL.Desenho = '" & Desenho & "' and CPL.Status_Item = 'N_RECEBIDO'"
'Debug.print StrSql

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly

If TBLISTA.EOF = False Then
Do While TBLISTA.EOF = False
    With Lista_comprados.ListItems
        .Add , , TBLISTA!IDlista
        .Item(.Count).SubItems(1) = TBLISTA!Pedido
        .Item(.Count).SubItems(2) = TBLISTA!Fornecedor
        .Item(.Count).SubItems(3) = Format(IIf(TBLISTA!Quant_Comp > 0, TBLISTA!Quant_Comp, 0), "###,##0.00")
        .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Prazo), "", TBLISTA!Prazo)
    End With
    TBLISTA.MoveNext
Loop
End If
End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Grid1_LeaveCell(ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
On Error GoTo tratar_erro

'USMsgBox NewRow & NewCol

ProcBuscaPedidos (NewRow)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Grid1_OwnerDrawCell(ByVal Row As Long, ByVal Col As Long, ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Handled As Boolean)
On Error GoTo tratar_erro

Dim i As Long, j As Long
Dim x As Long, Y As Long
Dim hPen As Long, hOldPen As Long
Dim hBrush As Long, hOldBrush As Long
Dim lngLevel As Long
Dim blnDrawLine As Boolean
Dim objNode As Node, tmpNode As Node
Dim intWidth As Integer
Dim intAdd As Integer

If Row < m_Row Or Col <> m_Col Then Exit Sub

intWidth = 20
intAdd = 26
    
Set objNode = m_Tree.FindNode(Row - m_Row + 2)
If Not objNode Is Nothing Then
    lngLevel = objNode.Level - 1

    'Tree lines
    hPen = CreatePen(0, 1, RGB(128, 128, 128))
    hOldPen = SelectObject(hdc, hPen)
    For i = 0 To lngLevel
        If i < lngLevel - 1 Then
            blnDrawLine = True
            Set tmpNode = objNode
            For j = i To lngLevel - 2
                Set tmpNode = tmpNode.Parent
            Next
            If tmpNode.NextNode Is Nothing Then
                blnDrawLine = False
            End If
            If blnDrawLine Then
                'All
                Call DrawLine(hdc, Left + intWidth * i + intAdd, Top - 1, Left + intWidth * i + intAdd, Bottom + 1)
            End If
        ElseIf i = lngLevel - 1 Then
            'Top
            Call DrawLine(hdc, Left + intWidth * i + intAdd, Top - 1, Left + intWidth * i + intAdd, Top + (Bottom - Top) / 2)
            If Not objNode.NextNode Is Nothing Then
                'Bottom
                Call DrawLine(hdc, Left + intWidth * i + intAdd, Top + (Bottom - Top) / 2, Left + intWidth * i + intAdd, Bottom + 1)
            End If
        ElseIf i = lngLevel Then
            'Top
            If objNode.VisibleNodesCount > 1 Then
                Call DrawLine(hdc, Left + intWidth * i + intAdd, Top + (Bottom - Top) / 2, Left + intWidth * i + intAdd, Bottom + 1)
            End If
        End If
        'Horizontal line
        If lngLevel > 0 Then
            Call DrawLine(hdc, Left + intWidth * (lngLevel - 1) + intAdd, Top + (Bottom - Top) / 2, Left + intWidth * (lngLevel - 1) + intAdd + 10, Top + (Bottom - Top) / 2)
        End If
    Next
    
    Call SelectObject(hdc, hOldPen)
    Call DeleteObject(hPen)

    '+/-
    If objNode.ChildrenCount > 0 Then
        hPen = CreatePen(0, 1, 0)
        hOldPen = SelectObject(hdc, hPen)
        hBrush = CreateSolidBrush(RGB(255, 255, 255))
        hOldPen = SelectObject(hdc, hBrush)
        
        x = Left + 2 + intWidth * lngLevel
        Y = Top + (Bottom - Top - 9) / 2
        
        Call Rectangle(hdc, x, Y, x + 9, Y + 9)
        If objNode.Expanded Then
            Call DrawLine(hdc, x + 2, Y + 4, x + 7, Y + 4)
        Else
            Call DrawLine(hdc, x + 2, Y + 4, x + 7, Y + 4)
            Call DrawLine(hdc, x + 4, Y + 2, x + 4, Y + 7)
        End If
    
        Call SelectObject(hdc, hOldPen)
        Call DeleteObject(hPen)
        Call SelectObject(hdc, hOldBrush)
        Call DeleteObject(hBrush)
    End If
    
    'Icon
    If objNode.ChildrenCount > 0 Then
        DrawIconEx hdc, Left + intWidth * lngLevel + 18, Top + (Bottom - Top - 16) / 2, imgFolder.Picture, 16, 16, 0, 0, DI_NORMAL
    Else
        DrawIconEx hdc, Left + intWidth * lngLevel + 18, Top + (Bottom - Top - 16) / 2, imgFile.Picture, 16, 16, 0, 0, DI_NORMAL
    End If
    
    'Text
    With Grid1.Cell(Row, Col)
        Dim rc As rect
        Call SetRect(rc, Left + intWidth * lngLevel + 37, Top, Right, Bottom)
        Call DrawText(hdc, .Text, -1, rc, DT_SINGLELINE Or DT_VCENTER)
    End With

    Handled = True
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Form_Load()
On Error GoTo tratar_erro
ProcCarregaToolBar1 Me, 17000, 4, True


ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro
StrSql = "Select VC.Desenho, VC.Unidade, VC.Quantidade, VC.Descricao, PP.codproduto,PP.compras,PP.vendas,PP.Producao, isnull(EC.Saldo,0) as Saldo  from vendas_Carteira VC Inner Join ProjProduto PP on PP.desenho = VC.desenho left outer join Estoque_Controle_Saldo_Item EC on VC.desenho = EC.desenho where cotacao = '" & frmVendas_PI.txtId & "'"
'Debug.print StrSql

    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
    If TBLISTA.EOF = False Then
        ProcCarregaEstrutura
    End If
    TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 2: 'ProcAjuda
    Case 4: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

