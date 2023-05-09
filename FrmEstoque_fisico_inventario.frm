VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form FrmEstoque_fisico_inventario 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Estoque | Importação de inventário"
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   7560
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   4500
      Width           =   6795
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   5205
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   714
      DibPicture      =   "FrmEstoque_fisico_inventario.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      EnableMaximizeButton=   0   'False
      EnableMinimizeButton=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "FrmEstoque_fisico_inventario.frx":65C5
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   300
      Top             =   870
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   525
      Left            =   360
      TabIndex        =   3
      Top             =   3840
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   926
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
   Begin DrawSuite2022.USButton cmdProximo 
      Height          =   1305
      Left            =   3660
      TabIndex        =   4
      Top             =   2460
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2302
      DibPicture      =   "FrmEstoque_fisico_inventario.frx":68DF
      Caption         =   "Próximo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      PicAlign        =   7
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
      ShowFocusRect   =   0   'False
   End
   Begin DrawSuite2022.USButton cmdAnterior 
      Height          =   1305
      Left            =   360
      TabIndex        =   5
      Top             =   2460
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   2302
      DibPicture      =   "FrmEstoque_fisico_inventario.frx":8A4E
      Caption         =   "Anterior"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      PicAlign        =   7
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
      ShowFocusRect   =   0   'False
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2 ° Escolha a planilha a ser importada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   390
      TabIndex        =   6
      Top             =   600
      Width           =   6765
      Begin VB.TextBox txtlocal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   390
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Arquivo excel a ser importado"
         Top             =   870
         Width           =   5565
      End
      Begin DrawSuite2022.USButton btnPlanilha 
         Height          =   435
         Left            =   5970
         TabIndex        =   13
         ToolTipText     =   "Localizar planilha de"
         Top             =   870
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         DibPicture      =   "FrmEstoque_fisico_inventario.frx":9F4F
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
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
         PicAlign        =   0
         PicSize         =   1
         Theme           =   5
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Planilha excell de inventário a ser importada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   2
         Left            =   1545
         TabIndex        =   8
         Top             =   600
         Width           =   3810
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1 ° Escolha a data do inventário"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   360
      TabIndex        =   1
      Top             =   630
      Width           =   6795
      Begin MSComCtl2.DTPicker Txt_data 
         Height          =   405
         Left            =   4380
         TabIndex        =   11
         ToolTipText     =   "Data."
         Top             =   660
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
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
         Format          =   198705153
         CurrentDate     =   43830
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atenção!!       Data do inventário"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   14
         Left            =   630
         TabIndex        =   12
         Top             =   720
         Width           =   3450
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "3 ° Iniciar a importação"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   360
      TabIndex        =   9
      Top             =   600
      Width           =   6795
      Begin DrawSuite2022.USButton btnImportar 
         Height          =   885
         Left            =   510
         TabIndex        =   10
         ToolTipText     =   "Importar planilha"
         Top             =   540
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   1561
         DibPicture      =   "FrmEstoque_fisico_inventario.frx":248FB
         Caption         =   "Importar planilha (Inventário)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
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
         PicAlign        =   7
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
   End
End
Attribute VB_Name = "FrmEstoque_fisico_inventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xl As New Excel.Application
Dim xlw As Excel.Workbook
Dim Passo1 As Boolean
Dim Passo2 As Boolean
Dim Passo3 As Boolean

Private Sub Cmd1_Click()
On Error GoTo tratar_erro
Dim Filter As String
'==============================================================
' Localiza arquivo excell
'==============================================================
Arquivo = ""
Filter = "(*.xlsx) | *.xlsx"
CD1.Filter = Filter
CD1.InitDir = App.Path
CD1.DefaultExt = "*.xlsx"
CD1.ShowOpen
Arquivo = CD1.filename
txtlocal = IIf(Arquivo = "", "", Arquivo)
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEntradaInventario()
On Error GoTo tratar_erro

Set TBItem = CreateObject("adodb.recordset")
StrSql = "SELECT IDESTOQUE, Sum(Entrada) - Sum(Saida) AS SALDO FROM Estoque_movimentacao WHERE Desenho = '" & CodigoLista & "' AND Data <= '" & Txt_data.Value & "'  GROUP BY IDESTOQUE HAVING Sum(Entrada) - Sum(Saida) > 0"
'Debug.print StrSql

TBItem.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly

'==================================================================
' Se não tiver nenhuma ficha no estoque controle, cria uma nova
'==================================================================
If TBItem.EOF = True Then
StrSql = "Select * from Estoque_Controle where Desenho = '" & CodigoLista & "'"
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = True Then
TBItem.AddNew

'TBItem!ID_empresa = frmestoque_fisico.Cmb_empresa.ItemData(frmestoque_fisico.Cmb_empresa.ListIndex)
TBItem!status = "ENTRADA_INVENTÁRIO"

TBItem!LOTE = "INV-" & Year(Txt_data.Value)
TBItem!Desenho = CodigoLista

TBItem!ID_empresa = frmestoque_fisico.Cmb_empresa.ItemData(frmestoque_fisico.Cmb_empresa.ListIndex)
TBItem!Bloqueado = "False"

Set TBCodigoDesc = CreateObject("adodb.recordset")
TBCodigoDesc.Open "Select * from projproduto where desenho = '" & CodigoLista & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBCodigoDesc.EOF = False Then
TBItem!Classe = TBCodigoDesc!Classe
End If

TBItem!Descricao = TBCodigoDesc!Descricao
TBItem!descricaotecnica = TBCodigoDesc!Descricao
TBItem!Data = Txt_data
TBItem!Corrida = "0"
TBItem!Certificado = "0"
TBItem!Un = TBCodigoDesc!Unidade

TBItem!valor_unitario = "0" 'Format(txtVlr_unit, "###.##0.00000")
TBItem!Valor_total = "0" 'Format(txtVlr_total, "###.##0.00")
'TBItem!local_armaz = cmbLocal_armaz

TBItem.Update



End If


If TBItem.EOF = False Then
'TBItem.MoveLast
'======================================================================
' Cria inventario no estoque fisico
'======================================================================
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Estoque_fisico", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
TBGravar!ID_empresa = frmestoque_fisico.Cmb_empresa.ItemData(frmestoque_fisico.Cmb_empresa.ListIndex)
TBGravar!Data = Txt_data
TBGravar!Responsavel = pubUsuario
'TBGravar!Codproduto = TBCodigoDesc!Codproduto
'TBGravar!Etiqueta = Txt_etiqueta
TBGravar!Novo_lote = False
TBGravar!LOTE = "INV-" & Year(Txt_data.Value)
TBGravar!Corrida = "0"
TBGravar!Certificado = "0"
TBGravar!IDEstoque = TBItem!IDEstoque
'TBGravar!local_armaz = txt_LA.Text
TBGravar!valor_unitario = "0" 'Txt_valor_unitario
TBGravar!Qtde_estoque = QTBaixar 'IIf(Txt_qtde_estoque <> "", Txt_qtde_estoque, 0)
TBGravar!Destino = "Interno"
TBGravar!qtde_fisica = QTBaixar
TBGravar!DtValidacao = Date
TBGravar!RespValidacao = pubUsuario
TBGravar.Update
TBGravar.Close

End If
End If
'======================================================================
' Cria movimentação no estoque
'======================================================================


Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Estoque_Movimentacao", Conexao, adOpenKeyset, adLockOptimistic
TBAbrir.AddNew

TBAbrir!ID_empresa = frmestoque_fisico.Cmb_empresa.ItemData(frmestoque_fisico.Cmb_empresa.ListIndex)
TBAbrir!Bloqueado = "False"
TBAbrir!Destino = "Interno"
TBAbrir!Terceiros = False
TBAbrir!LOTE = "INV-" & Year(Txt_data.Value)
TBAbrir!Documento = "INV-" & Year(Txt_data.Value)
TBAbrir!Desenho = CodigoLista

Set TBCodigoDesc = CreateObject("adodb.recordset")
TBCodigoDesc.Open "Select * from projproduto where desenho = '" & CodigoLista & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBCodigoDesc.EOF = False Then
TBAbrir!Familia = TBCodigoDesc!Classe
End If

TBAbrir!Operacao = "ENTRADA_INVENTÁRIO"
TBAbrir!Entrada = QTBaixar
TBAbrir!Saida = 0 'QTBaixar
TBAbrir!Entrada_PC = QTBaixar
TBAbrir!VlrUnit = 0 'Format(IIf(IsNull(TBAbrir!valor_unitario), 0, TBAbrir!valor_unitario), "###,##0.0000000000")
TBAbrir!vlrTotal = 0 'Format(TBEstoque!VlrUnit * qtdeliberada, "###,##0.00")
TBAbrir!Descricao = TBCodigoDesc!Descricao 'Produto
TBAbrir!Data = Txt_data.Value
TBAbrir!Responsavel = pubUsuario
TBAbrir!IDEstoque = TBItem!IDEstoque
TBAbrir!estoque_venda = QTBaixar

TBAbrir.Update
TBAbrir.Close
TBItem.Close
TBCodigoDesc.Close
'End If
'cmdImportar.Visible = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSaidaInventário()
On Error GoTo tratar_erro
QTBaixado = 0
'QTBaixar = 1
1:
Do While QTBaixar > 0.0001
Set TBItem = CreateObject("adodb.recordset")
'StrSql = "SELECT IDestoque as RE,Estoque_disponivel as Saldo from Estoque_produtos WHERE Desenho = '" & CodigoLista & "' AND Estoque_disponivel > '0' AND Data <= '" & Txt_data.Value & "'"
StrSql = "SELECT IDESTOQUE, Sum(Entrada) - Sum(Saida) AS SALDO FROM Estoque_movimentacao WHERE Desenho = '" & CodigoLista & "' AND Data <= '" & Txt_data.Value & "'  GROUP BY IDESTOQUE HAVING  Sum(Entrada) - Sum(Saida) > 0"

'StrSql = "SELECT IDestoque as RE,Estoque_disponivel as Saldo from Estoque_produtos WHERE (Desenho = N'" & CodigoLista & "') AND (Estoque_disponivel > 0)"
'Debug.print StrSql
TBItem.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
'Do while TBItem.EOF = False
'TBItem.Open "Select RE,Saldo from Estoque_Controle_Saldo_RE where Codigo = '" & CodigoLista & "' and Saldo >= '0'", Conexao, adOpenKeyset, adLockReadOnly


If TBItem.EOF = False Then

'Do While TBItem!Saldo <= 0 And TBItem.RecordCount > 0
'
'
'TBItem.MoveNext
'
'
'Loop


If TBItem!Saldo < QTBaixar And TBItem!Saldo <> 0 Then
QTBaixado = QTBaixado + TBItem!Saldo
Else
QTBaixado = QTBaixar
End If
'Conexao.Execute ("update Estoque_Controle Set Estoque_real = ECSRE.Saldo, Estoque_venda = ECSRE.Saldo, Qtde_fisica = ECSRE.saldo from Estoque_Controle EC inner Join Estoque_controle_Saldo_RE ECSRE on EC.IdEstoque = ECSRE.RE where EC.IdEstoque = " & TBItem!RE)
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Estoque_Movimentacao", Conexao, adOpenKeyset, adLockOptimistic
TBAbrir.AddNew


TBAbrir!Destino = "Interno"
TBAbrir!Terceiros = False
TBAbrir!LOTE = "INV-" & Year(Txt_data.Value)
TBAbrir!Documento = "INV-" & Year(Txt_data.Value)
TBAbrir!Desenho = CodigoLista
Set TBCodigoDesc = CreateObject("adodb.recordset")
TBCodigoDesc.Open "Select * from projproduto where desenho = '" & CodigoLista & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBCodigoDesc.EOF = False Then
TBAbrir!Familia = TBCodigoDesc!Classe
End If

TBAbrir!Operacao = "SAIDA_INVENTÁRIO"
If TBItem!Saldo > QTBaixar Then
TBAbrir!Saida = QTBaixar
TBAbrir!Saida_PC = QTBaixar
TBAbrir!estoque_venda = QTBaixar
QTBaixado = QTBaixar
Else
TBAbrir!Saida = IIf(TBItem!Saldo > 0, TBItem!Saldo, QTBaixar)
TBAbrir!Saida_PC = IIf(TBItem!Saldo > 0, TBItem!Saldo, QTBaixar)
TBAbrir!estoque_venda = TBItem!Saldo
QTBaixado = TBItem!Saldo
End If

TBAbrir!VlrUnit = 0 'Format(IIf(IsNull(TBAbrir!valor_unitario), 0, TBAbrir!valor_unitario), "###,##0.0000000000")
TBAbrir!vlrTotal = 0 'Format(TBEstoque!VlrUnit * qtdeliberada, "###,##0.00")
TBAbrir!Descricao = TBCodigoDesc!Descricao 'Produto
TBAbrir!Data = Txt_data.Value
TBAbrir!Responsavel = pubUsuario
TBAbrir!IDEstoque = TBItem!IDEstoque

TBAbrir.Update
TBAbrir.Close
TBCodigoDesc.Close
QTBaixar = QTBaixar - QTBaixado
Else
ProcEntradaInventario
'Se precisar faz uma entrada com nova RE pra ter saldo pra baixar
GoTo 1
Exit Sub
End If
TBItem.Close
Loop

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnImportar_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente executar o inventário através de importação de planilha na data de " & Txt_data.Value & "?", vbYesNo, "CAPRIND v5.0") = vbNo Then
Passo2 = True
Passo1 = False
Passo3 = False
Exit Sub
End If

txtStatus.Text = ""

Cont = 0
'==============================================================
' Abrir o arquivo do Excel
'==============================================================
Set xlw = xl.Workbooks.Open(Arquivo)
'==============================================================
' Conta a quantidade de linhas na planilha
'==============================================================
Do While xlw.Application.Cells(Cont + 1, 1).Value <> ""
    Cont = Cont + 1
Loop
PBLista.Max = Cont
Contador = 2
'==============================================================
' Inicio do inventário
'==============================================================
Do While Contador <= Cont
'==============================================================
' Pega o codigo interno do item na lista
'==============================================================
CodigoLista = xlw.Application.Cells(Contador, 1).Value
'txtStatus.Text = txtStatus.Text & vbCrLf & Now & " Item :" & contador - 1 & " - Codigo :" & CodigoLista
Set TBCodigoDesc = CreateObject("adodb.recordset")
TBCodigoDesc.Open "Select * from projproduto where desenho = '" & CodigoLista & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBCodigoDesc.EOF = True Then
USMsgBox "Não existe cadastro do item " & CodigoLista & " no sistema", vbCritical, "CAPRIND v5.0"
GoTo Proximo
End If

 If CodigoLista <> "" Then
'==============================================================
' Pega dados do item na tabela excell
'==============================================================
'CodigoLista = xlw.Application.Cells(contador, 1).Value
Un = xlw.Application.Cells(Contador, 3).Value
Produto = xlw.Application.Cells(Contador, 2).Value
QTBaixar = xlw.Application.Cells(Contador, 5).Value
       
'txtStatus.Text = txtStatus.Text & " | inventário :" & QTBaixar
txtStatus.Text = Now & " Item :" & Contador - 1 & " - Codigo :" & CodigoLista & " | inventário :" & QTBaixar
       
'==============================================================
' Compara Estoque sistema com estoque planilha e toma decisão
' Se maior retira do estoque, se menor acrescenta no estoque
'==============================================================
Set TBEstoque = CreateObject("adodb.recordset")

'StrSql = "SELECT Sum(Estoque_disponivel) as Saldo from Estoque_produtos WHERE Data <  '" & Txt_data.Value & "' AND (Desenho = N'" & CodigoLista & "') AND (Estoque_disponivel > 0)"
'StrSql = "Select sum(EP.Entrada) as TTEntrada, sum(EP.Saida) as TTSaida,sum(EP.entrada-EP.saida) as Saldo,EP.idestoque, EP.Etiqueta, EP.Data, EP.LOTE, EP.Desenho, EP.Ref, EP.Descricao, EP.classe, EP.local_armaz, EP.Corrida, EP.Certificado, EP.Numero_serie, EP.Fornecedor, EP.Cliente, EP.Unidade, EP.estoque_real, EP.Qtde_Empenhada, EP.Estoque_disponivel, EP.estoque_real_PC, EP.valor_unitario, EP.Valor_Total, EP.Liberado, EP.Status from Estoque_produtos EP LEFT JOIN Projproduto_fabricante PFAB " _
'& "ON PFAB.Codproduto = EP.codproduto where EP.desenho = '" & CodigoLista & "' and EP.data >= '31/12/2010' and EP.Data <= '" & Txt_data.Value & "' AND  (ID_empresa = '1' or ID_empresa is null)  and EP.Estoque_disponivel > 0 and EP.bloqueado = 'False' group by EP.idestoque, EP.Etiqueta, EP.Data, EP.LOTE, EP.Desenho, EP.Ref, EP.Descricao, EP.classe, EP.local_armaz, EP.Corrida, EP.Certificado, EP.Numero_serie, EP.Fornecedor, EP.Cliente, EP.Unidade, EP.estoque_real, EP.Qtde_Empenhada, EP.Estoque_disponivel, EP.estoque_real_PC, EP.valor_unitario," _
'& "EP.Valor_Total, EP.Liberado, EP.Status order by Desenho"
StrSql = "Select sum(EM.entrada) - sum(EM.saida) as Saldo from Estoque_Movimentacao EM where EM.desenho = '" & CodigoLista & "' and EM.data >= '31/12/2000' and EM.Data <= '" & Txt_data.Value & "'"

'
'StrSql = "SELECT sum(entrada) - Sum(saida) as Saldo from Estoque_movimentacao where Desenho = '" & CodigoLista & "' And Data <= '" & Txt_data.Value & "'"
'Debug.print StrSql
TBEstoque.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly


If TBEstoque.EOF = False Then
'==============================================================
'Se no sistema estiver menor que a planilha faz uma saida
'==============================================================
'Debug.print TBEstoque!ttEntrada
'Debug.print TBEstoque!ttsaida
txtStatus.Text = txtStatus.Text & " | Sistema :" & IIf(IsNull(TBEstoque!Saldo = True), 0, TBEstoque!Saldo)

If TBEstoque!Saldo <> "" Then
EstoqueSaldo = TBEstoque!Saldo
Else
EstoqueSaldo = 0
End If

If EstoqueSaldo < QTBaixar Then
QTBaixar = QTBaixar - EstoqueSaldo
ProcEntradaInventario
QTBaixar = EstoqueSaldo + QTBaixar
End If
'==============================================================
'Se no sistema estiver maior que a planilha faz uma entrada
'==============================================================
If EstoqueSaldo > QTBaixar Then
QTBaixar = EstoqueSaldo - QTBaixar
ProcSaidaInventário
End If

If EstoqueSaldo = 0 Then
Conexao.Execute ("update Estoque_Controle Set Estoque_real = '0' where Desenho = '" & CodigoLista & "'")
Conexao.Execute ("update Estoque_Controle Set Estoque_venda = '0' where Desenho = '" & CodigoLista & "'")
Conexao.Execute ("update Estoque_Controle Set Qtde_fisica = '0' where Desenho = '" & CodigoLista & "'")
End If
End If
TBEstoque.Close
End If
Proximo:
Contador = Contador + 1
PBLista.Value = Contador
Loop

USMsgBox "Inventário por importação executado com sucesso!", vbInformation, "CAPRIND v5.0"
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnPlanilha_Click()
On Error GoTo tratar_erro
Dim Filter As String

'==============================================================
' Localiza arquivo excell
'==============================================================
Arquivo = ""
Filter = "(*.xlsx) | *.xlsx"
CD1.Filter = Filter
CD1.InitDir = App.Path
CD1.DefaultExt = "*.xlsx"
CD1.ShowOpen
Arquivo = CD1.filename
txtlocal = IIf(Arquivo = "", "", Arquivo)
Passo2 = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdAnterior_Click()
On Error GoTo tratar_erro
txtlocal.Text = ""

If Passo1 = False And Passo2 = False And Passo3 = False Then
cmdAnterior.Enabled = False
cmdProximo.Enabled = True

Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Passo1 = False
Passo2 = False
Passo3 = False
End If


If Passo1 = True And Passo2 = False And Passo3 = False Then
cmdAnterior.Enabled = True
cmdProximo.Enabled = True
Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
Passo1 = False
Passo2 = False
Passo3 = False
End If

If Passo1 = False And Passo2 = True And Passo3 = False Then
cmdAnterior.Visible = True
cmdProximo.Enabled = True

Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
Passo1 = False
Passo2 = False
Passo3 = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdProximo_Click()
On Error GoTo tratar_erro

'==================================================================
' Escolher o local
'==================================================================
If Passo1 = True And Passo2 = False And Passo3 = False Then
cmdAnterior.Enabled = True
cmdProximo.Visible = True
txtlocal.Text = ""
'btnPlanilha.SetFocus

Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
Passo1 = False
Passo2 = False
Passo3 = False
txtStatus.Text = Now & " - Inventário na data: " & Txt_data.Value
End If

'==================================================================
' Iniciar o inventário
'==================================================================
If txtlocal.Text <> "" And Passo1 = False And Passo2 = True And Passo3 = False Then
cmdAnterior.Visible = True
cmdProximo.Enabled = False
txtStatus.Text = txtStatus.Text & vbCrLf & Now & " - Com o arquivo : " & CD1.FileTitle

Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
Passo1 = False
Passo2 = True
Passo3 = False
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Txt_data.Value = Date
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False

Passo1 = False
Passo2 = False
Passo3 = False

cmdAnterior.Enabled = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_data_Change()
On Error GoTo tratar_erro

Passo1 = True
Passo2 = False
Passo3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_data_Click()
On Error GoTo tratar_erro

Passo1 = True
Passo2 = False
Passo3 = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
