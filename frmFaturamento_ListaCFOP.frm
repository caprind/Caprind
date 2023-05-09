VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFaturamento_ListaCFOP 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Faturamento | Lista de CFOP permitidas"
   ClientHeight    =   8100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNatureza 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1170
      TabIndex        =   15
      Top             =   630
      Width           =   5835
   End
   Begin VB.TextBox txtPIS 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   5160
      Width           =   7125
   End
   Begin VB.TextBox txtCofins 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   6150
      Width           =   7125
   End
   Begin VB.TextBox txtIPI 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   4200
      Width           =   7125
   End
   Begin VB.TextBox txtICMS 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   3240
      Width           =   7125
   End
   Begin DrawSuite2022.USButton btnCST 
      Height          =   645
      Left            =   4710
      TabIndex        =   7
      ToolTipText     =   "Carregar CST"
      Top             =   6870
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1138
      DibPicture      =   "frmFaturamento_ListaCFOP.frx":0000
      Caption         =   "Carregar CFOP"
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
      PicAlign        =   7
      ShowFocusRect   =   0   'False
      Theme           =   4
   End
   Begin DrawSuite2022.USLabel USLabel1 
      Height          =   585
      Left            =   300
      TabIndex        =   17
      Top             =   6870
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   1032
      Caption         =   $"frmFaturamento_ListaCFOP.frx":A1AD
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   128
      NoHTMLCaption   =   $"frmFaturamento_ListaCFOP.frx":A236
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   7695
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   688
      DibPicture      =   "frmFaturamento_ListaCFOP.frx":A2C1
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
      Icon            =   "frmFaturamento_ListaCFOP.frx":1446E
   End
   Begin MSComctlLib.ListView ListaCFOP 
      Height          =   1710
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Lista de CST's permitidas pela CFOP"
      Top             =   1200
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   3016
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483641
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   530
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "CFOP"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Natureza de operação"
         Object.Width           =   6615
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "ICMS"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "IPI"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "PIS"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "COFINS"
         Object.Width           =   1412
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Devolucao"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Remessa"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Retorno"
         Object.Width           =   529
      EndProperty
   End
   Begin DrawSuite2022.USTextBoxEx txtCFOP 
      Height          =   315
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "CFOP"
      Top             =   600
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      Alignment       =   2
      BackColorOver   =   12640511
      BorderColor     =   16764835
      BorderColorOver =   15971960
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "@.@@@"
      MaxLength       =   5
      Text            =   " .   "
   End
   Begin DrawSuite2022.USButton cmdNatureza 
      Height          =   315
      Left            =   7050
      TabIndex        =   16
      ToolTipText     =   "Buscar natureza de operação CFOP"
      Top             =   600
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      DibPicture      =   "frmFaturamento_ListaCFOP.frx":14788
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      PicAlign        =   0
      ShowFocusRect   =   0   'False
      ToolTipTitle    =   "CAPRIND v5.0"
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de CFOP(s)"
      Height          =   345
      Left            =   270
      TabIndex        =   10
      Top             =   990
      Width           =   3945
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "CFOP                         (Natureza da operação) digite o código ou a descrição para filtrar"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   330
      TabIndex        =   6
      Top             =   420
      Width           =   6405
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cobrança do Cofins"
      Height          =   195
      Left            =   3090
      TabIndex        =   5
      Top             =   5940
      Width           =   1425
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cobrança do PIS"
      Height          =   195
      Left            =   3195
      TabIndex        =   4
      Top             =   4950
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cobrança do IPI"
      Height          =   195
      Left            =   3210
      TabIndex        =   3
      Top             =   3990
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cobrança do ICMS"
      Height          =   195
      Left            =   3120
      TabIndex        =   2
      Top             =   3000
      Width           =   1365
   End
End
Attribute VB_Name = "frmFaturamento_ListaCFOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcBuscaCST()
On Error GoTo tratar_erro

If ListaCFOP.ListItems.Count = 0 Then Exit Sub
If Len(ListaCFOP.SelectedItem.ListSubItems.Item(3).Text) = 4 Then
TextoIcms = Right(ListaCFOP.SelectedItem.ListSubItems.Item(3).Text, 3)
Else
TextoIcms = Right(ListaCFOP.SelectedItem.ListSubItems.Item(3).Text, 2)
End If

Select Case TextoIcms
'=================================================================
'Lucro presumido e lucro real
'=================================================================
    Case "00": txtICMS.Text = "Tributada integralmente"
    Case "10": txtICMS.Text = "Tributada e com cobrança do ICMS por substituição"
    Case "20": txtICMS.Text = "Com redução de base de cálculo"
    Case "40": txtICMS.Text = "Isenta"
    Case "41": txtICMS.Text = "Não tributada"
    Case "50": txtICMS.Text = "Suspensão"
    Case "51": txtICMS.Text = "Diferimento"
    Case "60": txtICMS.Text = "ICMS cobrado anteriormente por substituição tributária"
    Case "70": txtICMS.Text = "Com redução de base de cálculo e cobrança do ICMS por substituição tributária"
    Case "90": txtICMS.Text = "Outras"
'=================================================================
'Simples nacional
'=================================================================
    Case "101": txtICMS.Text = "Tributada pelo Simples Nacional com permissão de crédito"
    Case "102": txtICMS.Text = "Tributada pelo Simples Nacional sem permissão de crédito"
    Case "103": txtICMS.Text = "Isenção do ICMS no Simples Nacional para faixa de receita bruta"
    Case "201": txtICMS.Text = "Tributada pelo Simples Nacional com permissão de crédito e com cobrança do ICMS por Substituição Tributária"
    Case "202": txtICMS.Text = "Tributada pelo Simples Nacional sem permissão de crédito e com cobrança do ICMS por Substituição Tributária"
    Case "203": txtICMS.Text = "Isenção do ICMS nos Simples Nacional para faixa de receita bruta e com cobrança do ICMS por Substituição Tributária"
    Case "300": txtICMS.Text = "Imune"
    Case "400": txtICMS.Text = "Não tributada pelo Simples Nacional"
    Case "500": txtICMS.Text = "ICMS cobrado anteriormente por substituição tributária (substituído) ou por antecipação"
    Case "900": txtICMS.Text = "Outros"
End Select

'=================================================================
'CST do IPI
'=================================================================
TextoIPI = ListaCFOP.SelectedItem.ListSubItems.Item(4).Text

Select Case TextoIPI
    Case "00": txtIPI.Text = "Entrada com recuperação de crédito"
    Case "01": txtIPI.Text = "Entrada tributada com alíquota zero"
    Case "02": txtIPI.Text = "Entrada isenta"
    Case "03": txtIPI.Text = "Entrada não-tributada"
    Case "04": txtIPI.Text = "Entrada imune"
    Case "05": txtIPI.Text = "Entrada com suspensão"
    Case "49": txtIPI.Text = "Outras entradas"
    Case "50": txtIPI.Text = "Saída tributada"
    Case "51": txtIPI.Text = "Saída tributada com alíquota zero"
    Case "52": txtIPI.Text = "Saída isenta"
    Case "53": txtIPI.Text = "Saída não-tributada"
    Case "54": txtIPI.Text = "Saída imune"
    Case "55": txtIPI.Text = "Saída com suspensão"
    Case "99": txtIPI.Text = "Outras saídas"
'=================================================================
End Select

'=================================================================
'CST do PIS
'=================================================================
TextoPIS = ListaCFOP.SelectedItem.ListSubItems.Item(5).Text

Select Case TextoPIS
    Case "01": txtPIS.Text = "Operação Tributável (base de cálculo = valor da operação alíquota normal (cumulativo/não cumulativo))"
    Case "02": txtPIS.Text = "Operação Tributável (base de cálculo = valor da operação (alíquota diferenciada))"
    Case "03": txtPIS.Text = "Operação Tributável (base de cálculo = quantidade vendida x alíquota por unidade de produto)"
    Case "04": txtPIS.Text = "Operação Tributável (tributação monofásica (alíquota zero))"
    Case "05": txtPIS.Text = "Operação Tributável (Substituição Tributária)"
    Case "06": txtPIS.Text = "Operação Tributável (alíquota zero)"
    Case "07": txtPIS.Text = "Operação Isenta da Contribuição"
    Case "08": txtPIS.Text = "Operação Sem Incidência da Contribuição"
    Case "09": txtPIS.Text = "Operação com Suspensão da Contribuição"
    Case "49": txtPIS.Text = "Outras Operações de Saída"
    Case "50": txtPIS.Text = "Operação com Direito a Crédito - Vinculada Exclusivamente a Receita Tributada no Mercado Interno"
    Case "51": txtPIS.Text = "Operação com Direito a Crédito - Vinculada Exclusivamente a Receita Não Tributada no Mercado Interno"
    Case "52": txtPIS.Text = "Operação com Direito a Crédito - Vinculada Exclusivamente a Receita de Exportação"
    Case "53": txtPIS.Text = "Operação com Direito a Crédito - Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno"
    Case "54": txtPIS.Text = "Operação com Direito a Crédito - Vinculada a Receitas Tributadas no Mercado Interno e de Exportação"
    Case "55": txtPIS.Text = "Operação com Direito a Crédito - Vinculada a Receitas Não-Tributadas no Mercado Interno e de Exportação"
    Case "56": txtPIS.Text = "Operação com Direito a Crédito - Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno, e de Exportação"
    Case "60": txtPIS.Text = "Crédito Presumido - Operação de Aquisição Vinculada Exclusivamente a Receita Tributada no Mercado Interno"
    Case "61": txtPIS.Text = "Crédito Presumido - Operação de Aquisição Vinculada Exclusivamente a Receita Não Tributada no Mercado Interno"
    Case "62": txtPIS.Text = "Crédito Presumido - Operação de Aquisição Vinculada Exclusivamente a Receita de Exportação"
    Case "63": txtPIS.Text = "Crédito Presumido - Operação de Aquisição Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno"
    Case "64": txtPIS.Text = "Crédito Presumido - Operação de Aquisição Vinculada a Receitas Tributadas no Mercado Interno e de Exportação"
    Case "65": txtPIS.Text = "Crédito Presumido - Operação de Aquisição Vinculada a Receitas Não-Tributadas no Mercado Interno e de Exportação"
    Case "66": txtPIS.Text = "Crédito Presumido - Operação de Aquisição Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno, e de Exportação"
    Case "67": txtPIS.Text = "Crédito Presumido - Outras Operações"
    Case "70": txtPIS.Text = "Operação de Aquisição sem Direito a Crédito"
    Case "71": txtPIS.Text = "Operação de Aquisição com Isenção"
    Case "72": txtPIS.Text = "Operação de Aquisição com Suspensão"
    Case "73": txtPIS.Text = "Operação de Aquisição a Alíquota Zero"
    Case "74": txtPIS.Text = "Operação de Aquisição sem Incidência da Contribuição"
    Case "75": txtPIS.Text = "Operação de Aquisição por Substituição Tributária"
    Case "98": txtPIS.Text = "Outras Operações de Entrada"
    Case "99": txtPIS.Text = "Outras Operações"
End Select

'=================================================================
'CST do Cofins
'=================================================================
TextoCofins = ListaCFOP.SelectedItem.ListSubItems.Item(6).Text

Select Case TextoCofins

    Case "01": txtCofins.Text = "Operação Tributável (base de cálculo = valor da operação alíquota normal (cumulativo/não cumulativo))"
    Case "02": txtCofins.Text = "Operação Tributável (base de cálculo = valor da operação (alíquota diferenciada))"
    Case "03": txtCofins.Text = "Operação Tributável (base de cálculo = quantidade vendida x alíquota por unidade de produto)"
    Case "04": txtCofins.Text = "Operação Tributável (tributação monofásica (alíquota zero))"
    Case "05": txtCofins.Text = "Operação Tributável (Substituição Tributária)"
    Case "06": txtCofins.Text = "Operação Tributável (alíquota zero)"
    Case "07": txtCofins.Text = "Operação Isenta da Contribuição"
    Case "08": txtCofins.Text = "Operação Sem Incidência da Contribuição"
    Case "09": txtCofins.Text = "Operação com Suspensão da Contribuição"
    Case "49": txtCofins.Text = "Outras Operações de Saída"
    Case "50": txtCofins.Text = "Operação com Direito a Crédito - Vinculada Exclusivamente a Receita Tributada no Mercado Interno"
    Case "51": txtCofins.Text = "Operação com Direito a Crédito - Vinculada Exclusivamente a Receita Não Tributada no Mercado Interno"
    Case "52": txtCofins.Text = "Operação com Direito a Crédito - Vinculada Exclusivamente a Receita de Exportação"
    Case "53": txtCofins.Text = "Operação com Direito a Crédito - Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno"
    Case "54": txtCofins.Text = "Operação com Direito a Crédito - Vinculada a Receitas Tributadas no Mercado Interno e de Exportação"
    Case "55": txtCofins.Text = "Operação com Direito a Crédito - Vinculada a Receitas Não-Tributadas no Mercado Interno e de Exportação"
    Case "56": txtCofins.Text = "Operação com Direito a Crédito - Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno, e de Exportação"
    Case "60": txtCofins.Text = "Crédito Presumido - Operação de Aquisição Vinculada Exclusivamente a Receita Tributada no Mercado Interno"
    Case "61": txtCofins.Text = "Crédito Presumido - Operação de Aquisição Vinculada Exclusivamente a Receita Não Tributada no Mercado Interno"
    Case "62": txtCofins.Text = "Crédito Presumido - Operação de Aquisição Vinculada Exclusivamente a Receita de Exportação"
    Case "63": txtCofins.Text = "Crédito Presumido - Operação de Aquisição Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno"
    Case "64": txtCofins.Text = "Crédito Presumido - Operação de Aquisição Vinculada a Receitas Tributadas no Mercado Interno e de Exportação"
    Case "65": txtCofins.Text = "Crédito Presumido - Operação de Aquisição Vinculada a Receitas Não-Tributadas no Mercado Interno e de Exportação"
    Case "66": txtCofins.Text = "Crédito Presumido - Operação de Aquisição Vinculada a Receitas Tributadas e Não-Tributadas no Mercado Interno, e de Exportação"
    Case "67": txtCofins.Text = "Crédito Presumido - Outras Operações"
    Case "70": txtCofins.Text = "Operação de Aquisição sem Direito a Crédito"
    Case "71": txtCofins.Text = "Operação de Aquisição com Isenção"
    Case "72": txtCofins.Text = "Operação de Aquisição com Suspensão"
    Case "73": txtCofins.Text = "Operação de Aquisição a Alíquota Zero"
    Case "74": txtCofins.Text = "Operação de Aquisição sem Incidência da Contribuição"
    Case "75": txtCofins.Text = "Operação de Aquisição por Substituição Tributária"
    Case "98": txtCofins.Text = "Outras Operações de Entrada"
    Case "99": txtCofins.Text = "Outras Operações"
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaCFOP()
On Error GoTo tratar_erro

ListaCFOP.ListItems.Clear

Set TBLISTA = CreateObject("adodb.recordset")


TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic


If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
        With ListaCFOP.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!CFOP), "", TBLISTA!CFOP)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Descricao), "", TBLISTA!Descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!ICMS), "", TBLISTA!ICMS)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!IPI), "", TBLISTA!IPI)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!PIS), "", TBLISTA!PIS)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Cofins), "", TBLISTA!Cofins)
            .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Remessa), "", TBLISTA!Remessa)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!retorno), "", TBLISTA!retorno)
            .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Devolucao), "", TBLISTA!Devolucao)
        End With
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnCST_Click()
On Error GoTo tratar_erro

If ListaCFOP.ListItems.Count = 0 Then Exit Sub

With frmFaturamento_Prod_Serv
    .Txt_ID_CFOP_prod.Text = ListaCFOP.SelectedItem
    .Txt_CFOP_prod.Text = ListaCFOP.SelectedItem.ListSubItems.Item(1).Text
    .Txt_natureza_operacao_prod = ListaCFOP.SelectedItem.ListSubItems.Item(2).Text
    .txtCST_ICMS.Text = ListaCFOP.SelectedItem.ListSubItems.Item(3).Text
    .txtCST_IPI.Text = ListaCFOP.SelectedItem.ListSubItems.Item(4).Text
    .txtCST_PIS.Text = ListaCFOP.SelectedItem.ListSubItems.Item(5).Text
    .txtCST_Cofins.Text = ListaCFOP.SelectedItem.ListSubItems.Item(6).Text
End With

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdNatureza_Click()
On Error GoTo tratar_erro

txtICMS = ""
txtIPI.Text = ""
txtPIS.Text = ""
txtCofins = ""

StrSql = "select NOP.IDCountCfop AS ID,NOP.id_CFOP as CFOP ,NOP.txt_Descricao as Descricao,NOP.Remessa,NOP.Retorno,NOP.Devolucao, CST.CST_ICMS as ICMS, CST.CST_IPI as IPI,CST.CST_PIS as PIS,CST.CST_Cofins as Cofins from tbl_NaturezaOperacao_CST CST right join tbl_NaturezaOperacao NOP on NOP.IDCountCfop = CST.ID_CFOP WHERE NOP.DtValidacao IS NOT NULL and NOP.txt_Descricao like '%" & txtNatureza & "%'"
'Debug.print StrSql

If Len(txtNatureza.Text) > 1 Then
ProcCarregaListaCFOP
End If

'txtCFOP.Text = " .  "

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
If frmFaturamento_Prod_Serv.Txt_ID_CFOP_prod.Text <> "" Then
ID_CFOP = frmFaturamento_Prod_Serv.Txt_ID_CFOP_prod.Text
txtCFOP.Text = frmFaturamento_Prod_Serv.Txt_CFOP_prod.Text
txtNatureza.Text = frmFaturamento_Prod_Serv.Txt_natureza_operacao_prod.Text

ProcBuscaCST
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listacfop_DblClick()
On Error GoTo tratar_erro

If ListaCFOP.ListItems.Count = 0 Then Exit Sub

With frmFaturamento_Prod_Serv
    .Txt_ID_CFOP_prod.Text = ListaCFOP.SelectedItem
    .Txt_CFOP_prod.Text = ListaCFOP.SelectedItem.ListSubItems.Item(1).Text
    .Txt_natureza_operacao_prod = ListaCFOP.SelectedItem.ListSubItems.Item(2).Text
    .txtCST_ICMS.Text = ListaCFOP.SelectedItem.ListSubItems.Item(3).Text
    .txtCST_IPI.Text = ListaCFOP.SelectedItem.ListSubItems.Item(4).Text
    .txtCST_PIS.Text = ListaCFOP.SelectedItem.ListSubItems.Item(5).Text
    .txtCST_Cofins.Text = ListaCFOP.SelectedItem.ListSubItems.Item(6).Text
    .chkRemessa.Value = IIf(ListaCFOP.SelectedItem.ListSubItems.Item(7).Text = "True", 1, 0)
    .chkRetorno.Value = IIf(ListaCFOP.SelectedItem.ListSubItems.Item(8).Text = "True", 1, 0)
    Devolucao = ListaCFOP.SelectedItem.ListSubItems.Item(9).Text
End With

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaCFOP_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

ID_CFOP = ListaCFOP.SelectedItem
txtNatureza.Text = ListaCFOP.SelectedItem.ListSubItems.Item(2).Text
ProcBuscaCST

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtCFOP_Change()
On Error GoTo tratar_erro

txtNatureza.Text = ""

txtICMS = ""
txtIPI.Text = ""
txtPIS.Text = ""
txtCofins = ""

StrSql = "select NOP.IDCountCfop AS ID,NOP.id_CFOP as CFOP ,NOP.txt_Descricao as Descricao,NOP.Remessa,NOP.Retorno,NOP.Devolucao, CST.CST_ICMS as ICMS, CST.CST_IPI as IPI,CST.CST_PIS as PIS,CST.CST_Cofins as Cofins from tbl_NaturezaOperacao_CST CST right join tbl_NaturezaOperacao NOP on NOP.IDCountCfop = CST.ID_CFOP WHERE NOP.DtValidacao IS NOT NULL and NOP.ID_CFOP like '%" & txtCFOP & "'"
'Debug.print StrSql

ProcCarregaListaCFOP

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

