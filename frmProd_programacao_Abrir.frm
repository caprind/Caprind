VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmProd_programacao_abrir 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PCP - Programação - Localizar"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmProd_programacao_Abrir.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmProd_programacao_Abrir.frx":1042
   MousePointer    =   99  'Custom
   ScaleHeight     =   2445
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   55
      TabIndex        =   6
      Top             =   900
      Width           =   8805
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   4620
         TabIndex        =   9
         Top             =   210
         Width           =   3975
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   180
            Width           =   1155
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   180
            Width           =   1275
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio frase"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1470
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.ComboBox cmbfiltrarpor 
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
         Height          =   330
         ItemData        =   "frmProd_programacao_Abrir.frx":134C
         Left            =   180
         List            =   "frmProd_programacao_Abrir.frx":136E
         MousePointer    =   99  'Custom
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4365
      End
      Begin VB.TextBox txtTexto 
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
         TabIndex        =   1
         ToolTipText     =   "``"
         Top             =   1050
         Width           =   8415
      End
      Begin VB.ComboBox cmbfamilia 
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
         Height          =   330
         Left            =   180
         MousePointer    =   99  'Custom
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Status."
         Top             =   1050
         Width           =   8415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
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
         Left            =   3645
         TabIndex        =   8
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
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
         Left            =   1935
         TabIndex        =   7
         Top             =   180
         Width           =   840
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   885
      Left            =   60
      TabIndex        =   10
      Top             =   0
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   1561
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
      ButtonCaption1  =   "Filtrar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Filtrar (F2)"
      ButtonKey1      =   "2"
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
      ButtonHeight2   =   48
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "14"
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
      ButtonKey4      =   "15"
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
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "16"
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
      ButtonLeft5     =   110
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   5850
         Top             =   150
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmProd_programacao_Abrir.frx":13DE
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmProd_programacao_abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbFamilia_Click()
On Error GoTo tratar_erro

If cmbfamilia <> "" Then txtTexto = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

With cmbfamilia
    If cmbfiltrarpor = "Família" Or cmbfiltrarpor = "Status" Or cmbfiltrarpor = "Tipo" Then
        txtTexto.Visible = False
        .Visible = True
        .Clear
        .AddItem ""
        If cmbfiltrarpor = "Família" Then
            ProcCarregaComboFamilia cmbfamilia, "familia <> 'Null' and vendas = 'true'", False
        ElseIf cmbfiltrarpor = "Status" Then
                .AddItem "Aberta"
                .AddItem "Produzindo"
                .AddItem "Concluída"
                .AddItem "Cancelada"
                .AddItem "Aguardando"
                .AddItem "Entregue"
            Else
                .AddItem "Componente"
                .AddItem "Subconjunto"
                .AddItem "Produto final"
        End If
    Else
        txtTexto.Visible = True
        .Visible = False
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case KeyCode
    'Case vbKeyF1: Ajuda
    Case vbKeyF2: ProcFiltrar
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

cmbfiltrarpor = "Posto de trabalho"
Optinicio.Value = True
txtTexto.Visible = True

ProcCarregaToolBar1 Me, 15195, 4, True


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

With frmProd_programacao
    .Lista.ListItems.Clear
    .FormulaRel_Ordem_Programacao = ""
    If txtTexto.Visible = True And txtTexto <> "" Or cmbfamilia.Visible = True And cmbfamilia <> "" Then
        If cmbfiltrarpor = "Código interno" Then
            If Optinicio.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN producao on PCP_programacao_ordem.ordem = producao.ordem) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where producao.desenho like '" & txtTexto.Text & "%' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{producao.desenho} like '" & txtTexto & "*'"
            End If
            If Optmeio.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN producao on PCP_programacao_ordem.ordem = producao.ordem) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where producao.desenho like '%" & txtTexto.Text & "%' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{producao.desenho} like '*" & txtTexto & "*'"
            End If
            If Optfim.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN producao on PCP_programacao_ordem.ordem = producao.ordem) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where producao.desenho like '%" & txtTexto.Text & "' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{producao.desenho} like '*" & txtTexto & "'"
            End If
        End If
        If cmbfiltrarpor = "Código referencia" Then
            If Optinicio.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN producao on PCP_programacao_ordem.ordem = producao.ordem) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where producao.N_Referencia like '" & txtTexto.Text & "%' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{producao.N_Referencia} like '" & txtTexto & "*'"
            End If
            If Optmeio.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN producao on PCP_programacao_ordem.ordem = producao.ordem) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where producao.N_Referencia like '%" & txtTexto.Text & "%' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{producao.N_Referencia} like '*" & txtTexto & "*'"
            End If
            If Optfim.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN producao on PCP_programacao_ordem.ordem = producao.ordem) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where producao.N_Referencia like '%" & txtTexto.Text & "' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{producao.N_Referencia} like '*" & txtTexto & "'"
            End If
        End If
        If cmbfiltrarpor = "Descrição" Then
            If Optinicio.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN producao on PCP_programacao_ordem.ordem = producao.ordem) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where producao.produto like '" & txtTexto.Text & "%' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{producao.produto} like '" & txtTexto & "*'"
            End If
            If Optmeio.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN producao on PCP_programacao_ordem.ordem = producao.ordem) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where producao.produto like '%" & txtTexto.Text & "%' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{producao.produto} like '*" & txtTexto & "*'"
            End If
            If Optfim.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN producao on PCP_programacao_ordem.ordem = producao.ordem) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where producao.produto like '%" & txtTexto.Text & "' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{producao.produto} like '*" & txtTexto & "'"
            End If
        End If
        If cmbfiltrarpor = "Cliente" Then
            If Optinicio.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN producao on PCP_programacao_ordem.ordem = producao.ordem) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where producao.Cliente like '" & txtTexto.Text & "%' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{producao.cliente} like '" & txtTexto & "*'"
            End If
            If Optmeio.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN producao on PCP_programacao_ordem.ordem = producao.ordem) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where producao.Cliente like '%" & txtTexto.Text & "%' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{producao.cliente} like '*" & txtTexto & "*'"
            End If
            If Optfim.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN producao on PCP_programacao_ordem.ordem = producao.ordem) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where producao.Cliente like '%" & txtTexto.Text & "' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{producao.cliente} like '*" & txtTexto & "'"
            End If
        End If
        If cmbfiltrarpor = "Ordem" Then
            If Optinicio.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from PCP_programacao INNER JOIN PCP_programacao_ordem on PCP_programacao.id = PCP_programacao_ordem.idprogramacao where PCP_programacao_ordem.Ordem like '" & txtTexto.Text & "%' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{producao.ordem} like '" & txtTexto & "*'"
            End If
            If Optmeio.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from PCP_programacao INNER JOIN PCP_programacao_ordem on PCP_programacao.id = PCP_programacao_ordem.idprogramacao where PCP_programacao_ordem.Ordem like '%" & txtTexto.Text & "%' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{producao.ordem} like '*" & txtTexto & "*'"
            End If
            If Optfim.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from PCP_programacao INNER JOIN PCP_programacao_ordem on PCP_programacao.id = PCP_programacao_ordem.idprogramacao where PCP_programacao_ordem.Ordem like '%" & txtTexto.Text & "' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{producao.ordem} like '*" & txtTexto & "'"
            End If
        End If
        If cmbfiltrarpor = "OS" Then
            If Optinicio.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN ordemservico on PCP_programacao_ordem.OS = ordemservico.idproducao) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where ordemservico.idproducao like '" & cmbfamilia.Text & "%' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{ordemservico.idproducao} like '" & txtTexto & "*'"
            End If
            If Optmeio.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN ordemservico on PCP_programacao_ordem.OS = ordemservico.idproducao) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where ordemservico.idproducao = '%" & cmbfamilia.Text & "%' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{ordemservico.idproducao} like '*" & txtTexto & "*'"
            End If
            If Optfim.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN ordemservico on PCP_programacao_ordem.OS = ordemservico.idproducao) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where ordemservico.idproducao = '%" & cmbfamilia.Text & "' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{ordemservico.idproducao} like '*" & txtTexto & "'"
            End If
        End If
        If cmbfiltrarpor = "Posto de trabalho" Then
            If Optinicio.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from PCP_programacao INNER JOIN CadMaquinas on PCP_programacao.idmaquina = CadMaquinas.idmaquina where CadMaquinas.maquina = '" & txtTexto.Text & "%' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{CadMaquinas.maquina} like '" & txtTexto & "*'"
            End If
            If Optmeio.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from PCP_programacao INNER JOIN CadMaquinas on PCP_programacao.idmaquina = CadMaquinas.idmaquina where CadMaquinas.maquina = '%" & txtTexto.Text & "%' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{CadMaquinas.maquina} like '*" & txtTexto & "*'"
            End If
            If Optfim.Value = True Then
                .StrSql_Ordem_programacao = "Select PCP_programacao.* from PCP_programacao INNER JOIN CadMaquinas on PCP_programacao.idmaquina = CadMaquinas.idmaquina where CadMaquinas.maquina = '%" & txtTexto.Text & "' order by pcp_programacao.ID desc"
                .FormulaRel_Ordem_Programacao = "{CadMaquinas.maquina} like '*" & txtTexto & "'"
            End If
        End If
        If cmbfiltrarpor = "Status" Then
            .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN producao on PCP_programacao_ordem.ordem = producao.ordem) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where producao.status = '" & cmbfamilia.Text & "' order by pcp_programacao.ID desc"
            .FormulaRel_Ordem_Programacao = "{producao.status} = '" & cmbfamilia & "'"
        End If
        If cmbfiltrarpor = "Tipo" Then
            .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN producao on PCP_programacao_ordem.ordem = producao.ordem) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where producao.Tipo = '" & cmbfamilia.Text & "' order by pcp_programacao.ID desc"
            .FormulaRel_Ordem_Programacao = "{producao.Tipo} = '" & cmbfamilia & "'"
        End If
        If cmbfiltrarpor = "Família" Then
            Desenho = ""
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from projproduto where Classe = '" & cmbfamilia & "' order by classe", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                Do While TBAbrir.EOF = False
                    If Desenho <> TBAbrir!Desenho Then
                        .StrSql_Ordem_programacao = "Select PCP_programacao.* from (PCP_programacao_ordem INNER JOIN producao on PCP_programacao_ordem.ordem = producao.ordem) INNER JOIN PCP_programacao on PCP_programacao.id = pcp_programacao_Ordem.IDprogramacao where producao.desenho = '" & TBAbrir!Desenho & "' order by pcp_programacao.ID desc"
                        .ProcCarregaLista
                    End If
                    Desenho = TBAbrir!Desenho
                Loop
                .FormulaRel_Ordem_Programacao = "{projproduto.familia} = '" & cmbfamilia & "'"
            End If
            TBAbrir.Close
            Exit Sub
        End If
    Else
        .StrSql_Ordem_programacao = "Select * from PCP_programacao where idmaquina <> 0 order by ID desc"
    End If
    .ProcCarregaLista
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro
    
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTexto_Change()
On Error GoTo tratar_erro

If txtTexto <> "" Then cmbfamilia.ListIndex = -1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal Key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcFiltrar
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

