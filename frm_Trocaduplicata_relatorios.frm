VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frm_Trocaduplicata_relatorios 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Desconto de duplicata - Relatórios"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox optPeriodo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "De :"
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
      Left            =   840
      TabIndex        =   2
      Top             =   2730
      Width           =   615
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   11085
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancela e fecha formulário (Esc)"
      Top             =   165
      Width           =   570
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   675
      Left            =   55
      TabIndex        =   8
      Top             =   2490
      Width           =   4845
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   3360
         TabIndex        =   4
         ToolTipText     =   "Data de emissão da nota fiscal."
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   487456769
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   1470
         TabIndex        =   3
         ToolTipText     =   "Data de emissão da nota fiscal."
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   487456769
         CurrentDate     =   39057
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Até :"
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
         Left            =   2895
         TabIndex        =   9
         Top             =   240
         Width           =   360
      End
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   55
      TabIndex        =   11
      Top             =   0
      Width           =   4845
      _ExtentX        =   8546
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
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   4020
      Top             =   1470
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frm_Trocaduplicata_relatorios.frx":0000
      Count           =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1485
      Left            =   55
      TabIndex        =   6
      Top             =   990
      Width           =   4845
      Begin VB.ComboBox Cmb_empresa 
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
         ItemData        =   "frm_Trocaduplicata_relatorios.frx":22CB
         Left            =   180
         List            =   "frm_Trocaduplicata_relatorios.frx":22CD
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa."
         Top             =   390
         Width           =   4485
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
         ItemData        =   "frm_Trocaduplicata_relatorios.frx":22CF
         Left            =   180
         List            =   "frm_Trocaduplicata_relatorios.frx":22DC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Opções para filtro."
         Top             =   990
         Width           =   4485
      End
      Begin VB.Label Label27 
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
         Left            =   2070
         TabIndex        =   10
         Top             =   180
         Width           =   735
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
         Left            =   1995
         TabIndex        =   7
         Top             =   780
         Width           =   840
      End
   End
End
Attribute VB_Name = "frm_Trocaduplicata_relatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Borderô" Then
    optPeriodo.Value = 0
    optPeriodo.Enabled = False
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
Else
    optPeriodo.Enabled = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If cmbfiltrarpor = "Borderô" Then
    If frm_trocaduplicata.txtBordero = "" Then
        NomeCampo = "o borderô"
        Acao = "visualizar impressão"
        ProcVerificaAcao
        Unload Me
        Exit Sub
    End If
    If frm_trocaduplicata.Lista.ListItems.Count = 0 Then Exit Sub
    NomeRel = "Contas_receber_bordero.rpt"
    ProcImprimirRel "{troca_titulo.id} = " & frm_trocaduplicata.txtBordero, ""
End If
With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
ProcGravarDataFiltroRel msk_fltInicio, msk_fltFim, True, Cmb_empresa.ItemData(Cmb_empresa.ListIndex), ""

If cmbfiltrarpor = "Operações efetuadas detalhado" Then
    NomeRel = "Contas_receber_duplicatasoperacoes_detalhado.rpt"
    If optPeriodo.Value = 1 Then
        ProcImprimirRel "{troca_titulo.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {troca_titulo.Data}>=Date(" & Year(msk_fltInicio.Value) & "," & Month(msk_fltInicio.Value) & "," & Day(msk_fltInicio.Value) & ") and {troca_titulo.Data}<= Date(" & _
                        Year(msk_fltFim.Value) & "," & Month(msk_fltFim.Value) & "," & Day(msk_fltFim.Value) & ") and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""
    Else
        ProcImprimirRel "{troca_titulo.ID_empresa} = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""
    End If
End If
If cmbfiltrarpor = "Operações efetuadas resumido" Then
    Conexao.Execute "DELETE from Troca_titulo_relatorio where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'"
    If optPeriodo.Value = 1 Then
        Data_liquidadas = "(tbl_contas_receber.Data_pagamento) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
        Data_vencidas = "(tbl_contas_receber.Vencimento) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
        Data_vencidas_recompra = "(tbl_contas_receber.Vencimento) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
        Data_vencer = "tbl_contas_receber.Vencimento > '" & Format(msk_fltFim.Value, "Short Date") & "'"
    Else
        Data_liquidadas = "(tbl_contas_receber.Data_pagamento) <= '" & Format(Date, "Short Date") & "'"
        Data_vencidas = "(tbl_contas_receber.Vencimento) < '" & Format(Date, "Short Date") & "'"
        Data_vencidas_recompra = "(tbl_contas_receber.Vencimento) < '" & Format(Date, "Short Date") & "'"
        Data_vencer = "tbl_contas_receber.Vencimento > '" & Format(Date, "Short Date") & "'"
    End If
    Campos = "tbl_contas_receber.Data_pagamento, tbl_contas_receber.Vencimento, tbl_contas_receber.valortitulorecebido, troca_titulo_valores.valor_enviado, troca_titulo.Local_troca"
    
    'Liquidadas
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select " & Campos & " from (tbl_contas_receber inner join troca_titulo_valores on tbl_contas_receber.IDIntconta = troca_titulo_valores.n_conta) INNER JOIN troca_titulo on troca_titulo.id = troca_titulo_valores.IDduplicata where tbl_contas_receber.status = 'DUPLICATA DESCONTADA LIQUIDADA' and " & Data_liquidadas & " and troca_titulo.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by tbl_contas_receber.Data_pagamento, tbl_contas_receber.IDIntconta", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            ProcGravarLiquidados
            TBAbrir.MoveNext
        Loop
    End If
    
    'Vencidas
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select " & Campos & " from (tbl_contas_receber inner join troca_titulo_valores on tbl_contas_receber.IDIntconta = troca_titulo_valores.n_conta) INNER JOIN troca_titulo on troca_titulo.id = troca_titulo_valores.IDduplicata where tbl_contas_receber.status = 'DUPLICATA DESCONTADA EM ABERTO' and  " & Data_vencidas & " and troca_titulo.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by tbl_contas_receber.Vencimento, tbl_contas_receber.IDIntconta", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            ProcGravarVencidas
            TBAbrir.MoveNext
        Loop
    End If
    
    'Vencidas a recompra
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select " & Campos & " from ((tbl_contas_receber inner join troca_titulo_valores on tbl_contas_receber.IDIntconta = troca_titulo_valores.n_conta) INNER JOIN troca_titulo on troca_titulo.id = troca_titulo_valores.IDduplicata) inner join tbl_ContasPagar on tbl_contas_receber.IDIntconta = tbl_ContasPagar.IdContaReceber where tbl_ContasPagar.status = 'N' and tbl_contas_receber.status = 'DUPLICATA DESCONTADA RECOMPRADA' and " & Data_vencidas_recompra & " and troca_titulo.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by tbl_contas_receber.Vencimento, tbl_contas_receber.IDIntconta", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            ProcGravarVencidasRecompra
            TBAbrir.MoveNext
        Loop
    End If
    
    'Vencer
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select " & Campos & " from (tbl_contas_receber inner join troca_titulo_valores on tbl_contas_receber.IDIntconta = troca_titulo_valores.n_conta) INNER JOIN troca_titulo on troca_titulo.id = troca_titulo_valores.IDduplicata where tbl_contas_receber.status = 'DUPLICATA DESCONTADA EM ABERTO' and " & Data_vencer & " and troca_titulo.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) & " order by tbl_contas_receber.Vencimento, tbl_contas_receber.IDIntconta", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            ProcGravarVencer
            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
    
    NomeRel = "Contas_receber_duplicatasoperacoes_resumido.rpt"
    ProcImprimirRel "{Troca_titulo_relatorio.Responsavel} = '" & pubUsuario & "' and {Troca_titulo_relatorio.Modulo} = '" & Formulario & "' and {Producao_Relatorios_Total.Responsavel} = '" & pubUsuario & "' and {Producao_Relatorios_Total.Modulo} = '" & Formulario & "'", ""
End If

Exit Sub
tratar_erro:
    If Err.Number = "3075" Then
        Aspas = ("'")
        USMsgBox ("Não será possível visualizar o relatório, pois o local de troca está cadastrado com o seguinte caractere " & Aspas & "... no borderô " & TBAbrir!ID & "."), vbInformation, "CAPRIND v5.0"
        Exit Sub
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarLiquidados()
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Troca_titulo_relatorio where month (data) = " & Month(TBAbrir!Data_pagamento) & " and Local_troca = '" & TBAbrir!local_troca & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    TBGravar!Data_final = TBAbrir!Data_pagamento
Else
    TBGravar.AddNew
    If IsNull(TBGravar!Data_inicial) = True Or TBGravar!Data_inicial = "" Then TBGravar!Data_inicial = TBAbrir!Data_pagamento
    If IsNull(TBGravar!Data_final) = True Or TBGravar!Data_final = "" Then TBGravar!Data_final = TBAbrir!Data_pagamento
End If
TBGravar!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBGravar!Data = TBAbrir!Data_pagamento
TBGravar!local_troca = TBAbrir!local_troca
TBGravar!Liquidadas = TBGravar!Liquidadas + TBAbrir!valortitulorecebido
TBGravar!Responsavel = pubUsuario
TBGravar!Modulo = Formulario
TBGravar.Update
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarVencidas()
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Troca_titulo_relatorio where month (data) = " & Month(TBAbrir!Vencimento) & " and Local_troca = '" & TBAbrir!local_troca & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    TBGravar!Data_final = TBAbrir!Vencimento
Else
    TBGravar.AddNew
    If IsNull(TBGravar!Data_inicial) = True Or TBGravar!Data_inicial = "" Then TBGravar!Data_inicial = TBAbrir!Vencimento
    If IsNull(TBGravar!Data_final) = True Or TBGravar!Data_final = "" Then TBGravar!Data_final = TBAbrir!Vencimento
End If
TBGravar!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBGravar!Data = TBAbrir!Vencimento
TBGravar!local_troca = TBAbrir!local_troca
TBGravar!Vencidas = TBGravar!Vencidas + TBAbrir!valor_enviado
TBGravar!Responsavel = pubUsuario
TBGravar!Modulo = Formulario
TBGravar.Update
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarVencidasRecompra()
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Troca_titulo_relatorio where month (data) = " & Month(TBAbrir!Vencimento) & " and Local_troca = '" & TBAbrir!local_troca & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = False Then
    TBGravar!Data_final = TBAbrir!Vencimento
Else
    TBGravar.AddNew
    If IsNull(TBGravar!Data_inicial) = True Or TBGravar!Data_inicial = "" Then TBGravar!Data_inicial = TBAbrir!Vencimento
    If IsNull(TBGravar!Data_final) = True Or TBGravar!Data_final = "" Then TBGravar!Data_final = TBAbrir!Vencimento
End If
TBGravar!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBGravar!Data = TBAbrir!Vencimento
TBGravar!local_troca = TBAbrir!local_troca
TBGravar!Vencidasrec = TBGravar!Vencidasrec + TBAbrir!valor_enviado
TBGravar!Responsavel = pubUsuario
TBGravar!Modulo = Formulario
TBGravar.Update
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarVencer()
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Troca_titulo_relatorio order by Data", Conexao, adOpenKeyset, adLockOptimistic
Set TBGravar = CreateObject("adodb.recordset")
If TBFIltro.EOF = False Then
    TBGravar.Open "Select * from Troca_titulo_relatorio where month (data) = " & Month(TBFIltro!Data) & " and Local_troca = '" & TBAbrir!local_troca & "'", Conexao, adOpenKeyset, adLockOptimistic
    Data = TBFIltro!Data
Else
    TBGravar.Open "Select * from Troca_titulo_relatorio where month (data) = " & Month(TBAbrir!Vencimento) & " and Local_troca = '" & TBAbrir!local_troca & "'", Conexao, adOpenKeyset, adLockOptimistic
    Data = TBAbrir!Vencimento
End If
If TBGravar.EOF = False Then
    TBGravar!Data_final = TBAbrir!Vencimento
Else
    TBGravar.AddNew
    If IsNull(TBGravar!Data_inicial) = True Or TBGravar!Data_inicial = "" Then TBGravar!Data_inicial = TBAbrir!Vencimento
    If IsNull(TBGravar!Data_final) = True Or TBGravar!Data_final = "" Then TBGravar!Data_final = TBAbrir!Vencimento
End If
TBGravar!ID_empresa = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBGravar!Data = Data
TBGravar!local_troca = TBAbrir!local_troca
TBGravar!Vencer = TBGravar!Vencer + TBAbrir!valor_enviado
TBGravar!Responsavel = pubUsuario
TBGravar!Modulo = Formulario
TBGravar.Update
TBGravar.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF5: ProcImprimir
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

ProcCarregaToolBar1 Me, 4845, 5, True

cmbfiltrarpor = "Borderô"
msk_fltFim.Value = Date
msk_fltInicio.Value = Date
ProcCarregaComboEmpresa Cmb_empresa, False

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

Private Sub optPeriodo_Click()
On Error GoTo tratar_erro

If optPeriodo.Value = 1 Then
    Frame2.Enabled = True
    msk_fltInicio.SetFocus
Else
    Frame2.Enabled = False
    msk_fltInicio.Value = Date
    msk_fltFim.Value = Date
End If

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
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
