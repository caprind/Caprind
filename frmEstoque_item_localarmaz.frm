VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmEstoque_item_localarmaz 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'None
   Caption         =   "Estoque - Movimentação - Alterar local de armazenamento"
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   11160
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   847
      DibPicture      =   "frmEstoque_item_localarmaz.frx":0000
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
      Icon            =   "frmEstoque_item_localarmaz.frx":1CAD
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   24
      Top             =   6285
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   714
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   60
      TabIndex        =   11
      Top             =   5340
      Width           =   11025
      Begin VB.TextBox txtUn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   23
         ToolTipText     =   "Unidade."
         Top             =   390
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox txtSaldoEmpenho 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   9450
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "Saldo."
         Top             =   420
         Width           =   1305
      End
      Begin VB.TextBox txtEmpenho 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   6780
         Locked          =   -1  'True
         TabIndex        =   15
         ToolTipText     =   "Empenho."
         Top             =   420
         Width           =   1305
      End
      Begin VB.TextBox txtEmpenhoAlterar 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   8115
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Empenho a alterar."
         Top             =   420
         Width           =   1305
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
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
         Left            =   9900
         TabIndex        =   18
         Top             =   210
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empenho"
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
         Left            =   7095
         TabIndex        =   16
         Top             =   210
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empenho alterar"
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
         Left            =   8175
         TabIndex        =   13
         Top             =   210
         Width           =   1185
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   60
      TabIndex        =   5
      Top             =   510
      Width           =   11025
      Begin DrawSuite2022.USButton btnAlterar 
         Height          =   555
         Left            =   8880
         TabIndex        =   26
         ToolTipText     =   "Gravar novo local de armazenamento"
         Top             =   210
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   979
         DibPicture      =   "frmEstoque_item_localarmaz.frx":1FC7
         Caption         =   "Gravar local de armazenamento"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         Theme           =   3
      End
      Begin VB.TextBox txtSaldoPC 
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
         Left            =   12810
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "0,0000"
         ToolTipText     =   "Saldo peças."
         Top             =   1530
         Width           =   1275
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   7305
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Saldo."
         Top             =   420
         Width           =   1275
      End
      Begin VB.TextBox txtQtdePC 
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
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "0,0000"
         ToolTipText     =   "Quantidade de peças."
         Top             =   1530
         Width           =   1275
      End
      Begin VB.TextBox txtQtdeAlterarPC 
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
         Left            =   11504
         TabIndex        =   4
         Text            =   "0,0000"
         ToolTipText     =   "Quantidade a alterar peças."
         Top             =   1530
         Width           =   1275
      End
      Begin VB.TextBox txtQtdeAlterar 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   6000
         TabIndex        =   2
         ToolTipText     =   "Quantidade a alterar."
         Top             =   420
         Width           =   1275
      End
      Begin VB.TextBox txtQtde 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   4695
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade."
         Top             =   420
         Width           =   1275
      End
      Begin VB.ComboBox cmbLocal_armaz 
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
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Local de armazenamento."
         Top             =   420
         Width           =   3780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo PÇ"
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
         Left            =   13125
         TabIndex        =   22
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo final"
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
         Left            =   7590
         TabIndex        =   21
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade PÇ"
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
         Left            =   10305
         TabIndex        =   10
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde alterar PÇ"
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
         Left            =   11610
         TabIndex        =   9
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo atual"
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
         Left            =   4935
         TabIndex        =   8
         Top             =   210
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde transferir"
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
         Left            =   6090
         TabIndex        =   7
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Novo local de armazenamento"
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
         Left            =   990
         TabIndex        =   6
         Top             =   210
         Width           =   2160
      End
   End
   Begin DrawSuite2022.USFlexGrid FlexGrid 
      Height          =   3915
      Left            =   60
      TabIndex        =   14
      Top             =   1410
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   6906
      BackColorEvenRows=   14737632
      BackColorSelected1=   16643298
      BackColorSelected2=   16643298
      FocusRectColor  =   15181413
      GridColor       =   16247519
      HeaderGradientColor2=   12632256
      ColumnSeparatorColorOver=   15048022
      ColumnSeparatorColorDown=   15381630
      ProgressBarColor2=   2277891
      ForeColorSelected=   0
      AllowColumnResizing=   -1  'True
      CaptionHeight   =   28
      ColumnHeaderSmall=   -1  'True
      ColumnSort      =   -1  'True
      Editable        =   -1  'True
      FocusRowHighlightKeepTextForeColor=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFormatString=   $"frmEstoque_item_localarmaz.frx":3C74
      MinRowHeight    =   14
      TotalLineShow   =   0   'False
   End
End
Attribute VB_Name = "frmEstoque_item_localarmaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim desenhoLocal As String
Dim IdMovimentacao As Long

Private Sub ProcSalvar()
On Error GoTo tratar_erro

With frmestoque_item
    Acao = "salvar"
    If cmbLocal_armaz = "" Then
        NomeCampo = "o local de armazenamento"
        ProcVerificaAcao
        cmbLocal_armaz.SetFocus
        Exit Sub
    End If
    
    Qtd = IIf(txtQtdeAlterar = "", 0, txtQtdeAlterar)
    Qtde = IIf(txtQtde = "", 0, txtQtde)
    ValorPagar = IIf(txtQtdeAlterarPC = "", 0, txtQtdeAlterarPC)
    ValorTotalPagar = IIf(txtQtdePC = "", 0, txtQtdePC)
    If txtQtdeAlterar.Locked = False Then
        If txtQtdeAlterar = "" Or Qtd <= 0 Then
            NomeCampo = "a quantidade a alterar"
            ProcVerificaAcao
            txtQtdeAlterar.SetFocus
            Exit Sub
        End If
        
        If Qtd > Qtde Then
            USMsgBox "A quantidade a ser alterada não pode ser maior que a quantidade em estoque.", vbExclamation, "CAPRIND v5.0"
            txtQtdeAlterar.SetFocus
            Exit Sub
        End If
    Else
        If txtQtdeAlterarPC = "" Or ValorPagar <= 0 Then
            NomeCampo = "a quantidade peça a alterar"
            ProcVerificaAcao
            txtQtdeAlterar.SetFocus
            Exit Sub
        End If
        
        If ValorPagar > ValorTotalPagar Then
            USMsgBox "A quantidade peça a ser alterada não pode ser maior que a quantidade peça em estoque.", vbExclamation, "CAPRIND v5.0"
            txtQtdeAlterar.SetFocus
            Exit Sub
        End If
    End If
    
    quantidade = IIf(txtEmpenhoAlterar = "", 0, txtEmpenhoAlterar)
    If quantidade > Qtde Then
        USMsgBox ("o empenho a ser alterado não pode ser maior que a quantidade alterada."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    
    Saldo_Atual = IIf(txtSaldo = "", 0, txtSaldo)
    Saldo_Anterior = IIf(txtSaldoEmpenho = "", 0, txtSaldoEmpenho)
    If Saldo_Anterior > Saldo_Atual Then
        USMsgBox ("O saldo do empenho não pode ser maior que o saldo do RE alterado."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    
    Contador = 0
    With FlexGrid
        For InitFor = 1 To (.rows)
            If .CellChecked(Contador, 2) = True And .CellText(Contador, 9) = "" Then
                USMsgBox ("Informe a quantidade a ser alterada para o(s) empenho(s) selecionado(s)."), vbExclamation, "CAPRIND v5.0"
                Exit Sub
            End If
            
            If .CellChecked(Contador, 2) = False And .CellText(Contador, 9) <> "" Then
                USMsgBox ("Selecione o(s) empenho(s) com quantidade na lista."), vbExclamation, "CAPRIND v5.0"
                Exit Sub
            End If
            Contador = Contador + 1
        Next InitFor
    End With

    If USMsgBox("Deseja realmente alterar o local de armazenamento deste RE?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Set TBComponente = CreateObject("adodb.recordset")
        TBComponente.Open "SELECT * from estoque_controle where idestoque = " & .Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
        If TBComponente.EOF = False Then
            If TBComponente!local_armaz = cmbLocal_armaz Then
                USMsgBox ("Não é permitido alterar o local de armazenamento deste RE, pois o local de armazenamento continua o mesmo."), vbExclamation, "CAPRIND v5.0"
                TBComponente.Close
                Exit Sub
            End If

            If Qtde = Qtd Then
                TBComponente!local_armaz = cmbLocal_armaz
            
                Conexao.Execute "Update ECR Set ECR.local_armaz = '" & cmbLocal_armaz & "' from Estoque_controle_recebimento ECR INNER JOIN estoque_movimentacao EM on EM.IDEstoque_recebimento = ECR.ID where EM.IDestoque = " & TBComponente!IDEstoque
                Conexao.Execute "Update Estoque_fisico set Local_armaz = '" & cmbLocal_armaz & "' where IDestoque = " & TBComponente!IDEstoque
                Conexao.Execute "Update EF set EF.Local_armaz = '" & cmbLocal_armaz & "' from Estoque_fisico EF INNER JOIN Estoque_movimentacao EM ON EM.ID_inventario = EF.ID where EM.IDestoque = " & TBComponente!IDEstoque & " and EM.ID_inventario IS NOT NULL and EM.ID_inventario <> 0"
            Else
                IdMovimentacao = 0
                Conexao.Execute "Update EC set EC.estoque_real = '" & Qtde - Qtd & "' from Estoque_controle EC where EC.IDestoque = " & TBComponente!IDEstoque & ""
                ValorTotal = (Qtde - Qtd) * TBComponente!valor_unitario
                Conexao.Execute "Update EC set EC.Valor_total = '" & Replace(ValorTotal, ",", ".") & "' from Estoque_controle EC where EC.IDestoque = " & TBComponente!IDEstoque & ""

                'TBComponente!estoque_real = Qtde - Qtd
                'TBComponente!estoque_real_PC = ValorTotalPagar - ValorPagar
                'TBComponente!Valor_total = Format((Qtde - Qtd) * TBComponente!valor_unitario, "###,##0.00")
                procCriaMoventacao False
                procCriaRE
            End If
            TBComponente.Update
        End If
        TBComponente.Close
        
        USMsgBox ("Alteração efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = "Estoque/Movimentação"
        Evento = "Alterar local de armazenamento"
        ID_documento = .Lista.SelectedItem
        Documento = "Cód. interno: " & desenhoLocal
        Documento1 = ""
        ProcGravaEvento
        '==================================
        .Lista_Movimentacao.ListItems.Clear
        .ProcAtualizalista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
    End If
End With
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnAlterar_Click()
On Error GoTo tratar_erro


  ' If USMsgBox("Deseja realmente alterar o local de armazenamento?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    ProcSalvar
   ' End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long, vNewValue As String, Cancel As Boolean)
On Error GoTo tratar_erro

With FlexGrid
    If Col = 9 Then
        With FlexGrid
            Qtde = IIf(IsNumeric(vNewValue), vNewValue, 0)
            Qtd = .CellText(Row, 8)
            If Qtde > Qtd Then
                USMsgBox "O empenho a ser alterado não pode ser maior que o saldo disponivel.", vbInformation, "CAPRIND v5.0"
                vNewValue = ""
                Qtde = 0
            Else
                vNewValue = Format(vNewValue, "###,##0.0000")
            End If
        End With
    Else
        Qtde = 0
        If vNewValue = True Then Qtde = IIf(IsNumeric(.CellText(Row, 9)), .CellText(Row, 9), 0) Else .CellText(Row, 9) = ""

        Contador = 0
        For InitFor = 1 To (.rows)
            If Contador <> Row And .CellText(Contador, 9) <> "" And IsNumeric(.CellText(Contador, 9)) = True And .CellChecked(Contador, 2) = True Then Qtde = Qtde + .CellText(Contador, 9)
            Contador = Contador + 1
        Next InitFor
        
        txtEmpenhoAlterar = Format(Qtde, "###,##0.0000")
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo tratar_erro

If Col = 9 Then
    With FlexGrid
        .CellChecked(Row, 2) = False
        .CellText(Row, 9) = ""
        
        Qtde = 0
        Contador = 0
        For InitFor = 1 To (.rows)
            If Contador <> Row And .CellText(Contador, 9) <> "" And IsNumeric(.CellText(Contador, 9)) = True And .CellChecked(Contador, 2) = True Then Qtde = Qtde + .CellText(Contador, 9)
            Contador = Contador + 1
        Next InitFor
        
        txtEmpenhoAlterar = Format(Qtde, "###,##0.0000")
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaComboLA cmbLocal_armaz, False, False

Set TBNivel15 = CreateObject("adodb.recordset")

'StrSql = "SELECT Estoque_Real, Desenho, estoque_real_PC, Un FROM estoque_controle WHERE idestoque = " & frmestoque_item.Lista.SelectedItem
StrSql = "SELECT estoque_disponivel as Estoque_Real, Desenho, estoque_real_PC, Unidade FROM Estoque_produtos WHERE idestoque = " & frmestoque_item.Lista.SelectedItem

'Debug.print StrSql

TBNivel15.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

If TBNivel15.EOF = False Then
    txtQtde = IIf(IsNull(TBNivel15!estoque_real), "0,0000", Format(TBNivel15!estoque_real, "###,##0.0000"))
    txtQtdePC = IIf(IsNull(TBNivel15!estoque_real_PC), "0,0000", Format(TBNivel15!estoque_real_PC, "###,##0.0000"))
    desenhoLocal = IIf(IsNull(TBNivel15!Desenho), 0, TBNivel15!Desenho)
    
    Set TBMaterial = CreateObject("adodb.recordset")
    TBMaterial.Open "Select Movimentar_estoque_pc from Empresa where Movimentar_estoque_pc = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBMaterial.EOF = False And txtQtdePC > 0 Then
        With txtQtdeAlterar
            .Locked = True
            .TabStop = False
        End With
        With txtQtdeAlterarPC
            .Locked = False
            .TabStop = True
        End With
    Else
        With txtQtdeAlterar
            .Locked = False
            .TabStop = True
        End With
        With txtQtdeAlterarPC
            .Locked = True
            .TabStop = False
        End With
    End If
    
    Permitido = False
    ProcCarregaLista
    If Permitido = True Then
        Height = 6700
    Else
        Height = 1800
    End If
    
    txtQtdeAlterar = IIf(IsNull(TBNivel15!estoque_real), "0,0000", Format(TBNivel15!estoque_real, "###,##0.0000"))
    txtQtdeAlterarPC = IIf(IsNull(TBNivel15!estoque_real_PC), "0,0000", Format(TBNivel15!estoque_real_PC, "###,##0.0000"))
    
    txtUN = TBNivel15!Unidade
End If
TBNivel15.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtEmpenhoAlterar_Change()
On Error GoTo tratar_erro

Qtd = IIf(txtEmpenho = "", 0, txtEmpenho)
Qtde = IIf(txtEmpenhoAlterar = "", 0, txtEmpenhoAlterar)
txtSaldoEmpenho = Format(Qtd - Qtde, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtQtdeAlterar_Change()
On Error GoTo tratar_erro

If txtQtdeAlterar.Text <> "" Then
    VerifNumero = txtQtdeAlterar
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtdeAlterar.Text = ""
        txtQtdeAlterar.SetFocus
        Exit Sub
    End If
    If txtQtdeAlterarPC.Locked = True Then txtQtdeAlterarPC = Format(FunCalculaQtdePC(desenhoLocal, txtQtdeAlterar, True, txtUN), "###,##0.0000")
End If

Qtd = IIf(txtQtde = "", 0, txtQtde)
Qtde = IIf(txtQtdeAlterar = "", 0, txtQtdeAlterar)
txtSaldo = Format(Qtd - Qtde, "###,##0.0000")

procLiberaLista IIf(IsNumeric(txtQtdeAlterar), txtQtdeAlterar, 0), txtQtde

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtQtdeAlterar_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtQtdeAlterar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtQtdeAlterar_LostFocus()
On Error GoTo tratar_erro

txtQtdeAlterar = Format(txtQtdeAlterar, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtQtdeAlterarPC_Change()
On Error GoTo tratar_erro

If txtQtdeAlterarPC.Text <> "" Then
    VerifNumero = txtQtdeAlterarPC.Text
    ProcVerificaNumero
    If VerifNumero = False Then
        txtQtdeAlterarPC.Text = ""
        txtQtdeAlterarPC.SetFocus
        Exit Sub
    End If
    If txtQtdeAlterarPC.Locked = False Then txtQtdeAlterar = Format(FunCalculaQtdePC(desenhoLocal, txtQtdeAlterarPC, False, txtUN), "###,##0.0000")
End If

Qtd = IIf(txtQtdePC = "", 0, txtQtdePC)
Qtde = IIf(txtQtdeAlterarPC = "", 0, txtQtdeAlterarPC)
txtSaldoPC = Format(Qtd - Qtde, "###,##0.0000")

If txtQtdeAlterarPC.Locked = False Then procLiberaLista IIf(IsNumeric(txtQtdeAlterar), txtQtdeAlterar, 0), txtQtde

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtQtdeAlterarPC_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtQtdeAlterarPC

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtQtdeAlterarPC_LostFocus()
On Error GoTo tratar_erro

txtQtdeAlterarPC = Format(txtQtdeAlterarPC, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub


Public Sub procCriaRE()
On Error GoTo tratar_erro

'Cria nova RE
Set TBAfericao = CreateObject("adodb.recordset")
TBAfericao.Open "SELECT * from estoque_controle", Conexao, adOpenKeyset, adLockOptimistic
TBAfericao.AddNew
TBAfericao!LOTE = TBComponente!LOTE
TBAfericao!Desenho = TBComponente!Desenho
TBAfericao!Descricao = TBComponente!Descricao
TBAfericao!estoque_real = txtQtdeAlterar
TBAfericao!estoque_venda = txtQtdeAlterar
TBAfericao!Estoque_minimo = TBComponente!Estoque_minimo
TBAfericao!Un = TBComponente!Un
TBAfericao!Data = Date
TBAfericao!Responsavel = pubUsuario
TBAfericao!Fornecedor = TBComponente!Fornecedor
TBAfericao!Certificado = TBComponente!Certificado
TBAfericao!Classe = TBComponente!Classe
TBAfericao!descricaotecnica = TBComponente!descricaotecnica
TBAfericao!peso_unit = TBComponente!peso_unit
TBAfericao!local_armaz = cmbLocal_armaz
TBAfericao!imagem = TBComponente!imagem
TBAfericao!Pedido = TBComponente!Pedido
TBAfericao!status = "ENTRADA_LOCAL_DE_ARMAZENAMENTO"
TBAfericao!Qtde = txtQtdeAlterar
TBAfericao!NF = TBComponente!NF
TBAfericao!ID_Cliente = TBComponente!ID_Cliente
TBAfericao!Cliente = TBComponente!Cliente
TBAfericao!Ref = TBComponente!Ref
TBAfericao!Corrida = TBComponente!Corrida
TBAfericao!emissaonf = TBComponente!emissaonf
TBAfericao!Consignacao = TBComponente!Consignacao
TBAfericao!valor_unitario = TBComponente!valor_unitario
TBAfericao!Valor_total = Format(TBComponente!valor_unitario * TBAfericao!estoque_real, "###,##0.00")
TBAfericao!Desenho_sucata = TBComponente!Desenho_sucata
TBAfericao!idLote_sucata = TBComponente!idLote_sucata
TBAfericao!qtde_fisica = Qtd
TBAfericao!Etiqueta = TBComponente!Etiqueta
TBAfericao!ID_empresa = TBComponente!ID_empresa
TBAfericao!estoque_real_PC = txtQtdeAlterarPC
TBAfericao!qtde_fisica_PC = txtQtdeAlterarPC
TBAfericao!Dt_ult_mov = TBComponente!Dt_ult_mov
TBAfericao!Bloqueado = TBComponente!Bloqueado
TBAfericao!obs_Status = TBComponente!obs_Status
TBAfericao!resp_Status = TBComponente!resp_Status
TBAfericao!Numero_serie = TBComponente!Numero_serie
TBAfericao!Tipodest_NFcons = TBComponente!Tipodest_NFcons
TBAfericao!Serie = TBComponente!Serie
TBAfericao.Update

procCriaMoventacao True 'Cria movimentação
ProcEmpenho 'Cria os empenhos
TBAfericao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procCriaMoventacao(EntradaMov As Boolean)
On Error GoTo tratar_erro

'Cria movimentação
Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
TBCotacao.AddNew
TBCotacao!Destino = "Interno"
TBCotacao!Terceiros = False
TBCotacao!LOTE = TBComponente!LOTE
TBCotacao!Documento = TBComponente!LOTE
TBCotacao!Desenho = TBComponente!Desenho
TBCotacao!Familia = TBComponente!Classe
If EntradaMov = True Then
    TBCotacao!Operacao = "ENTRADA_LOCAL_DE_ARMAZENAMENTO"
    TBCotacao!IDEstoque = TBAfericao!IDEstoque 'ID da RE nova
    TBCotacao!Entrada = txtQtdeAlterar
    TBCotacao!Entrada_PC = txtQtdeAlterarPC
    TBCotacao!IdTrocaLocal = IdMovimentacao
Else
    TBCotacao!Operacao = "SAIDA_LOCAL_DE_ARMAZENAMENTO"
    TBCotacao!IDEstoque = TBComponente!IDEstoque 'ID da RE antiga
    TBCotacao!Saida = txtQtdeAlterar
    TBCotacao!Saida_PC = txtQtdeAlterarPC
End If
TBCotacao!VlrUnit = TBComponente!valor_unitario
TBCotacao!vlrTotal = Format(TBCotacao!VlrUnit * Qtd, "###,##0.00")
TBCotacao!Descricao = TBComponente!Descricao
TBCotacao!Data = Date
TBCotacao!Responsavel = pubUsuario
TBCotacao!estoque_venda = txtQtdeAlterar
TBCotacao.Update
If EntradaMov = False Then IdMovimentacao = TBCotacao!IDoperacao
TBCotacao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub ProcEmpenho()
On Error GoTo tratar_erro

Contador = 0
With FlexGrid
    For InitFor = 1 To (.rows)
        If .CellChecked(Contador, 2) = True Then
            quantidade = IIf(.CellText(Contador, 9) = "", 0, .CellText(Contador, 9))
            If .CellText(Contador, 1) = "O" Then
                'Empenho de ordem
                Set TBExecucao = CreateObject("adodb.recordset")
                TBExecucao.Open "SELECT *, (quantidade - ISNULL(qtde_saida, 0)) as qtdeEmpenho FROM Producao_NF_Consignada WHERE id = " & .CellText(Contador, 0) & " AND (quantidade - ISNULL(qtde_saida, 0)) > 0", Conexao, adOpenKeyset, adLockOptimistic
                If TBExecucao.EOF = False Then
                    If IsNull(TBExecucao!Qtde_saida) = True Or TBExecucao!Qtde_saida = 0 And quantidade >= TBExecucao!qtdeEmpenho Then
                        'Se não tiver saida apenas troca o RE
                        TBExecucao!IDEstoque = TBAfericao!IDEstoque
                        TBExecucao!IdAntigoLocal = Null
                    Else
                        'Cria novo empenho para nova RE
                        procCriaEmpenhoOrdem quantidade 'Salva o saldo que sobrou
                        TBExecucao!quantidade = TBExecucao!quantidade - quantidade
                    End If
                    TBExecucao.Update
                End If
                TBExecucao.Close
            Else
                'Empenho de pedido interno
                Set TBExecucao = CreateObject("adodb.recordset")
                TBExecucao.Open "Select *, (qtde_empenhada - ISNULL(qtde_saida, 0)) as qtdeEmpenho from Estoque_Controle_Empenho_Vendas where id_estoque = " & .CellText(Contador, 0) & " and (qtde_empenhada - ISNULL(qtde_saida, 0)) > 0", Conexao, adOpenKeyset, adLockOptimistic
                If TBExecucao.EOF = False Then
                    If IsNull(TBExecucao!Qtde_saida) = True Or TBExecucao!Qtde_saida = 0 And quantidade >= TBExecucao!qtdeEmpenho Then
                        'Se não tiver saida apenas troca o RE
                        TBExecucao!ID_estoque = TBAfericao!IDEstoque
                        TBExecucao!IdAntigoLocal = Null
                    Else
                        'Cria novo empenho para nova RE com quantidade que sobrou
                        procCriaEmpenhoVendas quantidade 'Salva o saldo que sobrou
                        TBExecucao!Qtde_empenhada = TBExecucao!Qtde_empenhada - quantidade
                    End If
                    TBExecucao.Update
                End If
                TBExecucao.Close
            End If
        End If
        Contador = Contador + 1
    Next InitFor
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procCriaEmpenhoVendas(qtdeEmpenhoVendas As Double)
On Error GoTo tratar_erro

Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "SELECT * FROM Estoque_Controle_Empenho_Vendas", Conexao, adOpenKeyset, adLockOptimistic
TBCFOP.AddNew
TBCFOP!ID_estoque = TBAfericao!IDEstoque
TBCFOP!ID_carteira = TBExecucao!ID_carteira
TBCFOP!Qtde_empenhada = qtdeEmpenhoVendas
TBCFOP!Qtde_saida = 0
TBCFOP!Data = TBExecucao!Data
TBCFOP!Responsavel = TBExecucao!Responsavel
TBCFOP!ID_Faturamento = TBExecucao!ID_Faturamento
TBCFOP!IdAntigoLocal = TBExecucao!ID
TBCFOP.Update
TBCFOP.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procCriaEmpenhoOrdem(qtdeEmpenhoOrdem As Double)
On Error GoTo tratar_erro

Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "SELECT * FROM Producao_NF_Consignada", Conexao, adOpenKeyset, adLockOptimistic
TBCFOP.AddNew
TBCFOP!IDEstoque = TBAfericao!IDEstoque
TBCFOP!Ordem = TBExecucao!Ordem
TBCFOP!Codinterno = TBExecucao!Codinterno
TBCFOP!quantidade = qtdeEmpenhoOrdem
TBCFOP!Quantidade_PC = FunCalculaQtdePC(desenhoLocal, qtdeEmpenhoOrdem, True, txtUN)
TBCFOP!Qtde_saida = 0
TBCFOP!Qtde_saida_PC = 0
TBCFOP!Data = TBExecucao!Data
TBCFOP!Responsavel = TBExecucao!Responsavel
TBCFOP!IdAntigoLocal = TBExecucao!ID
TBCFOP.Update
TBCFOP.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

Qtde = 0
FlexGrid.Clear
Contador = 1
Permitido = False
Set TBLISTA = CreateObject("adodb.recordset")
StrSql = "Select PNFC.ID, PNFC.Data as DataEmpenho, PNFC.Responsavel as ResponsavelEmpenho, PNFC.IDestoque, PNFC.Quantidade, PNFC.Qtde_saida, P.* from Producao_NF_Consignada PNFC INNER JOIN Producao P ON PNFC.Ordem = P.Ordem where PNFC.IDestoque = " & frmestoque_item.Lista.SelectedItem & " and P.Status <> 'Cancelada' and P.DtValidacao_custo IS NULL and PNFC.Quantidade - ISNULL(PNFC.Qtde_saida, 0) > 0"
'Debug.print StrSql

TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    Permitido = True
    Do While TBLISTA.EOF = False
        With FlexGrid
            L = .AddItem(TBLISTA!ID)
            .RowData(L) = Contador
            .CellText(L, 1) = "O"
            .CellText(L, 3) = IIf(IsNull(TBLISTA!Ordem), "", TBLISTA!Ordem)
            .CellText(L, 4) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
            .CellText(L, 5) = IIf(IsNull(TBLISTA!Quant), "", TBLISTA!Quant)
            valor = IIf(IsNull(TBLISTA!quantidade), 0, TBLISTA!quantidade)
            .CellText(L, 6) = Format(valor, "###,##0.0000")
            Valor1 = IIf(IsNull(TBLISTA!Qtde_saida), 0, TBLISTA!Qtde_saida)
            .CellText(L, 7) = Format(Valor1, "###,##0.0000")
            .CellText(L, 8) = Format(valor - Valor1, "###,##0.0000")
            Qtde = Qtde + (valor - Valor1)
        End With
        Contador = Contador + 1
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close

CamposFiltro = "EE.ID, EE.Data as Dataemp, EE.Responsavel as Respemp, VC.CODIGO, EE.ID_estoque, Sum(EE.Qtde_empenhada) as qtdeliberar, Sum(EE.Qtde_saida) as qtdeliberada, VC.Unidade, VP.Ncotacao, VP.Revisao, OPCP.Requisicaotexto, VP.vend_int, VP.vend_ext, VC.Qtde_produzir, VC.qtdeexpedida, VC.Prazofinal, VC.Desenho, VC.Rev_codinterno, VC.N_Referencia, VC.descricao_tecnica, VP.Cliente"
CamposFiltroGrupo = "EE.ID, EE.Data, EE.Responsavel, VC.CODIGO, EE.ID_estoque, VC.Unidade, VP.Ncotacao, VP.Revisao, OPCP.Requisicaotexto, VP.vend_int, VP.vend_ext, VC.Qtde_produzir, VC.qtdeexpedida, VC.Prazofinal, VC.Desenho, VC.Rev_codinterno, VC.N_Referencia, VC.descricao_tecnica, VP.Cliente"
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select " & CamposFiltro & " from ((vendas_carteira VC INNER JOIN Estoque_Controle_Empenho_Vendas EE ON VC.Codigo = EE.ID_carteira) LEFT JOIN Vendas_proposta VP ON VP.Cotacao = VC.Cotacao) LEFT JOIN Outros_SolicitacaoPCP OPCP ON OPCP.ID = VC.ID_Solicitacao group by " & CamposFiltroGrupo & " HAVING EE.ID_estoque = " & frmestoque_item.Lista.SelectedItem & " and Sum(EE.Qtde_empenhada) - Sum(ISNULL(EE.Qtde_saida, 0)) > 0", Conexao, adOpenKeyset, adLockReadOnly
If TBLISTA.EOF = False Then
    Permitido = True
    Do While TBLISTA.EOF = False
        With FlexGrid
            L = .AddItem(TBLISTA!ID)
            .RowData(L) = Contador
            .CellText(L, 1) = "P"
            If IsNull(TBLISTA!Ncotacao) = False Then
                .CellText(L, 3) = TBLISTA!Ncotacao
                .CellText(L, 4) = IIf(IsNull(TBLISTA!Cliente), "", TBLISTA!Cliente)
            Else
                .CellText(L, 3) = TBLISTA!Requisicaotexto
            End If
            
            .CellText(L, 5) = IIf(IsNull(TBLISTA!Qtde_produzir), "", TBLISTA!Qtde_produzir)
            valor = IIf(IsNull(TBLISTA!qtdeliberar), 0, TBLISTA!qtdeliberar)
            .CellText(L, 6) = Format(valor, "###,##0.0000")
            Valor1 = IIf(IsNull(TBLISTA!qtdeliberada), 0, TBLISTA!qtdeliberada)
            .CellText(L, 7) = Format(Valor1, "###,##0.0000")
            .CellText(L, 8) = Format(valor - Valor1, "###,##0.0000")
            Qtde = Qtde + (valor - Valor1)
        End With
        TBLISTA.MoveNext
        Contador = Contador + 1
    Loop
End If
TBLISTA.Close

FlexGrid.Redraw = True
txtEmpenho = Format(Qtde, "###,##0.0000")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Public Sub procLiberaLista(qtdeAlterar As Double, qtdeNormal As Double)
On Error GoTo tratar_erro

If qtdeAlterar < qtdeNormal Then
    Contador = 0
    With FlexGrid
        .Enabled = True
        For InitFor = 1 To (.rows)
            .CellChecked(Contador, 2) = False
            .CellText(Contador, 9) = ""
            Contador = Contador + 1
        Next InitFor
        txtEmpenhoAlterar = "0,0000"
    End With
Else
    Contador = 0
    Qtd = 0
    With FlexGrid
        .Enabled = False
        For InitFor = 1 To (.rows)
            .CellChecked(Contador, 2) = True
            .CellText(Contador, 9) = .CellText(Contador, 8)
            Qtd = Qtd + .CellText(Contador, 8)
            Contador = Contador + 1
        Next InitFor
        txtEmpenhoAlterar = Format(Qtd, "###,##0.0000")
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
