VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Begin VB.Form frmimprimir 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Visualizador de relatórios"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmimprimir.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   7140
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   1080
      ScreenWidth     =   2560
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10500
      FormWidthDT     =   15390
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15270
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer CrystalActiveXReportViewer1 
      CausesValidation=   0   'False
      Height          =   10035
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15165
      _cx             =   26749
      _cy             =   17701
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   0   'False
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1046
   End
End
Attribute VB_Name = "frmimprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CrystalActiveXReportViewer1_ExportButtonClicked(UseDefault As Boolean)
On Error GoTo tratar_erro

If Formulario = "Compras/Cotação" Or Formulario = "Compras/Pedido" And (Sit_REG = 1 Or Sit_REG = 2) Then
    If Formulario = "Compras/Cotação" Then
        TextoFiltro = "Cotação"
        With frmcompras_reqcot
            IDlista = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
            If Sit_REG = 1 Then Nome_anexo = Replace(.txtidcotacao, "/", "-") & ".pdf" Else Nome_anexo = Replace(.txtidcotacao, "/", "-") & " - " & frmCompras_reqcot_imprimir.cmbforn & ".pdf"
        End With
    Else
        TextoFiltro = "Pedido de compra"
        With frmCompras_Pedido
            IDlista = .Cmb_empresa.ItemData(.Cmb_empresa.ListIndex)
            Nome_anexo = Replace(.txtPedido, "/", "-") & ".pdf"
        End With
    End If
            
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Caminho from Empresa_armazenamento_PDF where ID_empresa = " & IDlista & " and Relatorio = '" & TextoFiltro & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If USMsgBox("Deseja utilizar a exportação em PDF configurada?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            UseDefault = False
            'Gerar arquivo em PDF
            Set crxExport = Report.ExportOptions
            
            If Len(TBAbrir!caminho) = 3 Then caminho = TBAbrir!caminho Else caminho = TBAbrir!caminho & "\"
            crxExport.DiskFileName = caminho & Nome_anexo
            
            crxExport.DestinationType = crEDTDiskFile
            crxExport.PDFExportAllPages = True
            crxExport.FormatType = crEFTPortableDocFormat
            Report.Export False
            
            USMsgBox ("Exportação efetuada com sucesso."), vbInformation, "CAPRIND v5.0"
        End If
    End If
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub CrystalActiveXReportViewer1_PrintButtonClicked(UseDefault As Boolean)
On Error GoTo tratar_erro

procImpressao

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

frmimprimir.WindowState = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procImpressao()
On Error GoTo tratar_erro

Select Case Formulario
    Case "PCP/Gerenciamento de ordem":
        With frmprod
            ProcSalvarViaOrdem IIf(.txtof = "", OF, .txtof), True
        End With
    Case "Compras/Programação":
        Set TBAliquota = CreateObject("adodb.recordset")
        TBAliquota.Open "Select * from Compras_Programa where id = " & frmCompras_programacao.txtId, Conexao, adOpenKeyset, adLockOptimistic
        If TBAliquota.EOF = False Then
            TBAliquota!via = IIf(IsNull(TBAliquota!via), 0, TBAliquota!via) + 1
            TBAliquota.Update
        End If
        TBAliquota.Close
End Select
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

