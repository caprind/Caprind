VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmVendas_Comissoes_Metas_Exportar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Exportar relatório"
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3105
      Left            =   330
      TabIndex        =   2
      Top             =   630
      Width           =   4155
      Begin DrawSuite2022.USButton btnExcell 
         Height          =   1125
         Left            =   210
         TabIndex        =   3
         Top             =   1770
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1984
         DibPicture      =   "frmVendas_Comissoes_Metas_Exportar.frx":0000
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Exportar para Excell"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   8
         PicSize         =   4
         PicSizeH        =   48
         PicSizeW        =   48
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
      End
      Begin DrawSuite2022.USButton btnPDF 
         Height          =   1185
         Left            =   210
         TabIndex        =   4
         Top             =   360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   2090
         DibPicture      =   "frmVendas_Comissoes_Metas_Exportar.frx":65C5
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Exportar para PDF"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   7
         PicSize         =   4
         PicSizeH        =   48
         PicSizeW        =   48
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   4065
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   767
      DibPicture      =   "frmVendas_Comissoes_Metas_Exportar.frx":1CFDD
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
      Icon            =   "frmVendas_Comissoes_Metas_Exportar.frx":1FB00
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmVendas_Comissoes_Metas_Exportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnExcell_Click()
On Error GoTo tratar_erro

With frmVendas_Comissoes_Metas.GridLista

        .PageSetup.Orientation = cellLandscape
        .PageSetup.HeaderAlignment = cellCenter
        .PageSetup.HeaderFont.Name = "Tahoma"
        .PageSetup.HeaderFont.size = 12
        .PageSetup.PrintCellBorders = True
        .PageSetup.PrintTitleColumns = True
        .PageSetup.PrintFixedColumn = True
        .PageSetup.PrintFixedRow = True
        .PageSetup.PrintGridlines = True
        .PageSetup.ThinBorderLineWidth = 1
        .PageSetup.Header = "Comissões por meta do mês de " & frmVendas_Comissoes_Metas.cmbdoMes.Text & " de " & frmVendas_Comissoes_Metas.cmbdoAno.Text
        .PageSetup.PaperSize = cellPaperA4
        .PageSetup.LeftMargin = 1
        .PageSetup.TopMargin = 2
        .PageSetup.RightMargin = 1
        .PageSetup.BottomMargin = 2
        .PageSetup.HeaderMargin = 1
        .PageSetup.FooterMargin = 1
        .PageSetup.Footer = "Pag &P de &N"
        .PageSetup.FooterAlignment = cellRight
        .PageSetup.FooterFont.Name = "Tahoma"
        .PageSetup.FooterFont.size = 8
        
    If .ExportToExcel("") Then
        USMsgBox "Relatório exportado com sucesso!", vbExclamation, "CAPRIND v5.0"
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnPDF_Click()
On Error GoTo tratar_erro

With frmVendas_Comissoes_Metas.GridLista

        .PageSetup.Orientation = cellLandscape
        .PageSetup.HeaderAlignment = cellCenter
        .PageSetup.HeaderFont.Name = "Tahoma"
        .PageSetup.HeaderFont.size = 12
        .PageSetup.PrintCellBorders = True
        .PageSetup.PrintTitleColumns = True
        .PageSetup.PrintFixedColumn = True
        .PageSetup.PrintFixedRow = True
        .PageSetup.PrintGridlines = True
        .PageSetup.ThinBorderLineWidth = 1
        .PageSetup.Header = "Comissões por meta do mês de " & frmVendas_Comissoes_Metas.cmbdoMes.Text & " no ano de " & frmVendas_Comissoes_Metas.cmbdoAno.Text
        .PageSetup.PaperSize = cellPaperA4
        .PageSetup.LeftMargin = 1
        .PageSetup.TopMargin = 2
        .PageSetup.RightMargin = 1
        .PageSetup.BottomMargin = 2
        .PageSetup.HeaderMargin = 1
        .PageSetup.FooterMargin = 1
        .PageSetup.Footer = "Pag &P de &N"
        .PageSetup.FooterAlignment = cellRight
        .PageSetup.FooterFont.Name = "Tahoma"
        .PageSetup.FooterFont.size = 8


    If .ExportToPDF("") Then
        USMsgBox "Relatório exportado com sucesso!", vbExclamation, "CAPRIND v5.0"
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
