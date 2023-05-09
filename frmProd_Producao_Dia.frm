VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmProd_Producao_Dia 
   Caption         =   "Form1"
   ClientHeight    =   10005
   ClientLeft      =   480
   ClientTop       =   405
   ClientWidth     =   15480
   ControlBox      =   0   'False
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
   MDIChild        =   -1  'True
   ScaleHeight     =   10005
   ScaleWidth      =   15480
   WindowState     =   2  'Maximizado
   Begin FlexCell.Grid Grid 
      Height          =   8415
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   14843
      BackColorBkg    =   16777215
      Cols            =   10
      DefaultFontSize =   8.25
      GridColor       =   14737632
      Rows            =   32
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      Height          =   1515
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15435
      Begin DrawSuite2022.USButton USButton1 
         Height          =   495
         Left            =   13890
         TabIndex        =   2
         Top             =   570
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   873
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         Caption         =   "Sair"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         ForeColorDown   =   16777215
         ForeColorOver   =   16777215
         GradientColor1  =   5263559
         GradientColor2  =   5263559
         GradientColor3  =   5263559
         GradientColor4  =   5263559
         GradientColorDisabled1=   13160660
         GradientColorDisabled2=   13160660
         GradientColorDisabled3=   13160660
         GradientColorDisabled4=   13160660
         GradientColorDown1=   4013465
         GradientColorDown2=   4013465
         GradientColorDown3=   4013465
         GradientColorDown4=   4013465
         GradientColorOver1=   4408288
         GradientColorOver2=   4408288
         GradientColorOver3=   4408288
         GradientColorOver4=   4408288
         Theme           =   4
      End
      Begin ActiveResizeCtl.ActiveResize ActiveResize1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         Resolution      =   99
         ScreenHeight    =   768
         ScreenWidth     =   1360
         ScreenHeightDT  =   1080
         ScreenWidthDT   =   1920
         AutoResizeOnLoad=   0   'False
         ApplicationName =   "Active Resize Control Professional"
         FormHeightDT    =   10470
         FormWidthDT     =   15600
         FormScaleHeightDT=   10005
         FormScaleWidthDT=   15480
         ResizeFormBackground=   -1  'True
         ResizePictureBoxContents=   -1  'True
      End
   End
End
Attribute VB_Name = "frmProd_Producao_Dia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ProcAjustaGrid()
On Error GoTo tratar_erro

    Grid.AllowUserPaste = cellTextOnly
    Grid.AllowUserResizing = False
    Grid.ExtendLastCol = True
    Grid.BoldFixedCell = False
    Grid.DisplayDateTimeMask = True
    Grid.DisplayFocusRect = False
    Grid.SelectionMode = cellSelectionByRow

    Grid.DrawMode = cellOwnerDraw
    Grid.Cols = 14
    Grid.Appearance = Flat
    Grid.ScrollBarStyle = Flat
    Grid.FixedRowColStyle = Flat
    Grid.Cell(0, 1).Text = "Data"
    Grid.Cell(0, 2).Text = "Recebido"
    Grid.Cell(0, 3).Text = "Programado"
    Grid.Cell(0, 4).Text = "Produzido"
    Grid.Cell(0, 5).Text = "% Atendimento"
    Grid.Cell(0, 6).Text = "Inspecionado"
    Grid.Cell(0, 7).Text = "% Inspecionado"
    Grid.Cell(0, 8).Text = "Aprovado"
    Grid.Cell(0, 9).Text = "% Aprovação"
    Grid.Cell(0, 10).Text = "Condicional"
    Grid.Cell(0, 11).Text = "% Condicional"
    Grid.Cell(0, 12).Text = "Reprovado"
    Grid.Cell(0, 13).Text = "% Reprova"
        
    Grid.Column(1).CellType = cellDate
    Grid.Column(1).Alignment = cellCenterCenter
    Grid.Column(1).FormatString = "DD/MM/YYYY"
    
    Grid.Column(2).CellType = cellTextBox
    Grid.Column(2).Alignment = cellCenterCenter
    
    Grid.Column(3).CellType = cellTextBox
    Grid.Column(3).Alignment = cellCenterCenter
    
    Grid.Column(4).CellType = cellTextBox
    Grid.Column(4).Alignment = cellCenterCenter
    
    Grid.Column(5).CellType = cellTextBox
    Grid.Column(5).Alignment = cellCenterCenter 'cellHyperLink
    
    Grid.Column(6).CellType = cellTextBox 'cellButton
    Grid.Column(6).Alignment = cellCenterCenter 'cellHyperLink
    
    Grid.Column(7).CellType = cellTextBox 'cellHyperLink
    Grid.Column(7).Alignment = cellCenterCenter 'cellHyperLink
    
    Grid.Column(8).CellType = cellTextBox 'cellHyperLink
    Grid.Column(8).Alignment = cellCenterCenter 'cellHyperLink
    
    Grid.Column(9).CellType = cellTextBox 'cellHyperLink
    Grid.Column(9).Alignment = cellCenterCenter 'cellHyperLink
    
    Grid.Column(10).CellType = cellTextBox 'cellHyperLink
    Grid.Column(10).Alignment = cellCenterCenter 'cellHyperLink
   
    Grid.Column(11).CellType = cellTextBox 'cellHyperLink
    Grid.Column(11).Alignment = cellCenterCenter 'cellHyperLink
   
    Grid.Column(12).CellType = cellTextBox 'cellHyperLink
    Grid.Column(12).Alignment = cellCenterCenter 'cellHyperLink
   
    Grid.Column(13).CellType = cellTextBox 'cellHyperLink
    Grid.Column(13).Alignment = cellCenterCenter 'cellHyperLink
   
 
    Grid.Column(0).Width = 10
    Grid.Column(1).Width = 80
    Grid.Column(2).Width = 80
    Grid.Column(3).Width = 80
    Grid.Column(4).Width = 80
    Grid.Column(5).Width = 80
    Grid.Column(6).Width = 80
    Grid.Column(7).Width = 80
    Grid.Column(8).Width = 80
    Grid.Column(9).Width = 80
    Grid.Column(10).Width = 80
    Grid.Column(11).Width = 80
    Grid.Column(12).Width = 80
    Grid.Column(13).Width = 80
   
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

ProcAjustaGrid

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USButton1_Click()
Unload Me
End Sub
