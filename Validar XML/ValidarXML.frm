VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13890
   LinkTopic       =   "Form1"
   ScaleHeight     =   9525
   ScaleWidth      =   13890
   StartUpPosition =   3  'Padrão Windows
   Begin VB.TextBox Text2 
      Height          =   7905
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "ValidarXML.frx":0000
      Top             =   1530
      Width           =   13365
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   13245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   210
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
strArquivo = "C:\CAPRIND\Nota fiscal\Danfe-XML\35190710766336000113550010000000641289562462-procNFe.xml"
   docNFe.Load strArquivo
   Text1.Text = "Validando o  Arquivo..." & vbCrLf
   strRetorno = FunValidaSchema(docNFe, "http://www.portalfiscal.inf.br/nfe", "G:\Caprind\Documentos NFe\PL_009_V4_00_NT_2019_001_v1.20a\PL_009_V4_00_NT_2019_001_v1.20a\nfe_v4.00.xsd", False)
Text2.Text = strRetorno
End Sub
