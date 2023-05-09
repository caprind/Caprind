VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmFluxoCaixaGrafico 
   Caption         =   "Form1"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   15375
   StartUpPosition =   3  'Padrão Windows
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   10005
      Left            =   0
      OleObjectBlob   =   "frmFluxoCaixa_Grafico.frx":0000
      TabIndex        =   0
      Top             =   30
      Width           =   15375
   End
End
Attribute VB_Name = "frmFluxoCaixaGrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
