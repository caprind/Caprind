VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmCompras_Recebimento_Medicoes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Recebimento | Medições"
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
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
   ScaleHeight     =   3780
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Valores encontrados"
      Height          =   1755
      Left            =   3000
      TabIndex        =   11
      Top             =   540
      Width           =   2325
      Begin VB.TextBox txtv3 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1110
         TabIndex        =   14
         Text            =   "0"
         ToolTipText     =   "Valor menor"
         Top             =   1290
         Width           =   1065
      End
      Begin VB.TextBox txtv2 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1110
         TabIndex        =   13
         Text            =   "0"
         ToolTipText     =   "Valor menor"
         Top             =   840
         Width           =   1065
      End
      Begin VB.TextBox txtv1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1110
         TabIndex        =   12
         Text            =   "0"
         ToolTipText     =   "Valor menor"
         Top             =   390
         Width           =   1065
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor 3 :"
         Height          =   225
         Left            =   330
         TabIndex        =   17
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor 2 :"
         Height          =   225
         Left            =   330
         TabIndex        =   16
         Top             =   870
         Width           =   765
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor 1 :"
         Height          =   225
         Left            =   330
         TabIndex        =   15
         Top             =   420
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Valores de referência"
      Height          =   1755
      Left            =   210
      TabIndex        =   4
      Top             =   540
      Width           =   2745
      Begin VB.TextBox txtvMaior 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1530
         TabIndex        =   7
         ToolTipText     =   "Valor menor"
         Top             =   1290
         Width           =   1065
      End
      Begin VB.TextBox txtvNominal 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1530
         TabIndex        =   6
         ToolTipText     =   "Valor menor"
         Top             =   840
         Width           =   1065
      End
      Begin VB.TextBox txtvMenor 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1530
         TabIndex        =   5
         ToolTipText     =   "Valor menor"
         Top             =   390
         Width           =   1065
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor maior :"
         Height          =   225
         Left            =   360
         TabIndex        =   10
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor nominal :"
         Height          =   225
         Left            =   180
         TabIndex        =   9
         Top             =   870
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor menor :"
         Height          =   225
         Left            =   300
         TabIndex        =   8
         Top             =   420
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   945
      Left            =   210
      TabIndex        =   2
      Top             =   2340
      Width           =   5115
      Begin VB.ComboBox cmbUnidade 
         Height          =   360
         Left            =   1500
         TabIndex        =   19
         ToolTipText     =   "Informe a unidade de medida para as medições"
         Top             =   390
         Width           =   1125
      End
      Begin DrawSuite2022.USButton btnSalvar 
         Height          =   525
         Left            =   3600
         TabIndex        =   18
         Top             =   270
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   926
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Salvar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Unid. medida :"
         Height          =   225
         Left            =   210
         TabIndex        =   3
         Top             =   450
         Width           =   1965
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   741
      DibPicture      =   "frmCompras_Recebimento_Medicoes.frx":0000
      CaptionDelimiter=   "|"
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
      Icon            =   "frmCompras_Recebimento_Medicoes.frx":62E4
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   3375
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   714
   End
End
Attribute VB_Name = "frmCompras_Recebimento_Medicoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
