VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Begin VB.Form frmFaturamento_Prod_serv_Danfe_xml 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'Nenhum
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   2460
   StartUpPosition =   3  'Padrão Windows
   Begin DrawSuite2014.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   556
      EnableMaximizeButton=   0   'False
      EnableMinimizeButton=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmFaturamento_Prod_serv_Danfe_xml.frx":0000
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opções "
      Height          =   1245
      Left            =   -30
      TabIndex        =   0
      Top             =   60
      Width           =   2475
      Begin DrawSuite2014.USButton USButton1 
         Height          =   345
         Left            =   300
         TabIndex        =   1
         Top             =   360
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   609
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Baixar DANFE e XML"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DrawSuite2014.USButton USButton2 
         Height          =   345
         Left            =   300
         TabIndex        =   2
         Top             =   720
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   609
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         Caption         =   "Enviar DANFE e XML"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmFaturamento_Prod_serv_Danfe_xml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
