VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Begin VB.Form Frm_Empresa_Certificado_Abrir 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'Nenhum
   Caption         =   "Certificado"
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   4860
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Empresa_Certificado_Abrir.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Padrão Windows
   Begin DrawSuite2014.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   661
      EnableCloseButton=   0   'False
      EnableMaximizeButton=   0   'False
      EnableMinimizeButton=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowClose       =   0   'False
      ShowControlBox  =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   1860
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   480
      Top             =   1740
   End
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   90
      TabIndex        =   0
      Top             =   540
      Width           =   4605
      Begin VB.Label Label1 
         Alignment       =   2  'Centralizar
         Caption         =   "Selecione o Certificado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4395
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3720
      Top             =   1320
   End
End
Attribute VB_Name = "Frm_Empresa_Certificado_Abrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Dim mTitulo As String
Public mCertificadoSel As String

Private Sub Timer1_Timer()
On Error GoTo tratar_erro
    
    Timer1.Enabled = False
    vv
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub vv()
On Error GoTo tratar_erro
Dim Certs As ICertificates, StoreSrc As New Store

    Call StoreSrc.Open(CAPICOM_CURRENT_USER_STORE, "My", CAPICOM_STORE_OPEN_EXISTING_ONLY)
    Set Certs = StoreSrc.Certificates

    '//Remove certificados sem a private key.
    If Certs.Count > 0 Then
        Set Certs = Certs.Find(CAPICOM_CERTIFICATE_FIND_EXTENDED_PROPERTY, CAPICOM_PROPID_KEY_PROV_INFO)
    End If
    '//Somente certificados com data válida.
    If Certs.Count > 0 Then
        Set Certs = Certs.Find(CAPICOM_CERTIFICATE_FIND_TIME_VALID)
    End If
    'MsgBox Certs.item(1).SubjectName
    Dim obj As Object
    On Error GoTo fim
    Set obj = Certs.Select(mTitulo, "Selecione o Certificado Digital para uso no aplicativo")
    mCertificadoSel = obj.Item(1).SerialNumber
    'MsgBox obj.Item(1).SubjectName
    Debug.Print obj.Item(1).Display
fim:
    Unload Me
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

