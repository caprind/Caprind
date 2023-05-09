VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Begin VB.Form Frm_empresa_Certificado 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   0  'Nenhum
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5565
   Icon            =   "Frm_Empresa_Certificado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Centralizar no Mestre
   Begin DrawSuite2014.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   714
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
   Begin VB.CommandButton Command1 
      Caption         =   "Seleciona certificado"
      Height          =   345
      Left            =   3510
      TabIndex        =   2
      Top             =   2670
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      Caption         =   "Certificado selecionado"
      Height          =   1905
      Left            =   150
      TabIndex        =   0
      Top             =   600
      Width           =   5235
      Begin VB.Label lbCerti 
         Caption         =   "--"
         Height          =   915
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   4995
      End
   End
End
Attribute VB_Name = "Frm_empresa_Certificado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mCertificadoSel As String

Private Sub Command1_Click()
On Error GoTo tratar_erro

'TimerClote.Enabled = False
    Frm_Empresa_Certificado_Abrir.Show vbModal
    mCertificadoSel = Frm_Empresa_Certificado_Abrir.mCertificadoSel
    frmOpcoesGeral.txtCertificadodigital.Text = mCertificadoSel
    'TimerClote.Enabled = True
    'AppIni.PutString "app", "cert-nfe-" & EmpresaID, mCertificadoSel
    
    ListaCert
fim:

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ListaCert()
On Error GoTo tratar_erro
Dim i As Long        ',: Cardinal;
Dim oNode As IXMLDOMNode
Dim SetT As New settings, Certs As ICertificates, StoreSrc As New Store
    'Dim StoreDst As New Store
Dim Cert As Certificate        ': OleVariant;
    'Dim oRps As IXMLDOMNodeList, oLote As IXMLDOMNodeList, oSigs As IXMLDOMNodeList
Dim s1 As String, s2 As String

    'Sett = CoSettings.Create
    On Error GoTo ListaCert_Error

    SetT.EnablePromptForCertificateUI = True
    'StoreSrc = CoStore.Create
    Call StoreSrc.Open(CAPICOM_CURRENT_USER_STORE, "My", CAPICOM_STORE_OPEN_EXISTING_ONLY)
    'StoreDst = CoStore.Create
    'Call StoreDst.Open(CAPICOM_CURRENT_USER_STORE, "TMP", CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED)
    Set Certs = StoreSrc.Certificates

    '//Remove certificados sem a private key.
    If Certs.Count > 0 Then
        '     Set Certs = Certs.Find(CAPICOM_CERTIFICATE_FIND_EXTENDED_PROPERTY, CAPICOM_PROPID_KEY_PROV_INFO)
    End If
    '//Somente certificados com data válida.
    If Certs.Count > 0 Then
        '    Set Certs = Certs.Find(CAPICOM_CERTIFICATE_FIND_TIME_VALID)
    End If
    'MsgBox Certs.item(1).SubjectName

    'Certs.Select
    lbCerti = "Nenhum certificado selecionado"
    For i = Certs.Count To 1 Step -1
        'MsgBox Certs.Item(i).SubjectName
        'MsgBox Certs.Item(i).SerialNumber
        If mCertificadoSel = Certs.Item(i).SerialNumber Then
            lbCerti = Certs.Item(i).SubjectName
        End If
        'MsgBox Cert.
    Next


    If Certs.Count = 0 Then
        lbCerti.Caption = "Sem certificados"
    Else
        Set Cert = Certs.Item(1)
    End If

    On Error GoTo 0
Exit Sub

ListaCert_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListaCert of Formulário frmMainNFe"

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

