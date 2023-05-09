Attribute VB_Name = "hNFe_Module"
Option Explicit
Dim FXMLDoc As DOMDocument50
Const DS = "xmlns:ds='http://www.w3.org/2000/09/xmldsig#'"

Function Assina(ByVal mm_xml As String, ByVal mSerial As String) As String
On Error GoTo tratar_erro

Dim doc As New DOMDocument50
Dim txt As String
Dim filename As String

    doc.validateOnParse = False
    doc.async = False
    doc.preserveWhiteSpace = False
    Assina = ""
    If doc.loadXML(mm_xml) Then
        filename = Format(Time, "hh-mm-ss")
        If SignXML(doc, mSerial) Then
            Assina = doc.XML
        End If
    Else
        USMsgBox "XML não carregado, " & doc.parseError
    End If
    
Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Function RemoveSignatures(D As DOMDocument50) As Boolean
On Error GoTo tratar_erro

Dim oSigs As IXMLDOMNodeList
    Call D.setProperty("SelectionNamespaces", DS)
    Set oSigs = D.selectNodes(".//ds:Signature")
    Do While (oSigs.Length <> 0)
        oSigs(0).parentNode.removeChild (oSigs(0))
        Set oSigs = D.selectNodes(".//ds:Signature")
    Loop
    Set oSigs = D.selectNodes(".//Signature")
    Do While (oSigs.Length <> 0)
        oSigs(0).parentNode.removeChild (oSigs(0))
        Set oSigs = D.selectNodes(".//Signature")
    Loop
    RemoveSignatures = (oSigs.Length = 0)

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Function AddSign(oDoc As DOMDocument50, n As IXMLDOMNode, Cert As Certificate, Store As Store) As Boolean
On Error GoTo tratar_erro

Dim sProvider As String
Dim sContainer As String
    'Dim oDoc As DOMDocument
Dim pKey As IXMLDSigKey, pKeyOut As IXMLDSigKey
Dim PrivateKey As PrivateKey
Dim oDSig As New MSXML2.MXDigitalSignature50


    oDoc.preserveWhiteSpace = False
    'Set oDoc = d
    Set oDSig.signature = n
    Set PrivateKey = Cert.PrivateKey

    '//Para criar uma KEY deve ser informado o Provider e o Container
    sProvider = Cert.PrivateKey.ProviderName
    sContainer = Cert.PrivateKey.ContainerName

    Set pKey = oDSig.createKeyFromCSP(CAPICOM_PROV_RSA_FULL, sProvider, sContainer, 0)
   ' Set oDSig.signature = n
    Set oDSig.Store = Store
    Set pKeyOut = oDSig.Sign(pKey, Certificates Or PURGE)

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Function SignXML(doc As DOMDocument50, ByVal mSerial As String) As Boolean
On Error GoTo tratar_erro

Dim i As Long
Dim oNode As IXMLDOMNode
Dim SetT As New settings, Certs, StoreSrc As New Store, StoreDst As New Store, Cert As Certificate
Dim oRps As IXMLDOMNodeList, oLote As IXMLDOMNodeList, oSigs As IXMLDOMNodeList
Dim s1 As String, s2 As String

'Sett = CoSettings.Create
SetT.EnablePromptForCertificateUI = True
'StoreSrc = CoStore.Create
Call StoreSrc.Open(CAPICOM_CURRENT_USER_STORE, "My", CAPICOM_STORE_OPEN_EXISTING_ONLY)
'StoreDst = CoStore.Create
Call StoreDst.Open(CAPICOM_CURRENT_USER_STORE, "TMP2", CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED)

For i = StoreDst.Certificates.Count To 1 Step -1
    StoreDst.Remove StoreDst.Certificates(i)
Next
If StoreDst.Certificates.Count > 0 Then
    USMsgBox "erro"
End If

'Set StoreDst = New Store
Set Certs = StoreSrc.Certificates

'//Remove certificados sem a private key.
If Certs.Count > 0 Then
    Set Certs = Certs.Find(CAPICOM_CERTIFICATE_FIND_EXTENDED_PROPERTY, CAPICOM_PROPID_KEY_PROV_INFO)
End If
'//Somente certificados com data válida.
If Certs.Count > 0 Then
    Set Certs = Certs.Find(CAPICOM_CERTIFICATE_FIND_TIME_VALID)
End If


If Certs.Count = 0 Then
    USMsgBox "Sem certificados"
Else
    Set Cert = Certs.Item(1)
    For i = Certs.Count To 1 Step -1
        If Certs.Item(i).SerialNumber = mSerial Then
            Set Cert = Certs.Item(i)
            'Debug.print Certs.Item(i).PrivateKey
            'Debug.print Certs.Item(i).SubjectName
            'Debug.print Certs.Item(i).HasPrivateKey
            'Debug.print Certs.Item(i).SerialNumber
            Exit For
        End If


    Next
End If
    
If Not Cert Is Nothing Then
    StoreDst.Certificates.Clear
    StoreDst.Add Cert
    Set FXMLDoc = doc
    Set oLote = FXMLDoc.selectNodes(".//NFe")
    Call doc.setProperty("SelectionNamespaces", DS)
    Set oSigs = doc.selectNodes(".//ds:Signature")
    
''Aqui ta colocando duas vezes a tag de assinatura
If oSigs.Length > 0 Then
    For i = (oSigs.Length - 1) To 0 Step -1
        Set oNode = oSigs.Item(i)
        Call AddSign(doc, oNode, Cert, StoreDst)
    Next
End If

End If
SignXML = True
Dim txt_out As String

'Debug.print txt_out

'doc.Save (DiretorioEnvio & "/" & NomeArquivo & ".xml")

txt_out = doc.XML
If doc.loadXML(txt_out) Then
End If

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function
