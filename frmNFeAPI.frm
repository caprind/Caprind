VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmNFeAPI 
   Caption         =   "NF-e API"
   ClientHeight    =   10035
   ClientLeft      =   6825
   ClientTop       =   1005
   ClientWidth     =   16545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   16545
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Fechar"
      Height          =   525
      Left            =   12390
      TabIndex        =   19
      Top             =   6540
      Width           =   1455
   End
   Begin VB.TextBox txt_Buscar 
      Height          =   375
      Left            =   9120
      TabIndex        =   18
      Top             =   6570
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Achar texto"
      Height          =   405
      Left            =   10650
      TabIndex        =   17
      Top             =   6570
      Width           =   1365
   End
   Begin RichTextLib.RichTextBox txtResult 
      Height          =   2445
      Left            =   180
      TabIndex        =   16
      Top             =   7140
      Width           =   16185
      _ExtentX        =   28549
      _ExtentY        =   4313
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmNFeAPI.frx":0000
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
      Caption         =   "Localizar XML"
      Height          =   405
      Left            =   7530
      TabIndex        =   15
      Top             =   6570
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtNFeXml 
      Height          =   5295
      Left            =   7440
      TabIndex        =   14
      Top             =   1200
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   9340
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmNFeAPI.frx":008B
   End
   Begin VB.TextBox txtConteudo 
      Appearance      =   0  'Flat
      Height          =   5295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Text            =   "frmNFeAPI.frx":0116
      Top             =   1200
      Width           =   7335
   End
   Begin VB.TextBox txtCaminho 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Text            =   "C:\Notas\"
      Top             =   360
      Width           =   5535
   End
   Begin VB.ComboBox cbTpConteudo 
      Height          =   315
      ItemData        =   "frmNFeAPI.frx":1819
      Left            =   8400
      List            =   "frmNFeAPI.frx":1826
      TabIndex        =   10
      Text            =   "txt"
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtTpAmb 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   5220
      TabIndex        =   8
      Text            =   "2"
      Top             =   6510
      Width           =   375
   End
   Begin VB.TextBox txtCNPJ 
      Height          =   315
      Left            =   5760
      TabIndex        =   6
      Text            =   "10766336000113"
      Top             =   360
      Width           =   2535
   End
   Begin VB.CheckBox checkExibir 
      Caption         =   "Exibir PDF"
      Height          =   255
      Left            =   5730
      TabIndex        =   4
      Top             =   6540
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.ComboBox cbTpDown 
      Height          =   315
      ItemData        =   "frmNFeAPI.frx":183A
      Left            =   1710
      List            =   "frmNFeAPI.frx":184D
      TabIndex        =   3
      Text            =   "XP"
      Top             =   6510
      Width           =   2055
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar XML para validar SEFAZ"
      Height          =   525
      Left            =   13890
      TabIndex        =   0
      Top             =   6540
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Salvar em:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Ambiente:"
      Height          =   195
      Left            =   3930
      TabIndex        =   9
      Top             =   6540
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ:"
      Height          =   195
      Left            =   5760
      TabIndex        =   7
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label13 
      Caption         =   "Tipo de Download:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   6540
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Resposta do Servidor"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   6870
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Conteudo"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   690
   End
End
Attribute VB_Name = "frmNFeAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ProcEnviarNotaSefaz()
    Dim retorno As String
    Dim token As String
    
'=======================================================================
'                      Token Caprind Sistemas                          =
'=======================================================================
'token = "Q0FQUklORCBTSVNURU1BSEs5c1o="
'=======================================================================
'                        Token FNL Tecnologia                          =
'=======================================================================
token = "RkFCSU8gQ0FSRE9TTyBSc2ZGcTI="
'=======================================================================
    If (txtCaminho.Text <> "") And (txtNFeXml.Text <> "") And (cbTpConteudo.Text <> "") And (cbTpDown.Text <> "") And (txtTpAmb.Text <> "") Then
        
        'Faz a emissão síncrona
        retorno = emitirNFeSincrono(Texto_Envio, cbTpConteudo.Text, txtcnpj.Text, cbTpDown.Text, txtTpAmb.Text, txtCaminho.Text, checkExibir.Value)
        txtResult.Text = retorno
        
        'Abaixo, confira um exemplo de tratamento de retorno da função emitirNFeSincrono
        
        Dim statusEnvio, statusConsulta, statusDownload, cStat, chNFe, nProt, motivo, nsNRec, erros As String
        
        'Lê o statusEnvio
        statusEnvio = LerDadosJSON(retorno, "statusEnvio", "", "")
        'Lê o statusConsulta
        statusConsulta = LerDadosJSON(retorno, "statusConsulta", "", "")
        'Lê o statusDownload
        statusDownload = LerDadosJSON(retorno, "statusDownload", "", "")
        'Lê o cStat
        cStat = LerDadosJSON(retorno, "cStat", "", "")
        'Lê a chNFe
        cStat = LerDadosJSON(retorno, "chNFe", "", "")
        'Lê o nProt
        nProt = LerDadosJSON(retorno, "nProt", "", "")
        'Lê o motivo
        motivo = LerDadosJSON(retorno, "motivo", "", "")
        'Lê o nsNRec
        nsNRec = LerDadosJSON(retorno, "nsNRec", "", "")
        'Lê os erros
        erros = LerDadosJSON(retorno, "erros", "", "")
        
        'Agora que você já leu os dados, é aconselhável que faça o salvamento de todos
        'eles no seu banco de dados antes de prosseguir para o teste abaixo
                 
        'Testa se houve sucesso na emissão
        If (statusEnvio = 200) Or (statusEnvio = -6) Then
            'Testa se houve sucesso na consulta
            If (statusConsulta = 200) Then
                'Testa se a nota foi autorizada
                If (cStat = 100) Then
                    'Aqui dentro você pode realizar procedimentos como desabilitar o botão de emitir, etc
                  USMsgBox (motivo)
                     
                    'Testa se o download teve problemas
                    If (statusDownload <> 200) Then
                      USMsgBox (motivo)
                    End If
                Else
                    'Aqui você pode mostrar alguma solução para o parceiro ou exibir opção de editar a nota
                  USMsgBox (motivo)
                End If
            'Caso tenha dado erro na consulta
            Else
                'Aqui você pode mostrar uma mensagem ao usuário
              USMsgBox (motivo + Chr(13) + erros)
            End If
        Else
            'Aqui você pode exibir para o usuário o erro que ocorreu no envio
          USMsgBox (motivo + Chr(13) + erros)
        End If
    Else
      USMsgBox ("Todos os campos devem ser preenchidos")
    End If
    
Exit Sub
SAI:
  USMsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, "CAPRIND v5.0", titleCTeAPI
End Sub


Sub BuscarTexto(txtNFeXml As Object, Optional ByVal PosIni As Integer)

    Dim Pos As Integer
'    Dim PalavraChave As String
    
    'TipoBusca corresponde si se busca Mayus y Minus identicas...
    Dim TipoBusca As Long
    
    'La variable Palavrachave toma el valor de txt_Buscar
    'PalavraChave = txt_Buscar.Text
    
    'Verificar si Palavrachave no esta vacía
    If Len(PalavraChave) Then
        'Verificar si Mayusculas y Minusculas esta desactivada
'        If Check1.Value = 0 Then
            TipoBusca = vbTextCompare
        'Else
       '     TipoBusca = vbBinaryCompare
        'End If
            
            'Busca desde la PosIni que se indico...
        Pos = InStr(PosIni + 1, txtNFeXml.Text, PalavraChave, TipoBusca)
        If Pos > 0 Then
             'Si devolvio mayor de 0...se encontro
                
             With txtNFXeml
                txtNFeXml.SelStart = Pos - 1
                txtNFeXml.SelLength = Len(PalavraChave)
                txtNFeXml.SetFocus
              End With
                USMsgBox "Palavra encontrada"
         Else
             'No se encoUsmsgbox =ntró
             txtNFeXml.SetFocus
             USMsgBox "Palavra não encontrada."
         End If
    End If

End Sub

Private Sub cmdEnviar_Click()
    On Error GoTo SAI
    Dim retorno As String
    Dim token As String
    
'=======================================================================
'                      Token Caprind Sistemas                          =
'=======================================================================
'token = "Q0FQUklORCBTSVNURU1BSEs5c1o="
'=======================================================================
'                        Token FNL Tecnologia                          =
'=======================================================================
token = "RkFCSU8gQ0FSRE9TTyBSc2ZGcTI="
'=======================================================================

    'Debug.print txtConteudo
    'Debug.print txtNFeXml
    
    If (txtCaminho.Text <> "") And (txtNFeXml.Text <> "") And (cbTpConteudo.Text <> "") And (cbTpDown.Text <> "") And (txtTpAmb.Text <> "") Then
        
        'Faz a emissão síncrona
        retorno = emitirNFeSincrono(txtNFeXml.Text, cbTpConteudo.Text, txtcnpj.Text, cbTpDown.Text, txtTpAmb.Text, txtCaminho.Text, checkExibir.Value)
        txtResult.Text = retorno
        
        'Abaixo, confira um exemplo de tratamento de retorno da função emitirNFeSincrono
        
        Dim statusEnvio, statusConsulta, statusDownload, cStat, chNFe, nProt, motivo, nsNRec, erros As String
        
        'Lê o statusEnvio
        statusEnvio = LerDadosJSON(retorno, "statusEnvio", "", "")
        'Lê o statusConsulta
        statusConsulta = LerDadosJSON(retorno, "statusConsulta", "", "")
        'Lê o statusDownload
        statusDownload = LerDadosJSON(retorno, "statusDownload", "", "")
        'Lê o cStat
        cStat = LerDadosJSON(retorno, "cStat", "", "")
        'Lê a chNFe
        cStat = LerDadosJSON(retorno, "chNFe", "", "")
        'Lê o nProt
        nProt = LerDadosJSON(retorno, "nProt", "", "")
        'Lê o motivo
        motivo = LerDadosJSON(retorno, "motivo", "", "")
        'Lê o nsNRec
        nsNRec = LerDadosJSON(retorno, "nsNRec", "", "")
        'Lê os erros
        erros = LerDadosJSON(retorno, "erros", "", "")
        
        'Agora que você já leu os dados, é aconselhável que faça o salvamento de todos
        'eles no seu banco de dados antes de prosseguir para o teste abaixo
                 
        'Testa se houve sucesso na emissão
        If (statusEnvio = 200) Or (statusEnvio = -6) Then
            'Testa se houve sucesso na consulta
            If (statusConsulta = 200) Then
                'Testa se a nota foi autorizada
                If (cStat = 100) Then
                    'Aqui dentro você pode realizar procedimentos como desabilitar o botão de emitir, etc
                  USMsgBox (motivo)
                     
                    'Testa se o download teve problemas
                    If (statusDownload <> 200) Then
                      USMsgBox (motivo)
                    End If
                Else
                    'Aqui você pode mostrar alguma solução para o parceiro ou exibir opção de editar a nota
                  USMsgBox (motivo)
                End If
            'Caso tenha dado erro na consulta
            Else
                'Aqui você pode mostrar uma mensagem ao usuário
              USMsgBox (motivo + Chr(13) + erros)
            End If
        Else
            'Aqui você pode exibir para o usuário o erro que ocorreu no envio
          USMsgBox (motivo + Chr(13) + erros)
        End If
    Else
      USMsgBox ("Todos os campos devem ser preenchidos")
    End If
    
    Exit Sub
SAI:
  USMsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, "CAPRIND v5.0", titleCTeAPI

End Sub

Private Sub Command1_Click()
    CommonDialog1.CancelError = True
    ' Set flags
    CommonDialog1.flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|RTF (*.rtf)|*.rtf|Text Files (*.txt)|*.txt"
    
    ' Display the Save dialog box
    CommonDialog1.filename = ""
    CommonDialog1.ShowOpen
    txtNFeXml.filename = CommonDialog1.filename
    Exit Sub
End Sub

Private Sub Command2_Click()
Dim PalavraChave As String

PalavraChave = InputBox("Informe a palavra a localizar :", "Busca", sAchar)

If PalavraChave = "" Then Exit Sub

'txtNFeXml.Find sAchar
'---------------------------------------------------------------------
    'Función que busca el texto escrito en el txt_Buscar
    Call BuscarTexto(txtNFeXml)

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
cbTpConteudo.Text = "xml"
End Sub
