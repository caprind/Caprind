VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{F15158C8-31F4-4D02-A18E-FFDF0FFFE433}#1.0#0"; "videocap.ocx"
Begin VB.Form frmCompras_Recebimento_Foto 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Recebimento | Foto do item no recebimento"
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   9195
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fotos da RE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   5160
      TabIndex        =   24
      Top             =   540
      Width           =   3915
      Begin VB.FileListBox DirFotos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         Left            =   0
         Pattern         =   "*.bmp"
         TabIndex        =   25
         Top             =   0
         Width           =   3915
      End
      Begin DrawSuite2022.USAlphaImage usFoto 
         Height          =   3240
         Left            =   -420
         TabIndex        =   26
         Top             =   2400
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   5715
         Image           =   "frmCompras_Recebimento_Foto.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   885
      Left            =   270
      TabIndex        =   16
      Top             =   540
      Width           =   4815
      Begin VB.TextBox txtFoto 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3210
         TabIndex        =   21
         Top             =   420
         Width           =   1455
      End
      Begin VB.TextBox txtAmostragem 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   420
         Width           =   1455
      End
      Begin VB.TextBox txtLote 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   420
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "n° foto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3675
         TabIndex        =   22
         Top             =   210
         Width           =   525
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amostragem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1935
         TabIndex        =   20
         Top             =   210
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   705
         TabIndex        =   18
         Top             =   210
         Width           =   345
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Diretório de fotos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   300
      TabIndex        =   13
      Top             =   5250
      Width           =   4785
      Begin VB.TextBox txtD3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   180
         TabIndex        =   14
         ToolTipText     =   "Abrir diretório de retorno..."
         Top             =   300
         Width           =   3465
      End
      Begin DrawSuite2022.USButton cmdD3 
         Height          =   315
         Left            =   3660
         TabIndex        =   15
         ToolTipText     =   "Abrir diretório de fotos do lote"
         Top             =   300
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmCompras_Recebimento_Foto.frx":0018
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   8421504
         BorderColorDisabled=   13160660
         BorderColorDown =   7907521
         BorderColorOver =   7907521
         GradientColor2  =   14737632
         GradientColor3  =   12632256
         GradientColor4  =   12632256
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   14417407
         GradientColorOver2=   12317439
         GradientColorOver3=   4838399
         GradientColorOver4=   9627391
         GradientColorDown1=   10802943
         GradientColorDown2=   7979263
         GradientColorDown3=   4370174
         GradientColorDown4=   7395582
         GradientColors  =   1
         PicAlign        =   0
         ShowFocusRect   =   0   'False
         Theme           =   1
         ToolTipTitle    =   "CAPRIND v5.0"
      End
      Begin DrawSuite2022.USButton btnFoto 
         Height          =   465
         Left            =   4170
         TabIndex        =   23
         ToolTipText     =   "Capturar foto do item..."
         Top             =   150
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   820
         DibPicture      =   "frmCompras_Recebimento_Foto.frx":1E11D
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColorDown =   15048022
         BorderColorOver =   15381630
         PicSize         =   2
         PicSizeH        =   24
         PicSizeW        =   24
         ShowFocusRect   =   0   'False
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   741
      DibPicture      =   "frmCompras_Recebimento_Foto.frx":1E95C
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmCompras_Recebimento_Foto.frx":2236D
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   11
      Top             =   6090
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   714
   End
   Begin VIDEOCAPLib.VideoCap VideoCap1 
      Height          =   3615
      Left            =   270
      TabIndex        =   10
      Top             =   1530
      Width           =   4815
      _Version        =   65536
      LicenseKey      =   "9980"
      _ExtentX        =   8493
      _ExtentY        =   6376
      _StockProps     =   0
   End
   Begin VB.CommandButton cmdFoto 
      BackColor       =   &H00E0E0E0&
      Caption         =   "tirar foto em JPG"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9810
      Picture         =   "frmCompras_Recebimento_Foto.frx":22687
      TabIndex        =   9
      Top             =   6450
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SnapShot To HBITMAP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      Picture         =   "frmCompras_Recebimento_Foto.frx":307C9
      TabIndex        =   8
      Top             =   8310
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SnapShot to Picture Box"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6900
      Picture         =   "frmCompras_Recebimento_Foto.frx":3E90B
      TabIndex        =   7
      Top             =   6390
      Width           =   2295
   End
   Begin VB.ComboBox cboVideoInput 
      Height          =   315
      Left            =   2760
      TabIndex        =   5
      Top             =   7350
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   270
      ScaleHeight     =   3615
      ScaleWidth      =   4815
      TabIndex        =   4
      Top             =   1530
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.ComboBox cbovideoformat 
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Top             =   7710
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SnapShot"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Picture         =   "frmCompras_Recebimento_Foto.frx":4CA4D
      TabIndex        =   1
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   8310
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Video Input"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   7350
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Video Format"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   7710
      Width           =   1215
   End
End
Attribute VB_Name = "frmCompras_Recebimento_Foto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Temp As Integer


Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public DiretorioFotos As String

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" ( _
      lpPictDesc As PictDesc, _
      riid As Guid, _
      ByVal fPictureOwnsHandle As Long, _
      ipic As IPicture _
    ) As Long

Private Sub btnFoto_Click()
On Error GoTo tratar_erro

'If VideoCap1.Visible = True Then

    If txtFoto.Text <> "" Then
        strfilename = txtD3.Text & "\" & txtFoto.Text & ".bmp"
        result = Me.VideoCap1.SnapShot(strfilename)
        Picture1.Picture = LoadPicture(strfilename, vbLPLarge, vbLPColor)
        
    Else
        USMsgBox "Informe o nome da foto a ser gravada!", vbInformation, "CAPRIND v5.0"
        txtFoto.SetFocus
        Exit Sub
    End If
 '   Picture1.Visible = True
 '   VideoCap1.Visible = False
'Else
   ' Picture1.Visible = False
   ' VideoCap1.Visible = True
'End If
DirFotos.Refresh
Seq = DirFotos.ListCount + 1

If Seq < 10 Then
txtFoto.Text = RE & "0" & Seq
Else
txtFoto.Text = RE & Seq
End If
txtFoto.Locked = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cbovideoformat_Click()
Call Command1_Click
End Sub

Private Sub cmdD3_Click()
On Error GoTo tratar_erro

  ShellExecute 0, "open", txtD3.Text, "", "", vbNormalFocus

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Command1_Click()

strVideoInput = cboVideoInput.List(cboVideoInput.ListIndex)
videoinputindex = Me.VideoCap1.VideoInputs.FindVideoInput(strVideoInput)

If videoinputindex <> -1 Then
        VideoCap1.VideoInput = videoinputindex
End If


strVideoFormat = cbovideoformat.List(cbovideoformat.ListIndex)
videoFormatIndex = Me.VideoCap1.VideoFormats.FindVideoFormat(strVideoFormat)

If videoFormatIndex <> -1 Then
        VideoCap1.VideoFormat = 6 'videoFormatIndex
End If

Me.VideoCap1.Start
End Sub

Private Sub Command2_Click()
strfilename = App.Path + "\" + "test" + ".bmp"
result = Me.VideoCap1.SnapShot(strfilename)
Picture1.Picture = LoadPicture(strfilename, vbLPLarge, vbLPColor)

End Sub

Sub FillVideoFormat()

    For Each myvideoformat In VideoCap1.VideoFormats

        Me.cbovideoformat.AddItem myvideoformat.Name

    Next


End Sub


Private Sub Command3_Click()


Picture1.Picture = VideoCap1.SnapShot2Picture

End Sub

Private Sub Command4_Click()

Picture1.Picture = BitmapToPicture(VideoCap1.SnapShot2HBITMAP)
End Sub

Private Sub Command5_Click()
VideoCap1.SnapShotJPEG "c:\test.jpg", 90

USMsgBox "Foto salva em c:\Foto.jpg"


End Sub

Public Sub ProcCriarPastaFotos()
On Error GoTo tratar_erro

DiretorioFotos = Localrel & "\Imagens\Fotos\" & RE

If DS.FileOrDirExists(DiretorioFotos) = False Then
MkDir DiretorioFotos
End If

txtD3.Text = DiretorioFotos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Screen.MousePointer = vbDefault
End Sub

Private Sub DirFotos_Click()
On Error GoTo tratar_erro

usFoto.LoadImage_FromFile txtD3.Text & "\" & DirFotos.filename

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

VideoCap1.LicenseKey = "11000"
ProcCriarPastaFotos

If txtD3.Text <> "" Then
    DirFotos.Path = txtD3.Text
    DirFotos.Refresh
End If

For Each myVideoInput In Me.VideoCap1.VideoInputs
        cboVideoInput.AddItem myVideoInput.Name
Next

If cboVideoInput.ListCount > 0 Then
        cboVideoInput.ListIndex = 0
 End If


FillVideoFormat

If cbovideoformat.ListCount > 0 Then
cbovideoformat.ListIndex = 0
End If

Temp = 1

txtAmostragem.Text = frmCompras_recebimento.Txtamostra.Text
txtLote.Text = frmCompras_recebimento.Txt_lote

DirFotos.Refresh
Seq = DirFotos.ListCount + 1

If Seq < 10 Then
txtFoto.Text = RE & "0" & Seq
Else
txtFoto.Text = RE & Seq
End If
txtFoto.Locked = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Screen.MousePointer = vbDefault
End Sub

Public Function BitmapToPicture(ByVal hBmp As Long) As IPicture

   If (hBmp = 0) Then Exit Function

   Dim NewPic As Picture, tPicConv As PictDesc, IGuid As Guid

   
   With tPicConv
      .cbSizeofStruct = Len(tPicConv)
      .picType = vbPicTypeBitmap
      .hImage = hBmp
   End With

   ' Fill in IDispatch Interface ID
   With IGuid
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With

   ' Create a picture object:
   OleCreatePictureIndirect tPicConv, IGuid, True, NewPic
   
   ' Return it:
   Set BitmapToPicture = NewPic

End Function

