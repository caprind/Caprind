VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{F15158C8-31F4-4D02-A18E-FFDF0FFFE433}#1.0#0"; "videocap.ocx"
Begin VB.Form frmFoto 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Foto do item no recebimento"
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2014.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   767
      DibPicture      =   "frmFoto.frx":0000
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
      Icon            =   "frmFoto.frx":3A11
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2014.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   11
      Top             =   6360
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   714
   End
   Begin VIDEOCAPLib.VideoCap VideoCap1 
      Height          =   5415
      Left            =   30
      TabIndex        =   10
      Top             =   450
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   9551
      _StockProps     =   0
   End
   Begin VB.CommandButton Command5 
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
      Left            =   4440
      Picture         =   "frmFoto.frx":3D2B
      TabIndex        =   9
      Top             =   5940
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
      Picture         =   "frmFoto.frx":11E6D
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
      Left            =   4080
      Picture         =   "frmFoto.frx":1FFAF
      TabIndex        =   7
      Top             =   8310
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
      Height          =   5385
      Left            =   6270
      ScaleHeight     =   5325
      ScaleWidth      =   5715
      TabIndex        =   4
      Top             =   450
      Width           =   5775
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
      Left            =   2760
      Picture         =   "frmFoto.frx":2E0F1
      TabIndex        =   1
      Top             =   8310
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
Attribute VB_Name = "frmFoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As Integer


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

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" ( _
      lpPictDesc As PictDesc, _
      riid As Guid, _
      ByVal fPictureOwnsHandle As Long, _
      ipic As IPicture _
    ) As Long

Private Sub cbovideoformat_Click()
Call Command1_Click
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
        VideoCap1.VideoFormat = videoFormatIndex
End If

Me.VideoCap1.Start
End Sub

Private Sub Command2_Click()
strFileName = App.Path + "\" + "test" + ".bmp"
result = Me.VideoCap1.SnapShot(strFileName)
Picture1.Picture = LoadPicture(strFileName, vbLPLarge, vbLPColor)

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

Private Sub Form_Load()
VideoCap1.LicenseKey = "9980"

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

temp = 1
End Sub

Private Sub Image1_Click()

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

