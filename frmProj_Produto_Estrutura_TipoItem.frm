VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmProj_Produto_Estrutura_TipoItem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Estrutura | Novo item"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Escolha o tipo do item na estrutura"
      Height          =   2715
      Left            =   300
      TabIndex        =   2
      Top             =   720
      Width           =   4065
      Begin DrawSuite2022.USButton cmdTitular 
         Height          =   1005
         Left            =   600
         TabIndex        =   3
         Top             =   390
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1773
         DibPicture      =   "frmProj_Produto_Estrutura_TipoItem.frx":0000
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         Caption         =   "Titular"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         ForeColorDown   =   16777215
         ForeColorOver   =   16777215
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         PicAlign        =   7
         PicSize         =   3
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   3
      End
      Begin DrawSuite2022.USButton cmdAlternativo 
         Height          =   1005
         Left            =   600
         TabIndex        =   4
         Top             =   1440
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1773
         DibPicture      =   "frmProj_Produto_Estrutura_TipoItem.frx":6415
         BorderColor     =   1154291
         BorderColorDisabled=   13160660
         BorderColorDown =   16576
         BorderColorOver =   8438015
         Caption         =   "Alternativo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         ForeColorDown   =   16777215
         ForeColorOver   =   16777215
         GradientColor1  =   1154291
         GradientColor2  =   1154291
         GradientColor3  =   1154291
         GradientColor4  =   1154291
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorDown1=   16576
         GradientColorDown2=   16576
         GradientColorDown3=   16576
         GradientColorDown4=   16576
         GradientColorOver1=   8438015
         GradientColorOver2=   8438015
         GradientColorOver3=   8438015
         GradientColorOver4=   8438015
         PicAlign        =   7
         PicSize         =   3
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         ShowFocusRectDown=   0   'False
         Theme           =   5
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   3810
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   873
      DibPicture      =   "frmProj_Produto_Estrutura_TipoItem.frx":C82A
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmProj_Produto_Estrutura_TipoItem.frx":D275
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
End
Attribute VB_Name = "frmProj_Produto_Estrutura_TipoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTitular_Click()
On Error GoTo tratar_erro

frmproj_EstruturaLocaliza_item.Show 1


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
