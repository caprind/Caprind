VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCQ_Certificado_Analise 
   Caption         =   "Qualidade - Certificados de analise"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9915
      Left            =   -30
      TabIndex        =   0
      Top             =   30
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   17489
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Certificado de analise"
      TabPicture(0)   =   "frmCQ_Certificado_Analise.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Lista"
      Tab(0).Control(1)=   "USToolBar1"
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(3)=   "USImageList1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Ensaios para liberação"
      TabPicture(1)   =   "frmCQ_Certificado_Analise.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Listaensaio"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "USImageList2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "USToolBar2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin DrawSuite2022.USToolBar USToolBar2 
         Height          =   975
         Left            =   60
         TabIndex        =   3
         Top             =   330
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   1720
         ButtonCount     =   8
         GradientColor2  =   14737632
         GradientColorOverRight1=   16315633
         GradientColorOverRight2=   15195350
         GripperColor    =   15195350
         IsStrech        =   -1  'True
         RightColor1     =   0
         RightColor2     =   0
         ShowEndPanel    =   0   'False
         Theme           =   1
         ButtonCaption1  =   "Novo"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Novo (Insert)"
         ButtonKey1      =   "1"
         ButtonAlignment1=   2
         BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   40
         ButtonHeight1   =   24
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Salvar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Salvar (F3)"
         ButtonKey2      =   "2"
         ButtonAlignment2=   2
         BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   44
         ButtonTop2      =   2
         ButtonWidth2    =   44
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Excluir"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Excluir (F4)"
         ButtonKey3      =   "3"
         ButtonAlignment3=   2
         BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   90
         ButtonTop3      =   2
         ButtonWidth3    =   45
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Relatório"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Relatório (F5)"
         ButtonKey4      =   "4"
         ButtonAlignment4=   2
         BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   137
         ButtonTop4      =   2
         ButtonWidth4    =   60
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonAlignment5=   2
         ButtonType5     =   1
         ButtonStyle5    =   -1
         BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState5    =   -1
         ButtonLeft5     =   199
         ButtonTop5      =   4
         ButtonWidth5    =   2
         ButtonHeight5   =   54
         ButtonCaption6  =   "Ajuda"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Ajuda (F1)"
         ButtonKey6      =   "6"
         ButtonAlignment6=   2
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   203
         ButtonTop6      =   2
         ButtonWidth6    =   41
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Sair"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Sair (Esc)"
         ButtonKey7      =   "7"
         ButtonAlignment7=   2
         BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft7     =   246
         ButtonTop7      =   2
         ButtonWidth7    =   30
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonKey8      =   "8"
         ButtonAlignment8=   2
         BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState8    =   5
         ButtonLeft8     =   278
         ButtonTop8      =   2
         ButtonWidth8    =   24
         ButtonHeight8   =   24
      End
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   -70320
         Top             =   600
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCQ_Certificado_Analise.frx":0038
         Count           =   1
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados do ensaio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   60
         TabIndex        =   19
         Top             =   1290
         Width           =   15225
         Begin VB.TextBox txtEncontrado 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
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
            Left            =   12600
            MaxLength       =   20
            TabIndex        =   38
            ToolTipText     =   "Gama de resultado"
            Top             =   600
            Width           =   1155
         End
         Begin VB.TextBox txtMinimo 
            Alignment       =   2  'Center
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
            Height          =   345
            Left            =   10230
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   22
            ToolTipText     =   "Resultado da analise"
            Top             =   600
            Width           =   1185
         End
         Begin VB.TextBox txtLaudo 
            Alignment       =   2  'Center
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
            Height          =   345
            Left            =   13785
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   24
            ToolTipText     =   "Observações"
            Top             =   600
            Width           =   1305
         End
         Begin VB.TextBox txtUnidade 
            Alignment       =   2  'Center
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
            Height          =   345
            Left            =   9090
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   21
            ToolTipText     =   "Unidade doensaio"
            Top             =   600
            Width           =   1125
         End
         Begin VB.TextBox txtMaximo 
            Alignment       =   2  'Center
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
            Height          =   345
            Left            =   11430
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   23
            ToolTipText     =   "Gama de resultado"
            Top             =   600
            Width           =   1155
         End
         Begin VB.TextBox txtEnsaio 
            Alignment       =   2  'Center
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
            Height          =   345
            Left            =   150
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   20
            ToolTipText     =   "Ensaio realizado"
            Top             =   600
            Width           =   8475
         End
         Begin VB.TextBox txtIDEnsaio 
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
            TabIndex        =   36
            Top             =   600
            Width           =   615
         End
         Begin DrawSuite2022.USButton cmdEnsaios 
            Height          =   315
            Left            =   8670
            TabIndex        =   37
            ToolTipText     =   "Ensaios para o certificado"
            Top             =   600
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   556
            DibPicture      =   "frmCQ_Certificado_Analise.frx":3EB5
            BorderColor     =   5263559
            BorderColorDisabled=   13160660
            BorderColorDown =   4013465
            BorderColorOver =   4408288
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
            ForeColor       =   16777215
            ForeColorDown   =   16777215
            ForeColorOver   =   16777215
            GradientColor1  =   5263559
            GradientColor2  =   5263559
            GradientColor3  =   5263559
            GradientColor4  =   5263559
            GradientColorDisabled1=   13160660
            GradientColorDisabled2=   13160660
            GradientColorDisabled3=   13160660
            GradientColorDisabled4=   13160660
            GradientColorDown1=   4013465
            GradientColorDown2=   4013465
            GradientColorDown3=   4013465
            GradientColorDown4=   4013465
            GradientColorOver1=   4408288
            GradientColorOver2=   4408288
            GradientColorOver3=   4408288
            GradientColorOver4=   4408288
            PicAlign        =   8
            ShowFocusRect   =   0   'False
            ShowFocusRectDown=   0   'False
            Theme           =   4
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Encontrado"
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
            Index           =   0
            Left            =   12795
            TabIndex        =   39
            Top             =   390
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minimo"
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
            Index           =   5
            Left            =   10575
            TabIndex        =   29
            Top             =   390
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unidade"
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
            Index           =   7
            Left            =   9420
            TabIndex        =   28
            Top             =   390
            Width           =   585
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Maximo"
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
            Index           =   4
            Left            =   11730
            TabIndex        =   27
            Top             =   390
            Width           =   540
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Laudo final"
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
            Index           =   4
            Left            =   14055
            TabIndex        =   26
            Top             =   390
            Width           =   780
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ensaio"
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
            Index           =   3
            Left            =   4155
            TabIndex        =   25
            Top             =   390
            Width           =   465
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dados do certificado de análise"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   -74940
         TabIndex        =   2
         Top             =   1290
         Width           =   15225
         Begin VB.TextBox txtDescricao 
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
            Left            =   4020
            Locked          =   -1  'True
            TabIndex        =   34
            ToolTipText     =   "Descrição do produto"
            Top             =   600
            Width           =   5385
         End
         Begin VB.TextBox TxtObservacoes 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   135
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            ToolTipText     =   "Observações da analise"
            Top             =   1350
            Width           =   14955
         End
         Begin VB.TextBox txtCA 
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
            ToolTipText     =   "Código do certificado de analise"
            Top             =   600
            Width           =   1155
         End
         Begin VB.TextBox txtCRQ 
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
            Left            =   12690
            TabIndex        =   15
            ToolTipText     =   "Código do CRQ"
            Top             =   600
            Width           =   1425
         End
         Begin VB.TextBox txtAnalista 
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
            Left            =   9420
            Locked          =   -1  'True
            TabIndex        =   14
            ToolTipText     =   "Analista"
            Top             =   600
            Width           =   2565
         End
         Begin VB.TextBox txtQuant_Env 
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
            Left            =   12000
            Locked          =   -1  'True
            TabIndex        =   13
            ToolTipText     =   "Quantidade analisada"
            Top             =   600
            Width           =   675
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
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   12
            ToolTipText     =   "Lote de fabricação do produto"
            Top             =   600
            Width           =   1035
         End
         Begin VB.TextBox txtProduto 
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
            Left            =   2370
            TabIndex        =   11
            ToolTipText     =   "Código do produto"
            Top             =   600
            Width           =   1635
         End
         Begin VB.TextBox txtData 
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
            Left            =   14130
            Locked          =   -1  'True
            TabIndex        =   10
            ToolTipText     =   "Data da analise"
            Top             =   600
            Width           =   945
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
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
            Index           =   1
            Left            =   6367
            TabIndex        =   35
            Top             =   420
            Width           =   690
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Observações"
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
            Index           =   8
            Left            =   300
            TabIndex        =   31
            Top             =   1140
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° CA"
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
            Index           =   2
            Left            =   540
            TabIndex        =   16
            Top             =   390
            Width           =   435
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data"
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
            Index           =   3
            Left            =   14460
            TabIndex        =   9
            Top             =   390
            Width           =   345
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° CRQ"
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
            Index           =   2
            Left            =   13140
            TabIndex        =   8
            Top             =   390
            Width           =   555
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Analista"
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
            Index           =   1
            Left            =   10417
            TabIndex        =   7
            Top             =   420
            Width           =   570
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Analisado"
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
            Index           =   0
            Left            =   12000
            TabIndex        =   6
            Top             =   390
            Width           =   690
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
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
            Index           =   0
            Left            =   1740
            TabIndex        =   5
            Top             =   390
            Width           =   315
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produto"
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
            Index           =   0
            Left            =   2910
            TabIndex        =   4
            Top             =   420
            Width           =   570
         End
      End
      Begin DrawSuite2022.USToolBar USToolBar1 
         Height          =   975
         Left            =   -74940
         TabIndex        =   1
         Top             =   330
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1720
         ButtonCount     =   8
         GradientColor2  =   14737632
         GradientColorOverRight1=   16315633
         GradientColorOverRight2=   15195350
         GripperColor    =   15195350
         IsStrech        =   -1  'True
         RightColor1     =   0
         RightColor2     =   0
         ShowEndPanel    =   0   'False
         Theme           =   1
         ButtonCaption1  =   "Novo"
         ButtonEnabled1  =   0   'False
         ButtonIconSize1 =   32
         ButtonToolTipText1=   "Novo (Insert)"
         ButtonKey1      =   "1"
         ButtonAlignment1=   2
         BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft1     =   2
         ButtonTop1      =   2
         ButtonWidth1    =   36
         ButtonHeight1   =   21
         ButtonUseMaskColor1=   0   'False
         ButtonCaption2  =   "Filtrar"
         ButtonEnabled2  =   0   'False
         ButtonIconSize2 =   32
         ButtonToolTipText2=   "Filtrar (F2)"
         ButtonKey2      =   "2"
         ButtonAlignment2=   2
         BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft2     =   40
         ButtonTop2      =   2
         ButtonWidth2    =   36
         ButtonHeight2   =   21
         ButtonUseMaskColor2=   0   'False
         ButtonCaption3  =   "Salvar"
         ButtonEnabled3  =   0   'False
         ButtonIconSize3 =   32
         ButtonToolTipText3=   "Salvar (F3)"
         ButtonKey3      =   "3"
         ButtonAlignment3=   2
         BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft3     =   78
         ButtonTop3      =   2
         ButtonWidth3    =   44
         ButtonHeight3   =   21
         ButtonUseMaskColor3=   0   'False
         ButtonCaption4  =   "Relatório"
         ButtonEnabled4  =   0   'False
         ButtonIconSize4 =   32
         ButtonToolTipText4=   "Relatório (F5)"
         ButtonKey4      =   "4"
         ButtonAlignment4=   2
         BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft4     =   124
         ButtonTop4      =   2
         ButtonWidth4    =   60
         ButtonHeight4   =   21
         ButtonUseMaskColor4=   0   'False
         ButtonEnabled5  =   0   'False
         ButtonIconSize5 =   32
         ButtonAlignment5=   2
         ButtonType5     =   1
         ButtonStyle5    =   -1
         BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState5    =   -1
         ButtonLeft5     =   186
         ButtonTop5      =   4
         ButtonWidth5    =   2
         ButtonHeight5   =   54
         ButtonCaption6  =   "Ajuda"
         ButtonEnabled6  =   0   'False
         ButtonIconSize6 =   32
         ButtonToolTipText6=   "Ajuda (F1)"
         ButtonKey6      =   "6"
         ButtonAlignment6=   2
         BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft6     =   190
         ButtonTop6      =   2
         ButtonWidth6    =   41
         ButtonHeight6   =   21
         ButtonUseMaskColor6=   0   'False
         ButtonCaption7  =   "Sair"
         ButtonEnabled7  =   0   'False
         ButtonIconSize7 =   32
         ButtonToolTipText7=   "Sair (Esc)"
         ButtonKey7      =   "7"
         ButtonAlignment7=   2
         BeginProperty ButtonFont7 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonLeft7     =   233
         ButtonTop7      =   2
         ButtonWidth7    =   30
         ButtonHeight7   =   21
         ButtonUseMaskColor7=   0   'False
         ButtonEnabled8  =   0   'False
         ButtonIconSize8 =   32
         ButtonKey8      =   "8"
         ButtonAlignment8=   2
         BeginProperty ButtonFont8 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonState8    =   5
         ButtonLeft8     =   265
         ButtonTop8      =   2
         ButtonWidth8    =   24
         ButtonHeight8   =   24
         Begin VB.TextBox txtID_Laudo 
            Height          =   315
            Left            =   2400
            TabIndex        =   33
            Top             =   1500
            Width           =   1095
         End
      End
      Begin DrawSuite2022.USImageList USImageList2 
         Left            =   12135
         Top             =   570
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmCQ_Certificado_Analise.frx":21FBA
         Count           =   1
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   5955
         Left            =   -74940
         TabIndex        =   18
         Top             =   3900
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   10504
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "IDLaudo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Certificado"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Lote"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Produto"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Descricao"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "N"
            Text            =   "Analista"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Enviado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "T"
            Text            =   "CRQ"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Object.Tag             =   "N"
            Text            =   "Data"
            Object.Width           =   1941
         EndProperty
      End
      Begin MSComctlLib.ListView Listaensaio 
         Height          =   7395
         Left            =   60
         TabIndex        =   30
         Top             =   2460
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   13044
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Ensaio"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Unidade"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Mínimo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Máximo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Encontrado"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Laudo"
            Object.Width           =   2646
         EndProperty
      End
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1366
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10410
      FormWidthDT     =   15480
      FormScaleHeightDT=   9945
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
End
Attribute VB_Name = "frmCQ_Certificado_Analise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Novo_Certificado As Boolean 'OK
Public StrSqlCQCertificadoLocalizar As String 'OK
Dim TBCQCertificado As ADODB.Recordset
Dim TBEnsaio As ADODB.Recordset

Private Sub ProcCarregaListaCertificado()
On Error GoTo tratar_erro
Lista.ListItems.Clear

StrSql = "Select * from CQ_Certificado order by IDLaudo"

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then


    Do While TBLISTA.EOF = False
        With Lista.ListItems
            .Add , , TBLISTA!IDLaudo
            .Item(.Count).SubItems(1) = TBLISTA!CodCertificado
            .Item(.Count).SubItems(2) = TBLISTA!LOTE
            .Item(.Count).SubItems(3) = TBLISTA!Produto
            .Item(.Count).SubItems(4) = TBLISTA!Descricao
            .Item(.Count).SubItems(5) = TBLISTA!Analista
            .Item(.Count).SubItems(6) = Format(TBLISTA!Quant_Env, "###,##0.00")
            .Item(.Count).SubItems(7) = TBLISTA!CRQ
            .Item(.Count).SubItems(8) = Format(TBLISTA!Data, "dd/mm/yyyy")
        End With
        TBLISTA.MoveNext
    Loop

End If
TBLISTA.Close


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaEnsaio()
On Error GoTo tratar_erro
Listaensaio.ListItems.Clear

StrSql = "select * from CQ_Certificado_Ensaios Where ID_laudo = " & txtID_Laudo.Text & ""

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then


    Do While TBLISTA.EOF = False
        With Listaensaio.ListItems
            .Add , , TBLISTA!ID_Ensaio
            .Item(.Count).SubItems(1) = TBLISTA!Criterio
            .Item(.Count).SubItems(2) = TBLISTA!Unidade
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Maximo), "0,0000", TBLISTA!Maximo)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Minimo), "0,0000", TBLISTA!Minimo)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!encontrado), "0,0000", TBLISTA!encontrado)
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Laudo), "0,0000", TBLISTA!Laudo)
        End With
        TBLISTA.MoveNext
    Loop

End If
TBLISTA.Close


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcSalvarCertificado()
On Error GoTo tratar_erro

If txtLote.Text = "" Then
USMsgBox "Falta informar o lote para o laudo", vbInformation, "CAPRIND v5.0"
ProcBuscaLote
Exit Sub
End If

If txtCRQ.Text = "" Then
USMsgBox "Falta informar o CRQ do analista para o laudo", vbInformation, "CAPRIND v5.0"
txtCRQ.SetFocus
Exit Sub
End If

StrSql = "Select * from CQ_Certificado where IDLaudo = " & IIf(txtID_Laudo.Text = "", 0, txtID_Laudo.Text) & ""

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
TBAbrir.AddNew
End If
TBAbrir!CodCertificado = txtCA.Text
TBAbrir!LOTE = txtLote.Text
TBAbrir!Produto = txtProduto.Text
TBAbrir!Descricao = txtdescricao.Text
TBAbrir!Analista = txtAnalista.Text
TBAbrir!Quant_Env = txtQuant_Env.Text
TBAbrir!CRQ = txtCRQ.Text
TBAbrir!Data = txtData.Text
TBAbrir!Observacoes = txtObservacoes.Text
TBAbrir.Update
txtID_Laudo.Text = TBAbrir!IDLaudo
USMsgBox "Dados do certificado salvos com sucesso!", vbInformation, "CAPRIND v5.0"
TBAbrir.Close
ProcCarregaListaCertificado

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcBuscaDadosLaudo()
On Error GoTo tratar_erro


Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
txtCA.Text = TBAbrir!CodCertificado
txtLote.Text = TBAbrir!LOTE
txtProduto.Text = TBAbrir!Produto
txtdescricao.Text = TBAbrir!Descricao
txtAnalista.Text = TBAbrir!Analista
txtQuant_Env.Text = Format(TBAbrir!Quant_Env, "###,##0.00")
txtCRQ.Text = TBAbrir!CRQ
txtData.Text = Format(TBAbrir!Data, "dd/mm/yyyy")
txtObservacoes.Text = TBAbrir!Observacoes
TBAbrir.Close
Else
USMsgBox "Laudo não localizado, favor verificar", vbInformation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub cmdEnsaios_Click()
On Error GoTo tratar_erro
  
frmCQ_Certificado_Criterios.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
  
Select Case SSTab1.Tab
    Case 0:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovo
            Case vbKeyF2: ProcLocalizar
            Case vbKeyF3: ProcSalvarCertificado
            'Case vbKeyF5: ProcImprimir
            Case vbKeyEscape: Unload Me
        End Select
    Case 1:
        Select Case KeyCode
            Case vbKeyInsert: ProcNovoAnalise
            Case vbKeyF3: procSalvarEnsaio
           ' Case vbKeyF4: ProcExcluir
            'Case vbKeyF5: ProcImprimir
            Case vbKeyEscape: Unload Me
        End Select
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 8, True
ProcCarregaToolBar2 Me, 15195, 7, True

ProcRemoveObjetosResize Me
ProcCarregaListaCertificado
SSTab1.Tab = 0

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub

txtID_Laudo.Text = Lista.SelectedItem
If txtID_Laudo.Text <> "" Then
StrSql = "Select * from CQ_Certificado where IDLaudo = " & IIf(txtID_Laudo.Text = "", 0, txtID_Laudo.Text) & ""
ProcBuscaDadosLaudo
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Listaensaio_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If Listaensaio.ListItems.Count = 0 Then Exit Sub

txtIDEnsaio.Text = Listaensaio.SelectedItem

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from CQ_Certificado_Ensaios Where ID_Ensaio = " & txtIDEnsaio.Text & "", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then

txtID_Laudo = TBAbrir!ID_Laudo
txtEnsaio = TBAbrir!Criterio
txtunidade = TBAbrir!Unidade
txtMinimo = IIf(IsNull(TBAbrir!Minimo), "0,00", TBAbrir!Minimo)
txtMaximo = IIf(IsNull(TBAbrir!Maximo), "0,00", TBAbrir!Maximo)
txtEncontrado = TBAbrir!encontrado
txtLaudo = TBAbrir!Laudo
txtIDEnsaio = TBAbrir!ID_Ensaio
TBAbrir.Close

End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo tratar_erro

Select Case SSTab1.Tab

Case 1:
If txtID_Laudo = "" Then
SSTab1.Tab = 0
End If

ProcCarregaListaEnsaio
ProcLimpaCamposAnalise
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtEncontrado_Change()
On Error GoTo tratar_erro
Dim Maximo As Double
Dim Minimo As Double
Dim encontrado As Double

Minimo = IIf(txtMinimo.Text = "", 0, txtMinimo.Text)
Maximo = IIf(txtMaximo.Text = "", 0, txtMaximo)
encontrado = IIf(txtEncontrado.Text = "", 0, txtEncontrado)

If encontrado >= Minimo And encontrado <= Maximo Then
txtLaudo.Text = "APROVADO"
Else
txtLaudo.Text = "REPROVADO"
End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovo
    Case 2: ProcLocalizar
    Case 3: ProcSalvarCertificado
    Case 4: ProcImprimirCertificado
    'Case 6: ProcAjuda
    Case 7: Unload Me 'ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimirCertificado()
On Error GoTo tratar_erro

If txtID_Laudo.Text = "" Then Exit Sub
NomeRel = "LaudoAnalise.rpt"

ProcImprimirRel "{CQ_Certificado.IDLaudo} = " & txtID_Laudo.Text & "", ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub USToolBar2_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcNovoAnalise
    Case 2: procSalvarEnsaio
    Case 3: ProcExcluirEnsaio
'    'Case 4: ProcAjuda
    Case 5: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExcluirEnsaio()
On Error GoTo tratar_erro

If txtIDEnsaio.Text <> "" Then
    If USMsgBox("Deseja realmente excluir esse ensaio?", vbYesNo, "CAPRIND v5.0") = vbYes Then
        Conexao.Execute "Delete from CQ_Certificado_Ensaios where id_Ensaio = " & txtIDEnsaio.Text
        USMsgBox "Ensaio excluido com sucesso", vbInformation, "CAPRIND v5.0"
        ProcCarregaListaEnsaio
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procSalvarEnsaio()
On Error GoTo tratar_erro

If txtIDEnsaio.Text = "" Then
txtIDEnsaio.Text = 0
End If

If txtEnsaio.Text = "" Then
USMsgBox "Digite o ensaio", vbInformation, "CAPRIND v5.0"
txtEnsaio.SetFocus
Exit Sub
End If

If txtunidade.Text = "" Then
USMsgBox "Digite unidade do ensaio", vbInformation, "CAPRIND v5.0"
txtunidade.SetFocus
Exit Sub
End If

If txtEncontrado.Text = "" Then
USMsgBox "Digite a analise do ensaio", vbInformation, "CAPRIND v5.0"
'txtEncontrado.SetFocus
Exit Sub
End If

If txtLaudo.Text = "" Then
USMsgBox "Digite a Observacao do ensaio", vbInformation, "CAPRIND v5.0"
txtLaudo.SetFocus
Exit Sub
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "select * from CQ_Certificado_Ensaios Where ID_Ensaio = " & txtIDEnsaio.Text & "", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
TBAbrir.AddNew
End If

TBAbrir!ID_Laudo = txtID_Laudo
TBAbrir!Criterio = txtEnsaio
TBAbrir!Unidade = txtunidade
TBAbrir!Minimo = txtMinimo
TBAbrir!Maximo = txtMaximo
TBAbrir!encontrado = txtEncontrado
TBAbrir!Laudo = txtLaudo
'TBAbrir!observacoes = txtLaudo
TBAbrir.Update
txtIDEnsaio = TBAbrir!ID_Ensaio
TBAbrir.Close

USMsgBox "Dados salvos com sucesso!", vbInformation, "CAPRIND v5.0"
ProcCarregaListaEnsaio

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcGerarCodigoCA()
On Error GoTo tratar_erro

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "select * from CQ_Certificado order by CodCertificado", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
TBLISTA.MoveLast

   If TBLISTA!CodCertificado <> "" Then
   CodigoLaudo = TBLISTA!CodCertificado
   CodigoLaudo = Right(CodigoLaudo, 9)
   CodigoLaudo = Left(CodigoLaudo, 6)
   CodigoLaudo = Int(CodigoLaudo) + 1
   Else
   CodigoLaudo = 1
   End If
    Select Case Len(CodigoLaudo)
        Case 1: CodigoLaudo = "00000" & CodigoLaudo
        Case 2: CodigoLaudo = "0000" & CodigoLaudo
        Case 3: CodigoLaudo = "000" & CodigoLaudo
        Case 4: CodigoLaudo = "00" & CodigoLaudo
        Case 5: CodigoLaudo = "0" & CodigoLaudo
    End Select
    Ano = Right(Year(Date), 2)
CodigoLaudo = "CA" & CodigoLaudo & "/" & Right(Year(Date), 2)
Else
    CodigoLaudo = "CA000001" & "/" & Right(Year(Date), 2)
End If
TBLISTA.Close
txtCA.Text = CodigoLaudo

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcBuscaLote()
On Error GoTo tratar_erro

frmCQ_Certificado_Abrir.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcNovo()
On Error GoTo tratar_erro

ProcLimpaCampos
ProcGerarCodigoCA
ProcBuscaLote

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcNovoAnalise()
On Error GoTo tratar_erro

ProcLimpaCamposAnalise
cmdEnsaios_Click

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

With frmCQ_Certificado_Analise
'.txt_ID_Cliente = ""
'.txt_NomeRazao = ""
.txtAnalista.Text = ""
.txtCRQ = ""
.txtData = ""
.txtID_Laudo = ""
.txtLote = ""
'.txtNota_Fiscal = ""
.txtObservacoes = ""
.txtProduto.Text = ""
.txtQuant_Env = ""
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLimpaCamposAnalise()
On Error GoTo tratar_erro

With frmCQ_Certificado_Analise
.txtIDEnsaio = 0
.txtEnsaio.Text = ""
.txtunidade.Text = ""
.txtEncontrado.Text = ""
.txtEncontrado.Text = ""
.txtLaudo.Text = ""
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcLocalizar()
On Error GoTo tratar_erro

frmCQ_Certificado_Localizar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub

