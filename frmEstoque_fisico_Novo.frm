VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmEstoque_fisico_Novo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Estoque - Inventário | Novo"
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8520
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
   ScaleHeight     =   4890
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opões para inventário de estoque"
      Height          =   3465
      Left            =   270
      TabIndex        =   2
      Top             =   660
      Width           =   7965
      Begin VB.TextBox txt_SaldoRE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6360
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   1755
         Width           =   1395
      End
      Begin VB.TextBox txt_vlr_Unitario 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5070
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   1755
         Width           =   1275
      End
      Begin VB.TextBox txt_Corrida 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2670
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   1755
         Width           =   2385
      End
      Begin VB.TextBox txt_Certificado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   270
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   1755
         Width           =   2385
      End
      Begin VB.ComboBox cmbLocal_armaz 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         ToolTipText     =   "Local de armazenamento."
         Top             =   2895
         Width           =   5370
      End
      Begin VB.TextBox txt_LA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Família do item"
         Top             =   2895
         Width           =   5355
      End
      Begin VB.TextBox Txt_familia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Família do item"
         Top             =   2355
         Width           =   5355
      End
      Begin VB.TextBox txt_RE 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4500
         MaxLength       =   50
         TabIndex        =   18
         ToolTipText     =   "Numero da RE (Registro de estoque)"
         Top             =   585
         Width           =   915
      End
      Begin VB.TextBox Txt_lote 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5775
         MaxLength       =   60
         TabIndex        =   13
         ToolTipText     =   "Numero do lote no estoque"
         Top             =   585
         Width           =   1665
      End
      Begin VB.TextBox txt_cod_Referencia 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2190
         MaxLength       =   50
         TabIndex        =   12
         ToolTipText     =   "Código de referência do item"
         Top             =   585
         Width           =   1965
      End
      Begin VB.TextBox Txt_cod_interno 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         MaxLength       =   50
         TabIndex        =   6
         ToolTipText     =   "Código interno do item"
         Top             =   585
         Width           =   1605
      End
      Begin VB.TextBox Txt_descricao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Descrição do item"
         Top             =   1185
         Width           =   6585
      End
      Begin VB.TextBox Txt_un 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Unidade."
         Top             =   1185
         Width           =   915
      End
      Begin DrawSuite2022.USButton btnItem 
         Height          =   885
         Left            =   5640
         TabIndex        =   3
         Top             =   2340
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1561
         DibPicture      =   "frmEstoque_fisico_Novo.frx":0000
         Caption         =   "Iniciar inventário"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         GradientColor1  =   5263559
         GradientColor2  =   5263559
         GradientColor3  =   5263559
         GradientColor4  =   5263559
         GradientColorDisabled1=   13160660
         GradientColorDisabled2=   13160660
         GradientColorDisabled3=   13160660
         GradientColorDisabled4=   13160660
         GradientColorOver1=   4408288
         GradientColorOver2=   4408288
         GradientColorOver3=   4408288
         GradientColorOver4=   4408288
         GradientColorDown1=   4013465
         GradientColorDown2=   4013465
         GradientColorDown3=   4013465
         GradientColorDown4=   4013465
         PicAlign        =   7
         PicSize         =   3
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
      Begin DrawSuite2022.USButton btnCod 
         Height          =   315
         Left            =   1860
         TabIndex        =   7
         ToolTipText     =   "Filtrar por código do item"
         Top             =   585
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_fisico_Novo.frx":53F21
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
      Begin DrawSuite2022.USButton btnRef 
         Height          =   315
         Left            =   4170
         TabIndex        =   8
         ToolTipText     =   "Filtrar por código de referencia"
         Top             =   585
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_fisico_Novo.frx":5B0B4
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
      Begin DrawSuite2022.USButton btnLote 
         Height          =   315
         Left            =   7440
         TabIndex        =   17
         ToolTipText     =   "Fltrar pelo numero do lote no estoque"
         Top             =   585
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_fisico_Novo.frx":62247
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
      Begin DrawSuite2022.USButton btnRE 
         Height          =   315
         Left            =   5430
         TabIndex        =   19
         ToolTipText     =   "Fltrar pelo numero da RE (Registro de estoque)"
         Top             =   585
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         DibPicture      =   "frmEstoque_fisico_Novo.frx":693DA
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo RE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   6742
         TabIndex        =   32
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor unitário"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   5235
         TabIndex        =   30
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corrida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   3600
         TabIndex        =   28
         Top             =   1560
         Width           =   525
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Certificado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   1072
         TabIndex        =   26
         Top             =   1560
         Width           =   780
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local de armazenamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   2025
         TabIndex        =   23
         Top             =   2700
         Width           =   1785
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   2670
         TabIndex        =   21
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   16
         Top             =   390
         Width           =   990
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. de referência"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2497
         TabIndex        =   15
         Top             =   390
         Width           =   1350
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° do lote*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6195
         TabIndex        =   14
         Top             =   390
         Width           =   825
      End
      Begin VB.Label Label7 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   7005
         TabIndex        =   11
         Top             =   990
         Width           =   585
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   3495
         TabIndex        =   10
         Top             =   990
         Width           =   690
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° do RE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4635
         TabIndex        =   9
         Top             =   390
         Width           =   645
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   4485
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   714
      DibPicture      =   "frmEstoque_fisico_Novo.frx":7056D
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
      Icon            =   "frmEstoque_fisico_Novo.frx":7221A
   End
End
Attribute VB_Name = "frmEstoque_fisico_Novo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCod_Click()
On Error GoTo tratar_erro

txt_RE = ""
txt_cod_Referencia = ""
Txt_descricao = ""
Txt_un = ""
Txt_lote = ""
Txt_familia = ""

If Txt_cod_interno <> "" Then
    Set TBProduto = CreateObject("adodb.recordset")
    'StrSql = "Select P.Codproduto, P.Desenho, P.Descricao, P.Unidade, P.Classe FROM Projproduto P where P.Desenho = '" & Txt_cod_interno & "' and P.Estoque = 'True' and P.DtValidacao IS NOT NULL and P.bloqueado = 'False'"
    StrSql = "Select P.Codproduto, P.Desenho, P.Descricao, P.Unidade, P.Classe FROM Projproduto P where P.Desenho = '" & Txt_cod_interno & "' and P.Estoque = 'True' and P.DtValidacao IS NOT NULL and P.bloqueado = 'False'"
    'Debug.print StrSql
    
    TBProduto.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        Txt_cod_interno = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
        Txt_descricao = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
        Txt_un = IIf(IsNull(TBProduto!Unidade), "", TBProduto!Unidade)
        Txt_familia = IIf(IsNull(TBProduto!Classe), "", TBProduto!Classe)
        txt_LA.Visible = False
        cmbLocal_armaz.Visible = True
        ProcCarregaComboLA cmbLocal_armaz, False, False
    End If
    TBProduto.Close
End If
                 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub btnItem_Click()
On Error GoTo tratar_erro


If Txt_cod_interno.Text = "" Then
 USMsgBox "Favor informar o código interno para inventário.", vbInformation, "CAPRIND v5.0"
 Txt_cod_interno.SetFocus
 Exit Sub
End If

If Txt_lote.Text = "" Then
 USMsgBox "Favor informar o numero do Lote para inventário.", vbInformation, "CAPRIND v5.0"
 Txt_lote.SetFocus
 Exit Sub
End If

If txt_LA.Text = "" And cmbLocal_armaz.Text = "" Then
 USMsgBox "Favor informar o local de armazenamento para inventário.", vbInformation, "CAPRIND v5.0"
 cmbLocal_armaz.SetFocus
 Exit Sub
End If



With frmestoque_fisico
.Txt_cod_interno = Txt_cod_interno
.txt_RE.Text = txt_RE.Text
.Txt_descricao = Txt_descricao
.Txt_lote.Text = Txt_lote.Text
.Txt_un = Txt_un
.Txt_familia = Txt_familia
.txt_cod_Referencia = txt_cod_Referencia


If txt_RE <> "" Then
    
    Set TBEstoque = CreateObject("adodb.recordset")
    StrSql = "Select Sum(Entrada) - Sum(Saida) as Saldo FROM Estoque_movimentacao EC where IDESTOQUE = '" & txt_RE.Text & "'"
    
    TBEstoque.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
    Saldo_Atual = TBEstoque!Saldo
    End If
    TBEstoque.Close
    
    Set TBProduto = CreateObject("adodb.recordset")
    StrSql = "Select EC.valor_unitario, EC.Corrida,EC.Certificado FROM Estoque_controle EC where EC.IDESTOQUE = '" & txt_RE.Text & "'"
    'Debug.print StrSql
    
    TBProduto.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        .Txt_valor_unitario = IIf(IsNull(TBProduto!valor_unitario), "", TBProduto!valor_unitario)
        .txt_Corrida.Text = IIf(IsNull(TBProduto!Corrida), "", TBProduto!Corrida)
        .txt_Certificado = IIf(IsNull(TBProduto!Certificado), "", TBProduto!Certificado)
        .Txt_qtde_estoque = Format(Saldo_Atual, "###,##0.0000")
    End If
    TBProduto.Close
End If

If txt_LA.Text <> "" Then
.txt_LA.Text = txt_LA
Else
.txt_LA.Text = cmbLocal_armaz.Text
End If

End With
                
Unload Me
                
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnLote_Click()
On Error GoTo tratar_erro

txt_cod_Referencia = ""
Txt_descricao = ""
Txt_un = ""
Txt_familia = ""

If Txt_lote <> "" Then


    Set TBProduto = CreateObject("adodb.recordset")
    StrSql = "Select EC.IDestoque, EC.Lote,EC.local_armaz, P.Codproduto, P.Desenho, P.Descricao, P.Unidade, P.Classe FROM Projproduto P INNER JOIN Estoque_controle EC ON P.Desenho = EC.Desenho where EC.Lote = '" & Txt_lote.Text & "' and P.Estoque = 'True' and P.DtValidacao IS NOT NULL and P.bloqueado = 'False'"
    'Debug.print StrSql
    
    TBProduto.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        Txt_cod_interno = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
        txt_RE.Text = IIf(IsNull(TBProduto!IDEstoque), "", TBProduto!IDEstoque)
        Txt_descricao = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
        Txt_lote.Text = IIf(IsNull(TBProduto!LOTE), "", TBProduto!LOTE)
        Txt_un = IIf(IsNull(TBProduto!Unidade), "", TBProduto!Unidade)
        Txt_familia = IIf(IsNull(TBProduto!Classe), "", TBProduto!Classe)
        txt_LA.Visible = True
        cmbLocal_armaz.Visible = False
        txt_LA = IIf(IsNull(TBProduto!local_armaz), "", TBProduto!local_armaz)
    End If
    TBProduto.Close
End If
                 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnRE_Click()
On Error GoTo tratar_erro

txt_cod_Referencia = ""
Txt_descricao = ""
Txt_un = ""
Txt_familia = ""

If txt_RE <> "" Then

    Set TBEstoque = CreateObject("adodb.recordset")
    StrSql = "Select Sum(Entrada) - Sum(Saida) as Saldo FROM Estoque_movimentacao EC where IDESTOQUE = '" & txt_RE.Text & "'"
    
    TBEstoque.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBEstoque.EOF = False Then
    Saldo_Atual = TBEstoque!Saldo
    End If
    TBEstoque.Close

    Set TBProduto = CreateObject("adodb.recordset")
    StrSql = "Select EC.*,P.Codproduto, P.Desenho, P.Descricao, P.Unidade, P.Classe FROM Projproduto P INNER JOIN Estoque_controle EC ON P.Desenho = EC.Desenho where EC.IDESTOQUE = '" & txt_RE.Text & "' and P.Estoque = 'True' and P.DtValidacao IS NOT NULL and P.bloqueado = 'False'"
    'Debug.print StrSql
    
    TBProduto.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        Txt_cod_interno = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
        txt_RE.Text = IIf(IsNull(TBProduto!IDEstoque), "", TBProduto!IDEstoque)
        Txt_descricao = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
        Txt_lote.Text = IIf(IsNull(TBProduto!LOTE), "", TBProduto!LOTE)
        Txt_un = IIf(IsNull(TBProduto!Unidade), "", TBProduto!Unidade)
        Txt_familia = IIf(IsNull(TBProduto!Classe), "", TBProduto!Classe)
        txt_LA.Visible = True
        cmbLocal_armaz.Visible = False
        txt_LA = IIf(IsNull(TBProduto!local_armaz), "", TBProduto!local_armaz)
        txt_Certificado.Text = IIf(IsNull(TBProduto!Certificado), "", TBProduto!Certificado)
        txt_Corrida.Text = IIf(IsNull(TBProduto!Corrida), "", TBProduto!Corrida)
        txt_vlr_Unitario.Text = IIf(IsNull(TBProduto!valor_unitario), "", Format(TBProduto!valor_unitario, "###,##0.0000"))
        txt_SaldoRE.Text = Format(Saldo_Atual, "###,##0.0000")
    End If
    TBProduto.Close
End If
                 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub btnRef_Click()
On Error GoTo tratar_erro

txt_RE = ""
Txt_cod_interno = ""
Txt_descricao = ""
Txt_un = ""
Txt_lote = ""
Txt_familia = ""

If txt_cod_Referencia <> "" Then

    Set TBProduto = CreateObject("adodb.recordset")
    StrSql = "select * from item_aplicacoes IA inner join projproduto PP on IA.codproduto = PP.codproduto where IA.n_referencia = '" & txt_cod_Referencia & "'"
    'Debug.print StrSql
    
    TBProduto.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
    Txt_cod_interno = TBProduto!Desenho
    Else
    USMsgBox "Item não localizado, favor verificar informação.", vbCritical, "CAPRIND v5.0"
    Exit Sub
    End If
    TBProduto.Close


    Set TBProduto = CreateObject("adodb.recordset")
    StrSql = "Select P.Codproduto, P.Desenho, P.Descricao, P.Unidade, P.Classe FROM Projproduto P INNER JOIN Estoque_controle EC ON P.Desenho = EC.Desenho where P.Desenho = '" & Txt_cod_interno & "' and P.Estoque = 'True' and P.DtValidacao IS NOT NULL and P.bloqueado = 'False'"
    'Debug.print StrSql
    
    TBProduto.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = False Then
        Txt_cod_interno = IIf(IsNull(TBProduto!Desenho), "", TBProduto!Desenho)
        Txt_descricao = IIf(IsNull(TBProduto!Descricao), "", TBProduto!Descricao)
        Txt_un = IIf(IsNull(TBProduto!Unidade), "", TBProduto!Unidade)
        Txt_familia = IIf(IsNull(TBProduto!Classe), "", TBProduto!Classe)
        txt_LA.Visible = False
        cmbLocal_armaz.Visible = True
        ProcCarregaComboLA cmbLocal_armaz, False, False
    End If
    TBProduto.Close
End If
                 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub

End Sub
