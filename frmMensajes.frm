VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMensajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13500
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMensajes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   13500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameReclamaciones 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   0
      TabIndex        =   246
      Top             =   0
      Width           =   9795
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   3
         Left            =   8370
         TabIndex        =   247
         Top             =   4770
         Width           =   1095
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   1320
         Top             =   5760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":57FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":6210
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView9 
         Height          =   3735
         Left            =   180
         TabIndex        =   248
         Top             =   840
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nivel"
            Object.Width           =   2699
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Debe"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Haber"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Saldo"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label Label30 
         Caption         =   "Label30"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   210
         TabIndex        =   249
         Top             =   240
         Width           =   9195
      End
   End
   Begin VB.Frame FrameDescuadre 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   0
      TabIndex        =   242
      Top             =   0
      Width           =   8865
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   2
         Left            =   7620
         TabIndex        =   244
         Top             =   4800
         Width           =   1095
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   1320
         Top             =   5760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":6662
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":BE54
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":C866
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   375
         Left            =   120
         TabIndex        =   243
         Top             =   4800
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ListView ListView8 
         Height          =   3735
         Left            =   120
         TabIndex        =   245
         Top             =   840
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nivel"
            Object.Width           =   2699
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Debe"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Haber"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Saldo"
            Object.Width           =   2999
         EndProperty
      End
   End
   Begin VB.Frame frameCalculoSaldos 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6975
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1320
         Top             =   5760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":CCB8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":124AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":12EBC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   4800
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5640
         TabIndex        =   16
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Iniciar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   15
         Top             =   4800
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3735
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nivel"
            Object.Width           =   2699
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Debe"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Haber"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Saldo"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "C�lculo de saldos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame FrameAsientoLiquida 
      Height          =   6720
      Left            =   0
      TabIndex        =   237
      Top             =   0
      Width           =   12270
      Begin VB.CommandButton CmdContabilizar 
         Caption         =   "Contabilizar"
         Height          =   375
         Left            =   9060
         TabIndex        =   241
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Left            =   10650
         TabIndex        =   238
         Top             =   6000
         Width           =   1365
      End
      Begin MSComctlLib.ListView ListView7 
         Height          =   4905
         Left            =   225
         TabIndex        =   239
         Top             =   1005
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   8652
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label54 
         Caption         =   "Asiento Contable"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   240
         Top             =   390
         Width           =   10185
      End
   End
   Begin VB.Frame frameCtasBalance 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   8205
      Begin VB.CheckBox chkResta 
         Caption         =   "Se resta "
         Height          =   255
         Left            =   1650
         TabIndex        =   155
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CommandButton cmdCtaBalan 
         Caption         =   "&Cancelar"
         Height          =   435
         Index           =   1
         Left            =   6450
         TabIndex        =   75
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdCtaBalan 
         Caption         =   "Command4"
         Height          =   435
         Index           =   0
         Left            =   5010
         TabIndex        =   74
         Top             =   2640
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Haber"
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   73
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Debe"
         Height          =   255
         Index           =   1
         Left            =   5340
         TabIndex        =   72
         Top             =   1920
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "SALDO"
         Height          =   255
         Index           =   0
         Left            =   3900
         TabIndex        =   71
         Top             =   1920
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.TextBox Text3 
         Height          =   360
         Left            =   1650
         TabIndex        =   69
         Text            =   "Text2"
         Top             =   1860
         Width           =   1275
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1650
         TabIndex        =   68
         Text            =   "Text2"
         Top             =   900
         Width           =   6045
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   1290
         Picture         =   "frmMensajes.frx":1330E
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label21 
         Caption         =   "Cuenta"
         Height          =   255
         Left            =   420
         TabIndex        =   70
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "Grupo"
         Height          =   255
         Left            =   420
         TabIndex        =   67
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "MODIFICAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   420
         TabIndex        =   66
         Top             =   240
         Width           =   4875
      End
   End
   Begin VB.Frame frameBalance 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4905
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   13455
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   480
         MaxLength       =   10
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CheckBox chkPintar 
         Caption         =   "Escribir si el resultado es negativo"
         Height          =   255
         Left            =   2430
         TabIndex        =   77
         Top             =   3630
         Width           =   3915
      End
      Begin VB.CheckBox chkCero 
         Caption         =   "Poner a CERO si el resultado es negativo"
         Height          =   255
         Left            =   6630
         TabIndex        =   76
         Top             =   3630
         Width           =   4575
      End
      Begin VB.CheckBox chkNegrita 
         Caption         =   "Negrita"
         Height          =   255
         Left            =   11460
         TabIndex        =   62
         Top             =   3630
         Width           =   1035
      End
      Begin VB.CommandButton cmdBalance 
         Caption         =   "Cancelar"
         Height          =   435
         Index           =   1
         Left            =   12000
         TabIndex        =   60
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdBalance 
         Caption         =   "Aceptar"
         Height          =   435
         Index           =   0
         Left            =   10680
         TabIndex        =   59
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   480
         MaxLength       =   200
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   2760
         Width           =   12675
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   480
         MaxLength       =   100
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   1980
         Width           =   12705
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   480
         MaxLength       =   100
         TabIndex        =   52
         Text            =   "WWWWWWWWWWFFFFFFFFFFWWWWWWWWWWFFFFFFFFFFWWWWWWWWWWFFFFFFFFFFWWWWWWWWWWFFFFFFFFFF"
         Top             =   1080
         Width           =   12735
      End
      Begin VB.Label Label15 
         Caption         =   "C�digo oficial balance"
         Height          =   315
         Index           =   3
         Left            =   480
         TabIndex        =   111
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "MODIFICAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   510
         TabIndex        =   61
         Top             =   300
         Width           =   4875
      End
      Begin VB.Label Label15 
         Caption         =   "Formula"
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   58
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Texto cuentas"
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   55
         Top             =   1740
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Nombre"
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   53
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Frame FrameCobros 
      Height          =   6720
      Left            =   0
      TabIndex        =   233
      Top             =   -30
      Width           =   13410
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   12000
         TabIndex        =   234
         Top             =   6000
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   4905
         Left            =   225
         TabIndex        =   235
         Top             =   1005
         Width           =   13035
         _ExtentX        =   22992
         _ExtentY        =   8652
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label52 
         Caption         =   "Cobros de la factura "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   236
         Top             =   390
         Width           =   10185
      End
   End
   Begin VB.Frame FrameShowProcess 
      Height          =   6720
      Left            =   0
      TabIndex        =   229
      Top             =   0
      Width           =   10950
      Begin VB.CommandButton CmdRegresar 
         Caption         =   "Salir"
         Height          =   375
         Left            =   9240
         TabIndex        =   230
         Top             =   6120
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   4905
         Left            =   225
         TabIndex        =   231
         Top             =   1005
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   8652
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label53 
         Caption         =   "Usuarios conectados a"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   232
         Top             =   390
         Width           =   10185
      End
   End
   Begin VB.Frame FrameInformeBBDD 
      Height          =   6720
      Left            =   0
      TabIndex        =   219
      Top             =   0
      Width           =   10950
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Salir"
         Height          =   375
         Left            =   9240
         TabIndex        =   220
         Top             =   6120
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   4905
         Left            =   225
         TabIndex        =   221
         Top             =   1005
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   8652
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8370
         TabIndex        =   228
         Top             =   660
         Width           =   795
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9240
         TabIndex        =   227
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   226
         Top             =   660
         Width           =   795
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5910
         TabIndex        =   225
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         Caption         =   "Ejercicio Siguiente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         TabIndex        =   224
         Top             =   300
         Width           =   3435
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         Caption         =   "Ejercicio Actual"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3210
         TabIndex        =   223
         Top             =   300
         Width           =   3705
      End
      Begin VB.Label Label46 
         Caption         =   "Concepto"
         Height          =   255
         Left            =   270
         TabIndex        =   222
         Top             =   660
         Width           =   2355
      End
   End
   Begin VB.Frame FrameIconosVisibles 
      Height          =   6720
      Left            =   0
      TabIndex        =   214
      Top             =   -60
      Width           =   7050
      Begin VB.CommandButton cmdAcepIconos 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   217
         Top             =   6060
         Width           =   1215
      End
      Begin VB.CommandButton cmdCanIconos 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5520
         TabIndex        =   216
         Top             =   6060
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   5235
         Left            =   225
         TabIndex        =   215
         Top             =   675
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   9234
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "C�digo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase "
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.Label Label44 
         Caption         =   "Accesos Directos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   218
         Top             =   270
         Width           =   5145
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   6480
         Picture         =   "frmMensajes.frx":13D10
         Top             =   330
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   6120
         Picture         =   "frmMensajes.frx":13E5A
         Top             =   330
         Width           =   240
      End
   End
   Begin VB.Frame FrameImpPunteo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   30
      TabIndex        =   157
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   8
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   173
         Text            =   "Text9"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   7
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   172
         Text            =   "Text9"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   6
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   171
         Text            =   "Text9"
         Top             =   2160
         Width           =   1755
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   5
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   168
         Text            =   "Text9"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   4
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   167
         Text            =   "Text9"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   3
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   166
         Text            =   "Text9"
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   2
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   165
         Text            =   "Text9"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   1
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   164
         Text            =   "Text9"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   0
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   163
         Text            =   "Text9"
         Top             =   1200
         Width           =   1755
      End
      Begin VB.CommandButton cmdPunteo 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   4950
         TabIndex        =   158
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Line Line7 
         BorderWidth     =   3
         X1              =   120
         X2              =   6180
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Label Label22 
         Caption         =   "Haber"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   170
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label22 
         Caption         =   "Debe"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   169
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   240
         Index           =   13
         Left            =   5490
         TabIndex        =   162
         Top             =   840
         Width           =   510
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Sin puntear"
         Height          =   240
         Index           =   12
         Left            =   2730
         TabIndex        =   161
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label22 
         Caption         =   "Punteada"
         Height          =   195
         Index           =   11
         Left            =   810
         TabIndex        =   160
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label37 
         Caption         =   "Importes punteo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   159
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame frameSaldosHco 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   125
         Top             =   3330
         Width           =   5625
         Begin VB.TextBox txtsaldo 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   9
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   127
            Text            =   "Text1"
            Top             =   120
            Width           =   1815
         End
         Begin VB.TextBox txtsaldo 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   8
            Left            =   1860
            Locked          =   -1  'True
            TabIndex        =   126
            Text            =   "Text1"
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label28 
            Caption         =   "SALDO PERIODO"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   128
            Top             =   120
            Width           =   1755
         End
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   7
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   123
         Text            =   "Text1"
         Top             =   2970
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   6
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   122
         Text            =   "Text1"
         Top             =   2970
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   5
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "Text1"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   4
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   112
         Text            =   "Text1"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   3
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   2
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   1
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtsaldo 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   0
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   0
         Left            =   4590
         TabIndex        =   1
         Top             =   3930
         Width           =   1095
      End
      Begin VB.Image Image6 
         Height          =   240
         Index           =   0
         Left            =   5820
         Picture         =   "frmMensajes.frx":13FA4
         Top             =   1605
         Width           =   240
      End
      Begin VB.Image Image6 
         Height          =   240
         Index           =   1
         Left            =   5820
         Picture         =   "frmMensajes.frx":149A6
         Top             =   2070
         Width           =   240
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5760
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label28 
         Caption         =   "SALDO"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   124
         Top             =   2970
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "TOTALES"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   6855
      End
      Begin VB.Label Label5 
         Caption         =   "PENDIENTE"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "PUNTEADA"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "HABER"
         Height          =   255
         Left            =   4380
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "DEBE"
         Height          =   255
         Left            =   2580
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Saldos hist�rico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.Frame FrameVerObservacionesCuentas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6105
      Left            =   0
      TabIndex        =   184
      Top             =   0
      Width           =   9375
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   5
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   189
         Tag             =   "Observaciones|T|S|||cuentas|obsdatos|||"
         Text            =   "frmMensajes.frx":153A8
         Top             =   720
         Width           =   7665
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   4
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   188
         Tag             =   "Observaciones|T|S|||cuentas|obsdatos|||"
         Text            =   "frmMensajes.frx":153AE
         Top             =   720
         Width           =   1005
      End
      Begin VB.CommandButton cmdVerObservaciones 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   7920
         TabIndex        =   187
         Top             =   5460
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   3915
         Index           =   6
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   185
         Tag             =   "Observaciones|T|S|||cuentas|obsdatos|||"
         Text            =   "frmMensajes.frx":153B4
         Top             =   1320
         Width           =   8775
      End
      Begin VB.Label Label22 
         Caption         =   "Descripci�n cuentas Plan General Contable  2008"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   19
         Left            =   360
         TabIndex        =   186
         Top             =   360
         Width           =   5010
      End
   End
   Begin VB.Frame Frame347DatExt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      TabIndex        =   174
      Top             =   0
      Width           =   7455
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   2
         Left            =   5040
         MaxLength       =   1
         TabIndex        =   181
         Text            =   "G"
         Top             =   2097
         Width           =   375
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   1
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   180
         Text            =   "F"
         Top             =   2097
         Width           =   375
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   0
         Left            =   240
         TabIndex        =   179
         Top             =   1320
         Width           =   6735
      End
      Begin VB.CommandButton cmd347DatExt 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5880
         TabIndex        =   176
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmd347DatExt 
         Caption         =   "Importar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4560
         TabIndex        =   175
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Letra proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   2880
         TabIndex        =   183
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label Label22 
         Caption         =   "Letra clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   182
         Top             =   2160
         Width           =   1530
      End
      Begin VB.Image Image4 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmMensajes.frx":153BA
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label22 
         Caption         =   "Fichero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   178
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label Label37 
         Caption         =   "Importar datos externos 347"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Index           =   2
         Left            =   240
         TabIndex        =   177
         Top             =   360
         Width           =   6495
      End
   End
   Begin VB.Frame FrameAyuda 
      BackColor       =   &H00E3FEFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      TabIndex        =   205
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton CmdAyuda2 
         Caption         =   "OTROS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1290
         TabIndex        =   207
         Top             =   0
         Width           =   1365
      End
      Begin VB.CommandButton CmdAyuda1 
         Caption         =   "TECLAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   206
         Top             =   0
         Width           =   1275
      End
      Begin VB.Frame FrameAyuda1 
         BackColor       =   &H00E3FEFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3555
         Left            =   0
         TabIndex        =   208
         Top             =   390
         Width           =   6225
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E3FEFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   2445
            Left            =   330
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   210
            Text            =   "frmMensajes.frx":15DBC
            Top             =   840
            Width           =   6015
         End
         Begin VB.Label Label42 
            BackColor       =   &H00E3FEFF&
            Caption         =   "Teclas R�pidas en Mantenimientos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Left            =   330
            TabIndex        =   209
            Top             =   270
            Width           =   5865
         End
      End
      Begin VB.Frame FrameAyuda2 
         BackColor       =   &H00E3FEFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3555
         Left            =   0
         TabIndex        =   211
         Top             =   420
         Width           =   6225
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E3FEFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   2445
            Left            =   330
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   213
            Top             =   750
            Width           =   5715
         End
         Begin VB.Label Label43 
            BackColor       =   &H00E3FEFF&
            Caption         =   "Otros"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Left            =   330
            TabIndex        =   212
            Top             =   270
            Width           =   5805
         End
      End
   End
   Begin VB.Frame FrameCarta347 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   0
      TabIndex        =   78
      Top             =   0
      Width           =   10215
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   6
         Left            =   6900
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   88
         Tag             =   "#Despedida"
         Text            =   "frmMensajes.frx":15E2F
         Top             =   4860
         Width           =   3075
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   7680
         MaxLength       =   100
         TabIndex        =   82
         Tag             =   "#Referencia"
         Text            =   "Text4"
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   240
         MaxLength       =   100
         TabIndex        =   81
         Tag             =   "#Asunto"
         Text            =   "Text4"
         Top             =   1740
         Width           =   6615
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Index           =   5
         Left            =   300
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   86
         Tag             =   "#Parrafo4"
         Text            =   "frmMensajes.frx":15E35
         Top             =   4860
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Index           =   4
         Left            =   3600
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   87
         Tag             =   "#Parrafo5"
         Text            =   "frmMensajes.frx":15E3B
         Top             =   4860
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Index           =   3
         Left            =   6840
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   85
         Tag             =   "#Parrafo3"
         Text            =   "frmMensajes.frx":15E41
         Top             =   2580
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Index           =   2
         Left            =   3540
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   84
         Tag             =   "#Parrafo2"
         Text            =   "frmMensajes.frx":15F3E
         Top             =   2580
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Index           =   1
         Left            =   240
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   83
         Tag             =   "#Parrafo1"
         Text            =   "frmMensajes.frx":15F44
         Top             =   2580
         Width           =   3135
      End
      Begin VB.CommandButton cmd347 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   8820
         TabIndex        =   91
         Top             =   6360
         Width           =   915
      End
      Begin VB.CommandButton cmd347 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   7620
         TabIndex        =   90
         Top             =   6360
         Width           =   915
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   100
         TabIndex        =   80
         Tag             =   "#Saludos"
         Text            =   "Text4"
         Top             =   1080
         Width           =   6615
      End
      Begin VB.Label Label22 
         Caption         =   "Despedida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   6900
         TabIndex        =   99
         Top             =   4620
         Width           =   930
      End
      Begin VB.Label Label22 
         Caption         =   "Referencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   7680
         TabIndex        =   98
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label Label22 
         Caption         =   "Asunto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   97
         Top             =   1500
         Width           =   690
      End
      Begin VB.Label Label22 
         Caption         =   "Parrafo 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   3600
         TabIndex        =   96
         Top             =   4620
         Width           =   795
      End
      Begin VB.Label Label22 
         Caption         =   "Parrafo 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   95
         Top             =   4620
         Width           =   795
      End
      Begin VB.Label Label22 
         Caption         =   "Parrafo 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   6900
         TabIndex        =   94
         Top             =   2340
         Width           =   795
      End
      Begin VB.Label Label22 
         Caption         =   "Parrafo 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3480
         TabIndex        =   93
         Top             =   2340
         Width           =   795
      End
      Begin VB.Label Label22 
         Caption         =   "Parrafo 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   92
         Top             =   2340
         Width           =   795
      End
      Begin VB.Label Label23 
         Caption         =   "Datos carta modelo 347"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   300
         TabIndex        =   89
         Top             =   240
         Width           =   4875
      End
      Begin VB.Label Label22 
         Caption         =   "Saludos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   79
         Top             =   840
         Width           =   690
      End
   End
   Begin VB.Frame FrameBloqueoEmpresas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   0
      TabIndex        =   190
      Top             =   0
      Visible         =   0   'False
      Width           =   11415
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   201
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5520
         TabIndex        =   200
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   199
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdBlEmp 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   196
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdBloqEmpre 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   9840
         TabIndex        =   194
         Top             =   6840
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5775
         Index           =   0
         Left            =   210
         TabIndex        =   192
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Empresa"
            Object.Width           =   5644
         EndProperty
      End
      Begin VB.CommandButton cmdBloqEmpre 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   8400
         TabIndex        =   191
         Top             =   6840
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5775
         Index           =   1
         Left            =   6240
         TabIndex        =   193
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Empresa"
            Object.Width           =   5644
         EndProperty
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Bloqueadas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   10050
         TabIndex        =   198
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label Label41 
         Caption         =   "Permitidas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   197
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Bloqueo de empresas por usuario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   495
         Index           =   2
         Left            =   2880
         TabIndex        =   195
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame frameSaltos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   0
      TabIndex        =   100
      Top             =   0
      Width           =   9045
      Begin VB.CommandButton cmdCabError 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   1
         Left            =   7770
         TabIndex        =   107
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdCabError 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   6690
         TabIndex        =   106
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   3255
         Left            =   4440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   103
         Text            =   "frmMensajes.frx":15F4A
         Top             =   900
         Width           =   4365
      End
      Begin VB.TextBox Text5 
         Height          =   3255
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   102
         Text            =   "frmMensajes.frx":15F50
         Top             =   900
         Width           =   4125
      End
      Begin VB.Label Label22 
         Caption         =   "Salto"
         Height          =   195
         Index           =   10
         Left            =   4440
         TabIndex        =   105
         Top             =   660
         Width           =   930
      End
      Begin VB.Label Label22 
         Caption         =   "Repetidos"
         Height          =   225
         Index           =   9
         Left            =   180
         TabIndex        =   104
         Top             =   630
         Width           =   1020
      End
      Begin VB.Label Label24 
         Caption         =   "Asientos Err�neos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   180
         TabIndex        =   101
         Top             =   240
         Width           =   4875
      End
   End
   Begin VB.Frame frameAcercaDE 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Left            =   -60
      TabIndex        =   48
      Top             =   0
      Width           =   5355
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax: 963 80 55 79"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3480
         TabIndex        =   110
         Top             =   3540
         Width           =   1560
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tfno: 902 88 88 78  -  96 380 55 79"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   240
         TabIndex        =   109
         Top             =   3540
         Width           =   3075
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "ARICONTA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   915
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   81.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1725
         Left            =   3780
         TabIndex        =   108
         Top             =   60
         Width           =   1350
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "46007 - VALENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3240
         TabIndex        =   64
         Top             =   3120
         Width           =   1620
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C/ Uruguay 11,710"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   300
         TabIndex        =   63
         Top             =   3120
         Width           =   1635
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   50
         Top             =   1200
         Width           =   3795
      End
      Begin VB.Image Image2 
         Height          =   540
         Left            =   1740
         Top             =   2460
         Width           =   2880
      End
      Begin VB.Image Image1 
         Height          =   4395
         Left            =   0
         Stretch         =   -1  'True
         Top             =   -1200
         Width           =   5355
      End
   End
   Begin VB.Frame FrameErrorRestore 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   240
      TabIndex        =   202
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4215
         Left            =   120
         TabIndex        =   203
         Top             =   600
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   7435
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label Label29 
         Caption         =   "Cambio caracteres recupera backup"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   204
         Top             =   240
         Width           =   4935
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   4920
         Picture         =   "frmMensajes.frx":15F56
         ToolTipText     =   "Quitar seleccion"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   4920
         Picture         =   "frmMensajes.frx":160A0
         ToolTipText     =   "Todos"
         Top             =   1080
         Width           =   240
      End
   End
   Begin VB.Timer tCuadre 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   6420
      Top             =   5400
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameImpCta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   115
      Top             =   0
      Width           =   6495
      Begin VB.CheckBox chkCrear 
         Caption         =   "Crear cuentas si no existen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3600
         TabIndex        =   129
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdImpCta 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   118
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtImpCta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   117
         Text            =   "Text2"
         Top             =   720
         Width           =   5655
      End
      Begin VB.CommandButton cmdImpCta 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   116
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Left            =   2640
         Picture         =   "frmMensajes.frx":161EA
         Top             =   360
         Width           =   240
      End
      Begin VB.Label lblImpCta 
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   121
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblimpCta2 
         Caption         =   "Lineas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   120
         Top             =   1320
         Width           =   615
      End
      Begin VB.Image Image4 
         Height          =   240
         Index           =   0
         Left            =   2280
         Picture         =   "frmMensajes.frx":16BEC
         Top             =   360
         Width           =   240
      End
      Begin VB.Label lblDescFich 
         Caption         =   "Fichero con datos fiscales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   119
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame FrameeMPRESAS 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmdEmpresa 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   47
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdEmpresa 
         Caption         =   "Regresar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   46
         Top             =   4800
         Width           =   975
      End
      Begin MSComctlLib.ListView lwE 
         Height          =   3615
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "dsdsd"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Empresas en el sistema"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame framaLlevarFacturas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      TabIndex        =   142
      Top             =   0
      Width           =   5775
      Begin VB.Frame FrameImportarFechas 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   147
         Top             =   1560
         Width           =   5535
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   152
            Text            =   "Text8"
            Top             =   600
            Width           =   5295
         End
         Begin VB.Image Image5 
            Height          =   240
            Left            =   840
            Picture         =   "frmMensajes.frx":175EE
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label39 
            Caption         =   "Fichero"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   153
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.TextBox txtFecha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   150
         Text            =   "Text7"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   148
         Text            =   "Text7"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox chkImportarFacturas 
         Caption         =   "Eliminar ficheros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   146
         Top             =   3960
         Width           =   2535
      End
      Begin VB.CommandButton cmdImportarFacuras 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   145
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdImportarFacuras 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   144
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Line Line6 
         BorderWidth     =   3
         X1              =   120
         X2              =   5640
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label38 
         Caption         =   "Label38"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   240
         TabIndex        =   156
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label Label40 
         Caption         =   "Label40"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   154
         Top             =   3240
         Width           =   5295
      End
      Begin VB.Label Label35 
         Caption         =   "Fecha Inicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   151
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label35 
         Caption         =   "Fecha Inicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   149
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Caption         =   "Label38"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   143
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame FrameCambioPWD 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   0
      TabIndex        =   130
      Top             =   0
      Width           =   4575
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   134
         Text            =   "Text7"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   133
         Text            =   "Text7"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   132
         Text            =   "Text7"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   131
         Text            =   "Text7"
         Top             =   960
         Width           =   4095
      End
      Begin VB.CommandButton cmdCambioPwd 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   136
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdCambioPwd 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   135
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Caption         =   "Cambio clave"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   141
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label36 
         Caption         =   "Reescribalo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   140
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label36 
         Caption         =   "Nuevo password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   139
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label36 
         Caption         =   "Password actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   138
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label36 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   137
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame frameamort 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton Command3 
         Caption         =   "Salir"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   42
         Top             =   5760
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "porcentaje"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   4080
         TabIndex        =   41
         Top             =   4920
         Width           =   1065
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Coefi. maximo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   5760
         TabIndex        =   40
         Top             =   3720
         Width           =   1200
      End
      Begin VB.Label Label11 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   2040
         TabIndex        =   39
         Top             =   5040
         Width           =   120
      End
      Begin VB.Line Line4 
         Index           =   4
         X1              =   4080
         X2              =   5280
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label Label10 
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   4440
         TabIndex        =   38
         Top             =   5280
         Width           =   465
      End
      Begin VB.Label Label10 
         Caption         =   "Valor adquisici�n   x  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   2280
         TabIndex        =   37
         Top             =   5040
         Width           =   1755
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "PORCENTAJE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   36
         Top             =   4440
         Width           =   1635
      End
      Begin VB.Label Label11 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   2040
         TabIndex        =   35
         Top             =   3840
         Width           =   120
      End
      Begin VB.Line Line4 
         Index           =   3
         X1              =   5880
         X2              =   6840
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label10 
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   6120
         TabIndex        =   34
         Top             =   4080
         Width           =   345
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "(Valor adquisici�n -amort. acumulada)  x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   2280
         TabIndex        =   33
         Top             =   3840
         Width           =   3435
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "DEGRESIVO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   32
         Top             =   3240
         Width           =   1440
      End
      Begin VB.Label Label11 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   27
         Top             =   2760
         Width           =   120
      End
      Begin VB.Label Label11 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   26
         Top             =   1560
         Width           =   120
      End
      Begin VB.Line Line4 
         Index           =   1
         X1              =   2280
         X2              =   5280
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label10 
         Caption         =   "a�os de vida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   25
         Top             =   3000
         Width           =   1065
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Valor adquisici�n - valor residual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   24
         Top             =   2640
         Width           =   2745
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "LINEAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   975
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   2280
         X2              =   4200
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label10 
         Caption         =   "a�os de vida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   22
         Top             =   1800
         Width           =   1065
      End
      Begin VB.Label Label10 
         Caption         =   "Valor adquisici�n"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2640
         TabIndex        =   21
         Top             =   1440
         Width           =   1425
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "TABLAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000080&
         X1              =   5280
         X2              =   2040
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   1200
         X2              =   1800
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipos de amortizaci�n"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   360
         Width           =   3450
      End
   End
   Begin VB.Label Label11 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1800
      TabIndex        =   31
      Top             =   600
      Width           =   120
   End
   Begin VB.Line Line4 
      Index           =   2
      X1              =   2040
      X2              =   3960
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label10 
      Caption         =   "a�os de vida"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2520
      TabIndex        =   30
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label Label10 
      Caption         =   "Valor adquisici�n"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2400
      TabIndex        =   29
      Top             =   480
      Width           =   1425
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TABLAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '1.- Saldos historico
    '2.- Comprobar saldos
    '3.- Mostrar tipos de amortizacion
    '4.- Seleccionar empresas
    
    '5.- Es, como si fuera comprobar saldos , pero se lanza y se cierra autmaticamente
    '6.- El acerca DE
    
    '7.- Nueva linea en configuracion balances
    '8.- Modificar linea balances
    
    
    '9.- Nueva CTA de configuracion balances
    '10- MODIFICAR  "   "             "
        
    '11- Carta modelo 347
    '12- Asientos con saltos y/o repetidos
    
    '13- Importar datos fiscales de las cuentas
    
    '15- Cambio Password
    
        
    '16- Traspaso de facturas entre PC's. EXP
    '17-   "                  "           IMPORTAR
    
    
     '18- Importes punteo
     '19- Copiar de un balance a OTRO
     '20- Importar fichero datos 347 externo
     '21- Ver OBSERVACIONES cuentas
     '22- Ver empresas bloquedas
     '23- Menu de ayuda
     '24- Iconos de pantalla principal
     '25- Informe de base de datos
     '26- Show processlist
     
     '27- Cobros de la factura
     '28- Pagos de la factura
     
     '29- Asiento de liquidacion
     
     '30- Asientos descuadrados
     
     '***** TESORERIA *****
     '50- Facturas de Reclamaciones
    
Public Parametros As String
    '1.- Vendran empipados: Cuenta, PunteadoD, punteadoH, pdteD,PdteH

Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private PrimeraVez As Boolean

Dim i As Integer
Dim SQL As String
Dim RS As Recordset
Dim ItmX As ListItem
Dim Errores As String
Dim NE As Integer
Dim Ok As Integer

Dim CampoOrden As String
Dim Orden As Boolean


Private Sub cmd347_Click(Index As Integer)
    If Index = 0 Then
        If Not GuardarDatosCarta Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmd347DatExt_Click(Index As Integer)
Dim B As Boolean
    If Index = 0 Then
        If Text9(0).Text = "" Then Exit Sub
        
        
        If Dir(Text9(0).Text, vbArchive) = "" Then
            MsgBox "Fichero no encontrado", vbExclamation
            Exit Sub
        End If
                
        If Text9(1).Text = "" Or Text1(2).Text = "" Then
            MsgBox "Ponga las letras para clientes / proveedores", vbExclamation
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        B = ImportarDatosExternos347
        Screen.MousePointer = vbDefault
        
    Else
        B = True
    End If
    If B Then Unload Me
End Sub

Private Sub cmdAcepIconos_Click()
Dim i As Integer
Dim SQL As String
Dim CadenaIconos As String


Dim K As Integer
Dim J As Integer
Dim Px As Single
Dim Py As Single


Dim H As Integer
Dim MargenX As Single
Dim MargenY As Single
Dim Ocupado As Boolean
    'Ponemos los que ha desmarcado a cero
    CadenaIconos = ""
    For i = 1 To ListView6.ListItems.Count
        If Not ListView6.ListItems(i).Checked Then
            CadenaIconos = CadenaIconos & ListView6.ListItems(i).Text & ","
        End If
    Next i
    If CadenaIconos <> "" Then
        CadenaIconos = Mid(CadenaIconos, 1, Len(CadenaIconos) - 1)
        SQL = "update menus_usuarios set posx = 0, posy = 0, vericono = 0 where aplicacion = 'ariconta' and codusu = " & vUsu.Id & " and codigo in (" & CadenaIconos & ")"
        Conn.Execute SQL
    End If

    

    For i = 1 To ListView6.ListItems.Count
        If ListView6.ListItems(i).Checked Then
            'SI NO ERA VISIBLE le busco el hueco
            If ListView6.ListItems(i).SubItems(2) = "0" Then
                SQL = ""
                For J = 1 To 8
                    For K = 1 To 5
                        DevuelCoordenadasCuadricula J, K, Px, Py
                        Ocupado = False
                        'Busco hueco
                        For H = 1 To ListView6.ListItems.Count
                            If ListView6.ListItems(H).SubItems(2) = "1" Then
                                MargenX = Abs(Px - CSng(ListView6.ListItems(H).SubItems(3)))
                                MargenY = Abs(Py - CSng(ListView6.ListItems(H).SubItems(4)))
                                
                                If MargenX < 300 And MargenY < 300 Then
                                    'HUECO
                                    Ocupado = True
                                    Exit For
                                End If
                            End If
                        Next H
                        
                        If Not Ocupado Then
                            'OK. Este es. Lo ponemos a true y actualizamos BD
                            ListView6.ListItems(i).SubItems(2) = "1"
                            ListView6.ListItems(i).SubItems(3) = Px
                            ListView6.ListItems(i).SubItems(4) = Py
                            SQL = "update menus_usuarios set posx = " & DBSet(Px, "N")
                            SQL = SQL & ", posy = " & DBSet(Py, "N") & ", vericono = 1 where "
                            SQL = SQL & "aplicacion = 'ariconta' and codusu = " & vUsu.Id
                            SQL = SQL & " and codigo =" & DBSet(ListView6.ListItems(i).Text, "T")
                            Conn.Execute SQL
                           Exit For
                        End If
                    Next K
                    If SQL <> "" Then Exit For
                Next J
            End If
           
        End If
    Next i
    Reorganizar = True
    
    Unload Me
End Sub





Private Sub CmdAyuda1_Click()
    Me.FrameAyuda1.Visible = True
    Me.FrameAyuda2.Visible = False
End Sub

Private Sub CmdAyuda2_Click()
    Me.FrameAyuda1.Visible = False
    Me.FrameAyuda2.Visible = True
End Sub



Private Sub cmdBalance_Click(Index As Integer)
    If Index = 1 Then
        CadenaDesdeOtroForm = ""
        Unload Me
    Else
        If Text1(0).Text = "" Then
            MsgBox "Primer campo obligatorio", vbExclamation
            Exit Sub
        End If
        If InsertarModificar Then Unload Me
    End If
End Sub

Private Sub cmdBlEmp_Click(Index As Integer)

    Select Case Index
    Case 0, 1
        'Index Me dira que listview
        For Ok = ListView2(Index).ListItems.Count To 1 Step -1
            If ListView2(Index).ListItems(Ok).Selected Then
                i = ListView2(Index).ListItems(Ok).Index
                PasarUnaEmpresaBloqueada Index = 0, i
            End If
        Next Ok
    Case Else
        If Index = 2 Then
            Ok = 0
        Else
            Ok = 1
        End If
        For NumRegElim = ListView2(Ok).ListItems.Count To 1 Step -1
            PasarUnaEmpresaBloqueada Ok = 0, ListView2(Ok).ListItems(NumRegElim).Index
        Next NumRegElim
        Ok = 0
    End Select
End Sub



Private Sub PasarUnaEmpresaBloqueada(ABLoquedas As Boolean, indice As Integer)
Dim Origen As Integer
Dim Destino As Integer
Dim IT
    If ABLoquedas Then
        Origen = 0
        Destino = 1
        NE = 2
    Else
        Origen = 1
        Destino = 0
        NE = 1 'icono
    End If
    
    SQL = ListView2(Origen).ListItems(indice).Key
    Set IT = ListView2(Destino).ListItems.Add(, SQL)
    IT.SmallIcon = NE
    IT.Text = ListView2(Origen).ListItems(indice).Text
    IT.SubItems(1) = ListView2(Origen).ListItems(indice).SubItems(1)

    'Borramos en origen
    ListView2(Origen).ListItems.Remove indice
End Sub

Private Sub cmdBloqEmpre_Click(Index As Integer)
    If Index = 0 Then
        SQL = "DELETE FROM usuarios.usuarioempresa WHERE codusu =" & Parametros
        Conn.Execute SQL
        SQL = ""
        For i = 1 To ListView2(1).ListItems.Count
            SQL = SQL & ", (" & Parametros & "," & Val(Mid(ListView2(1).ListItems(i).Key, 2)) & ")"
        Next i
        If SQL <> "" Then
            'Quitmos la primera coma
            SQL = Mid(SQL, 2)
            SQL = "INSERT INTO usuarios.usuarioempresa(codusu,codempre) VALUES " & SQL
            If Not EjecutaSQL(SQL) Then MsgBox "Se han producido errores insertando datos", vbExclamation
        End If
    End If
    Unload Me
End Sub

Private Sub cmdCabError_Click(Index As Integer)
Dim RS As ADODB.Recordset
Dim J As Long
Dim ii As Long
Dim Anyo As Integer

    If Index = 1 Then
        Unload Me
    Else
        Screen.MousePointer = vbHourglass
        Anyo = 0
        i = 0
        Do
          
            SQL = "select numasien,fechaent from hcabapu where fechaent >= '"
            SQL = SQL & Format(DateAdd("yyyy", Anyo, vParam.fechaini), FormatoFecha)
            SQL = SQL & "' AND fechaent <= '" & Format(DateAdd("yyyy", Anyo, vParam.fechafin), FormatoFecha) & "' ORDER By NumAsien"
            Set RS = New ADODB.Recordset
            RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            ii = 0
            While Not RS.EOF
               J = RS.Fields(0)
               'Igual
                If J - ii = 0 Then
                    
                    SQL = Format(J, "00000")
                    SQL = SQL & "  -  " & Format(RS!FechaEnt, "dd/mm/yyyy")
                    Text5.Text = Text5.Text & SQL & vbCrLf
                    i = i + 1
                Else
                    If J - ii > 1 Then
                        If J - ii = 2 Then
                            SQL = Format(J - 1, "00000")
                        Else
                            SQL = "Entre " & Format(ii, "00000") & "  y  " & Format(J, "00000")
                        End If
                        SQL = SQL & " (" & CStr(Year(vParam.fechaini) + Anyo) & ")"
                        Text6.Text = Text6.Text & SQL & vbCrLf
                        i = i + 1
                    End If
                End If
                ii = J
                'Refrescamos
                If i > 50 Then
                    Text5.Refresh
                    Text6.Refresh
                    i = 0
                End If
                
                '
                RS.MoveNext
            Wend
            RS.Close
            Anyo = Anyo + 1
        Loop Until Anyo > 1
        Me.Refresh
        Screen.MousePointer = vbDefault
        cmdCabError(0).Enabled = False
    End If
End Sub

Private Sub cmdCambioPwd_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    For i = 1 To Text7.Count - 1
        Text7(i).Text = Trim(Text7(i).Text)
        If Text7(i).Text = "" Then
            MsgBox "Hay que rellenar todos los campos", vbExclamation
            Exit Sub
        End If
    Next i
    
    
    'Todos rellenados
    'Ha puesto la clave actual real
    If Text7(1).Text <> vUsu.PasswdPROPIO Then
        MsgBox "Clave actual incorrecta", vbExclamation
        Exit Sub
    End If
    
    If Text7(2).Text <> Text7(3).Text Then
        MsgBox "Mal reescrita la clave nueva", vbExclamation
        Exit Sub
    End If
    
    
    If InStr(1, Text7(2).Text, "'") > 0 Then
        MsgBox "Clave nueva contiene caracter no permitido", vbExclamation
        Exit Sub
    End If
    
    
    'UPDATEAMOS
    On Error Resume Next
    SQL = "UPDATE Usuarios.Usuarios Set passwordpropio='" & Text7(2).Text
    SQL = SQL & "' WHERE codusu = " & (vUsu.Codigo Mod 1000)
    
    Conn.Execute SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cambio clave"
    Else
        vUsu.PasswdPROPIO = Text7(2).Text
        MsgBox "Cambio de clave realizado con �xito", vbInformation
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub cmdCanIconos_Click()
    Reorganizar = False
    Unload Me
End Sub

Private Sub CmdContabilizar_Click()
    CadenaDesdeOtroForm = "OK"
    Unload Me
End Sub

Private Sub cmdCtaBalan_Click(Index As Integer)
    If Index = 1 Then
        CadenaDesdeOtroForm = ""
    Else
        If Text3.Text = "" Then
            MsgBox "la cuenta no puede estar en blanco", vbExclamation
            Exit Sub
        End If
        If Not IsNumeric(Text3.Text) Then
            MsgBox "La cuenta debe ser num�rica", vbExclamation
            Exit Sub
        End If
        'Esto es el OPTION
        SQL = ""
        For i = 0 To 2
            If Option1(i).Value Then SQL = SQL & Mid(Option1(i).Caption, 1, 1)
        Next i
        If SQL = "" Then
            MsgBox "Seleccione una opci�n de la cuenta (Saldo - Debe - Haber )"
            Exit Sub
        End If
        
        'RESTA y la resta
        SQL = SQL & "|" & Abs(Me.chkResta.Value)
        CadenaDesdeOtroForm = Text3.Text & "|" & SQL & "|"
    End If
    Unload Me
End Sub

Private Sub cmdEmpresa_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    If Index = 0 Then
        SQL = ""
        Parametros = ""
        For i = 1 To lwE.ListItems.Count
            If Me.lwE.ListItems(i).Checked Then
                SQL = SQL & Me.lwE.ListItems(i).Text & "|"
                Parametros = Parametros & "1" 'Contador
            End If
        Next i
        CadenaDesdeOtroForm = Len(Parametros) & "|" & SQL
        'Vemos las conta
        SQL = ""
        For i = 1 To lwE.ListItems.Count
            If Me.lwE.ListItems(i).Checked Then
                SQL = SQL & Me.lwE.ListItems(i).Tag & "|"
            End If
        Next i
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & SQL
    End If
    Unload Me
End Sub

Private Sub cmdImpCta_Click(Index As Integer)
Dim Cad As String

    If Index = 1 Then
        Unload Me
        Exit Sub
    End If

    txtImpCta.Text = Trim(txtImpCta.Text)
    
    If txtImpCta.Text = "" Then Exit Sub
    
    If Dir(txtImpCta.Text) = "" Then
        MsgBox "El fichero: " & txtImpCta.Text & " NO existe.", vbExclamation
        Exit Sub
    End If
    
    Cad = "Seguro que desa continuar con la importaci�n de los datos fiscales?"
    If MsgBox(Cad, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    'Ya esta preparado comprobemos k podemos abrir la conexion

        lblimpCta2.Caption = "Lineas"
        lblImpCta.Caption = "0"
        Errores = ""
        NE = 0
        Ok = 0
        Me.Refresh
        CadenaDesdeOtroForm = ""
        'La contabilidad existe
        HacerImportacion
        'Si hay errores
        If NE > 0 Then
            Errores = Ok & " lineas pasadas con exito." & vbCrLf & vbCrLf & Errores
            ImprimeFichero
        Else
            MsgBox Ok & " lineas pasadas con exito", vbInformation
        End If
        CadenaDesdeOtroForm = ""
    Screen.MousePointer = vbDefault
    lblimpCta2.Caption = ""
    lblImpCta.Caption = ""
End Sub

Private Sub cmdImportarFacuras_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    'COMPROBACIONES
    If Opcion = 17 Then
        If Text8.Text = "" Then
            MsgBox "Fichero en blanco", vbExclamation
            Exit Sub
        End If
        
        If Dir(Text8.Text, vbArchive) = "" Then
            MsgBox "Fichero no existe.", vbExclamation
            Exit Sub
        End If
        
        SQL = "Va a realizar la importaci�n de datos  de facturas en la empresa: " & vbCrLf & vbCrLf & vEmpresa.nomempre
        SQL = SQL & "(" & vEmpresa.nomresum & ") - Conta: " & vEmpresa.codempre & vbCrLf & vbCrLf & "�Desea continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
    End If
    
    
    If Opcion = 16 Then
        ExportarFactur
    Else
        ImportarFicheroFac
    End If
End Sub


Private Sub ExportarFactur()
    On Error GoTo EExportarDatosF
        'Primero borramos el temporal facturas
        Errores = App.Path & "\tmpexpdatos.tmp"
        If Dir(Errores, vbArchive) <> "" Then Kill Errores
        NE = FreeFile
        Open Errores For Output As NE
        'Primero proveedores
        ExportarDatosFacturas True
        'Clientes
        ExportarDatosFacturas False
        Close NE
        
        
        
        Text8.Text = ""
        Image4_Click 0
        If Text8.Text <> "" Then
        'Hay k copiar el archivo
            Errores = App.Path & "\tmpexpdatos.tmp"
            CopiarArchivo
        Else
           ' MsgBox "Opcion cancelada", vbExclamation
        End If
        
        
        
        
        Exit Sub



EExportarDatosF:
    MuestraError Err.Number, Err.Description
    On Error Resume Next
    Close NE
End Sub



Private Sub cmdPunteo_Click()
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdRegresar_Click()
    Unload Me
End Sub

Private Sub cmdVerObservaciones_Click()
    Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Command2_Click()
Dim Digitos As Integer
    ListView1.ListItems.Clear
    Me.ProgressBar1.Value = 0
    Me.ProgressBar1.Max = vEmpresa.numnivel + 1
    Me.ProgressBar1.Visible = True
    Screen.MousePointer = vbHourglass
    Me.ProgressBar1.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        Case 5
            Screen.MousePointer = vbHourglass
            Me.tCuadre.Enabled = True
        Case 21
            cargarObservacionesCuenta
        Case 22
            cargaempresasbloquedas
            
        Case 24
            CargaIconosVisibles
            
        Case 25
            CargaInformeBBDD
        
        Case 26
            CargaShowProcessList
        
        Case 27
            CargaCobrosFactura
        Case 28
            CargaPagosFactura
            
        Case 29
            CargarAsiento
        Case 30
            CargarAsientosDescuadrados
        Case 31
            CargarFacturasSinAsientos
        Case 50
            CargarFacturasReclamaciones
        Case 51
            CargarFacturasRemesas
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And Opcion = 23 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
Dim W, H
    Me.tCuadre.Enabled = False
    PrimeraVez = True
    Me.frameSaldosHco.Visible = False
    Me.frameCalculoSaldos.Visible = False
    Me.frameamort.Visible = False
    Me.FrameeMPRESAS.Visible = False
    Me.frameAcercaDE.Visible = False
    Me.frameCtasBalance.Visible = False
    Me.FrameCarta347.Visible = False
    Me.frameSaltos.Visible = False
    frameBalance.Visible = False
    FrameImpCta.Visible = False
    Me.FrameCambioPWD.Visible = False
    framaLlevarFacturas.Visible = False
    Me.FrameImpPunteo.Visible = False
    Me.Frame347DatExt.Visible = False
    FrameVerObservacionesCuentas.Visible = False
    Me.FrameBloqueoEmpresas.Visible = False
    Me.FrameAyuda.Visible = False
    Me.FrameIconosVisibles.Visible = False
    Me.FrameInformeBBDD.Visible = False
    Me.FrameShowProcess.Visible = False
    Me.FrameCobros.Visible = False
    Me.FrameAsientoLiquida.Visible = False
    Me.FrameDescuadre.Visible = False
    Me.FrameReclamaciones.Visible = False
    
    Select Case Opcion
    Case 1
        Me.Caption = "C�lculo de saldo"
        W = frameSaldosHco.Width
        H = Me.frameSaldosHco.Height
        Me.frameSaldosHco.Visible = True
        
        CargaValoresHco
        Command1(0).Cancel = True
    Case 2
        Me.Caption = "Comprobacion saldos"
        W = Me.frameCalculoSaldos.Width
        H = Me.frameCalculoSaldos.Height + 150
        Me.frameCalculoSaldos.Visible = True
        Command1(1).Enabled = True
        Command2.Enabled = True
    Case 3
        Me.Caption = "Informaci�n tipo amortizaci�n"
        W = Me.frameamort.Width
        H = Me.frameamort.Height + 200
        Me.frameamort.Visible = True
    Case 4
        Me.Caption = "Seleccion"
        W = Me.FrameeMPRESAS.Width
        H = Me.FrameeMPRESAS.Height + 200
        Me.FrameeMPRESAS.Visible = True
        cargaempresas
    Case 5
        'Lanzar automaticamente la comprobaci�n de saldo
        Me.Caption = "Comprobacion saldos"
        W = Me.frameCalculoSaldos.Width
        H = Me.frameCalculoSaldos.Height
        Me.frameCalculoSaldos.Visible = True
        Command1(1).Enabled = False
        Command2.Enabled = False
    Case 6
        CargaImagen
        Me.Caption = "Acerca de ....."
        W = Me.frameAcercaDE.Width
        H = Me.frameAcercaDE.Height + 200
        Me.frameAcercaDE.Visible = True
        Label13.Caption = "Versi�n:  " & App.Major & "." & App.Minor & "." & App.Revision & " "
    Case 7, 8
        Me.Caption = "Lineas configuracion balance"
        W = Me.frameBalance.Width
        H = Me.frameBalance.Height + 300
        Me.frameBalance.Visible = True
        PonerCamposBalance
    Case 9, 10
        If Opcion = 9 Then
            Me.cmdCtaBalan(0).Caption = "Insertar"
        Else
            Me.cmdCtaBalan(0).Caption = "Modificar"
        End If
        Me.Caption = "Cuentas configuracion balances"
        W = Me.frameCtasBalance.Width
        H = Me.frameCtasBalance.Height + 300
        frameCtasBalance.Visible = True
        PonerCamposCtaBalance
        
    Case 11
        'Carta modelo 347
        Me.Caption = "Datos carta modelo 347"
        W = Me.FrameCarta347.Width
        H = Me.FrameCarta347.Height + 300
        Me.FrameCarta347.Visible = True
        CargarDatosCarta
    Case 12
        'Saltos y repedtidos
        Me.Caption = "B�squeda cabeceras asientos incorrectos"
        W = Me.frameSaltos.Width
        H = Me.frameSaltos.Height + 300
        Me.frameSaltos.Visible = True
        Me.cmdCabError(0).Enabled = True
        Text5.Text = ""
        Text6.Text = ""
        cmdCabError(1).Cancel = True
    Case 13
        Me.Caption = "Importar datos fiscales de las cuentas"
        W = Me.FrameImpCta.Width
        H = Me.FrameImpCta.Height + 450
        Me.FrameImpCta.Visible = True
        cmdImpCta(1).Cancel = True
        txtImpCta.Text = ""
        Me.lblImpCta.Caption = ""
        Me.lblimpCta2.Caption = ""
    Case 15
        'Cambio password usuario
        Me.Caption = "Cambio password"
        W = Me.FrameCambioPWD.Width
        H = Me.FrameCambioPWD.Height + 300
        Me.FrameCambioPWD.Visible = True
        Text7(0).Text = vUsu.Nombre
        For i = 1 To 3
            Text7(i).Text = ""
        Next i
        cmdCambioPwd(1).Cancel = True
    Case 16, 17
        Text8.Text = ""
        Caption = "UTIL. FACTURAS"
        W = Me.framaLlevarFacturas.Width
        H = Me.framaLlevarFacturas.Height + 300
        Me.framaLlevarFacturas.Visible = True
        chkImportarFacturas.Visible = Opcion = 17
        FrameImportarFechas.Visible = Opcion = 17
        If Opcion = 16 Then
            Label38(0).Caption = "EXPORTAR"
        Else
            Label38(0).Caption = "IMPORTAR"
        End If
        Label38(1).Caption = vEmpresa.nomempre & "   (" & vEmpresa.nomresum & ")"
        Me.txtFecha(2).Text = Format(vParam.fechaini, "dd/mm/yyyy")
        Me.txtFecha(3).Text = Format(Now, "dd/mm/yyyy")
        Label40.Caption = ""
        cmdImportarFacuras(1).Cancel = True
        
    Case 18
        Me.FrameImpPunteo.Visible = True
        Caption = "Importes"
        For i = 0 To 8
            Me.txtImporteP(i).Text = RecuperaValor(Parametros, i + 1)
        Next i
        W = Me.FrameImpPunteo.Width
        H = Me.FrameImpPunteo.Height + 300
        cmdPunteo.Cancel = True
    Case 20
        Me.Frame347DatExt.Visible = True
        Caption = "Importar datos 347"
        W = Me.Frame347DatExt.Width
        H = Me.Frame347DatExt.Height + 300
        Me.cmd347DatExt(1).Cancel = True
        
    Case 21
        'obseravaciones cuenta
        FrameVerObservacionesCuentas.Visible = True
        Caption = "Observaciones P.G.C."
        W = Me.FrameVerObservacionesCuentas.Width
        H = Me.FrameVerObservacionesCuentas.Height + 300
        
        cmdVerObservaciones.Cancel = True
        
        
    Case 22
        Me.FrameBloqueoEmpresas.Visible = True
        Caption = "Bloqueo empresas"
        W = Me.FrameBloqueoEmpresas.Width
        H = Me.FrameBloqueoEmpresas.Height + 300
        'Como cuando venga por esta opcion, viene llamado desde el manteusu
        Me.ListView2(0).SmallIcons = frmMantenusu.ImageList1
        Me.ListView2(1).SmallIcons = frmMantenusu.ImageList1
        Me.cmdBloqEmpre(1).Cancel = True
        
        
    Case 23
        Me.FrameAyuda.Visible = True
        Caption = "Ayuda Ariconta"
        W = Me.FrameAyuda.Width
        H = Me.FrameAyuda.Height + 300
        
        
    Case 24 ' iconos visbles
        Me.Caption = "Panel de Control"
        Me.FrameIconosVisibles.Visible = True
        W = Me.FrameIconosVisibles.Width
        H = Me.FrameIconosVisibles.Height + 300
        
    Case 25 ' informe de base de datos
        Me.Caption = "Informaci�n de Base de Datos"
        Me.FrameInformeBBDD.Visible = True
        W = Me.FrameInformeBBDD.Width
        H = Me.FrameInformeBBDD.Height + 300
        
        Me.Label47.Caption = "Ejercicio " & vParam.fechaini & " a " & vParam.fechafin
        Me.Label48.Caption = "Ejercicio " & DateAdd("yyyy", 1, vParam.fechaini) & " a " & DateAdd("yyyy", 1, vParam.fechafin)
        
    Case 26 ' show process list
        Me.Caption = "Informaci�n de Procesos del Sistema"
        Me.FrameShowProcess.Visible = True
        W = Me.FrameShowProcess.Width
        H = Me.FrameShowProcess.Height + 300
        
        Label53.Caption = Label53.Caption & " Ariconta" & vEmpresa.codempre & " (" & vEmpresa.nomempre & ")"
        
    Case 27 ' cobros de facturas
        Me.Caption = "Facturas de Cliente"
        Label52.Caption = "Cobros de la Factura " & RecuperaValor(Parametros, 1) & "-" & Format(RecuperaValor(Parametros, 2), "0000000") & " de fecha " & RecuperaValor(Parametros, 3)
        Me.FrameCobros.Visible = True
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height + 300
    
    Case 28 ' pagos de facturas
        Me.Caption = "Facturas de Proveedor"
        Label52.Caption = "Pagos de la Factura " & RecuperaValor(Parametros, 1) & "-" & RecuperaValor(Parametros, 3) & " de fecha " & RecuperaValor(Parametros, 4)
        Me.FrameCobros.Visible = True
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height + 300
        
        
    Case 29 ' asiento de liquidacion
        Me.Caption = "Asiento de Liquidaci�n"
        Me.FrameAsientoLiquida.Visible = True
        W = Me.FrameAsientoLiquida.Width
        H = Me.FrameAsientoLiquida.Height + 300
        
        
    Case 30 ' asientos descuadrados
        Me.Caption = "Asientos descuadrados"
        Me.FrameDescuadre.Visible = True
        W = Me.FrameDescuadre.Width
        H = Me.FrameDescuadre.Height + 300
        
    Case 31 ' facturas sin asientos
        Me.Caption = "Facturas sin asiento"
        Me.FrameDescuadre.Visible = True
        W = Me.FrameDescuadre.Width
        H = Me.FrameDescuadre.Height + 300
        
    Case 50 ' facturas de reclamaciones
        Me.Caption = "Facturas Reclamadas"
        Me.Label30.Caption = "Reclamaci�n a " & RecuperaValor(Parametros, 2) & " de fecha " & RecuperaValor(Parametros, 3)
        Me.FrameReclamaciones.Visible = True
        W = Me.FrameReclamaciones.Width
        H = Me.FrameReclamaciones.Height + 300
        
        Orden = True
        CampoOrden = "fecfactu"
    
    Case 51 ' facturas de remesas
        Me.Caption = "Facturas de Remesa"
        Me.Label30.Caption = "Remesa " & RecuperaValor(Parametros, 1) & " / " & RecuperaValor(Parametros, 2)
        Me.FrameReclamaciones.Visible = True
        W = Me.FrameReclamaciones.Width
        H = Me.FrameReclamaciones.Height + 300
    
        Orden = True
        CampoOrden = "fecfactu"
            
    End Select
    Me.Width = W + 120
    Me.Height = H + 120
End Sub




Private Sub CargaValoresHco()
'Lo que hace es dado el parametro scamos la cuenta, nomcuenta, saldos
'1.- Vendran empipados: Cuenta, PunteadoD, punteadoH, pdteD,PdteH

Label6.Caption = RecuperaValor(Parametros, 1)
For i = 0 To 3
    Me.txtsaldo(i).Text = RecuperaValor(Parametros, i + 2)
Next i
CalculaSaldosFinales
End Sub



Private Sub CalculaSaldosFinales()
Dim Importe As Currency
    For i = 0 To 3
        Importe = ImporteFormateado(txtsaldo(i).Text)
        txtsaldo(i).Tag = Importe
    Next i
    txtsaldo(4).Text = ""
    txtsaldo(5).Text = ""
    
    Importe = CCur(txtsaldo(1).Tag) + CCur(txtsaldo(3).Tag)
    txtsaldo(5).Tag = Importe
    txtsaldo(5).Text = Format(Importe, FormatoImporte)
    Importe = CCur(txtsaldo(0).Tag) + CCur(txtsaldo(2).Tag)
    txtsaldo(4).Tag = Importe
    txtsaldo(4).Text = Format(Importe, FormatoImporte)
    Importe = CCur(txtsaldo(5).Tag) - CCur(txtsaldo(4).Tag)
    txtsaldo(6).Text = ""
    txtsaldo(7).Text = ""
    If Importe <> 0 Then
        If Importe > 0 Then
            txtsaldo(6).Text = Format(Importe, FormatoImporte)
        Else
            txtsaldo(7).Text = Format(Abs(Importe), FormatoImporte)
        End If
    End If
    
    
    'Ahora veremos si tiene del periodo
    txtsaldo(8).Text = ""
    txtsaldo(9).Text = ""
    SQL = RecuperaValor(Parametros, 6)
    If SQL = "" Then
        NE = 0
        
    Else
        NE = 1
        Importe = CCur(SQL)
        If Importe >= 0 Then
            txtsaldo(8).Text = Format(Importe, FormatoImporte)
        Else
            txtsaldo(9).Text = Format(Abs(Importe), FormatoImporte)
        End If
    End If
    
    Label28(1).Visible = (NE = 1)
    txtsaldo(9).Visible = (NE = 1)
    txtsaldo(8).Visible = (NE = 1)
    
    'Descripcion cuenta
    SQL = Trim(RecuperaValor(Parametros, 7))   'Descripcion cuenta
    If SQL <> "" Then SQL = " - " & SQL
    Label6.Caption = Label6.Caption & SQL
    
    
    
    'NUEVO 14 Febrero... San valentin
    Importe = CCur(txtsaldo(2).Tag) - CCur(txtsaldo(3).Tag)
    Image6(0).ToolTipText = "Saldo punteado: " & Format(Importe, FormatoImporte)
    Importe = CCur(txtsaldo(0).Tag) - CCur(txtsaldo(1).Tag)
    Image6(1).ToolTipText = "Saldo pendiente: " & Format(Importe, FormatoImporte)
    
    
    
End Sub


Private Sub cargaempresas()
Dim Prohibidas As String
On Error GoTo Ecargaempresas

    VerEmresasProhibidas Prohibidas
    
    SQL = "Select * from Usuarios.Empresasariconta "
    If vUsu.Codigo > 0 Then SQL = SQL & " WHERE codempre<100 and conta like 'ariconta%'"
    SQL = SQL & " order by codempre"
    Set lwE.SmallIcons = Me.ImageList1
    lwE.ListItems.Clear
    Set RS = New ADODB.Recordset
    i = -1
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        SQL = "|" & RS!codempre & "|"
        If InStr(1, Prohibidas, SQL) = 0 Then
            Set ItmX = lwE.ListItems.Add(, , RS!nomempre, , 3)
            ItmX.Tag = RS!codempre
            If ItmX.Tag = vEmpresa.codempre Then
                If CadenaDesdeOtroForm = "" Then
                    ItmX.Checked = True
                    i = ItmX.Index
                End If
            End If
            ItmX.ToolTipText = RS!CONTA
        End If
        RS.MoveNext
    Wend
    RS.Close
    If i > 0 Then Set lwE.SelectedItem = lwE.ListItems(i)
    
    CadenaDesdeOtroForm = ""
    
Ecargaempresas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos empresas"
    Set RS = Nothing
End Sub

Private Sub VerEmresasProhibidas(ByRef VarProhibidas As String)

On Error GoTo EVerEmresasProhibidas
    VarProhibidas = "|"
    SQL = "Select codempre from Usuarios.usuarioempresa WHERE codusu = " & (vUsu.Codigo Mod 1000)
    SQL = SQL & " order by codempre"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
          VarProhibidas = VarProhibidas & RS!codempre & "|"
          RS.MoveNext
    Wend
    RS.Close
    Exit Sub
EVerEmresasProhibidas:
    MuestraError Err.Number, Err.Description & vbCrLf & " Consulte soporte t�cnico"
    Set RS = Nothing
End Sub




Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    Text3.Text = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub Image3_Click()
    Set frmC = New frmColCtas
    frmC.ConfigurarBalances = 1
    frmC.DatosADevolverBusqueda = "0|1"
    frmC.Show vbModal
    Set frmC = Nothing
End Sub



Private Sub Image4_Click(Index As Integer)
   On Error GoTo ELee
    
    With cd1
        .CancelError = True
        If Index = 1 Then
            'importacion datos externos 347
            .DialogTitle = "Fichero datos externos 347"
        Else
            If Opcion = 16 Then
                .DialogTitle = "DESTINO. Nuevo nombre de fichero."
            Else
                .DialogTitle = "Seleccione archivo importaci�n"
            End If
        End If
        .InitDir = "C:\"
        .ShowOpen
        Select Case Opcion
        Case 16
            'A mano. Es para el fichero de gaurdar datos
            Text8.Text = .FileName
        Case 17
            Text8.Text = .FileName
        Case 20
            Text9(0).Text = .FileName
        Case Else
            txtImpCta.Text = .FileName
        End Select
    End With
    
    Exit Sub
ELee:
    Err.Clear
End Sub

Private Sub Image5_Click()
    Image4_Click 0
End Sub

Private Sub Image6_Click(Index As Integer)
    MsgBox Image6(Index).ToolTipText, vbInformation
End Sub

Private Sub ImageAyudaImpcta_Click()
    'Ejemplo
    '43000001|SECUVE, S.L.|RIU VERT  N� 7|46600|ALZIRA|VALENCIA|B97301808|
    SQL = "Formato para la importaci�n de datos fiscales. " & vbCrLf & vbCrLf & vbCrLf
    SQL = SQL & "El fichero vendr� con cada campo separados por PIPES." & vbCrLf
    SQL = SQL & "Codigo cta contable |" & vbCrLf
    SQL = SQL & "Descripcion |" & vbCrLf
    SQL = SQL & "Direccion |" & vbCrLf
    SQL = SQL & "Cod. Postal |" & vbCrLf
    SQL = SQL & "Poblacion |" & vbCrLf
    SQL = SQL & "Provincia |" & vbCrLf
    SQL = SQL & "NIF|" & vbCrLf
    SQL = SQL & "Cta bancaria:   ENTIDAD|" & vbCrLf
    SQL = SQL & "Cta bancaria:   OFICINA|" & vbCrLf
    SQL = SQL & "Cta bancaria:   CC|" & vbCrLf
    SQL = SQL & "Cta bancaria:   CUENTA|" & vbCrLf
    SQL = SQL & "347:    0.- No    1.- Si|" & vbCrLf
    'Enero 2009
    SQL = SQL & "Forma pago|" & vbCrLf
    SQL = SQL & "Cta banco tesoreria|" & vbCrLf
    ' forpa y
    MsgBox SQL, vbInformation
End Sub

Private Sub imgCheck_Click(Index As Integer)
    For NE = 1 To TreeView1.Nodes.Count
        TreeView1.Nodes(NE).Checked = Index = 1
    Next
    
    Select Case Index
        ' ICONOS VISIBLES EN EL LISTVIEW DEL FRMPPAL
        Case 2 ' marcar todos
            For i = 1 To ListView6.ListItems.Count
                ListView6.ListItems(i).Checked = True
            Next i
        Case 3 ' desmarcar todos
            For i = 1 To ListView6.ListItems.Count
                ListView6.ListItems(i).Checked = False
            Next i
    End Select
        
    
End Sub

Private Sub ListView9_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    If Opcion = 51 Then
        Orden = Not Orden
        
        Select Case ColumnHeader
            Case "Serie"
                CampoOrden = "numserie"
            Case "Factura"
                CampoOrden = "numfactu"
            Case "Fecha"
                CampoOrden = "fecfactu"
            Case "Vto"
                CampoOrden = "numorden"
            Case "Fecha Vto"
                CampoOrden = "fecvenci"
            Case "Importe"
                CampoOrden = "importe"
        End Select
        
        CargarFacturasRemesas
    Else
        Orden = Not Orden
        
        Select Case ColumnHeader
            Case "Serie"
                CampoOrden = "numserie"
            Case "Factura"
                CampoOrden = "numfactu"
            Case "Fecha"
                CampoOrden = "fecfactu"
            Case "Vto"
                CampoOrden = "numorden"
            Case "Fecha Vto"
                CampoOrden = "fecvenci"
            Case "Importe"
                CampoOrden = "importe"
        End Select
        
        CargarFacturasReclamaciones
    
    End If
    
End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub tCuadre_Timer()
    tCuadre.Enabled = False
    Screen.MousePointer = vbHourglass
    Command2_Click
    Me.ListView1.Refresh
    Screen.MousePointer = vbHourglass
    espera 2
    Unload Me
End Sub


Private Sub CargaImagen()
On Error Resume Next
    Image2.Picture = LoadPicture(App.Path & "\minilogo.bmp")
    'Image1.Picture = LoadPicture(App.path & "\fondon.gif")
    Err.Clear
End Sub


Private Sub PonerCamposBalance()
    If Opcion = 7 Then
        Label16.Caption = "NUEVO"
        Label16.ForeColor = &H800000
        Me.chkPintar.Value = 1
        For i = 0 To 3
            Text1(i).Text = ""
        Next i

    Else
        'NumBalan|Pasivo|codigo|padre|Orden|tipo|deslinea|texlinea|formula|TienenCtas|Negrita|LibroCD|
        Text1(0).Text = RecuperaValor(Parametros, 7)
        Text1(1).Text = RecuperaValor(Parametros, 8)
        i = Val(RecuperaValor(Parametros, 10))
        If i = 1 Then
            'Tiene cuentas
            Text1(2).Text = ""
            Text1(2).Enabled = False
        Else
            Text1(2).Text = RecuperaValor(Parametros, 9)
        End If
        i = Val(RecuperaValor(Parametros, 11))
        chkNegrita.Value = i
        i = Val(RecuperaValor(Parametros, 12))
        chkCero.Value = i
        i = Val(RecuperaValor(Parametros, 13))
        chkPintar.Value = i
        Text1(3).Text = RecuperaValor(Parametros, 14)
    End If
End Sub



Private Sub PonerCamposCtaBalance()
    'EL grupo se le pasa siempre
    Text2.Text = RecuperaValor(Parametros, 1)
    
    
    If Opcion = 9 Then
        Label19.Caption = "NUEVO"
        Label19.ForeColor = &H800000
        Text3.Text = ""
        Text3.Enabled = True
        chkResta.Value = 0
    Else
        Text3.Enabled = False
        Text3.Text = RecuperaValor(Parametros, 2)
        i = Val(RecuperaValor(Parametros, 3))
        Option1(i).Value = True
        i = Val(RecuperaValor(Parametros, 4))
        chkResta.Value = i
    End If
End Sub




Private Function InsertarModificar() As Boolean
Dim Aux As String

On Error GoTo EInse

    InsertarModificar = False
    
    'Comprobamos el concpeto del libro a CD
     Text1(3).Text = UCase(Trim(Text1(3).Text))
    If Text1(3).Text <> "" Then
        If Not IsNumeric(Text1(3).Text) Then
            MsgBox "El campo 'Concepto Libro CD' debe ser num�rico", vbExclamation
            Exit Function
        End If
    End If
    
    'Hay k comprobar, si tiene formula k sea correcta
    Text1(2).Text = UCase(Trim(Text1(2).Text))
    If Text1(2).Text <> "" Then
        SQL = CompruebaFormulaConfigBalan(CInt(RecuperaValor(Parametros, 1)), Text1(2).Text)
        If SQL <> "" Then
            MsgBox SQL, vbExclamation
            Exit Function
        End If
    End If
    If Opcion = 7 Then
        SQL = "INSERT INTO balances_texto (NumBalan, Pasivo, codigo, padre, "
        SQL = SQL & "Orden, tipo, deslinea, texlinea, formula, TienenCtas, Negrita,A_Cero,Pintar,LibroCD) VALUES ("
        SQL = SQL & RecuperaValor(Parametros, 1)  'Numero
        SQL = SQL & ",'" & RecuperaValor(Parametros, 2) 'pasivo
        SQL = SQL & "'," & RecuperaValor(Parametros, 3)  'Codigo
        Aux = RecuperaValor(Parametros, 4) 'padre
        If Aux = "" Then
            Aux = ",NULL,"
        Else
            Aux = ",'" & Aux & "',"
        End If
        SQL = SQL & Aux
        SQL = SQL & RecuperaValor(Parametros, 5)
        If Text1(2).Text = "" Then
            Aux = "0"
        Else
            Aux = "1"
        End If
        SQL = SQL & "," & Aux
        SQL = SQL & ",'" & Text1(0).Text 'Text linea
        SQL = SQL & "','" & Text1(1).Text 'Desc linea
        SQL = SQL & "','" & Text1(2).Text 'Formula
        SQL = SQL & "',0," & chkNegrita.Value
        SQL = SQL & "," & Me.chkCero.Value
        SQL = SQL & "," & Me.chkPintar.Value
        SQL = SQL & ",'" & Text1(3).Text 'Libro CD
        SQL = SQL & "')"
    Else
        'Modificar
        'NumBalan|Pasivo|codigo|padre|Orden|tipo|deslinea|texlinea|formula|TienenCtas|Negrita|
        SQL = "UPDATE balances_texto SET "
        SQL = SQL & "deslinea='" & Text1(0).Text & "',"
        SQL = SQL & "texlinea='" & Text1(1).Text & "',"
        SQL = SQL & "formula='" & Text1(2).Text & "',"
        If Text1(2).Text = "" Then
            Aux = "0"
        Else
            Aux = "1"
        End If
        SQL = SQL & "Tipo =" & Aux & ","
        SQL = SQL & "Negrita = " & chkNegrita.Value
        SQL = SQL & ", A_Cero = " & Me.chkCero.Value
        SQL = SQL & ", Pintar = " & Me.chkPintar.Value
        SQL = SQL & ", LibroCD = '" & Text1(3).Text & "'"
        SQL = SQL & " WHERE numbalan =" & RecuperaValor(Parametros, 1)
        SQL = SQL & " AND Pasivo = '" & RecuperaValor(Parametros, 2)
        SQL = SQL & "' AND codigo = " & RecuperaValor(Parametros, 3)
        
    End If
    Conn.Execute SQL
    InsertarModificar = True
    'Ha insertado
    'Devuelve el texto, el texto auxiliar, y si es formula o no, descripcion cta y concepto oficial
    CadenaDesdeOtroForm = Text1(0).Text & "|" & Text1(1).Text & "|" & Aux & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text1(3).Text & "|"
    Exit Function
EInse:
    MuestraError Err.Number
End Function





Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

'--------------------------------------------------------------------
'
'       Carta IVA Clientes
'
Private Sub CargarDatosCarta()
Dim Limpiar As Boolean

    On Error GoTo ECargarDatos
    Limpiar = True
    SQL = App.Path & "\txt347.dat"
    If Dir(SQL) <> "" Then
        'Vamos a ir leyendo , y devoviendo cadena
        i = FreeFile
        Open SQL For Input As #i
        For NumRegElim = 0 To Text4.Count - 1
            'Obtenemos la cadena
           LeerCadenaFicheroTexto    'lo guarda en SQL
           Text4(NumRegElim) = SQL
        Next NumRegElim
        Close #i
        Limpiar = False
    End If
ECargarDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Carga fichero. " & Err.Description
    If Limpiar Then
        'No existe el fihero de configuracion
        For i = 0 To Text4.Count - 1
            Text4(i).Text = ""
        Next i
    End If
End Sub


Private Sub LeerCadenaFicheroTexto()
On Error GoTo ELeerCadenaFicheroTexto
    'Son dos lineas. La primaera indica k campo y la segunda el valor
    Line Input #i, SQL
    Line Input #i, SQL
    Exit Sub
ELeerCadenaFicheroTexto:
    SQL = ""
    Err.Clear
End Sub


Private Function GuardarDatosCarta()
    On Error GoTo Eguardardatoscarta
    SQL = App.Path & "\txt347.dat"
    i = FreeFile
    Open SQL For Output As #i
    For NumRegElim = 0 To Text4.Count - 1
        Print #i, Text4(NumRegElim).Tag
        Print #i, Text4(NumRegElim).Text
    Next NumRegElim
    Close #i
    Exit Function
Eguardardatoscarta:
    MuestraError Err.Number, "guardar datos carta"
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub





'------------- IMportar datos fiscales


Private Sub HacerImportacion()
Dim NF As Integer
Dim Linea As String

    On Error GoTo EHacere
    


    'Abrimos el fichero
    NF = FreeFile
    Open txtImpCta.Text For Input As #NF
    
  
    
    'Vamos linea a linea
    While Not EOF(NF)
        Line Input #NF, Linea
        Linea = Trim(Linea)
        If Linea <> "" Then
            If ProcesarLinea(Linea) Then Ok = Ok + 1
        End If
        lblImpCta.Caption = CStr(NE + Ok)
        lblImpCta.Refresh
    Wend
    
    'Cerramos
    Close (NF)
    
    Exit Sub
EHacere:
    MsgBox Err.Description, vbExclamation
End Sub





Private Function ProcesarLinea(Linea As String) As Boolean
'Dim Valores(6) As String
Dim Valores(13) As String    'ENero 2009. Dos campos msa. Total camppos=14. Vector(13)
Dim i As Integer
Dim Cad As String
Dim Crear As Boolean

    On Error GoTo EProcesarLinea
    ProcesarLinea = False
    
    
    'Orden en el k llegan
    For i = 0 To 13
        Valores(i) = RecuperaValor(Linea, i + 1)
    Next i
    
    'TRIM
    For i = 0 To 13
        Valores(i) = Trim(Valores(i))
    Next i
    
    'Comprobaciones
    '-----------------
    If Valores(0) = "" Or Valores(1) = "" Or Valores(6) = "" Then
        'Ni cta, ni nombre cta, ni NIF pueden ser nulos
        AnyadeErrores "Valores nulos ", Linea
        Exit Function
    End If
    
    'Cuenta NO puede ser numerica
    If Not IsNumeric(Valores(0)) Then
        AnyadeErrores "Cuenta: " & Valores(0), "No Numerica"
        Exit Function
    End If
    
    
    
    For i = 7 To 10
        If Valores(i) <> "" Then
            If Not IsNumeric(Valores(i)) Then
                AnyadeErrores "Cuenta bancaria", "CCC(" & i & "):   " & Valores(i)
                Exit Function
            End If
        End If
    Next i
    
    
    'Enero 2009
    'Si pone cta banco por defecto, comprobaremos que la lingitud es la correcta
    If Valores(13) <> "" Then
        If Len(Valores(13)) <> vEmpresa.DigitosUltimoNivel Then
            AnyadeErrores "Longitud cta banco tesoreria distinto ultimo nivel", Valores(13)
            Exit Function
        End If
    End If
    'Vemos si existe
    'Vemos si existe
    Crear = False
    If Not ExisteCuenta(Valores(0)) Then
        If Me.chkCrear.Value = 0 Then
            AnyadeErrores "Cuenta: " & Valores(0), "No existe"
            Exit Function
        Else
            Crear = True
        End If
    End If
    
    'Controlamos valores de Multibase para los textos, y las ' para la insercion
    For i = 1 To 5 'Sin NIF ni codmacta, 6 y 0 respectivamente
        If i <> 3 Then
            Cad = RevisaCaracterMultibase(Valores(i))
            NombreSQL Cad
            Valores(i) = Cad
        End If
    Next i
    
    
    
    
    '
    
    
     If Crear Then
        i = DigitosNivel(vEmpresa.numnivel - 1)
        Cad = Mid(Valores(0), 1, i)
        If Cad <> CadenaDesdeOtroForm Then
            If Not CreaSubcuentas(Valores(0), i, "IMPORTACION AUTOMATICA") Then
                AnyadeErrores "Cuenta: " & Valores(0), "GENERANDO SUBNIVELES"
                Exit Function
            End If
            CadenaDesdeOtroForm = Cad
        End If
    End If
    
    'Montamos el SQL
        'Montamos el SQL
    If Crear Then
        Cad = "INSERT INTO Cuentas (codmacta,nommacta,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos,"
        'NUEVO
        Cad = Cad & "entidad,oficina,CC,cuentaba,"
        Cad = Cad & "model347,apudirec,forpa ,ctabanco"
        Cad = Cad & ") VALUES ("
        Cad = Cad & "'" & Valores(0) & "',"
        Cad = Cad & "'" & Valores(1) & "',"
        Cad = Cad & "'" & Valores(1) & "',"
        Cad = Cad & "'" & Valores(2) & "',"
        Cad = Cad & "'" & Valores(3) & "',"
        Cad = Cad & "'" & Valores(4) & "',"
        Cad = Cad & "'" & Valores(5) & "',"
        Cad = Cad & "'" & Valores(6) & "',"
        For i = 7 To 10
            If Valores(i) = "" Then
                Cad = Cad & "NULL,"
            Else
                Cad = Cad & "'" & Valores(i) & "',"
            End If
        Next i
        If Valores(11) = "1" Then
            Cad = Cad & "1"
        Else
            Cad = Cad & "0"
        End If
        
        Cad = Cad & ",'S'"
        'Enerom2009
        'forpa ,ctabanco
        For i = 12 To 13
            If Valores(i) = "" Then
                Cad = Cad & ",NULL"
            Else
                Cad = Cad & ",'" & Valores(i) & "'"
            End If
        Next
        
        'Final
        Cad = Cad & ")"
    
    Else
        Cad = "UPDATE Cuentas SET "
        Cad = Cad & " nommacta = '" & Valores(1) & "',"
        Cad = Cad & " razosoci = '" & Valores(1) & "',"
        Cad = Cad & " dirdatos = '" & Valores(2) & "',"
        Cad = Cad & " codposta = '" & Valores(3) & "',"
        Cad = Cad & " despobla = '" & Valores(4) & "',"
        Cad = Cad & " desprovi = '" & Valores(5) & "',"
        Cad = Cad & " nifdatos = '" & Valores(6) & "',"
        'model347
        Cad = Cad & " model347 = "
        If Valores(11) = "1" Then
            Cad = Cad & "1"
        Else
            Cad = Cad & "0"
        End If
        
        'CCC
        Cad = Cad & ", entidad =" & ValorSQL(Valores(7))
        Cad = Cad & ", oficina =" & ValorSQL(Valores(8))
        Cad = Cad & ", CC =" & ValorSQL(Valores(9))
        Cad = Cad & ", cuentaba =" & ValorSQL(Valores(10))
            
        'Enero 2009
        Cad = Cad & ", forpa  =" & ValorSQL(Valores(12))
        Cad = Cad & ", ctabanco =" & ValorSQL(Valores(13))
            
        Cad = Cad & " WHERE codmacta ='" & Valores(0) & "'"
    End If
   
    If Not EjecutaSQL2(Cad) Then Exit Function
    ProcesarLinea = True
    Exit Function
EProcesarLinea:
    AnyadeErrores "Linea: " & Linea, Err.Description
    Err.Clear
    
End Function

Private Function ValorSQL(ByRef C As String) As String
    If C = "" Then
        ValorSQL = "NULL"
    Else
        ValorSQL = "'" & C & "'"
    End If
End Function
Private Function EjecutaSQL2(SQL As String) As Boolean
    EjecutaSQL2 = False
    On Error Resume Next
    Conn.Execute SQL
    If Err.Number <> 0 Then
        AnyadeErrores "SQL: " & SQL, Err.Description
        Err.Clear
    Else
        EjecutaSQL2 = True
    End If
End Function


Private Sub AnyadeErrores(L1 As String, L2 As String)
    NE = NE + 1
    Errores = Errores & "-----------------------------" & vbCrLf
    Errores = Errores & L1 & vbCrLf
    Errores = Errores & L2 & vbCrLf


End Sub

Private Sub ImprimeFichero()
Dim NF As Integer
    On Error GoTo EImprimeFichero
    NF = FreeFile
    Open App.Path & "\errimpdat.txt" For Output As #NF
    Print #NF, Errores
    Close (NF)
    Shell "notepad.exe " & App.Path & "\errimpdat.txt", vbMaximizedFocus
    Exit Sub
EImprimeFichero:
    MsgBox Err.Description & vbCrLf, vbCritical
    Err.Clear
End Sub


Private Function ExisteCuenta(Cta As String) As Boolean

    
    ExisteCuenta = False
    SQL = DevuelveDesdeBD("codmacta", "cuentas", "codmacta", Cta, "T")
    If SQL <> "" Then ExisteCuenta = True
    
End Function

Private Sub Text7_GotFocus(Index As Integer)
    Text7(Index).SelStart = 0
    Text7(Index).SelLength = Len(Text7(Index).Text)
End Sub

Private Sub Text7_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim N As Node
    'Si es padre
    If Node.Parent Is Nothing Then
        If Node.Children > 0 Then
            Set N = Node.Child
            Do
                N.Checked = Node.Checked
                Set N = N.Next
            Loop Until N Is Nothing
        End If
    End If
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    txtFecha(Index).SelStart = 0
    txtFecha(Index).SelLength = Len(txtFecha(Index).Text)
End Sub

Private Sub txtfecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtfecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index))
    If txtFecha(Index) = "" Then Exit Sub
    If Not EsFechaOK(txtFecha(Index)) Then
        MsgBox "Fecha incorrecta: " & txtFecha(Index), vbExclamation
        txtFecha(Index).Text = ""
        txtFecha(Index).SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------------
'
'
Private Function ExportarDatosFacturas(Proveedores As Boolean) As Boolean
Dim vOpc As String
Dim Aux As String

    'Comprobamos el RS
    ExportarDatosFacturas = False
    If Proveedores Then
        SQL = "prov"
        Parametros = "fecrecpr"
    Else
        Parametros = "fecfaccl"
        SQL = ""
    End If
    SQL = SQL & " where " & Parametros & " >= '" & Format(CDate(txtFecha(2).Text), FormatoFecha) & "'"
    SQL = SQL & " AND " & Parametros & " <= '" & Format(CDate(txtFecha(3).Text), FormatoFecha) & "'"
    Set RS = New ADODB.Recordset
    Errores = "select count(*) from cabfact" & SQL
    RS.Open Errores, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Ok = 0
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If RS.Fields(0) > 0 Then Ok = 1
        End If
    End If
    RS.Close
    
    If Ok = 0 Then
        SQL = "Ningun dato a traspasar de facturas "
        If Proveedores Then
            SQL = SQL & "proveedores"
        Else
            SQL = SQL & "clientes"
        End If
        MsgBox SQL, vbExclamation
        Exit Function
    End If
    

    

    
    '----------------------------------------------------------------------
    'OPCION
    vOpc = "OPCION"
    EncabezadoPieFact False, vOpc, 0
    If Proveedores Then
        Print #NE, 0
        Print #NE, "Proveedores"
    Else
        Print #NE, 1
        Print #NE, "Clientes"
    End If
    
    'Ultimo nivel de las cuentas contables
    Print #NE, vEmpresa.DigitosUltimoNivel
    EncabezadoPieFact True, vOpc, 1
    
    '----------------------------------------------------------------------
    'CUENTAS
    vOpc = "CUENTAS"
    Label40.Caption = "Cuentas"
    Label40.Refresh
    
    EncabezadoPieFact False, vOpc, 0
    Parametros = "DELETE FROM tmpcierre1 where codusu =" & vUsu.Codigo
    Conn.Execute Parametros
    
    
    'Cuentas que necesito
    
    Parametros = "Select distinct(codmacta) from cabfact" & SQL
    RS.Open Parametros, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        InsertaEnTmpCta
        RS.MoveNext
        
    Wend
    RS.Close
    
    Parametros = "Select distinct(Cuereten) from cabfact" & SQL & " and not (cuereten is null)"
    RS.Open Parametros, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        InsertaEnTmpCta
        RS.MoveNext
    Wend
    RS.Close
    
    
    'la cuentas de las lineas de factura
    Parametros = "Select codtbase from linfact"
    If Proveedores Then
        Parametros = Parametros & "prov where anofacpr >=" & Year(CDate(txtFecha(2).Text))
        Parametros = Parametros & " and anofacpr <=" & Year(CDate(txtFecha(3).Text))
    Else
        Parametros = Parametros & " where anofaccl >=" & Year(CDate(txtFecha(2).Text))
        Parametros = Parametros & " and anofaccl <=" & Year(CDate(txtFecha(3).Text))
    End If
    Parametros = Parametros & " GROUP BY codtbase"
    RS.Open Parametros, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        InsertaEnTmpCta
        RS.MoveNext
    Wend
    RS.Close
    
    
    'Ahora cojo todos los datos de tmpcierr1 y creo los inserts de las cuentas
    Parametros = "Select cuentas.* from cuentas,tmpcierre1 where cuentas.codmacta=tmpcierre1.cta "
    Parametros = Parametros & " and codusu =" & vUsu.Codigo
    RS.Open Parametros, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    BACKUP_TablaIzquierda RS, CadenaDesdeOtroForm
    Ok = 0
    While Not RS.EOF
        Label40.Caption = RS!codmacta
        Label40.Refresh
        Ok = Ok + 1
        BACKUP_Tabla RS, Parametros
        Parametros = "INSERT INTO Cuentas " & CadenaDesdeOtroForm & " VALUES " & Parametros & ";"
        Print #NE, Parametros
        RS.MoveNext
    Wend
    RS.Close
    
    EncabezadoPieFact True, vOpc, Ok


    '----------------------------------------------------------------------
    'OPCION
    vOpc = "CC"
    Label40.Caption = "C.C."
    Label40.Refresh
    EncabezadoPieFact False, vOpc, 0
    'Volvemos a utlizar la misma tabla
    Parametros = "DELETE FROM tmpcierre1 where codusu =" & vUsu.Codigo
    Conn.Execute Parametros

    Parametros = "Select codccost from linfact"
    If Proveedores Then
        Parametros = Parametros & "prov where anofacpr >=" & Year(CDate(txtFecha(2).Text))
        Parametros = Parametros & " and anofacpr <=" & Year(CDate(txtFecha(3).Text))
    Else
        Parametros = Parametros & " where anofaccl >=" & Year(CDate(txtFecha(2).Text))
        Parametros = Parametros & " and anofaccl <=" & Year(CDate(txtFecha(3).Text))
    End If
    Parametros = Parametros & " AND not (codccost is null)"
    Parametros = Parametros & " GROUP BY codccost"
    RS.Open Parametros, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Ok = 0
    While Not RS.EOF
        Ok = Ok + 1
        InsertaEnTmpCta
        RS.MoveNext
        
    Wend
    RS.Close
    
    
    
    
    If Ok > 0 Then
        'Ahora cojo todos los datos de tmpcierr1 y creo los inserts de las cuentas
        Parametros = "Select ccoste.* from ccoste,tmpcierre1 where ccoste.codccost=tmpcierre1.cta "
        Parametros = Parametros & " and codusu =" & vUsu.Codigo
        RS.Open Parametros, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        BACKUP_TablaIzquierda RS, CadenaDesdeOtroForm

        'Si k hay CC
        
        While Not RS.EOF
            Ok = Ok + 1
            BACKUP_Tabla RS, Parametros
            Parametros = "INSERT INTO ccoste " & CadenaDesdeOtroForm & " VALUES " & Parametros & ";"
            Print #NE, Parametros
            RS.MoveNext
        Wend
        RS.Close
            
    End If
    
    
    EncabezadoPieFact True, vOpc, 0
    
    
    
    
    '----------------------------------------------------------------------
    'OPCION
    vOpc = "IVA"
    EncabezadoPieFact False, vOpc, 0
    Parametros = "Select * from tiposiva"
    RS.Open Parametros, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    BACKUP_TablaIzquierda RS, CadenaDesdeOtroForm


    While Not RS.EOF
        Ok = Ok + 1
        BACKUP_Tabla RS, Parametros
        Parametros = "INSERT INTO tiposiva " & CadenaDesdeOtroForm & " VALUES " & Parametros & ";"
        Print #NE, Parametros
        RS.MoveNext
    Wend
    RS.Close

    
    EncabezadoPieFact True, vOpc, 1
    
    
    '------------------------------------
    'Para las facturas de clientes necesitare tb las series de factura
    If Not Proveedores Then
        vOpc = "CONTADORES"
        EncabezadoPieFact False, vOpc, 0
        Parametros = "Select * from contadores"
        RS.Open Parametros, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        BACKUP_TablaIzquierda RS, CadenaDesdeOtroForm


        While Not RS.EOF
            Ok = Ok + 1
            BACKUP_Tabla RS, Parametros
            Parametros = "INSERT INTO contadores " & CadenaDesdeOtroForm & " VALUES " & Parametros & ";"
            Print #NE, Parametros
            RS.MoveNext
        Wend
        RS.Close
    
        
        EncabezadoPieFact True, vOpc, 1
    End If
    
    
    
    '----------------------------------------------------------------------
    'FACTURAS
    'Grabaremos en cada linea
    '
    '  codigo |INSERT |UPDATE |base1|base2....
    '   Codigo: Para clientes ser�: numserie, codfacl, anofaccl
    
    vOpc = "FACTURAS"
    EncabezadoPieFact False, vOpc, 0
    
    
    If Not Proveedores Then
        SQL = "numserie,codfaccl,anofaccl,fecfaccl,codmacta,confaccl,ba1faccl,ba2faccl,ba3faccl,pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,fecliqcl"
        Parametros = "Select " & SQL & " from cabfact"
        Parametros = Parametros & " where fecfaccl >= '" & Format(CDate(txtFecha(2).Text), FormatoFecha) & "'"
        Parametros = Parametros & " and fecfaccl <= '" & Format(CDate(txtFecha(3).Text), FormatoFecha) & "'"
        Ok = 3
    Else
        SQL = "numregis,anofacpr,fecfacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,fecliqpr,nodeducible"
        Parametros = "Select " & SQL & " from cabfactprov"
        Parametros = Parametros & " where fecrecpr >= '" & Format(CDate(txtFecha(2).Text), FormatoFecha) & "'"
        Parametros = Parametros & " and fecrecpr <= '" & Format(CDate(txtFecha(3).Text), FormatoFecha) & "'"
        Ok = 2
    End If
    
    RS.Open Parametros, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    BACKUP_TablaIzquierda RS, CadenaDesdeOtroForm
    If Not Proveedores Then
        CadenaDesdeOtroForm = "INSERT INTO cabfact " & CadenaDesdeOtroForm & " VALUES "
    Else
        CadenaDesdeOtroForm = "INSERT INTO cabfactprov " & CadenaDesdeOtroForm & " VALUES "
    End If
        
        
    Set miRsAux = New ADODB.Recordset
    Errores = ""
    NumRegElim = 0
    While Not RS.EOF
        
        NumRegElim = NumRegElim + 1
        SQL = RS.Fields(0) & "|"
        If Not Proveedores Then SQL = SQL & "0" 'meto un 0 para que las facturas que coinciden con el a�o no den errores
        SQL = SQL & RS.Fields(1) & "|"
        If Not Proveedores Then SQL = SQL & RS.Fields(2) & "|"
        Label40.Caption = SQL
        Label40.Refresh
        'Cadena insert
        BACKUP_Tabla RS, Parametros
        Parametros = SQL & CadenaDesdeOtroForm & Parametros & ";|"
        
        
        'El UPDATE
        SQL = ""
        
        For i = Ok To RS.Fields.Count - 1
            If SQL <> "" Then SQL = SQL & ","
            SQL = SQL & RS.Fields(i).Name & " = "
            If IsNull(RS.Fields(i)) Then
                SQL = SQL & "NULL"
            Else
                Select Case RS.Fields(i).Type
                Case 133
                    SQL = SQL & "'" & Format(RS.Fields(i), FormatoFecha) & "'"
                
                Case 17
                    'numero
                    SQL = SQL & RS.Fields(i)
                    
                Case 131
                    SQL = SQL & TransformaComasPuntos(CStr(RS.Fields(i)))
                Case Else
                    SQL = SQL & "'" & DevNombreSQL(RS.Fields(i)) & "'"
                End Select
                
            End If
        Next i
      
        SQL = SQL & " WHERE "
        Aux = ""
        For i = 0 To Ok - 1
            Aux = Aux & RS.Fields(i).Name & " = '" & RS.Fields(i) & "' and "
        Next
        Aux = Mid(Aux, 1, Len(Aux) - 4)
        SQL = SQL & Aux
        If Not Proveedores Then
            SQL = "UPDATE cabfact SET " & SQL
        Else
            SQL = "UPDATE cabfactprov SET " & SQL
        End If
        Parametros = Parametros & SQL & "|"
        
        
        'Metemos una marca para separar las lineas
        Parametros = Parametros & "<>"
        
        'Las lineas
        '----------------------------
        
        SQL = "Select * from linfact"
        If Proveedores Then SQL = SQL & "prov"
        SQL = SQL & " WHERE " & Aux
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Errores = "" Then
            BACKUP_TablaIzquierda miRsAux, Errores
            SQL = "INSERT INTO linfact"
            If Proveedores Then SQL = SQL & "prov"
            Errores = SQL & "  " & Errores & " VALUES "
        End If
        While Not miRsAux.EOF
            BACKUP_Tabla miRsAux, SQL
            SQL = Errores & SQL & ";"
            Parametros = Parametros & SQL & "|"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Print #NE, Parametros
        
        RS.MoveNext
    Wend
    RS.Close
    Set miRsAux = Nothing
    
    EncabezadoPieFact True, vOpc, CInt(NumRegElim)
    
    
    
    
    'Y dejo limpio el tajo
    Parametros = "DELETE FROM tmpcierre1 where codusu =" & vUsu.Codigo
    Conn.Execute Parametros
    Label40.Caption = ""
    CadenaDesdeOtroForm = "2"
    Set RS = Nothing
    
    ExportarDatosFacturas = True
End Function


Private Sub CopiarArchivo()
On Error GoTo ECopiarArchivo

    If Dir(Text8.Text, vbArchive) <> "" Then Kill Text8.Text
    FileCopy Errores, Text8.Text
    
    Errores = "El fichero: " & Text8.Text & " se ha generado con �xito"
    MsgBox Errores, vbInformation
    Exit Sub
ECopiarArchivo:
    MuestraError Err.Number, "Copiar archivo"
End Sub



Private Sub EncabezadoPieFact(Pie As Boolean, ByVal Text As String, REG As Integer)
    If Pie Then
        Text = "[/" & Text & "]" & REG
    Else
        Text = "[" & Text & "]"
    End If
    Print #NE, Text
End Sub


Private Sub InsertaEnTmpCta()
On Error Resume Next
    
    Conn.Execute "INSERT INTO tmpcierre1 (codusu, cta) VALUES (" & vUsu.Codigo & ",'" & RS.Fields(0) & "')"
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub ImportarFicheroFac()
    NE = FreeFile
    Screen.MousePointer = vbHourglass
    Open Text8.Text For Input As #NE
    'Importamos el primer trozo. PROVEEDORES
    If ImportarDatosFacturas Then
        'CLIENTES
        If ImportarDatosFacturas Then
            Close #NE
            MsgBox "Proceso finalizado", vbExclamation
            If chkImportarFacturas.Value Then
                If Dir(Text8.Text, vbArchive) <> "" Then Kill Text8.Text
            End If
            cmdImportarFacuras(0).Enabled = False
        End If
    End If
    Label40.Caption = ""
    Screen.MousePointer = vbDefault
End Sub

'---------------------------------------------------------------------------------------
Private Function ImportarDatosFacturas() As Boolean
Dim Fin As Boolean
Dim Clientes As Boolean

    On Error GoTo EIM
    
    CadenaDesdeOtroForm = "Abriendo fichero. Datos basicos"
    Label40.Caption = CadenaDesdeOtroForm
    Label40.Refresh
    
    
    Line Input #NE, SQL   'OPCION
    If SQL <> "[OPCION]" Then
        MsgBox "Formato fichero incorrecto", vbExclamation
        Close NE
        Exit Function
    End If
    Line Input #NE, SQL   ' PRoveedores o clientes
    i = Val(SQL)
    If i = 1 Then
        'CLIENTES
        Clientes = True
    Else
        'PROVEEDORES
        Clientes = False
    End If
    Line Input #NE, SQL   ' Datos vacios
    Line Input #NE, SQL   ' digitos ultimo nivel
    i = Val(SQL)
    If i <> vEmpresa.DigitosUltimoNivel Then
        MsgBox "Ultimo nivel disitinto:" & i, vbExclamation
        Close NE
        Exit Function
    End If
    Line Input #NE, SQL   'FIN OPCION
    
    'CUENTAS
    CadenaDesdeOtroForm = "Cuentas"
    Label40.Caption = CadenaDesdeOtroForm
    Label40.Refresh
    Line Input #NE, SQL   'CUENTAS
    i = 0
    Fin = False
    Do
        Line Input #NE, SQL   'FIN OPCION
        If InStr(1, SQL, "[/CUENTAS]") > 0 Then
            'Fin
            Fin = True
            'Ver numero registros
            Ok = InStr(1, SQL, "]")
            SQL = Mid(SQL, Ok + 1)
            Ok = Val(SQL)
            If i <> Ok Then
            
            End If
        Else
            'Mandamos la linea a ejecutar
            Label40.Caption = "Cta: " & Mid(SQL, 155 + vEmpresa.DigitosUltimoNivel, 30)
            Label40.Refresh
            EjecutarSQL
            i = i + 1
            
        End If
    Loop Until Fin
    
    'CC
    CadenaDesdeOtroForm = "CC"
    Label40.Caption = CadenaDesdeOtroForm
    Label40.Refresh
    Line Input #NE, SQL   'CC
    i = 0
    Fin = False
    Do
        Line Input #NE, SQL   'FIN OPCION
        If InStr(1, SQL, "[/CC]") > 0 Then
            'Fin
            Fin = True
            'Ver numero registros
        Else
            'Mandamos la linea a ejecutar
            EjecutarSQL
            i = i + 1
            
        End If
    Loop Until Fin
    
    
    
    'IVA
    CadenaDesdeOtroForm = "IVA"
    Label40.Caption = CadenaDesdeOtroForm
    Label40.Refresh
    Line Input #NE, SQL   'IVA
    i = 0
    Fin = False
    Do
        Line Input #NE, SQL   'FIN OPCION
        If InStr(1, SQL, "[/IVA]") > 0 Then
            'Fin
            Fin = True
            'Ver numero registros
        Else
            'Mandamos la linea a ejecutar
            EjecutarSQL
            i = i + 1
            
        End If
    Loop Until Fin
    
    
    If Clientes Then
        'SOLO CLIENTES LLEVA CONTADORES
        CadenaDesdeOtroForm = "CONTADORES"
        Label40.Caption = CadenaDesdeOtroForm
        Label40.Refresh
        Line Input #NE, SQL   'CONTADPORES
        i = 0
        Fin = False
        Do
            Line Input #NE, SQL   'FIN OPCION
            If InStr(1, SQL, "[/CONTADORES]") > 0 Then
                'Fin
                Fin = True
                'Ver numero registros
            Else
                'Mandamos la linea a ejecutar
                EjecutarSQL
                i = i + 1
                
            End If
        Loop Until Fin
    End If
    
    
    
    
    'FACTURAS
    Set RS = New ADODB.Recordset
    CadenaDesdeOtroForm = "FACTURAS"
    Label40.Caption = CadenaDesdeOtroForm
    Label40.Refresh
    Line Input #NE, SQL   'FACTS
    i = 0
    Fin = False
    Do
        Line Input #NE, SQL   'FIN OPCION
        If InStr(1, SQL, "[/FACT") > 0 Then
            'Fin
            Fin = True
            'Ver numero registros
            Ok = InStr(1, SQL, "]")
            Ok = Val(Mid(SQL, Ok + 1))
            If Ok <> i Then MsgBox "Diferencia entre facturas procesadas. Fichero: " & Ok & " -> " & i, vbExclamation
        Else
            'Mandamos la linea a ejecutar
            ProcesarLineaFactura Clientes
            i = i + 1
        End If
    Loop Until Fin
    Label40.Caption = "Actualizando datos"
    Me.Refresh
    espera 1
    ImportarDatosFacturas = True
    Exit Function
    
EIM:
    
    If Err.Number <> 0 Then
        MuestraError Err.Number, CadenaDesdeOtroForm
        i = 1
    Else
        i = 0
    End If
    Close NE
    If i = 0 Then
        If Me.chkImportarFacturas.Value = 1 Then Kill Text8.Text
    End If
    Set RS = Nothing
End Function


Private Sub EjecutarSQL()
    On Error Resume Next
    
    Conn.Execute SQL
    If Err.Number <> 0 Then
        If Conn.Errors(0).Number = 1062 Then
            Err.Clear
        Else
            'MuestraError Err.Number, Err.Description
        End If
        Err.Clear
    End If
End Sub


Private Sub ProcesarLineaFactura(Clientes As Boolean)
Dim A�o As Integer
Dim numero As Long
Dim Serie As String
Dim J As Long
Dim Aux As String

    If Clientes Then
        Serie = RecuperaValor(SQL, 1)
        numero = RecuperaValor(SQL, 2)
        A�o = RecuperaValor(SQL, 3)
    Else
        Serie = ""
        numero = RecuperaValor(SQL, 1)
        A�o = RecuperaValor(SQL, 2)
    End If
    
        
    Label40.Caption = Serie & " " & numero & " / " & A�o
    Label40.Refresh
    DoEvents
    'Quitamos el La cadaena
    J = InStr(2, SQL, "|" & A�o & "|")
    If J = 0 Then
        MsgBox "Error en a�o factura", vbExclamation
        Exit Sub
    End If
    
    J = J + 6
    SQL = Mid(SQL, J)
    
    If Clientes Then
        Aux = "Select * from cabfact WHERE numserie = '" & Serie & "'"
        Aux = Aux & " and anofaccl = " & A�o & " and codfaccl =" & numero
    Else
        Aux = "Select * from cabfactprov WHERE "
        Aux = Aux & " anofacpr = " & A�o & " and numregis =" & numero
    End If
    RS.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    J = 1
    If Not RS.EOF Then
        J = 2
        'Borro las lineas
        If Clientes Then
            Aux = "DELETE from linfact WHERE numserie = '" & Serie & "'"
            Aux = Aux & " and anofaccl = " & A�o & " and codfaccl =" & numero
        Else
            Aux = "DELETE from linfactprov WHERE "
            Aux = Aux & " anofacpr = " & A�o & " and numregis =" & numero
        End If
        Conn.Execute Aux
    End If
    RS.Close
    
    Aux = RecuperaValor(SQL, CInt(J))
    Conn.Execute Aux
    
    '---------------------------------------------------
    J = InStr(1, SQL, "<>")
    If J = 0 Then
        MsgBox "Error lineas: " & Aux, vbExclamation
        Exit Sub
    End If
        
    SQL = Mid(SQL, J + 2)
    Do
        J = InStr(1, SQL, "|")
        If J > 0 Then
            Aux = Mid(SQL, 1, J - 1)
            SQL = Mid(SQL, J + 1)
            Conn.Execute Aux
        End If
    Loop Until J = 0
End Sub


Private Function ImportarDatosExternos347() As Boolean
On Error GoTo EImportarDatosExternos347
    ImportarDatosExternos347 = False
    'Abrimos el fichero
    NE = FreeFile
    Open Text9(0).Text For Input As #NE
    Line Input #NE, Errores
    Close #NE

    
    
    If Errores <> "" Then
        SQL = RecuperaValor(Errores, 1)
    Else
        SQL = ""
    End If
        
    If SQL = "" Then
        MsgBox "Error en fichero. Linea vacia o sin a�o importacion." & vbCrLf & SQL, vbExclamation
        Exit Function
    End If


    If Val(SQL) = 0 Then
        MsgBox "A�o incorrecto: " & SQL & vbCrLf & Errores, vbExclamation
        Exit Function
    End If


    SQL = "DELETE FROM datosext347 where a�o =" & SQL
    Conn.Execute SQL
    
    
    
    'Volvemos a abrir el fichero
    NE = FreeFile
    Open Text9(0).Text For Input As #NE
    i = 0
    
    
    While Not EOF(NE)
        Line Input #NE, Errores
            
        SQL = RecuperaValor(Errores, 1)
        Parametros = Trim(RecuperaValor(Errores, 2))
        If Parametros = "1" Then
            Ok = 1
        Else
            Ok = 2
        End If
        SQL = SQL & ",'" & Text9(Ok).Text & "'"
        For Ok = 3 To 8
            Parametros = RevisaCaracterMultibase(Trim(RecuperaValor(Errores, Ok)))
            Parametros = DevNombreSQL(Parametros)
            SQL = SQL & ",'" & Parametros & "'"
        Next Ok
        
        'El importe
        Parametros = TransformaComasPuntos((RecuperaValor(Errores, 9)))
        SQL = SQL & "," & Parametros & ")"
        SQL = "INSERT INTO datosext347 (a�o, letra, nif, nombre, direc, codposta, poblacion, provincia, importe) VALUES (" & SQL
        Conn.Execute SQL
        i = i + 1
    Wend
    Close #NE

    
    
    If i > 0 Then
        ImportarDatosExternos347 = True
        MsgBox "Proceso finalizado.   " & i & " registros insertados", vbInformation
    Else
        MsgBox "No se han importado datos", vbExclamation
    End If
    Exit Function
EImportarDatosExternos347:
    MuestraError Err.Number, SQL
    On Error Resume Next
        Close #NE
        Err.Clear
End Function


Private Sub cargarObservacionesCuenta()
    Set RS = New ADODB.Recordset
    SQL = "select codmacta,nommacta,obsdatos from cuentas where codmacta = '" & Parametros & "'"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        SQL = RS!codmacta & "|" & RS!Nommacta & "|" & DBMemo(RS!obsdatos) & "|"
    Else
        SQL = "Err|ERROR  LEYENDO DATOS CUENTAS | ****  ERROR ****|"
    End If
    RS.Close
    Set RS = Nothing
    For i = 1 To 3
        Text1(i + 3).Text = RecuperaValor(SQL, i)
    Next i
End Sub


Private Sub cargaempresasbloquedas()
Dim IT As ListItem
    On Error GoTo Ecargaempresasbloquedas
    Set RS = New ADODB.Recordset
    SQL = "select empresasariconta.codempre,nomempre,nomresum,usuarioempresasariconta.codempre bloqueada from usuarios.empresasariconta left join usuarios.usuarioempresasariconta on "
    SQL = SQL & " empresasariconta.codempre = usuarioempresasariconta.codempre And (usuarioempresasariconta.codusu = " & Parametros & " Or codusu Is Null)"
    '[Monica] solo ariconta
    SQL = SQL & " WHERE conta like 'ariconta%' "
    SQL = SQL & " ORDER BY empresasariconta.codempre"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        Errores = Format(RS!codempre, "00000")
        SQL = "C" & Errores
        
        If IsNull(RS!bloqueada) Then
            'Va al list de la derecha
            Set IT = ListView2(0).ListItems.Add(, SQL)
            IT.SmallIcon = 1
        Else
            Set IT = ListView2(1).ListItems.Add(, SQL)
            IT.SmallIcon = 2
        End If
        IT.Text = Errores
        IT.SubItems(1) = RS!nomempre
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    Errores = ""
    Exit Sub
Ecargaempresasbloquedas:
    MuestraError Err.Number, Err.Description
    Me.cmdBloqEmpre(0).Enabled = False
    Errores = ""
    Set RS = Nothing
End Sub










'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'Restore desde backup
'
'


Private Sub CargaIconosVisibles()
Dim IT As ListItem
Dim TotalArray  As Long
    On Error GoTo ECargaIconosVisibles
    
    Set RS = New ADODB.Recordset
    
    SQL = "select menus.codigo, menus.imagen, menus.descripcion, menus_usuarios.posx, menus_usuarios.posy, menus_usuarios.vericono "
    SQL = SQL & " from menus inner join menus_usuarios on menus.codigo = menus_usuarios.codigo and menus.aplicacion = menus_usuarios.aplicacion"
    SQL = SQL & " WHERE menus.aplicacion = 'ariconta' and menus_usuarios.codusu = " & vUsu.Id
    SQL = SQL & " and menus.imagen <> 0 " ' si tiene imagen puede estar en el listview para seleccionar
    SQL = SQL & " and menus_usuarios.ver = 1 "
    SQL = SQL & " ORDER BY menus.codigo "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView6.SmallIcons = frmPpal.ImageListPpal16
    
    
    ListView6.ColumnHeaders.Clear
    
    ListView6.ColumnHeaders.Add , , "C�digo", 1800.0631
    ListView6.ColumnHeaders.Add , , "Descripci�n", 4200.2522, 0
    ListView6.ColumnHeaders.Add , , "EraVisible", 0, 0
    ListView6.ColumnHeaders.Add , , "X", 0, 0
    ListView6.ColumnHeaders.Add , , "Y", 0, 0
    
    
    TotalArray = 0
    While Not RS.EOF
        Set IT = ListView6.ListItems.Add
        
        IT.SmallIcon = DBLet(RS!imagen, "N")
        
        IT.Text = Format(DBLet(RS!Codigo, "N"), "000000")
        IT.SubItems(1) = DBLet(RS!Descripcion, "T")
        If DBLet(RS!vericono, "N") <> 0 Then
            IT.Checked = True
            IT.SubItems(2) = 1
            IT.SubItems(3) = RS!PosX
            IT.SubItems(4) = RS!PosY
        Else
           
            IT.SubItems(2) = 0
            IT.SubItems(3) = 0
            IT.SubItems(4) = 0
            IT.Checked = False
        End If
        
        
        
        RS.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    RS.Close
    Exit Sub
    
ECargaIconosVisibles:
    MuestraError Err.Number, Err.Description
    Me.cmdBloqEmpre(0).Enabled = False
    Errores = ""
    Set RS = Nothing
End Sub


Private Sub CargaInformeBBDD()
Dim IT As ListItem
Dim TotalArray  As Long
    On Error GoTo ECargaInformeBBDD
    
    Set RS = New ADODB.Recordset
    
    SQL = "select * from tmpinfbbdd where codusu = " & vUsu.Codigo & " order by posicion "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    ListView3.ColumnHeaders.Clear
    ListView3.ColumnHeaders.Add , , "CONCEPTO", 3500.0631
    ListView3.ColumnHeaders.Add , , "count ACTUAL", 2250.2522, 1
    ListView3.ColumnHeaders.Add , , "porcen ACTUAL", 1000.2522, 1
    ListView3.ColumnHeaders.Add , , "count siguiente", 2250.2522, 1
    ListView3.ColumnHeaders.Add , , "porcen siguiente", 1000.2522, 1
    
    
    
    
    TotalArray = 0
    While Not RS.EOF
        Set IT = ListView3.ListItems.Add
        
        IT.Text = UCase(DBLet(RS!Concepto, "T"))
        
        If DBLet(RS!posicion, "N") > 2 Then
            IT.SubItems(1) = Format(DBLet(RS!nactual, "N"), "###,###,###,##0")
            IT.SubItems(2) = Format(DBLet(RS!Poractual, "N"), "##0.00") & "%"
            IT.SubItems(3) = Format(DBLet(RS!nsiguiente, "N"), "###,###,###,##0")
            IT.SubItems(4) = Format(DBLet(RS!Porsiguiente, "N"), "##0.00") & "%"
        Else
            IT.SubItems(1) = Format(DBLet(RS!nactual, "N"), "###,###,###,##0")
            IT.SubItems(3) = Format(DBLet(RS!nsiguiente, "N"), "###,###,###,##0")
        End If
        
        RS.MoveNext
    Wend
    
    RS.Close
    Exit Sub
    
ECargaInformeBBDD:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set RS = Nothing
End Sub


Private Sub CargaShowProcessList()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String

    On Error GoTo ECargaShowProcessList
    
    Set RS = New ADODB.Recordset
    
    ListView4.ColumnHeaders.Clear
    
    ListView4.ColumnHeaders.Add , , "ID", 1500.0631
    ListView4.ColumnHeaders.Add , , "User", 2250.2522, 1
    ListView4.ColumnHeaders.Add , , "Host", 3000.2522, 1
    ListView4.ColumnHeaders.Add , , "Tiempo espera", 3050.2522, 1
    
    
    Set RS = New ADODB.Recordset
    
    SERVER = Mid(Conn.ConnectionString, InStr(LCase(Conn.ConnectionString), "server=") + 7)
    SERVER = Mid(SERVER, 1, InStr(1, SERVER, ";"))
    
    EquipoConBD = (UCase(vUsu.PC) = UCase(SERVER)) Or (LCase(SERVER) = "localhost")
    
    Cad = "show full processlist"
    RS.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not RS.EOF
        If Not IsNull(RS.Fields(3)) Then
            If InStr(1, RS.Fields(3), "ariconta") <> 0 Then
                If UCase(RS.Fields(3)) = UCase(vUsu.CadenaConexion) Then
                    Equipo = RS.Fields(2)
                    'Primero quitamos los dos puntos del puerto
                    NumRegElim = InStr(1, Equipo, ":")
                    If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
                    
                    'El punto del dominio
                    NumRegElim = InStr(1, Equipo, ".")
                    If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
                    
                    Equipo = UCase(Equipo)
                    
                    
                    Set IT = ListView4.ListItems.Add
                    
                    IT.Text = RS.Fields(0)
                    IT.SubItems(1) = RS.Fields(1)
                    IT.SubItems(2) = Equipo
                    
                    'tiempo de espera
                    Dim FechaAnt As Date
                    FechaAnt = DateAdd("s", RS.Fields(5), Now)
                    IT.SubItems(3) = Format((Now - FechaAnt), "hh:mm:ss")
                End If
            End If
        End If
        
        'Siguiente
        RS.MoveNext
    Wend
    NumRegElim = 0
    RS.Close
    Set RS = Nothing
    
    Exit Sub
    
ECargaShowProcessList:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set RS = Nothing
End Sub



Private Sub CargaCobrosFactura()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String

    On Error GoTo ECargaCobrosFactura
    
    Set RS = New ADODB.Recordset
    
    ListView5.ColumnHeaders.Clear
    
    ListView5.ColumnHeaders.Add , , "Ord.", 800.0631
    ListView5.ColumnHeaders.Add , , "Forma de Pago", 3000.2522
    ListView5.ColumnHeaders.Add , , "Fecha Vto", 1450.2522
    ListView5.ColumnHeaders.Add , , "Importe Vto", 1550.2522, 1
    ListView5.ColumnHeaders.Add , , "Gastos", 1550.2522, 1
    ListView5.ColumnHeaders.Add , , "F.Ult.Cobro", 1450.2522
    ListView5.ColumnHeaders.Add , , "Imp.Pagado", 1550.2522, 1
    ListView5.ColumnHeaders.Add , , "Pendiente", 1550.2522, 1
    
    Set RS = New ADODB.Recordset
    
    ListView5.SmallIcons = frmPpal.imgListComun
    
    Cad = "select numorden, formapago.nomforpa, fecvenci, impvenci, gastos, fecultco, impcobro, (coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0)) pendiente, cobros.ctabanc1, numasien  "
    Cad = Cad & " from (cobros left join formapago on cobros.codforpa = formapago.codforpa) "
    Cad = Cad & " where cobros.numserie = " & DBSet(RecuperaValor(Parametros, 1), "T")
    Cad = Cad & " and cobros.numfactu = " & DBSet(RecuperaValor(Parametros, 2), "N")
    Cad = Cad & " and cobros.fecfactu = " & DBSet(RecuperaValor(Parametros, 3), "F")
    Cad = Cad & " order by numorden "
    
    RS.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not RS.EOF
                    
        Set IT = ListView5.ListItems.Add
        
        If DBLet(RS!NumAsien, "N") <> 0 Then IT.SmallIcon = 18
        
        IT.Text = DBLet(RS.Fields(0))
        IT.SubItems(1) = DBLet(RS.Fields(1))
        IT.SubItems(2) = DBLet(RS.Fields(2))
        IT.SubItems(3) = Format(DBLet(RS.Fields(3)), "###,###,##0.00")
        IT.SubItems(4) = Format(DBLet(RS.Fields(4)), "###,###,##0.00")
        IT.SubItems(5) = DBLet(RS.Fields(5))
        IT.SubItems(6) = Format(DBLet(RS.Fields(6)), "###,###,##0.00")
        IT.SubItems(7) = Format(DBLet(RS.Fields(7)), "###,###,##0.00")
        
        'Siguiente
        RS.MoveNext
    Wend
    NumRegElim = 0
    RS.Close
    Set RS = Nothing
    
    Exit Sub
    
ECargaCobrosFactura:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set RS = Nothing
End Sub



Private Sub CargaPagosFactura()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String

    On Error GoTo ECargaPagosFactura
    
    Set RS = New ADODB.Recordset
    
    ListView5.ColumnHeaders.Clear
    
    ListView5.ColumnHeaders.Add , , "Ord.", 800.0631
    ListView5.ColumnHeaders.Add , , "Forma de Pago", 3000.2522
    ListView5.ColumnHeaders.Add , , "Fecha Vto", 1450.2522
    ListView5.ColumnHeaders.Add , , "Importe Vto", 1550.2522, 1
    ListView5.ColumnHeaders.Add , , "F.Ult.Pago", 1450.2522
    ListView5.ColumnHeaders.Add , , "Imp.Pagado", 1550.2522, 1
    ListView5.ColumnHeaders.Add , , "Pendiente", 1550.2522, 1
    
    Set RS = New ADODB.Recordset
    
    ListView5.SmallIcons = frmPpal.imgListComun
    
    Cad = "select numorden, formapago.nomforpa, fecefect, impefect, fecultpa, imppagad, (coalesce(impefect,0)  - coalesce(imppagad,0)) pendiente, pagos.ctabanc1, numasien  "
    Cad = Cad & " from (pagos left join formapago on pagos.codforpa = formapago.codforpa) "
    Cad = Cad & " where pagos.numserie = " & DBSet(RecuperaValor(Parametros, 1), "T")
    Cad = Cad & " and pagos.codmacta = " & DBSet(RecuperaValor(Parametros, 2), "T")
    Cad = Cad & " and pagos.numfactu = " & DBSet(RecuperaValor(Parametros, 3), "T")
    Cad = Cad & " and pagos.fecfactu = " & DBSet(RecuperaValor(Parametros, 4), "F")
    Cad = Cad & " order by numorden "
    
    RS.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not RS.EOF
                    
        Set IT = ListView5.ListItems.Add
        
        If DBLet(RS!NumAsien, "N") <> 0 Then IT.SmallIcon = 18
        
        IT.Text = DBLet(RS.Fields(0))
        IT.SubItems(1) = DBLet(RS.Fields(1))
        IT.SubItems(2) = DBLet(RS.Fields(2))
        IT.SubItems(3) = Format(DBLet(RS.Fields(3)), "###,###,##0.00")
        IT.SubItems(4) = DBLet(RS.Fields(4))
        IT.SubItems(5) = Format(DBLet(RS.Fields(5)), "###,###,##0.00")
        IT.SubItems(6) = Format(DBLet(RS.Fields(6)), "###,###,##0.00")
        
        'Siguiente
        RS.MoveNext
    Wend
    NumRegElim = 0
    RS.Close
    Set RS = Nothing
    
    Exit Sub
    
ECargaPagosFactura:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set RS = Nothing
End Sub



Private Sub CargarAsiento()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarAsiento
    
    
    ListView7.ColumnHeaders.Clear
    ListView7.ListItems.Clear
    
    
'    ListView7.ColumnHeaders.Add , , "Ord.", 800.0631
    ListView7.ColumnHeaders.Add , , "Cuenta", 1500.2522
    ListView7.ColumnHeaders.Add , , "Denominaci�n", 4000.2522
    ListView7.ColumnHeaders.Add , , "Debe", 2050.2522, 1
    ListView7.ColumnHeaders.Add , , "Haber", 2050.2522, 1
    ListView7.ColumnHeaders.Add , , "Saldo", 2050.2522, 1
    
    Set RS = New ADODB.Recordset
    
    
    Pos = DevuelveValor("select max(pos) from tmpconext where codusu = " & DBSet(vUsu.Codigo, "N"))
    
    
    Cad = "select tmpconext.cta, cuentas.nommacta, tmpconext.timported, tmpconext.timporteh, acumtotT, tmpconext.pos"
    Cad = Cad & " from (ariconta" & NumConta & ".tmpconext inner join ariconta" & NumConta & ".cuentas on tmpconext.cta = cuentas.codmacta) "
    Cad = Cad & " left join ariconta" & NumConta & ".tmpconextcab on tmpconext.codusu = tmpconextcab.codusu and tmpconext.cta = tmpconextcab.cta"
    Cad = Cad & " where tmpconext.codusu = " & DBSet(vUsu.Codigo, "N")
    Cad = Cad & " order by pos "
    
    RS.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not RS.EOF
                    
        Set IT = ListView7.ListItems.Add
        
        
        IT.Text = DBLet(RS.Fields(0))
        IT.SubItems(1) = DBLet(RS.Fields(1))
        If DBLet(RS.Fields(2)) <> 0 Then
            IT.SubItems(2) = Format(DBLet(RS.Fields(2)), "###,###,##0.00")
        Else
            IT.SubItems(2) = ""
        End If
        If DBLet(RS.Fields(3)) <> 0 Then
            IT.SubItems(3) = Format(DBLet(RS.Fields(3)), "###,###,##0.00")
        Else
            IT.SubItems(3) = ""
        End If
        
        ' si no estamos en la �ltima l�nea mostramos el saldo de la cuenta
        If DBLet(RS.Fields(5)) <> Pos Then
            If DBLet(RS.Fields(4)) <> 0 Then
                IT.SubItems(4) = Format(DBLet(RS.Fields(4)), "###,###,##0.00")
            Else
                IT.SubItems(4) = ""
            End If
            IT.ListSubItems(4).ForeColor = &HAE8859   '&HEED68C
'            IT.ListSubItems(4).Bold = True
            
        End If
        
        If DBLet(RS.Fields(5)) = Pos Then
            IT.Bold = True
            IT.ListSubItems(1).Bold = True
            IT.ListSubItems(2).Bold = True
            IT.ListSubItems(3).Bold = True
'            IT.ListSubItems(4).Bold = True
        End If
        
        'Siguiente
        RS.MoveNext
    Wend
    NumRegElim = 0
    RS.Close
    Set RS = Nothing
    
    Exit Sub
    
ECargarAsiento:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set RS = Nothing
End Sub



Private Sub CargarAsientosDescuadrados()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarAsiento
    
    
    ListView8.ColumnHeaders.Clear
    ListView8.ListItems.Clear
    
    ListView8.ColumnHeaders.Add , , "Diario", 1000.2522
    ListView8.ColumnHeaders.Add , , "Asiento", 1500.2522
    ListView8.ColumnHeaders.Add , , "Fecha", 1500.2522
    ListView8.ColumnHeaders.Add , , "Debe", 2050.2522, 1
    ListView8.ColumnHeaders.Add , , "Haber", 2050.2522, 1
    
    Set RS = New ADODB.Recordset
    
    
    Cad = "select tmphistoapu.numdiari, tmphistoapu.numasien, tmphistoapu.fechaent, tmphistoapu.timported, tmphistoapu.timporteh, tmphistoapu.timported - tmphistoapu.timporteh "
    Cad = Cad & " from tmphistoapu "
    Cad = Cad & " where tmphistoapu.codusu = " & DBSet(vUsu.Codigo, "N")
    Cad = Cad & " order by numdiari, numasien, fechaent "
    
    RS.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not RS.EOF
                    
        Set IT = ListView8.ListItems.Add
        
        
        IT.Text = DBLet(RS.Fields(0))
        IT.SubItems(1) = DBLet(RS.Fields(1))
        IT.SubItems(2) = DBLet(RS.Fields(2))
        If DBLet(RS.Fields(3)) <> 0 Then
            IT.SubItems(3) = Format(DBLet(RS.Fields(3)), "###,###,##0.00")
        Else
            IT.SubItems(3) = ""
        End If
        If DBLet(RS.Fields(4)) <> 0 Then
            IT.SubItems(4) = Format(DBLet(RS.Fields(4)), "###,###,##0.00")
        Else
            IT.SubItems(4) = ""
        End If
        
        
        'Siguiente
        RS.MoveNext
    Wend
    NumRegElim = 0
    RS.Close
    Set RS = Nothing
    
    Exit Sub
    
ECargarAsiento:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set RS = Nothing
End Sub



Private Sub CargarFacturasSinAsientos()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarAsiento
    
    
    ListView8.ColumnHeaders.Clear
    ListView8.ListItems.Clear
    
    
    ListView8.ColumnHeaders.Add , , "Serie", 800.2522
    ListView8.ColumnHeaders.Add , , "Descripci�n", 2700.2522
    ListView8.ColumnHeaders.Add , , "Factura", 1500.2522
    ListView8.ColumnHeaders.Add , , "Fecha", 1500.2522
    ListView8.ColumnHeaders.Add , , "Total", 2000.2522, 1
    
    Set RS = New ADODB.Recordset
    
    Cad = "select tmpfaclin.numserie, tmpfaclin.nomserie, tmpfaclin.numfac, tmpfaclin.fecha, tmpfaclin.total "
    Cad = Cad & " from tmpfaclin "
    Cad = Cad & " where tmpfaclin.codusu = " & DBSet(vUsu.Codigo, "N")
    Cad = Cad & " order by numserie, numfac, fecha "
    
    RS.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not RS.EOF
                    
        Set IT = ListView8.ListItems.Add
        
        IT.Text = DBLet(RS.Fields(0))
        IT.SubItems(1) = DBLet(RS.Fields(1))
        IT.SubItems(2) = DBLet(RS.Fields(2))
        IT.SubItems(3) = DBLet(RS.Fields(3))
        If DBLet(RS.Fields(4)) <> 0 Then
            IT.SubItems(4) = Format(DBLet(RS.Fields(4)), "###,###,##0.00")
        Else
            IT.SubItems(4) = ""
        End If
        
        'Siguiente
        RS.MoveNext
    Wend
    NumRegElim = 0
    RS.Close
    Set RS = Nothing
    
    Exit Sub
    
ECargarAsiento:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set RS = Nothing
End Sub




Private Sub CargarFacturasReclamaciones()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarFacturas
    
    
    ListView9.ColumnHeaders.Clear
    ListView9.ListItems.Clear
    
    
    ListView9.ColumnHeaders.Add , , "Serie", 800.2522
    ListView9.ColumnHeaders.Add , , "Factura", 2000.2522
    ListView9.ColumnHeaders.Add , , "Fecha", 2000.2522
    ListView9.ColumnHeaders.Add , , "Vto", 1500.2522
    ListView9.ColumnHeaders.Add , , "Importe", 2000.2522, 1
    
    Set RS = New ADODB.Recordset
    
    Cad = "select numserie, numfactu, fecfactu, numorden, impvenci importe "
    Cad = Cad & " from reclama_facturas "
    Cad = Cad & " where codigo = " & DBSet(RecuperaValor(Parametros, 1), "N")
    
    If CampoOrden = "" Then CampoOrden = "fecfactu"
    Cad = Cad & " ORDER BY " & CampoOrden
    If Orden Then Cad = Cad & " DESC"
    
'    Cad = Cad & " order by numserie, numfactu, fecfactu "
    
    RS.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not RS.EOF
                    
        Set IT = ListView9.ListItems.Add
        
        IT.Text = DBLet(RS.Fields(0))
        IT.SubItems(1) = DBLet(RS.Fields(1))
        IT.SubItems(2) = DBLet(RS.Fields(2))
        IT.SubItems(3) = DBLet(RS.Fields(3))
        If DBLet(RS.Fields(4)) <> 0 Then
            IT.SubItems(4) = Format(DBLet(RS.Fields(4)), "###,###,##0.00")
        Else
            IT.SubItems(4) = ""
        End If
        
        'Siguiente
        RS.MoveNext
    Wend
    NumRegElim = 0
    RS.Close
    Set RS = Nothing
    
    Exit Sub
    
ECargarFacturas:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set RS = Nothing
End Sub


Private Sub CargarFacturasRemesas()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim Cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarFacturas
    
    
    ListView9.ColumnHeaders.Clear
    ListView9.ListItems.Clear
    
    
    ListView9.ColumnHeaders.Add , , "Serie", 800.2522
    ListView9.ColumnHeaders.Add , , "Factura", 1500.2522
    ListView9.ColumnHeaders.Add , , "Fecha", 1500.2522
    ListView9.ColumnHeaders.Add , , "Vto", 700.2522
    ListView9.ColumnHeaders.Add , , "Fecha Vto", 1500.2522
    ListView9.ColumnHeaders.Add , , "Importe", 2000.2522, 1
    
    Set RS = New ADODB.Recordset
    
    Cad = "select cobros.numserie, cobros.numfactu, cobros.fecfactu, cobros.numorden, cobros.fecvenci, cobros.impvenci importe, ' ' devol"
    Cad = Cad & " from cobros "
    Cad = Cad & " where (cobros.codrem = " & DBSet(RecuperaValor(Parametros, 1), "N") & " and cobros.anyorem = " & DBSet(RecuperaValor(Parametros, 2), "N") & ") "
    Cad = Cad & " union "
    Cad = Cad & " select cobros_realizados.numserie, cobros_realizados.numfactu, cobros_realizados.fecfactu, cobros_realizados.numorden, cobros.fecvenci, cobros_realizados.impcobro * (-1) importe, '*' devol"
    Cad = Cad & " from cobros left join cobros_realizados on cobros.numserie = cobros_realizados.numserie and cobros.numfactu = cobros_realizados.numfactu and cobros.fecfactu = cobros_realizados.fecfactu and cobros.numorden = cobros_realizados.numorden "
    Cad = Cad & " where (cobros_realizados.codrem = " & DBSet(RecuperaValor(Parametros, 1), "N") & " and cobros_realizados.anyorem = " & DBSet(RecuperaValor(Parametros, 2), "N") & ") "
    
    If CampoOrden = "" Then CampoOrden = "fecfactu"
    Cad = Cad & " ORDER BY " & CampoOrden
    If Orden Then Cad = Cad & " DESC"
    
    
    RS.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Cad = ""
    While Not RS.EOF
                    
        Set IT = ListView9.ListItems.Add
        
        IT.Text = DBLet(RS.Fields(0))
        IT.SubItems(1) = DBLet(RS.Fields(1))
        IT.SubItems(2) = DBLet(RS.Fields(2))
        IT.SubItems(3) = DBLet(RS.Fields(3))
        IT.SubItems(4) = DBLet(RS.Fields(4))
        
        If DBLet(RS.Fields(5)) <> 0 Then
            IT.SubItems(5) = Format(DBLet(RS.Fields(5)), "###,###,##0.00")
        Else
            IT.SubItems(5) = ""
        End If
        'Siguiente
        
        If RS.Fields(6) = "*" Then
            IT.Bold = True
            IT.ForeColor = vbRed
        End If
        
        
        
        RS.MoveNext
    Wend
    NumRegElim = 0
    RS.Close
    Set RS = Nothing
    
    Exit Sub
    
ECargarFacturas:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set RS = Nothing
End Sub





