VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMensajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14130
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
   ScaleWidth      =   14130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameTalPagPdtes 
      Height          =   6720
      Left            =   0
      TabIndex        =   252
      Top             =   0
      Width           =   14010
      Begin VB.TextBox txtSuma 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEF7E4&
         Enabled         =   0   'False
         Height          =   360
         Left            =   7950
         TabIndex        =   259
         Top             =   6030
         Width           =   1575
      End
      Begin VB.CommandButton CmdCancelTalPag 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   12300
         TabIndex        =   254
         Top             =   6000
         Width           =   1365
      End
      Begin VB.CommandButton CmdAcepTalPag 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   10710
         TabIndex        =   253
         Top             =   6000
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView12 
         Height          =   4905
         Left            =   225
         TabIndex        =   255
         Top             =   1005
         Width           =   13545
         _ExtentX        =   23892
         _ExtentY        =   8652
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   360
         TabIndex        =   257
         Top             =   6000
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Dividir Vencimiento"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblSuma 
         Caption         =   "Suma"
         Height          =   255
         Left            =   7260
         TabIndex        =   260
         Top             =   6060
         Width           =   615
      End
      Begin VB.Label Label34 
         Caption         =   "boton de dividir vto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   810
         TabIndex        =   258
         Top             =   6030
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   6
         Left            =   13440
         Picture         =   "frmMensajes.frx":000C
         ToolTipText     =   "Puntear al Debe"
         Top             =   690
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   7
         Left            =   13080
         Picture         =   "frmMensajes.frx":0156
         ToolTipText     =   "Quitar al Debe"
         Top             =   690
         Width           =   240
      End
      Begin VB.Label Label33 
         Caption         =   "Talones / Pagarés Pendientes "
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
         TabIndex        =   256
         Top             =   390
         Width           =   12465
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
   Begin VB.Frame FrameBancosRemesas 
      Height          =   6720
      Left            =   0
      TabIndex        =   242
      Top             =   0
      Width           =   8670
      Begin VB.CommandButton CmdCancelBancoRem 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6750
         TabIndex        =   244
         Top             =   6060
         Width           =   1365
      End
      Begin VB.CommandButton CmdAcepBancoRem 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   5160
         TabIndex        =   243
         Top             =   6060
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView10 
         Height          =   4905
         Left            =   225
         TabIndex        =   245
         Top             =   1005
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   8652
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
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   5
         Left            =   7620
         Picture         =   "frmMensajes.frx":02A0
         ToolTipText     =   "Quitar al Debe"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   4
         Left            =   7980
         Picture         =   "frmMensajes.frx":03EA
         ToolTipText     =   "Puntear al Debe"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label31 
         Caption         =   "Importe por Banco"
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
         TabIndex        =   246
         Top             =   390
         Width           =   8025
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
         TabIndex        =   147
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
         Picture         =   "frmMensajes.frx":0534
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
         Caption         =   "Código oficial balance"
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
   Begin VB.Frame FrameShowProcess 
      Height          =   6720
      Left            =   0
      TabIndex        =   221
      Top             =   0
      Width           =   10950
      Begin VB.CommandButton CmdRegresar 
         Caption         =   "Salir"
         Height          =   375
         Left            =   9240
         TabIndex        =   222
         Top             =   6120
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   4905
         Left            =   225
         TabIndex        =   223
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
         TabIndex        =   224
         Top             =   390
         Width           =   10185
      End
   End
   Begin VB.Frame FrameInformeBBDD 
      Height          =   6720
      Left            =   0
      TabIndex        =   211
      Top             =   0
      Width           =   10950
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Salir"
         Height          =   375
         Left            =   9240
         TabIndex        =   212
         Top             =   6120
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   4905
         Left            =   225
         TabIndex        =   213
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
         TabIndex        =   220
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
         TabIndex        =   219
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
         TabIndex        =   218
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
         TabIndex        =   217
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
         TabIndex        =   216
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
         TabIndex        =   215
         Top             =   300
         Width           =   3705
      End
      Begin VB.Label Label46 
         Caption         =   "Concepto"
         Height          =   255
         Left            =   270
         TabIndex        =   214
         Top             =   660
         Width           =   2355
      End
   End
   Begin VB.Frame FrameIconosVisibles 
      Height          =   6720
      Left            =   0
      TabIndex        =   206
      Top             =   -60
      Width           =   7050
      Begin VB.CommandButton cmdAcepIconos 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   209
         Top             =   6060
         Width           =   1215
      End
      Begin VB.CommandButton cmdCanIconos 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5520
         TabIndex        =   208
         Top             =   6060
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   5235
         Left            =   225
         TabIndex        =   207
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
            Text            =   "Código"
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
         TabIndex        =   210
         Top             =   270
         Width           =   5145
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   6480
         Picture         =   "frmMensajes.frx":0F36
         Top             =   330
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   6120
         Picture         =   "frmMensajes.frx":1080
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
      TabIndex        =   149
      Top             =   0
      Width           =   6495
      Begin VB.TextBox txtImporteP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   8
         Left            =   4380
         Locked          =   -1  'True
         TabIndex        =   165
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
         TabIndex        =   164
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
         TabIndex        =   163
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
         TabIndex        =   160
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
         TabIndex        =   159
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
         TabIndex        =   158
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
         TabIndex        =   157
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
         TabIndex        =   156
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
         TabIndex        =   155
         Text            =   "Text9"
         Top             =   1200
         Width           =   1755
      End
      Begin VB.CommandButton cmdPunteo 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   4950
         TabIndex        =   150
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
         TabIndex        =   162
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label22 
         Caption         =   "Debe"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   161
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   240
         Index           =   13
         Left            =   5490
         TabIndex        =   154
         Top             =   840
         Width           =   510
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Sin puntear"
         Height          =   240
         Index           =   12
         Left            =   2730
         TabIndex        =   153
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label22 
         Caption         =   "Punteada"
         Height          =   195
         Index           =   11
         Left            =   810
         TabIndex        =   152
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
         TabIndex        =   151
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
         TabIndex        =   118
         Top             =   3330
         Width           =   5625
         Begin VB.TextBox txtsaldo 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   9
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   120
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
            TabIndex        =   119
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
            TabIndex        =   121
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
         TabIndex        =   116
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
         TabIndex        =   115
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
         Picture         =   "frmMensajes.frx":11CA
         Top             =   1605
         Width           =   240
      End
      Begin VB.Image Image6 
         Height          =   240
         Index           =   1
         Left            =   5820
         Picture         =   "frmMensajes.frx":1BCC
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
         TabIndex        =   117
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
         Caption         =   "Saldos histórico"
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
      TabIndex        =   176
      Top             =   0
      Width           =   9375
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   5
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   181
         Tag             =   "Observaciones|T|S|||cuentas|obsdatos|||"
         Text            =   "frmMensajes.frx":25CE
         Top             =   720
         Width           =   7665
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   4
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   180
         Tag             =   "Observaciones|T|S|||cuentas|obsdatos|||"
         Text            =   "frmMensajes.frx":25D4
         Top             =   720
         Width           =   1005
      End
      Begin VB.CommandButton cmdVerObservaciones 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   7920
         TabIndex        =   179
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
         TabIndex        =   177
         Tag             =   "Observaciones|T|S|||cuentas|obsdatos|||"
         Text            =   "frmMensajes.frx":25DA
         Top             =   1320
         Width           =   8775
      End
      Begin VB.Label Label22 
         Caption         =   "Descripción cuentas Plan General Contable  2008"
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
         TabIndex        =   178
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
      TabIndex        =   166
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
         TabIndex        =   173
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
         TabIndex        =   172
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
         TabIndex        =   171
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
         TabIndex        =   168
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
         TabIndex        =   167
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
         TabIndex        =   175
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
         TabIndex        =   174
         Top             =   2160
         Width           =   1530
      End
      Begin VB.Image Image4 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmMensajes.frx":25E0
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
         TabIndex        =   170
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
         TabIndex        =   169
         Top             =   360
         Width           =   6495
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
         Text            =   "frmMensajes.frx":2FE2
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
         Text            =   "frmMensajes.frx":2FE8
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
         Text            =   "frmMensajes.frx":2FEE
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
         Text            =   "frmMensajes.frx":2FF4
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
         Text            =   "frmMensajes.frx":30F1
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
         Text            =   "frmMensajes.frx":30F7
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
      TabIndex        =   197
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
         TabIndex        =   199
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
         TabIndex        =   198
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
         TabIndex        =   200
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
            TabIndex        =   202
            Text            =   "frmMensajes.frx":30FD
            Top             =   840
            Width           =   6015
         End
         Begin VB.Label Label42 
            BackColor       =   &H00E3FEFF&
            Caption         =   "Teclas Rápidas en Mantenimientos"
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
            TabIndex        =   201
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
         TabIndex        =   203
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
            TabIndex        =   205
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
            TabIndex        =   204
            Top             =   270
            Width           =   5805
         End
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
         Text            =   "frmMensajes.frx":3170
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
         Text            =   "frmMensajes.frx":3176
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
         Caption         =   "Asientos Erróneos"
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
      TabIndex        =   194
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4215
         Left            =   120
         TabIndex        =   195
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
         TabIndex        =   196
         Top             =   240
         Width           =   4935
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   4920
         Picture         =   "frmMensajes.frx":317C
         ToolTipText     =   "Quitar seleccion"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   4920
         Picture         =   "frmMensajes.frx":32C6
         ToolTipText     =   "Todos"
         Top             =   1080
         Width           =   240
      End
   End
   Begin VB.Frame FrameCobros 
      Height          =   6720
      Left            =   0
      TabIndex        =   225
      Top             =   -30
      Width           =   13410
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   12000
         TabIndex        =   226
         Top             =   6000
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   4905
         Left            =   225
         TabIndex        =   227
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
         TabIndex        =   228
         Top             =   390
         Width           =   10185
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
      TabIndex        =   182
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
         TabIndex        =   193
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
         TabIndex        =   192
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
         TabIndex        =   191
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
         TabIndex        =   188
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdBloqEmpre 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   9840
         TabIndex        =   186
         Top             =   6840
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5775
         Index           =   0
         Left            =   210
         TabIndex        =   184
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
         TabIndex        =   183
         Top             =   6840
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5775
         Index           =   1
         Left            =   6240
         TabIndex        =   185
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
         TabIndex        =   190
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
         TabIndex        =   189
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
         TabIndex        =   187
         Top             =   240
         Width           =   5295
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
         Caption         =   "Valor adquisición   x  "
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
         Caption         =   "(Valor adquisición -amort. acumulada)  x"
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
         Caption         =   "años de vida"
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
         Caption         =   "Valor adquisición - valor residual"
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
         Caption         =   "años de vida"
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
         Caption         =   "Valor adquisición"
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
         Caption         =   "Tipos de amortización"
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
   Begin VB.Frame FrameAsientoLiquida 
      Height          =   6720
      Left            =   0
      TabIndex        =   229
      Top             =   0
      Width           =   12270
      Begin VB.CommandButton CmdContabilizar 
         Caption         =   "Contabilizar"
         Height          =   375
         Left            =   9060
         TabIndex        =   233
         Top             =   6000
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Left            =   10650
         TabIndex        =   230
         Top             =   6000
         Width           =   1365
      End
      Begin MSComctlLib.ListView ListView7 
         Height          =   4905
         Left            =   225
         TabIndex        =   231
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
         TabIndex        =   232
         Top             =   390
         Width           =   10185
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
      TabIndex        =   134
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
         TabIndex        =   139
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
            TabIndex        =   144
            Text            =   "Text8"
            Top             =   600
            Width           =   5295
         End
         Begin VB.Image Image5 
            Height          =   240
            Left            =   840
            Picture         =   "frmMensajes.frx":3410
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
            TabIndex        =   145
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
         TabIndex        =   142
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
         TabIndex        =   140
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
         TabIndex        =   138
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
         TabIndex        =   137
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
         TabIndex        =   136
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
         TabIndex        =   148
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
         TabIndex        =   146
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
         TabIndex        =   143
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
         TabIndex        =   141
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
         TabIndex        =   135
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
      TabIndex        =   122
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
         TabIndex        =   126
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
         TabIndex        =   125
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
         TabIndex        =   124
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
         TabIndex        =   123
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
         TabIndex        =   128
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
         TabIndex        =   127
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
         TabIndex        =   133
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
         TabIndex        =   132
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
         TabIndex        =   131
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
         TabIndex        =   130
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
         TabIndex        =   129
         Top             =   720
         Width           =   975
      End
   End
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
      TabIndex        =   238
      Top             =   0
      Width           =   9795
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   3
         Left            =   8370
         TabIndex        =   239
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
               Picture         =   "frmMensajes.frx":3E12
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":9604
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":A016
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView9 
         Height          =   3735
         Left            =   180
         TabIndex        =   240
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
         TabIndex        =   241
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
      TabIndex        =   234
      Top             =   0
      Width           =   8865
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   2
         Left            =   7620
         TabIndex        =   236
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
               Picture         =   "frmMensajes.frx":A468
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":FC5A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":1066C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   375
         Left            =   120
         TabIndex        =   235
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
         TabIndex        =   237
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
               Picture         =   "frmMensajes.frx":10ABE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":162B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMensajes.frx":16CC2
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
         Caption         =   "Cálculo de saldos"
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
   Begin VB.Frame FrameRecibos 
      Height          =   6720
      Left            =   0
      TabIndex        =   247
      Top             =   0
      Width           =   8670
      Begin VB.CommandButton CmdAcepRecibos 
         Caption         =   "Continuar"
         Height          =   375
         Left            =   5160
         TabIndex        =   249
         Top             =   6060
         Width           =   1455
      End
      Begin VB.CommandButton CmdCanRecibos 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6750
         TabIndex        =   248
         Top             =   6060
         Width           =   1365
      End
      Begin MSComctlLib.ListView ListView11 
         Height          =   4905
         Left            =   225
         TabIndex        =   250
         Top             =   1005
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   8652
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   0
      End
      Begin VB.Label Label32 
         Caption         =   "Recibos con cobros parciales"
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
         TabIndex        =   251
         Top             =   390
         Width           =   8025
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
      Caption         =   "años de vida"
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
      Caption         =   "Valor adquisición"
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
     
     '51- Facturas remesas
     '52- Bancos remesas
     '53- Recibos con cobros parciales
     '54- Cobros pendientes en recepcion de documentos (talones/pagares)
    
Public Parametros As String
    '1.- Vendran empipados: Cuenta, PunteadoD, punteadoH, pdteD,PdteH

'recepcion de talon/pagare
Public Importe As Currency
Public Codigo As String
Public Tipo As String
Public FecCobro As String
Public FecVenci As String
Public Banco As String
Public Referencia As String


Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private PrimeraVez As Boolean

Dim I As Integer
Dim Sql As String
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

Private Sub CmdAcepBancoRem_Click()
Dim I As Integer

    CadenaDesdeOtroForm = ""

    For I = 1 To ListView10.ListItems.Count
        If ListView10.ListItems(I).Checked Then
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "'" & Trim(ListView10.ListItems(I).Text) & "',"
        End If
    Next I
        
    If CadenaDesdeOtroForm <> "" Then CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, 1, Len(CadenaDesdeOtroForm) - 1)
    
    Unload Me
End Sub

Private Sub cmdAcepIconos_Click()
Dim I As Integer
Dim Sql As String
Dim CadenaIconos As String


Dim k As Integer
Dim J As Integer
Dim Px As Single
Dim Py As Single


Dim H As Integer
Dim MargenX As Single
Dim MargenY As Single
Dim Ocupado As Boolean
    'Ponemos los que ha desmarcado a cero
    CadenaIconos = ""
    For I = 1 To ListView6.ListItems.Count
        If Not ListView6.ListItems(I).Checked Then
            CadenaIconos = CadenaIconos & ListView6.ListItems(I).Text & ","
        End If
    Next I
    If CadenaIconos <> "" Then
        CadenaIconos = Mid(CadenaIconos, 1, Len(CadenaIconos) - 1)
        Sql = "update menus_usuarios set posx = 0, posy = 0, vericono = 0 where aplicacion = 'ariconta' and codusu = " & vUsu.Id & " and codigo in (" & CadenaIconos & ")"
        Conn.Execute Sql
    End If

    

    For I = 1 To ListView6.ListItems.Count
        If ListView6.ListItems(I).Checked Then
            'SI NO ERA VISIBLE le busco el hueco
            If ListView6.ListItems(I).SubItems(2) = "0" Then
                Sql = ""
                For J = 1 To 8
                    For k = 1 To 5
                        DevuelCoordenadasCuadricula J, k, Px, Py
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
                            ListView6.ListItems(I).SubItems(2) = "1"
                            ListView6.ListItems(I).SubItems(3) = Px
                            ListView6.ListItems(I).SubItems(4) = Py
                            Sql = "update menus_usuarios set posx = " & DBSet(Px, "N")
                            Sql = Sql & ", posy = " & DBSet(Py, "N") & ", vericono = 1 where "
                            Sql = Sql & "aplicacion = 'ariconta' and codusu = " & vUsu.Id
                            Sql = Sql & " and codigo =" & DBSet(ListView6.ListItems(I).Text, "T")
                            Conn.Execute Sql
                           Exit For
                        End If
                    Next k
                    If Sql <> "" Then Exit For
                Next J
            End If
           
        End If
    Next I
    Reorganizar = True
    
    Unload Me
End Sub





Private Sub CmdAcepRecibos_Click()
    CadenaDesdeOtroForm = "OK"
    Unload Me
End Sub

Private Sub CmdAcepTalPag_Click()

    If Not DatosOK Then Exit Sub

    
    InsertarModificarTalones
    
    Unload Me
End Sub

Private Sub InsertarModificarTalones()
Dim Importe As Currency
Dim importe2 As Currency
Dim TipForpa As String
Dim J As Integer
        
    On Error GoTo EInsertarModificar
    
    CadenaDesdeOtroForm = ""
    
    Conn.BeginTrans
    
    For I = 1 To Me.ListView12.ListItems.Count
        If ListView12.ListItems(I).Checked Then
        
            Sql = "select * from talones_facturas where codigo = " & DBSet(Codigo, "N") & " and " & MontaWhere(I, False)
        
            If TotalRegistrosConsulta(Sql) = 0 Then
                
                Sql = "insert into `talones_facturas` (`codigo`,`numserie`,`numfactu`,`fecfactu`, "
                Sql = Sql & "`numorden`,`importe`,`contabilizado`) VALUES ("
                Sql = Sql & DBSet(Codigo, "N") & "," & DBSet(ListView12.ListItems(I).Text, "T") & "," & DBSet(ListView12.ListItems(I).SubItems(1), "N") & ","
                Sql = Sql & DBSet(ListView12.ListItems(I).SubItems(2), "F") & "," & DBSet(ListView12.ListItems(I).SubItems(4), "N") & ","
                Sql = Sql & DBSet(ListView12.ListItems(I).SubItems(9), "N") & ",0) "

                Conn.Execute Sql
                
                ' ahora en cobros
                ' sacamos el tipforpa
                Sql = MontaWhere(I, False)
                
                Sql = " WHERE cobros.codforpa=formapago.codforpa and " & Sql
                Sql = "select tipforpa from cobros,formapago " & Sql
                
                TipForpa = DevuelveValor(Sql)
                
                'actualizamos cobros
                Sql = "UPDATE cobros SET recedocu=1 "
                Sql = Sql & ", impcobro = coalesce(impcobro,0) + " & DBSet(ListView12.ListItems(I).SubItems(9), "N")
                Sql = Sql & ", fecultco = " & DBSet(FecCobro, "F")
                'Fecha vencimiento tb le pongo la de la recpcion
                Sql = Sql & ", fecvenci = " & DBSet(FecVenci, "F")
                'BANCO LO PONGO EN OBSERVACION
                Sql = Sql & ", observa = " & DBSet(Banco, "T")
                'Si no era forma de pago talon/pagare la pongo
                If Tipo = 0 Then
                    J = vbPagare
                Else
                    J = vbTalon
                End If
                
                If TipForpa <> J Then
                    'AQUI BUSCARE una forma de pago
                    J = Val(DevuelveDesdeBD("codforpa", "formapago", "tipforpa", CStr(J)))
                    If J > 0 Then Sql = Sql & ", codforpa = " & J
                    
                End If
                            
                Conn.Execute Sql & MontaWhere(I, True)
            Else
                Sql = "UPDATE cobros SET recedocu=1 "
                
                Importe = DevuelveValor("select importe from talones_facturas where codigo = " & DBSet(Codigo, "N") & " and " & MontaWhere(I, False)) + ListView12.ListItems(I).SubItems(9)

                
                Sql = Sql & ", impcobro = coalesce(impcobro,0) +  " & DBSet(Importe, "N")
                Sql = Sql & ", fecultco = " & DBSet(FecCobro, "F")
                'Fecha vencimiento tb le pongo la de la recpcion
                Sql = Sql & ", fecvenci = " & DBSet(FecVenci, "F")
                'BANCO LO PONGO EN OBSERVACION
                Sql = Sql & ", observa = " & DBSet(Banco, "T")
                'Si no era forma de pago talon/pagare la pongo
                If Tipo = 0 Then
                    J = vbPagare
                Else
                    J = vbTalon
                End If
                
                If TipForpa <> J Then
                    'AQUI BUSCARE una forma de pago
                    J = Val(DevuelveDesdeBD("codforpa", "formapago", "tipforpa", CStr(J)))
                    If J > 0 Then Sql = Sql & ", codforpa = " & J
                    
                End If
                            
                Conn.Execute Sql & MontaWhere(I, True)
            
            End If
        
            CadenaDesdeOtroForm = CadenaDesdeOtroForm & "(" & DBSet(ListView12.ListItems(I).Text, "T") & "," & DBSet(ListView12.ListItems(I).SubItems(1), "N") & "," & DBSet(ListView12.ListItems(I).SubItems(2), "F") & "," & DBSet(ListView12.ListItems(I).SubItems(4), "N") & "),"
            
        Else
            'Obtengo el importe del vto
            Importe = DevuelveValor("select impcobro from cobros " & MontaWhere(I, True))
            importe2 = DevuelveValor("select importe from talones_facturas where codigo = " & DBSet(Codigo, "N") & " and " & MontaWhere(I, False))
            
            If Importe <> importe2 Then
                'TODO EL IMPORTE estaba en la linea. Fecultco a NULL
                J = 1
                Importe = Importe - importe2
            Else
                J = 0
            End If
            
            Sql = "Delete from talones_facturas"
            Sql = Sql & " WHERE codigo =" & DBSet(Codigo, "N")
            Sql = Sql & " AND " & MontaWhere(I, False)
            Conn.Execute Sql
            
            'Updateo en cobros reestableciendo los valores
            
            Sql = "UPDATE cobros SET recedocu=0"
            If J = 0 Then
                Sql = Sql & ", impcobro = NULL, fecultco = NULL"
            Else
                Sql = Sql & ", impcobro = " & TransformaComasPuntos(CStr(Importe))  'NO somos capace sde ver cual fue la utlima fecha de amortizacion
            End If
            Sql = Sql & ", observa = NULL"
            Sql = Sql & " WHERE " & MontaWhere(I, False)
            
            Conn.Execute Sql
        End If
        
        'ponemos la situacion
        Sql = "update cobros set situacion = if(impvenci + coalesce(gastos,0) - coalesce(impcobro,0) = 0,1,0) "
        Sql = Sql & " where " & MontaWhere(I, False)
        Conn.Execute Sql
        
        
    Next I
    Conn.CommitTrans
    Exit Sub
    
EInsertarModificar:
    Conn.RollbackTrans
    MuestraError Err.Number, "Insertar Modificar Talones/Pagarés", Err.Description
End Sub



Private Function MontaWhere(I As Integer, conWhere As Boolean) As String
Dim Sql As String
    
    If conWhere Then
        Sql = " WHERE "
    Else
        Sql = " "
    End If
    Sql = Sql & " numserie = " & DBSet(ListView12.ListItems(I).Text, "T")
    Sql = Sql & " and numfactu = " & DBSet(ListView12.ListItems(I).SubItems(1), "N") & " and fecfactu = " & DBSet(ListView12.ListItems(I).SubItems(2), "F")
    Sql = Sql & " and numorden = " & DBSet(ListView12.ListItems(I).SubItems(4), "N")
    
    MontaWhere = Sql

End Function


Private Function DatosOK() As Boolean
Dim B As Boolean

    DatosOK = False
    
    If CCur(ComprobarCero(txtSuma)) > Importe Then
        If MsgBox("El importe de Talón/Pagaré es inferior a la suma de las facturas seleccionadas. " & vbCrLf & "Deberia dividir vencimiento." & vbCrLf & vbCrLf & " ¿ Continuar ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            B = False
        Else
            B = True
        End If
    Else
        If CCur(ComprobarCero(txtSuma)) < Importe Then
            If MsgBox("El importe de Talón/Pagaré es superior a la suma de las facturas seleccionadas. " & vbCrLf & vbCrLf & " ¿ Desea continuar ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                B = False
            Else
                B = True
            End If
        Else
            B = True
        End If
    End If
    
    DatosOK = B
    
End Function


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
                I = ListView2(Index).ListItems(Ok).Index
                PasarUnaEmpresaBloqueada Index = 0, I
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



Private Sub PasarUnaEmpresaBloqueada(ABLoquedas As Boolean, Indice As Integer)
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
    
    Sql = ListView2(Origen).ListItems(Indice).Key
    Set IT = ListView2(Destino).ListItems.Add(, Sql)
    IT.SmallIcon = NE
    IT.Text = ListView2(Origen).ListItems(Indice).Text
    IT.SubItems(1) = ListView2(Origen).ListItems(Indice).SubItems(1)

    'Borramos en origen
    ListView2(Origen).ListItems.Remove Indice
End Sub

Private Sub cmdBloqEmpre_Click(Index As Integer)
    If Index = 0 Then
        Sql = "DELETE FROM usuarios.usuarioempresa WHERE codusu =" & Parametros
        Conn.Execute Sql
        Sql = ""
        For I = 1 To ListView2(1).ListItems.Count
            Sql = Sql & ", (" & Parametros & "," & Val(Mid(ListView2(1).ListItems(I).Key, 2)) & ")"
        Next I
        If Sql <> "" Then
            'Quitmos la primera coma
            Sql = Mid(Sql, 2)
            Sql = "INSERT INTO usuarios.usuarioempresa(codusu,codempre) VALUES " & Sql
            If Not EjecutaSQL(Sql) Then MsgBox "Se han producido errores insertando datos", vbExclamation
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
        I = 0
        Do
          
            Sql = "select numasien,fechaent from hcabapu where fechaent >= '"
            Sql = Sql & Format(DateAdd("yyyy", Anyo, vParam.fechaini), FormatoFecha)
            Sql = Sql & "' AND fechaent <= '" & Format(DateAdd("yyyy", Anyo, vParam.fechafin), FormatoFecha) & "' ORDER By NumAsien"
            Set RS = New ADODB.Recordset
            RS.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            ii = 0
            While Not RS.EOF
               J = RS.Fields(0)
               'Igual
                If J - ii = 0 Then
                    
                    Sql = Format(J, "00000")
                    Sql = Sql & "  -  " & Format(RS!FechaEnt, "dd/mm/yyyy")
                    Text5.Text = Text5.Text & Sql & vbCrLf
                    I = I + 1
                Else
                    If J - ii > 1 Then
                        If J - ii = 2 Then
                            Sql = Format(J - 1, "00000")
                        Else
                            Sql = "Entre " & Format(ii, "00000") & "  y  " & Format(J, "00000")
                        End If
                        Sql = Sql & " (" & CStr(Year(vParam.fechaini) + Anyo) & ")"
                        Text6.Text = Text6.Text & Sql & vbCrLf
                        I = I + 1
                    End If
                End If
                ii = J
                'Refrescamos
                If I > 50 Then
                    Text5.Refresh
                    Text6.Refresh
                    I = 0
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
    
    For I = 1 To Text7.Count - 1
        Text7(I).Text = Trim(Text7(I).Text)
        If Text7(I).Text = "" Then
            MsgBox "Hay que rellenar todos los campos", vbExclamation
            Exit Sub
        End If
    Next I
    
    
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
    Sql = "UPDATE Usuarios.Usuarios Set passwordpropio='" & Text7(2).Text
    Sql = Sql & "' WHERE codusu = " & (vUsu.Codigo Mod 1000)
    
    Conn.Execute Sql
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cambio clave"
    Else
        vUsu.PasswdPROPIO = Text7(2).Text
        MsgBox "Cambio de clave realizado con éxito", vbInformation
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

Private Sub CmdCancelBancoRem_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub CmdCancelTalPag_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub cmdCanIconos_Click()
    Reorganizar = False
    Unload Me
End Sub

Private Sub CmdCanRecibos_Click()
    CadenaDesdeOtroForm = ""
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
            MsgBox "La cuenta debe ser numérica", vbExclamation
            Exit Sub
        End If
        'Esto es el OPTION
        Sql = ""
        For I = 0 To 2
            If Option1(I).Value Then Sql = Sql & Mid(Option1(I).Caption, 1, 1)
        Next I
        If Sql = "" Then
            MsgBox "Seleccione una opción de la cuenta (Saldo - Debe - Haber )"
            Exit Sub
        End If
        
        'RESTA y la resta
        Sql = Sql & "|" & Abs(Me.chkResta.Value)
        CadenaDesdeOtroForm = Text3.Text & "|" & Sql & "|"
    End If
    Unload Me
End Sub

Private Sub cmdEmpresa_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    If Index = 0 Then
        Sql = ""
        Parametros = ""
        For I = 1 To lwE.ListItems.Count
            If Me.lwE.ListItems(I).Checked Then
                Sql = Sql & Me.lwE.ListItems(I).Text & "|"
                Parametros = Parametros & "1" 'Contador
            End If
        Next I
        CadenaDesdeOtroForm = Len(Parametros) & "|" & Sql
        'Vemos las conta
        Sql = ""
        For I = 1 To lwE.ListItems.Count
            If Me.lwE.ListItems(I).Checked Then
                Sql = Sql & Me.lwE.ListItems(I).Tag & "|"
            End If
        Next I
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & Sql
    End If
    Unload Me
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
        
        Sql = "Va a realizar la importación de datos  de facturas en la empresa: " & vbCrLf & vbCrLf & vEmpresa.nomempre
        Sql = Sql & "(" & vEmpresa.nomresum & ") - Conta: " & vEmpresa.codempre & vbCrLf & vbCrLf & "¿Desea continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
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
        Case 52
            CargarBancosRemesas
        Case 53
            CargarRecibosConCobrosParciales
        Case 54
            CargarTalonesPagaresPendientes
            
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
    Me.FrameAmort.Visible = False
    Me.FrameeMPRESAS.Visible = False
    Me.frameAcercaDE.Visible = False
    Me.frameCtasBalance.Visible = False
    Me.FrameCarta347.Visible = False
    Me.frameSaltos.Visible = False
    frameBalance.Visible = False
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
    Me.FrameBancosRemesas.Visible = False
    Me.FrameRecibos.Visible = False
    Me.FrameTalPagPdtes.Visible = False
    
    
    ' botón de dividir vencimiento cuando estamos en talones/pagares pendientes
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 44
    End With
    
    
    Select Case Opcion
    Case 1
        Me.Caption = "Cálculo de saldo"
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
        Me.Caption = "Información tipo amortización"
        W = Me.FrameAmort.Width
        H = Me.FrameAmort.Height + 200
        Me.FrameAmort.Visible = True
    Case 4
        Me.Caption = "Seleccion"
        W = Me.FrameeMPRESAS.Width
        H = Me.FrameeMPRESAS.Height + 200
        Me.FrameeMPRESAS.Visible = True
        cargaempresas
    Case 5
        'Lanzar automaticamente la comprobación de saldo
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
        Label13.Caption = "Versión:  " & App.Major & "." & App.Minor & "." & App.Revision & " "
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
        Me.Caption = "Búsqueda cabeceras asientos incorrectos"
        W = Me.frameSaltos.Width
        H = Me.frameSaltos.Height + 300
        Me.frameSaltos.Visible = True
        Me.cmdCabError(0).Enabled = True
        Text5.Text = ""
        Text6.Text = ""
        cmdCabError(1).Cancel = True
    Case 15
        'Cambio password usuario
        Me.Caption = "Cambio password"
        W = Me.FrameCambioPWD.Width
        H = Me.FrameCambioPWD.Height + 300
        Me.FrameCambioPWD.Visible = True
        Text7(0).Text = vUsu.Nombre
        For I = 1 To 3
            Text7(I).Text = ""
        Next I
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
        For I = 0 To 8
            Me.txtImporteP(I).Text = RecuperaValor(Parametros, I + 1)
        Next I
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
        Me.Caption = "Información de Base de Datos"
        Me.FrameInformeBBDD.Visible = True
        W = Me.FrameInformeBBDD.Width
        H = Me.FrameInformeBBDD.Height + 300
        
        Me.Label47.Caption = "Ejercicio " & vParam.fechaini & " a " & vParam.fechafin
        Me.Label48.Caption = "Ejercicio " & DateAdd("yyyy", 1, vParam.fechaini) & " a " & DateAdd("yyyy", 1, vParam.fechafin)
        
    Case 26 ' show process list
        Me.Caption = "Información de Procesos del Sistema"
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
        Me.Caption = "Asiento de Liquidación"
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
        Me.Label30.Caption = "Reclamación a " & RecuperaValor(Parametros, 2) & " de fecha " & RecuperaValor(Parametros, 3)
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
            
    Case 52 ' bancos de remesas
        Me.Caption = "Remesas"
        Me.FrameBancosRemesas.Visible = True
        W = Me.FrameBancosRemesas.Width
        H = Me.FrameBancosRemesas.Height + 300
    
    Case 53 ' recibos con cobros parciales
        Me.Caption = "Recibos "
        Me.FrameRecibos.Visible = True
        W = Me.FrameRecibos.Width
        H = Me.FrameRecibos.Height + 300
            
    Case 54 ' cobros talones / pagares pendientes
        Me.Caption = "Talones/Pagarés Pendientes"
        Me.FrameTalPagPdtes.Visible = True
        W = Me.FrameTalPagPdtes.Width
        H = Me.FrameTalPagPdtes.Height + 300
            
    End Select
    Me.Width = W + 120
    Me.Height = H + 120
End Sub




Private Sub CargaValoresHco()
'Lo que hace es dado el parametro scamos la cuenta, nomcuenta, saldos
'1.- Vendran empipados: Cuenta, PunteadoD, punteadoH, pdteD,PdteH

Label6.Caption = RecuperaValor(Parametros, 1)
For I = 0 To 3
    Me.txtsaldo(I).Text = RecuperaValor(Parametros, I + 2)
Next I
CalculaSaldosFinales
End Sub



Private Sub CalculaSaldosFinales()
Dim Importe As Currency
    For I = 0 To 3
        Importe = ImporteFormateado(txtsaldo(I).Text)
        txtsaldo(I).Tag = Importe
    Next I
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
    Sql = RecuperaValor(Parametros, 6)
    If Sql = "" Then
        NE = 0
        
    Else
        NE = 1
        Importe = CCur(Sql)
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
    Sql = Trim(RecuperaValor(Parametros, 7))   'Descripcion cuenta
    If Sql <> "" Then Sql = " - " & Sql
    Label6.Caption = Label6.Caption & Sql
    
    
    
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
    
    Sql = "Select * from Usuarios.Empresasariconta "
    If vUsu.Codigo > 0 Then Sql = Sql & " WHERE codempre<100 and conta like 'ariconta%'"
    Sql = Sql & " order by codempre"
    Set lwE.SmallIcons = Me.ImageList1
    lwE.ListItems.Clear
    Set RS = New ADODB.Recordset
    I = -1
    RS.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        Sql = "|" & RS!codempre & "|"
        If InStr(1, Prohibidas, Sql) = 0 Then
            Set ItmX = lwE.ListItems.Add(, , RS!nomempre, , 3)
            ItmX.Tag = RS!codempre
            If ItmX.Tag = vEmpresa.codempre Then
                If CadenaDesdeOtroForm = "" Then
                    ItmX.Checked = True
                    I = ItmX.Index
                End If
            End If
            ItmX.ToolTipText = RS!CONTA
        End If
        RS.MoveNext
    Wend
    RS.Close
    If I > 0 Then Set lwE.SelectedItem = lwE.ListItems(I)
    
    CadenaDesdeOtroForm = ""
    
Ecargaempresas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos empresas"
    Set RS = Nothing
End Sub

Private Sub VerEmresasProhibidas(ByRef VarProhibidas As String)

On Error GoTo EVerEmresasProhibidas
    VarProhibidas = "|"
    Sql = "Select codempre from Usuarios.usuarioempresa WHERE codusu = " & (vUsu.Codigo Mod 1000)
    Sql = Sql & " order by codempre"
    Set RS = New ADODB.Recordset
    RS.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
          VarProhibidas = VarProhibidas & RS!codempre & "|"
          RS.MoveNext
    Wend
    RS.Close
    Exit Sub
EVerEmresasProhibidas:
    MuestraError Err.Number, Err.Description & vbCrLf & " Consulte soporte técnico"
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
                .DialogTitle = "Seleccione archivo importación"
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


Private Sub imgCheck_Click(Index As Integer)
    For NE = 1 To TreeView1.Nodes.Count
        TreeView1.Nodes(NE).Checked = Index = 1
    Next
    
    Select Case Index
        ' ICONOS VISIBLES EN EL LISTVIEW DEL FRMPPAL
        Case 2 ' marcar todos
            For I = 1 To ListView6.ListItems.Count
                ListView6.ListItems(I).Checked = True
            Next I
        Case 3 ' desmarcar todos
            For I = 1 To ListView6.ListItems.Count
                ListView6.ListItems(I).Checked = False
            Next I
            
        ' bancos remesados
        Case 4 ' marcar todos
            For I = 1 To ListView10.ListItems.Count
                ListView10.ListItems(I).Checked = True
            Next I
        Case 5 ' desmarcar todos
            For I = 1 To ListView10.ListItems.Count
                If ListView10.ListItems(I).Text <> Parametros Then ListView10.ListItems(I).Checked = False
            Next I
            
        ' talones/pagares pendientes
        Case 6 ' marcar todos
            For I = 1 To ListView12.ListItems.Count
                ListView12.ListItems(I).Checked = True
            Next I
            SumaTotales
        Case 7 ' desmarcar todos
            For I = 1 To ListView12.ListItems.Count
                ListView12.ListItems(I).Checked = False
            Next I
            txtSuma.Text = ""
            
    End Select
        
    
End Sub


Private Sub ListView10_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Text = Parametros Then
        If Not Item.Checked Then
            MsgBox "El banco por defecto no puede ser desmarcado. ", vbExclamation
            Item.Checked = True
        End If
    End If
End Sub


Private Sub ListView12_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    SumaTotales
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
        For I = 0 To 3
            Text1(I).Text = ""
        Next I

    Else
        'NumBalan|Pasivo|codigo|padre|Orden|tipo|deslinea|texlinea|formula|TienenCtas|Negrita|LibroCD|
        Text1(0).Text = RecuperaValor(Parametros, 7)
        Text1(1).Text = RecuperaValor(Parametros, 8)
        I = Val(RecuperaValor(Parametros, 10))
        If I = 1 Then
            'Tiene cuentas
            Text1(2).Text = ""
            Text1(2).Enabled = False
        Else
            Text1(2).Text = RecuperaValor(Parametros, 9)
        End If
        I = Val(RecuperaValor(Parametros, 11))
        chkNegrita.Value = I
        I = Val(RecuperaValor(Parametros, 12))
        chkCero.Value = I
        I = Val(RecuperaValor(Parametros, 13))
        chkPintar.Value = I
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
        I = Val(RecuperaValor(Parametros, 3))
        Option1(I).Value = True
        I = Val(RecuperaValor(Parametros, 4))
        chkResta.Value = I
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
            MsgBox "El campo 'Concepto Libro CD' debe ser numérico", vbExclamation
            Exit Function
        End If
    End If
    
    'Hay k comprobar, si tiene formula k sea correcta
    Text1(2).Text = UCase(Trim(Text1(2).Text))
    If Text1(2).Text <> "" Then
        Sql = CompruebaFormulaConfigBalan(CInt(RecuperaValor(Parametros, 1)), Text1(2).Text)
        If Sql <> "" Then
            MsgBox Sql, vbExclamation
            Exit Function
        End If
    End If
    If Opcion = 7 Then
        Sql = "INSERT INTO balances_texto (NumBalan, Pasivo, codigo, padre, "
        Sql = Sql & "Orden, tipo, deslinea, texlinea, formula, TienenCtas, Negrita,A_Cero,Pintar,LibroCD) VALUES ("
        Sql = Sql & RecuperaValor(Parametros, 1)  'Numero
        Sql = Sql & ",'" & RecuperaValor(Parametros, 2) 'pasivo
        Sql = Sql & "'," & RecuperaValor(Parametros, 3)  'Codigo
        Aux = RecuperaValor(Parametros, 4) 'padre
        If Aux = "" Then
            Aux = ",NULL,"
        Else
            Aux = ",'" & Aux & "',"
        End If
        Sql = Sql & Aux
        Sql = Sql & RecuperaValor(Parametros, 5)
        If Text1(2).Text = "" Then
            Aux = "0"
        Else
            Aux = "1"
        End If
        Sql = Sql & "," & Aux
        Sql = Sql & ",'" & Text1(0).Text 'Text linea
        Sql = Sql & "','" & Text1(1).Text 'Desc linea
        Sql = Sql & "','" & Text1(2).Text 'Formula
        Sql = Sql & "',0," & chkNegrita.Value
        Sql = Sql & "," & Me.chkCero.Value
        Sql = Sql & "," & Me.chkPintar.Value
        Sql = Sql & ",'" & Text1(3).Text 'Libro CD
        Sql = Sql & "')"
    Else
        'Modificar
        'NumBalan|Pasivo|codigo|padre|Orden|tipo|deslinea|texlinea|formula|TienenCtas|Negrita|
        Sql = "UPDATE balances_texto SET "
        Sql = Sql & "deslinea='" & Text1(0).Text & "',"
        Sql = Sql & "texlinea='" & Text1(1).Text & "',"
        Sql = Sql & "formula='" & Text1(2).Text & "',"
        If Text1(2).Text = "" Then
            Aux = "0"
        Else
            Aux = "1"
        End If
        Sql = Sql & "Tipo =" & Aux & ","
        Sql = Sql & "Negrita = " & chkNegrita.Value
        Sql = Sql & ", A_Cero = " & Me.chkCero.Value
        Sql = Sql & ", Pintar = " & Me.chkPintar.Value
        Sql = Sql & ", LibroCD = '" & Text1(3).Text & "'"
        Sql = Sql & " WHERE numbalan =" & RecuperaValor(Parametros, 1)
        Sql = Sql & " AND Pasivo = '" & RecuperaValor(Parametros, 2)
        Sql = Sql & "' AND codigo = " & RecuperaValor(Parametros, 3)
        
    End If
    Conn.Execute Sql
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
    Sql = App.Path & "\txt347.dat"
    If Dir(Sql) <> "" Then
        'Vamos a ir leyendo , y devoviendo cadena
        I = FreeFile
        Open Sql For Input As #I
        For NumRegElim = 0 To Text4.Count - 1
            'Obtenemos la cadena
           LeerCadenaFicheroTexto    'lo guarda en SQL
           Text4(NumRegElim) = Sql
        Next NumRegElim
        Close #I
        Limpiar = False
    End If
ECargarDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Carga fichero. " & Err.Description
    If Limpiar Then
        'No existe el fihero de configuracion
        For I = 0 To Text4.Count - 1
            Text4(I).Text = ""
        Next I
    End If
End Sub


Private Sub LeerCadenaFicheroTexto()
On Error GoTo ELeerCadenaFicheroTexto
    'Son dos lineas. La primaera indica k campo y la segunda el valor
    Line Input #I, Sql
    Line Input #I, Sql
    Exit Sub
ELeerCadenaFicheroTexto:
    Sql = ""
    Err.Clear
End Sub


Private Function GuardarDatosCarta()
    On Error GoTo Eguardardatoscarta
    Sql = App.Path & "\txt347.dat"
    I = FreeFile
    Open Sql For Output As #I
    For NumRegElim = 0 To Text4.Count - 1
        Print #I, Text4(NumRegElim).Tag
        Print #I, Text4(NumRegElim).Text
    Next NumRegElim
    Close #I
    Exit Function
Eguardardatoscarta:
    MuestraError Err.Number, "guardar datos carta"
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Function ValorSQL(ByRef C As String) As String
    If C = "" Then
        ValorSQL = "NULL"
    Else
        ValorSQL = "'" & C & "'"
    End If
End Function
Private Function EjecutaSQL2(Sql As String) As Boolean
    EjecutaSQL2 = False
    On Error Resume Next
    Conn.Execute Sql
    If Err.Number <> 0 Then
        AnyadeErrores "SQL: " & Sql, Err.Description
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
    Sql = DevuelveDesdeBD("codmacta", "cuentas", "codmacta", Cta, "T")
    If Sql <> "" Then ExisteCuenta = True
    
End Function

Private Sub Text7_GotFocus(Index As Integer)
    Text7(Index).SelStart = 0
    Text7(Index).SelLength = Len(Text7(Index).Text)
End Sub

Private Sub Text7_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    'dividir vencimientos
    If Me.ListView12.SelectedItem Is Nothing Then Exit Sub
    
    DividirVencimiento

End Sub

Private Sub DividirVencimiento()
Dim Im As Currency

    If ListView12.ListItems.Count = 0 Then Exit Sub
    If ListView12.SelectedItem Is Nothing Then Exit Sub
        
    
    
    'Si esta totalmente cobrado pues no podemos desdoblar ekl vto
    Im = ImporteFormateado(ListView12.SelectedItem.SubItems(9))
    If Im <= 0 Then
        MsgBox "NO puede dividir el vencimiento. Importe totalmente cobrado", vbExclamation
        Exit Sub
    End If
    
       'CadenaDesdeOtroForm. Pipes
        '           1.- cadenaSQL numfac,numsere,fecfac
        '           2.- Numero vto
        '           3.- Importe maximo
    
    CadenaDesdeOtroForm = "numserie = " & DBSet(ListView12.SelectedItem.Text, "T") & " AND numfactu = " & ListView12.SelectedItem.SubItems(1)
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & " AND fecfactu = " & DBSet(ListView12.SelectedItem.SubItems(2), "F") & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & ListView12.SelectedItem.SubItems(4) & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & CStr(Im) & "|"
    
    
    'Ok, Ahora pongo los labels
    frmTESCobrosDivVto.Opcion = 27
    frmTESCobrosDivVto.Label4(56).Caption = Label33.Caption
    frmTESCobrosDivVto.txtcodigo(2).Text = ListView12.SelectedItem.SubItems(3)
    
    frmTESCobrosDivVto.Label4(57).Caption = ListView12.SelectedItem.Text & Format(ListView12.SelectedItem.SubItems(1), "0000000") & " / " & ListView12.SelectedItem.SubItems(4) & "      de " & Format(ListView12.SelectedItem.SubItems(2), "dd/mm/yyyy")
    
    'Si ya ha cobrado algo...
    Im = DBLet(ComprobarCero(Trim(ListView12.SelectedItem.SubItems(8))), "N")
    If Im > 0 Then frmTESCobrosDivVto.txtcodigo(1).Text = ListView12.SelectedItem.SubItems(9)
    
    If Text1(0).Text = "" Then
        MsgBox "El cobro no tiene forma de pago. Revise.", vbExclamation
        Exit Sub
    End If
    
    
    frmTESCobrosDivVto.Show vbModal
    
    
    
    
    If CadenaDesdeOtroForm <> "" Then
        CargarTalonesPagaresPendientes
    End If
    
    
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
        Sql = "prov"
        Parametros = "fecrecpr"
    Else
        Parametros = "fecfaccl"
        Sql = ""
    End If
    Sql = Sql & " where " & Parametros & " >= '" & Format(CDate(txtFecha(2).Text), FormatoFecha) & "'"
    Sql = Sql & " AND " & Parametros & " <= '" & Format(CDate(txtFecha(3).Text), FormatoFecha) & "'"
    Set RS = New ADODB.Recordset
    Errores = "select count(*) from cabfact" & Sql
    RS.Open Errores, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Ok = 0
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If RS.Fields(0) > 0 Then Ok = 1
        End If
    End If
    RS.Close
    
    If Ok = 0 Then
        Sql = "Ningun dato a traspasar de facturas "
        If Proveedores Then
            Sql = Sql & "proveedores"
        Else
            Sql = Sql & "clientes"
        End If
        MsgBox Sql, vbExclamation
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
    
    Parametros = "Select distinct(codmacta) from cabfact" & Sql
    RS.Open Parametros, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        InsertaEnTmpCta
        RS.MoveNext
        
    Wend
    RS.Close
    
    Parametros = "Select distinct(Cuereten) from cabfact" & Sql & " and not (cuereten is null)"
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
    '   Codigo: Para clientes será: numserie, codfacl, anofaccl
    
    vOpc = "FACTURAS"
    EncabezadoPieFact False, vOpc, 0
    
    
    If Not Proveedores Then
        Sql = "numserie,codfaccl,anofaccl,fecfaccl,codmacta,confaccl,ba1faccl,ba2faccl,ba3faccl,pi1faccl,pi2faccl,pi3faccl,pr1faccl,pr2faccl,pr3faccl,ti1faccl,ti2faccl,ti3faccl,tr1faccl,tr2faccl,tr3faccl,totfaccl,tp1faccl,tp2faccl,tp3faccl,intracom,retfaccl,trefaccl,cuereten,fecliqcl"
        Parametros = "Select " & Sql & " from cabfact"
        Parametros = Parametros & " where fecfaccl >= '" & Format(CDate(txtFecha(2).Text), FormatoFecha) & "'"
        Parametros = Parametros & " and fecfaccl <= '" & Format(CDate(txtFecha(3).Text), FormatoFecha) & "'"
        Ok = 3
    Else
        Sql = "numregis,anofacpr,fecfacpr,fecrecpr,numfacpr,codmacta,confacpr,ba1facpr,ba2facpr,ba3facpr,pi1facpr,pi2facpr,pi3facpr,pr1facpr,pr2facpr,pr3facpr,ti1facpr,ti2facpr,ti3facpr,tr1facpr,tr2facpr,tr3facpr,totfacpr,tp1facpr,tp2facpr,tp3facpr,extranje,retfacpr,trefacpr,cuereten,fecliqpr,nodeducible"
        Parametros = "Select " & Sql & " from cabfactprov"
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
        Sql = RS.Fields(0) & "|"
        If Not Proveedores Then Sql = Sql & "0" 'meto un 0 para que las facturas que coinciden con el año no den errores
        Sql = Sql & RS.Fields(1) & "|"
        If Not Proveedores Then Sql = Sql & RS.Fields(2) & "|"
        Label40.Caption = Sql
        Label40.Refresh
        'Cadena insert
        BACKUP_Tabla RS, Parametros
        Parametros = Sql & CadenaDesdeOtroForm & Parametros & ";|"
        
        
        'El UPDATE
        Sql = ""
        
        For I = Ok To RS.Fields.Count - 1
            If Sql <> "" Then Sql = Sql & ","
            Sql = Sql & RS.Fields(I).Name & " = "
            If IsNull(RS.Fields(I)) Then
                Sql = Sql & "NULL"
            Else
                Select Case RS.Fields(I).Type
                Case 133
                    Sql = Sql & "'" & Format(RS.Fields(I), FormatoFecha) & "'"
                
                Case 17
                    'numero
                    Sql = Sql & RS.Fields(I)
                    
                Case 131
                    Sql = Sql & TransformaComasPuntos(CStr(RS.Fields(I)))
                Case Else
                    Sql = Sql & "'" & DevNombreSQL(RS.Fields(I)) & "'"
                End Select
                
            End If
        Next I
      
        Sql = Sql & " WHERE "
        Aux = ""
        For I = 0 To Ok - 1
            Aux = Aux & RS.Fields(I).Name & " = '" & RS.Fields(I) & "' and "
        Next
        Aux = Mid(Aux, 1, Len(Aux) - 4)
        Sql = Sql & Aux
        If Not Proveedores Then
            Sql = "UPDATE cabfact SET " & Sql
        Else
            Sql = "UPDATE cabfactprov SET " & Sql
        End If
        Parametros = Parametros & Sql & "|"
        
        
        'Metemos una marca para separar las lineas
        Parametros = Parametros & "<>"
        
        'Las lineas
        '----------------------------
        
        Sql = "Select * from linfact"
        If Proveedores Then Sql = Sql & "prov"
        Sql = Sql & " WHERE " & Aux
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Errores = "" Then
            BACKUP_TablaIzquierda miRsAux, Errores
            Sql = "INSERT INTO linfact"
            If Proveedores Then Sql = Sql & "prov"
            Errores = Sql & "  " & Errores & " VALUES "
        End If
        While Not miRsAux.EOF
            BACKUP_Tabla miRsAux, Sql
            Sql = Errores & Sql & ";"
            Parametros = Parametros & Sql & "|"
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
    
    Errores = "El fichero: " & Text8.Text & " se ha generado con éxito"
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
    
    
    Line Input #NE, Sql   'OPCION
    If Sql <> "[OPCION]" Then
        MsgBox "Formato fichero incorrecto", vbExclamation
        Close NE
        Exit Function
    End If
    Line Input #NE, Sql   ' PRoveedores o clientes
    I = Val(Sql)
    If I = 1 Then
        'CLIENTES
        Clientes = True
    Else
        'PROVEEDORES
        Clientes = False
    End If
    Line Input #NE, Sql   ' Datos vacios
    Line Input #NE, Sql   ' digitos ultimo nivel
    I = Val(Sql)
    If I <> vEmpresa.DigitosUltimoNivel Then
        MsgBox "Ultimo nivel disitinto:" & I, vbExclamation
        Close NE
        Exit Function
    End If
    Line Input #NE, Sql   'FIN OPCION
    
    'CUENTAS
    CadenaDesdeOtroForm = "Cuentas"
    Label40.Caption = CadenaDesdeOtroForm
    Label40.Refresh
    Line Input #NE, Sql   'CUENTAS
    I = 0
    Fin = False
    Do
        Line Input #NE, Sql   'FIN OPCION
        If InStr(1, Sql, "[/CUENTAS]") > 0 Then
            'Fin
            Fin = True
            'Ver numero registros
            Ok = InStr(1, Sql, "]")
            Sql = Mid(Sql, Ok + 1)
            Ok = Val(Sql)
            If I <> Ok Then
            
            End If
        Else
            'Mandamos la linea a ejecutar
            Label40.Caption = "Cta: " & Mid(Sql, 155 + vEmpresa.DigitosUltimoNivel, 30)
            Label40.Refresh
            EjecutarSQL
            I = I + 1
            
        End If
    Loop Until Fin
    
    'CC
    CadenaDesdeOtroForm = "CC"
    Label40.Caption = CadenaDesdeOtroForm
    Label40.Refresh
    Line Input #NE, Sql   'CC
    I = 0
    Fin = False
    Do
        Line Input #NE, Sql   'FIN OPCION
        If InStr(1, Sql, "[/CC]") > 0 Then
            'Fin
            Fin = True
            'Ver numero registros
        Else
            'Mandamos la linea a ejecutar
            EjecutarSQL
            I = I + 1
            
        End If
    Loop Until Fin
    
    
    
    'IVA
    CadenaDesdeOtroForm = "IVA"
    Label40.Caption = CadenaDesdeOtroForm
    Label40.Refresh
    Line Input #NE, Sql   'IVA
    I = 0
    Fin = False
    Do
        Line Input #NE, Sql   'FIN OPCION
        If InStr(1, Sql, "[/IVA]") > 0 Then
            'Fin
            Fin = True
            'Ver numero registros
        Else
            'Mandamos la linea a ejecutar
            EjecutarSQL
            I = I + 1
            
        End If
    Loop Until Fin
    
    
    If Clientes Then
        'SOLO CLIENTES LLEVA CONTADORES
        CadenaDesdeOtroForm = "CONTADORES"
        Label40.Caption = CadenaDesdeOtroForm
        Label40.Refresh
        Line Input #NE, Sql   'CONTADPORES
        I = 0
        Fin = False
        Do
            Line Input #NE, Sql   'FIN OPCION
            If InStr(1, Sql, "[/CONTADORES]") > 0 Then
                'Fin
                Fin = True
                'Ver numero registros
            Else
                'Mandamos la linea a ejecutar
                EjecutarSQL
                I = I + 1
                
            End If
        Loop Until Fin
    End If
    
    
    
    
    'FACTURAS
    Set RS = New ADODB.Recordset
    CadenaDesdeOtroForm = "FACTURAS"
    Label40.Caption = CadenaDesdeOtroForm
    Label40.Refresh
    Line Input #NE, Sql   'FACTS
    I = 0
    Fin = False
    Do
        Line Input #NE, Sql   'FIN OPCION
        If InStr(1, Sql, "[/FACT") > 0 Then
            'Fin
            Fin = True
            'Ver numero registros
            Ok = InStr(1, Sql, "]")
            Ok = Val(Mid(Sql, Ok + 1))
            If Ok <> I Then MsgBox "Diferencia entre facturas procesadas. Fichero: " & Ok & " -> " & I, vbExclamation
        Else
            'Mandamos la linea a ejecutar
            ProcesarLineaFactura Clientes
            I = I + 1
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
        I = 1
    Else
        I = 0
    End If
    Close NE
    If I = 0 Then
        If Me.chkImportarFacturas.Value = 1 Then Kill Text8.Text
    End If
    Set RS = Nothing
End Function


Private Sub EjecutarSQL()
    On Error Resume Next
    
    Conn.Execute Sql
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
Dim Año As Integer
Dim numero As Long
Dim Serie As String
Dim J As Long
Dim Aux As String

    If Clientes Then
        Serie = RecuperaValor(Sql, 1)
        numero = RecuperaValor(Sql, 2)
        Año = RecuperaValor(Sql, 3)
    Else
        Serie = ""
        numero = RecuperaValor(Sql, 1)
        Año = RecuperaValor(Sql, 2)
    End If
    
        
    Label40.Caption = Serie & " " & numero & " / " & Año
    Label40.Refresh
    DoEvents
    'Quitamos el La cadaena
    J = InStr(2, Sql, "|" & Año & "|")
    If J = 0 Then
        MsgBox "Error en año factura", vbExclamation
        Exit Sub
    End If
    
    J = J + 6
    Sql = Mid(Sql, J)
    
    If Clientes Then
        Aux = "Select * from cabfact WHERE numserie = '" & Serie & "'"
        Aux = Aux & " and anofaccl = " & Año & " and codfaccl =" & numero
    Else
        Aux = "Select * from cabfactprov WHERE "
        Aux = Aux & " anofacpr = " & Año & " and numregis =" & numero
    End If
    RS.Open Aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    J = 1
    If Not RS.EOF Then
        J = 2
        'Borro las lineas
        If Clientes Then
            Aux = "DELETE from linfact WHERE numserie = '" & Serie & "'"
            Aux = Aux & " and anofaccl = " & Año & " and codfaccl =" & numero
        Else
            Aux = "DELETE from linfactprov WHERE "
            Aux = Aux & " anofacpr = " & Año & " and numregis =" & numero
        End If
        Conn.Execute Aux
    End If
    RS.Close
    
    Aux = RecuperaValor(Sql, CInt(J))
    Conn.Execute Aux
    
    '---------------------------------------------------
    J = InStr(1, Sql, "<>")
    If J = 0 Then
        MsgBox "Error lineas: " & Aux, vbExclamation
        Exit Sub
    End If
        
    Sql = Mid(Sql, J + 2)
    Do
        J = InStr(1, Sql, "|")
        If J > 0 Then
            Aux = Mid(Sql, 1, J - 1)
            Sql = Mid(Sql, J + 1)
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
        Sql = RecuperaValor(Errores, 1)
    Else
        Sql = ""
    End If
        
    If Sql = "" Then
        MsgBox "Error en fichero. Linea vacia o sin año importacion." & vbCrLf & Sql, vbExclamation
        Exit Function
    End If


    If Val(Sql) = 0 Then
        MsgBox "Año incorrecto: " & Sql & vbCrLf & Errores, vbExclamation
        Exit Function
    End If


    Sql = "DELETE FROM datosext347 where año =" & Sql
    Conn.Execute Sql
    
    
    
    'Volvemos a abrir el fichero
    NE = FreeFile
    Open Text9(0).Text For Input As #NE
    I = 0
    
    
    While Not EOF(NE)
        Line Input #NE, Errores
            
        Sql = RecuperaValor(Errores, 1)
        Parametros = Trim(RecuperaValor(Errores, 2))
        If Parametros = "1" Then
            Ok = 1
        Else
            Ok = 2
        End If
        Sql = Sql & ",'" & Text9(Ok).Text & "'"
        For Ok = 3 To 8
            Parametros = RevisaCaracterMultibase(Trim(RecuperaValor(Errores, Ok)))
            Parametros = DevNombreSQL(Parametros)
            Sql = Sql & ",'" & Parametros & "'"
        Next Ok
        
        'El importe
        Parametros = TransformaComasPuntos((RecuperaValor(Errores, 9)))
        Sql = Sql & "," & Parametros & ")"
        Sql = "INSERT INTO datosext347 (año, letra, nif, nombre, direc, codposta, poblacion, provincia, importe) VALUES (" & Sql
        Conn.Execute Sql
        I = I + 1
    Wend
    Close #NE

    
    
    If I > 0 Then
        ImportarDatosExternos347 = True
        MsgBox "Proceso finalizado.   " & I & " registros insertados", vbInformation
    Else
        MsgBox "No se han importado datos", vbExclamation
    End If
    Exit Function
EImportarDatosExternos347:
    MuestraError Err.Number, Sql
    On Error Resume Next
        Close #NE
        Err.Clear
End Function


Private Sub cargarObservacionesCuenta()
    Set RS = New ADODB.Recordset
    Sql = "select codmacta,nommacta,obsdatos from cuentas where codmacta = '" & Parametros & "'"
    RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        Sql = RS!codmacta & "|" & RS!Nommacta & "|" & DBMemo(RS!obsdatos) & "|"
    Else
        Sql = "Err|ERROR  LEYENDO DATOS CUENTAS | ****  ERROR ****|"
    End If
    RS.Close
    Set RS = Nothing
    For I = 1 To 3
        Text1(I + 3).Text = RecuperaValor(Sql, I)
    Next I
End Sub


Private Sub cargaempresasbloquedas()
Dim IT As ListItem
    On Error GoTo Ecargaempresasbloquedas
    Set RS = New ADODB.Recordset
    Sql = "select empresasariconta.codempre,nomempre,nomresum,usuarioempresasariconta.codempre bloqueada from usuarios.empresasariconta left join usuarios.usuarioempresasariconta on "
    Sql = Sql & " empresasariconta.codempre = usuarioempresasariconta.codempre And (usuarioempresasariconta.codusu = " & Parametros & " Or codusu Is Null)"
    '[Monica] solo ariconta
    Sql = Sql & " WHERE conta like 'ariconta%' "
    Sql = Sql & " ORDER BY empresasariconta.codempre"
    RS.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        Errores = Format(RS!codempre, "00000")
        Sql = "C" & Errores
        
        If IsNull(RS!bloqueada) Then
            'Va al list de la derecha
            Set IT = ListView2(0).ListItems.Add(, Sql)
            IT.SmallIcon = 1
        Else
            Set IT = ListView2(1).ListItems.Add(, Sql)
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
    
    Sql = "select menus.codigo, menus.imagen, menus.descripcion, menus_usuarios.posx, menus_usuarios.posy, menus_usuarios.vericono "
    Sql = Sql & " from menus inner join menus_usuarios on menus.codigo = menus_usuarios.codigo and menus.aplicacion = menus_usuarios.aplicacion"
    Sql = Sql & " WHERE menus.aplicacion = 'ariconta' and menus_usuarios.codusu = " & vUsu.Id
    Sql = Sql & " and menus.imagen <> 0 " ' si tiene imagen puede estar en el listview para seleccionar
    Sql = Sql & " and menus_usuarios.ver = 1 "
    Sql = Sql & " ORDER BY menus.codigo "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView6.SmallIcons = frmPpal.ImageListPpal16
    
    ListView6.ColumnHeaders.Clear
    
    ListView6.ColumnHeaders.Add , , "Código", 1800.0631
    ListView6.ColumnHeaders.Add , , "Descripción", 4200.2522, 0
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
    
    Sql = "select * from tmpinfbbdd where codusu = " & vUsu.Codigo & " order by posicion "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
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
Dim cad As String
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
    
    cad = "show full processlist"
    RS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
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


Private Function CobroContabilizado(Serie As String, FACTURA As String, Fecha As String, Orden As String) As Boolean
Dim Sql As String

    Sql = "select * from hlinapu where numserie = " & DBSet(Serie, "T") & " and numfaccl = " & DBSet(FACTURA, "N") & " and fecfactu = " & DBSet(Fecha, "F") & " and numorden = " & DBSet(Orden, "N")
    CobroContabilizado = (TotalRegistrosConsulta(Sql) <> 0)

End Function

Private Function PagoContabilizado(Serie As String, Proveedor As String, FACTURA As String, Fecha As String, Orden As String) As Boolean
Dim Sql As String

    Sql = "select * from hlinapu where numserie = " & DBSet(Serie, "T") & " and codmacta = " & DBSet(Proveedor, "T") & " and numfacpr = " & DBSet(FACTURA, "T") & " and fecfactu = " & DBSet(Fecha, "F") & " and numorden = " & DBSet(Orden, "N")
    PagoContabilizado = (TotalRegistrosConsulta(Sql) <> 0)

End Function



Private Sub CargaCobrosFactura()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim cad As String
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
    
    cad = "select numorden, formapago.nomforpa, fecvenci, impvenci, gastos, fecultco, impcobro, (coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0)) pendiente, cobros.ctabanc1  "
    cad = cad & " from (cobros left join formapago on cobros.codforpa = formapago.codforpa) "
    cad = cad & " where cobros.numserie = " & DBSet(RecuperaValor(Parametros, 1), "T")
    cad = cad & " and cobros.numfactu = " & DBSet(RecuperaValor(Parametros, 2), "N")
    cad = cad & " and cobros.fecfactu = " & DBSet(RecuperaValor(Parametros, 3), "F")
    cad = cad & " order by numorden "
    
    RS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
    While Not RS.EOF
                    
        Set IT = ListView5.ListItems.Add
        
        If CobroContabilizado(RecuperaValor(Parametros, 1), RecuperaValor(Parametros, 2), RecuperaValor(Parametros, 3), DBLet(RS.Fields(0))) Then IT.SmallIcon = 18
        
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
Dim cad As String
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
    
    cad = "select numorden, formapago.nomforpa, fecefect, impefect, fecultpa, imppagad, (coalesce(impefect,0)  - coalesce(imppagad,0)) pendiente, pagos.ctabanc1  "
    cad = cad & " from (pagos left join formapago on pagos.codforpa = formapago.codforpa) "
    cad = cad & " where pagos.numserie = " & DBSet(RecuperaValor(Parametros, 1), "T")
    cad = cad & " and pagos.codmacta = " & DBSet(RecuperaValor(Parametros, 2), "T")
    cad = cad & " and pagos.numfactu = " & DBSet(RecuperaValor(Parametros, 3), "T")
    cad = cad & " and pagos.fecfactu = " & DBSet(RecuperaValor(Parametros, 4), "F")
    cad = cad & " order by numorden "
    
    RS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
    While Not RS.EOF
                    
        Set IT = ListView5.ListItems.Add
        
        If PagoContabilizado(RecuperaValor(Parametros, 1), RecuperaValor(Parametros, 2), RecuperaValor(Parametros, 3), RecuperaValor(Parametros, 4), DBLet(RS.Fields(0))) Then IT.SmallIcon = 18

'        If DBLet(RS!NumAsien, "N") <> 0 Then IT.SmallIcon = 18
        
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
Dim cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarAsiento
    
    
    ListView7.ColumnHeaders.Clear
    ListView7.ListItems.Clear
    
    
'    ListView7.ColumnHeaders.Add , , "Ord.", 800.0631
    ListView7.ColumnHeaders.Add , , "Cuenta", 1500.2522
    ListView7.ColumnHeaders.Add , , "Denominación", 4000.2522
    ListView7.ColumnHeaders.Add , , "Debe", 2050.2522, 1
    ListView7.ColumnHeaders.Add , , "Haber", 2050.2522, 1
    ListView7.ColumnHeaders.Add , , "Saldo", 2050.2522, 1
    
    Set RS = New ADODB.Recordset
    
    
    Pos = DevuelveValor("select max(pos) from tmpconext where codusu = " & DBSet(vUsu.Codigo, "N"))
    
    
    cad = "select tmpconext.cta, cuentas.nommacta, tmpconext.timported, tmpconext.timporteh, acumtotT, tmpconext.pos"
    cad = cad & " from (ariconta" & NumConta & ".tmpconext inner join ariconta" & NumConta & ".cuentas on tmpconext.cta = cuentas.codmacta) "
    cad = cad & " left join ariconta" & NumConta & ".tmpconextcab on tmpconext.codusu = tmpconextcab.codusu and tmpconext.cta = tmpconextcab.cta"
    cad = cad & " where tmpconext.codusu = " & DBSet(vUsu.Codigo, "N")
    cad = cad & " order by pos "
    
    RS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
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
        
        ' si no estamos en la última línea mostramos el saldo de la cuenta
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
Dim cad As String
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
    
    
    cad = "select tmphistoapu.numdiari, tmphistoapu.numasien, tmphistoapu.fechaent, tmphistoapu.timported, tmphistoapu.timporteh, tmphistoapu.timported - tmphistoapu.timporteh "
    cad = cad & " from tmphistoapu "
    cad = cad & " where tmphistoapu.codusu = " & DBSet(vUsu.Codigo, "N")
    cad = cad & " order by numdiari, numasien, fechaent "
    
    RS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
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
Dim cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarAsiento
    
    
    ListView8.ColumnHeaders.Clear
    ListView8.ListItems.Clear
    
    
    ListView8.ColumnHeaders.Add , , "Serie", 800.2522
    ListView8.ColumnHeaders.Add , , "Descripción", 2700.2522
    ListView8.ColumnHeaders.Add , , "Factura", 1500.2522
    ListView8.ColumnHeaders.Add , , "Fecha", 1500.2522
    ListView8.ColumnHeaders.Add , , "Total", 2000.2522, 1
    
    Set RS = New ADODB.Recordset
    
    cad = "select tmpfaclin.numserie, tmpfaclin.nomserie, tmpfaclin.numfac, tmpfaclin.fecha, tmpfaclin.total "
    cad = cad & " from tmpfaclin "
    cad = cad & " where tmpfaclin.codusu = " & DBSet(vUsu.Codigo, "N")
    cad = cad & " order by numserie, numfac, fecha "
    
    RS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
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
Dim cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarFacturas
    
    Set ListView9.SmallIcons = frmPpal.imgListComun16
    
    ListView9.ColumnHeaders.Clear
    ListView9.ListItems.Clear
    
    
    ListView9.ColumnHeaders.Add , , "Serie", 800.2522
    ListView9.ColumnHeaders.Add , , "Factura", 2000.2522
    ListView9.ColumnHeaders.Add , , "Fecha", 2000.2522
    ListView9.ColumnHeaders.Add , , "Vto", 1500.2522
    ListView9.ColumnHeaders.Add , , "Importe", 2000.2522, 1
    
    Set RS = New ADODB.Recordset
    
    cad = "select numserie, numfactu, fecfactu, numorden, impvenci importe "
    cad = cad & " from reclama_facturas "
    cad = cad & " where codigo = " & DBSet(RecuperaValor(Parametros, 1), "N")
    
    If CampoOrden = "" Then CampoOrden = "fecfactu"
    cad = cad & " ORDER BY " & CampoOrden
    If Orden Then cad = cad & " DESC"
    
'    Cad = Cad & " order by numserie, numfactu, fecfactu "
    
    RS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
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
        
        Sql = "select devuelto from cobros where numserie = " & DBSet(RS.Fields(0), "T") & " and numfactu = " & DBSet(RS.Fields(1), "N")
        Sql = Sql & " and fecfactu = " & DBSet(RS.Fields(2), "F") & " and numorden = " & DBSet(RS.Fields(3), "N")
        
        If DevuelveValor(Sql) = 1 Then
            IT.SmallIcon = 42
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
Dim cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarFacturas
    
    Set ListView9.SmallIcons = frmPpal.imgListComun16
    
    ListView9.ColumnHeaders.Clear
    ListView9.ListItems.Clear
    
    
    ListView9.ColumnHeaders.Add , , "Serie", 800.2522
    ListView9.ColumnHeaders.Add , , "Factura", 1500.2522
    ListView9.ColumnHeaders.Add , , "Fecha", 1500.2522
    ListView9.ColumnHeaders.Add , , "Vto", 700.2522
    ListView9.ColumnHeaders.Add , , "Fecha Vto", 1500.2522
    ListView9.ColumnHeaders.Add , , "Gastos", 1000.2522, 1
    ListView9.ColumnHeaders.Add , , "Importe", 2000.2522, 1
    
    Set RS = New ADODB.Recordset
    
    cad = "select cobros.numserie, cobros.numfactu, cobros.fecfactu, cobros.numorden, cobros.fecvenci, cobros.gastos, cobros.impvenci  importe, ' ' devol"
    cad = cad & " from cobros "
    cad = cad & " where (cobros.codrem = " & DBSet(RecuperaValor(Parametros, 1), "N") & " and cobros.anyorem = " & DBSet(RecuperaValor(Parametros, 2), "N") & ") "
    cad = cad & " union "
    cad = cad & " select hlinapu.numserie, hlinapu.numfaccl, hlinapu.fecfactu, hlinapu.numorden, cobros.fecvenci, hlinapu.gastodev, coalesce(hlinapu.timporteh,0) - coalesce(hlinapu.timported,0) importe, '*' devol"
    cad = cad & " from cobros inner join hlinapu on cobros.numserie = hlinapu.numserie and cobros.numfactu = hlinapu.numfaccl and cobros.fecfactu = hlinapu.fecfactu and cobros.numorden = hlinapu.numorden "
    cad = cad & " where (hlinapu.codrem = " & DBSet(RecuperaValor(Parametros, 1), "N") & " and hlinapu.anyorem = " & DBSet(RecuperaValor(Parametros, 2), "N") & " and hlinapu.esdevolucion = 0) "
    
    If CampoOrden = "" Then CampoOrden = "fecfactu"
    cad = cad & " ORDER BY " & CampoOrden
    If Orden Then cad = cad & " DESC"
    
    
    RS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
    While Not RS.EOF
                    
        Set IT = ListView9.ListItems.Add
        
        IT.Text = DBLet(RS.Fields(0))
        IT.SubItems(1) = DBLet(RS.Fields(1))
        IT.SubItems(2) = DBLet(RS.Fields(2))
        IT.SubItems(3) = DBLet(RS.Fields(3))
        IT.SubItems(4) = DBLet(RS.Fields(4))
        
        'gastos
        If DBLet(RS.Fields(5), "N") <> 0 Then
            IT.SubItems(5) = Format(DBLet(RS.Fields(5)), "###,###,##0.00")
        Else
            IT.SubItems(5) = " "
        End If
        
        'importe
        If DBLet(RS.Fields(6), "N") <> 0 Then
            IT.SubItems(6) = Format(DBLet(RS.Fields(6)), "###,###,##0.00")
        Else
            IT.SubItems(6) = " "
        End If
        
        'Siguiente
        If RS.Fields(6) = "*" Then
            IT.SmallIcon = 42
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



Private Sub CargarBancosRemesas()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarBancosRemesas
    
    Set ListView10.SmallIcons = frmPpal.imgListComun16
    
    ListView10.ColumnHeaders.Clear
    ListView10.ListItems.Clear
    
    
    ListView10.ColumnHeaders.Add , , "Banco", 1900.2522
    ListView10.ColumnHeaders.Add , , "Nombre", 3600.2522
    ListView10.ColumnHeaders.Add , , "Importe", 2000.2522, 1
    
    Set RS = New ADODB.Recordset
    
    cad = "select cta, nomcta, acumperd from tmpcierre1 where codusu = " & vUsu.Codigo
    cad = cad & " ORDER BY 1 "
    
    
    RS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
    While Not RS.EOF
                    
        Set IT = ListView10.ListItems.Add
        
        IT.Text = DBLet(RS.Fields(0))
        IT.SubItems(1) = DBLet(RS.Fields(1))
        
        'importe
        If DBLet(RS.Fields(2), "N") <> 0 Then
            IT.SubItems(2) = Format(DBLet(RS.Fields(2)), "###,###,##0.00")
        Else
            IT.SubItems(2) = " "
        End If
        
        IT.Checked = True
        
        RS.MoveNext
    Wend
    NumRegElim = 0
    RS.Close
    Set RS = Nothing
    
    Exit Sub
    
ECargarBancosRemesas:
    MuestraError Err.Number, Err.Description
    Errores = ""
    Set RS = Nothing
End Sub


Private Sub CargarRecibosConCobrosParciales()
Dim IT As ListItem
Dim TotalArray  As Long
Dim SERVER As String
Dim EquipoConBD As Boolean
Dim cad As String
Dim Equipo As String
Dim Pos As Long

    On Error GoTo ECargarRecibosConCobrosParciales
    
    Set ListView11.SmallIcons = frmPpal.imgListComun16
    
    ListView11.ColumnHeaders.Clear
    ListView11.ListItems.Clear
    
    
    ListView11.ColumnHeaders.Add , , "Serie", 800.2522
    ListView11.ColumnHeaders.Add , , "Factura", 1300.2522
    ListView11.ColumnHeaders.Add , , "Fecha", 1500.2522, 1
    ListView11.ColumnHeaders.Add , , "Vto", 900.2522, 1
    ListView11.ColumnHeaders.Add , , "Importe Vto", 1500.2522, 1
    ListView11.ColumnHeaders.Add , , "Cobrado", 1500.2522, 1
    
    
    Set RS = New ADODB.Recordset
    
    ' le hemos pasado el select completo de cobros
    cad = Parametros
    
    
    RS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
    While Not RS.EOF
                    
        Set IT = ListView11.ListItems.Add
        
        IT.Text = DBLet(RS!NUmSerie)
        IT.SubItems(1) = DBLet(RS!NumFactu)
        IT.SubItems(2) = DBLet(RS!FecFactu)
        IT.SubItems(3) = DBLet(RS!numorden)
        
        
        'importe
        If DBLet(RS!ImpVenci, "N") <> 0 Then
            IT.SubItems(4) = Format(DBLet(RS!ImpVenci), "###,###,##0.00")
        Else
            IT.SubItems(4) = " "
        End If
        
        'importe cobrado
        If DBLet(RS!impcobro, "N") <> 0 Then
            IT.SubItems(5) = Format(DBLet(RS!impcobro), "###,###,##0.00")
        Else
            IT.SubItems(5) = " "
        End If
        
        RS.MoveNext
    Wend
    NumRegElim = 0
    RS.Close
    Set RS = Nothing
    
    Exit Sub
    
ECargarRecibosConCobrosParciales:
    MuestraError Err.Number, "Carga Recibos con cobros parciales " & Err.Description
    Errores = ""
    Set RS = Nothing
End Sub




Private Sub CargarTalonesPagaresPendientes()
Dim IT As ListItem
Dim cad As String
Dim Pendiente As Currency

    On Error GoTo ECargarTalonesPagaresPendientes
    
    Set ListView12.SmallIcons = frmPpal.imgListComun16
    
    ListView12.ColumnHeaders.Clear
    ListView12.ListItems.Clear
    
    
    ListView12.ColumnHeaders.Add , , "Serie", 800.2522
    ListView12.ColumnHeaders.Add , , "Factura", 1300.2522
    ListView12.ColumnHeaders.Add , , "Fecha", 1500.2522, 1
    ListView12.ColumnHeaders.Add , , "Fec.Vto", 1500.2522, 1
    ListView12.ColumnHeaders.Add , , "Vto", 900.2522, 1
    ListView12.ColumnHeaders.Add , , "Tipo", 900.2522, 1
    
    ListView12.ColumnHeaders.Add , , "Importe", 1500.2522, 1
    ListView12.ColumnHeaders.Add , , "Gasto", 1500.2522, 1
    ListView12.ColumnHeaders.Add , , "Cobrado", 1500.2522, 1
    ListView12.ColumnHeaders.Add , , "Pendiente", 1500.2522, 1
    
    
    Set RS = New ADODB.Recordset
    
    ' le hemos pasado el select completo de cobros
    cad = "SELECT cobros.*, formapago.nomforpa, tipofpago.descformapago, tipofpago.siglas, "
    cad = cad & " cobros.nomclien nommacta,cobros.codmacta,tipofpago.tipoformapago, 0 aaa "
    cad = cad & " FROM ((cobros INNER JOIN formapago ON cobros.codforpa = formapago.codforpa) "
    cad = cad & " INNER JOIN tipofpago ON formapago.tipforpa = tipofpago.tipoformapago) "
    
    If Parametros <> "" Then cad = cad & " WHERE " & Parametros
    
    cad = cad & " Union "
    cad = cad & " SELECT cobros.*, formapago.nomforpa, tipofpago.descformapago, tipofpago.siglas, "
    cad = cad & " cobros.nomclien nommacta,cobros.codmacta,tipofpago.tipoformapago, 1 aaa "
    cad = cad & " FROM ((cobros INNER JOIN formapago ON cobros.codforpa = formapago.codforpa) "
    cad = cad & " INNER JOIN tipofpago ON formapago.tipforpa = tipofpago.tipoformapago) "
    cad = cad & " where (numserie, numfactu, fecfactu, numorden) in (select numserie, numfactu, fecfactu, numorden from talones_facturas where codigo = " & DBSet(Codigo, "N") & ")"
    
    RS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
    While Not RS.EOF
                    
        Set IT = ListView12.ListItems.Add
        
        IT.Text = DBLet(RS!NUmSerie)
        IT.SubItems(1) = DBLet(RS!NumFactu)
        IT.SubItems(2) = DBLet(RS!FecFactu)
        IT.SubItems(3) = DBLet(RS!FecVenci)
        IT.SubItems(4) = DBLet(RS!numorden)
        IT.SubItems(5) = DBLet(RS!descformapago)
        
        
        'importe
        If DBLet(RS!ImpVenci, "N") <> 0 Then
            IT.SubItems(6) = Format(DBLet(RS!ImpVenci), "###,###,##0.00")
        Else
            IT.SubItems(6) = " "
        End If
        
        'importe gasto
        If DBLet(RS!Gastos, "N") <> 0 Then
            IT.SubItems(7) = Format(DBLet(RS!Gastos), "###,###,##0.00")
        Else
            IT.SubItems(7) = " "
        End If
        
        'importe cobrado
        If DBLet(RS!impcobro, "N") <> 0 Then
            IT.SubItems(8) = Format(DBLet(RS!impcobro), "###,###,##0.00")
        Else
            IT.SubItems(8) = " "
        End If
        
        Pendiente = DBLet(RS!ImpVenci) + DBLet(RS!Gastos, "N") - DBLet(RS!impcobro, "N")
        If Pendiente <> 0 Then
            IT.SubItems(9) = Format(Pendiente, "###,###,##0.00")
        Else
            IT.SubItems(9) = " "
        End If
        
        If DBLet(RS!aaa, "N") = 1 Then IT.Checked = True
        
        RS.MoveNext
    Wend
    NumRegElim = 0
    RS.Close
    Set RS = Nothing
    
    Exit Sub
    
ECargarTalonesPagaresPendientes:
    MuestraError Err.Number, "Carga Recibos Talones/Pagarés pendientes" & Err.Description
    Errores = ""
    Set RS = Nothing
End Sub



Private Function SumaTotales() As Boolean
Dim Sql As String
Dim Suma As Currency
    
    Suma = 0
    txtSuma.Text = ""
    
    For I = 1 To Me.ListView12.ListItems.Count
        If Me.ListView12.ListItems(I).Checked Then
            Suma = Suma + ComprobarCero(Trim(Me.ListView12.ListItems(I).SubItems(9)))
        End If
    Next I
    
    If Suma <> 0 Then txtSuma.Text = Format(Suma, "###,###,##0.00")

End Function
