VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESRealizarCobros 
   Caption         =   "Realizar Cobro"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   17685
   Icon            =   "frmTESRealizarCobros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   17685
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   750
      Top             =   2100
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
            Picture         =   "frmTESRealizarCobros.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTESRealizarCobros.frx":686E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTESRealizarCobros.frx":6B88
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frame 
      Height          =   2325
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   17535
      Begin VB.Frame FrameRemesar 
         BorderStyle     =   0  'None
         Height          =   1905
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   17265
         Begin VB.ComboBox Combo2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmTESRealizarCobros.frx":6EA2
            Left            =   11460
            List            =   "frmTESRealizarCobros.frx":6EA4
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Tag             =   "Tipo de pago|N|N|||formapago|tipforpa|||"
            Top             =   330
            Width           =   2265
         End
         Begin VB.CheckBox chkSoloBancoPrev 
            Caption         =   "Sólo Banco previsto"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11460
            TabIndex        =   35
            Top             =   1200
            Width           =   2415
         End
         Begin VB.CheckBox chkImprimir 
            Caption         =   "Imprimir Recibos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11460
            TabIndex        =   32
            Top             =   900
            Width           =   2175
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   6
            Left            =   10020
            TabIndex        =   30
            Text            =   "0000000000"
            Top             =   930
            Width           =   1305
         End
         Begin VB.CheckBox chkVerPendiente 
            Caption         =   "Ver lo Pdte del cliente"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8520
            TabIndex        =   29
            Top             =   1560
            Width           =   2775
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmTESRealizarCobros.frx":6EA6
            Left            =   8520
            List            =   "frmTESRealizarCobros.frx":6EA8
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Tag             =   "Tipo de pago|N|N|||formapago|tipforpa|||"
            Top             =   330
            Width           =   2835
         End
         Begin VB.TextBox txtCta 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   5
            Left            =   1260
            TabIndex        =   5
            Text            =   "0000000000"
            Top             =   1470
            Width           =   1305
         End
         Begin VB.TextBox txtDCta 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   2670
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "Text2"
            Top             =   930
            Width           =   5715
         End
         Begin VB.TextBox txtCta 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   1260
            TabIndex        =   4
            Text            =   "0000000000"
            Top             =   930
            Width           =   1305
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   4230
            TabIndex        =   1
            Text            =   "0000000000"
            Top             =   330
            Width           =   1305
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   7050
            TabIndex        =   2
            Text            =   "0000000000"
            Top             =   330
            Width           =   1305
         End
         Begin VB.CheckBox chkVtoCuenta 
            Caption         =   "Agrupar por Cliente"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   13920
            TabIndex        =   23
            Top             =   600
            Width           =   2745
         End
         Begin VB.CheckBox chkGenerico 
            Caption         =   "Cuenta genérica"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   13920
            TabIndex        =   22
            Top             =   1500
            Width           =   2175
         End
         Begin VB.CheckBox chkPorFechaVenci 
            Caption         =   "Contab. por fecha vto."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   13920
            TabIndex        =   21
            Top             =   1200
            Width           =   2865
         End
         Begin VB.CheckBox chkContrapar 
            Caption         =   "Agrupar por Banco"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   13920
            TabIndex        =   20
            Top             =   900
            Width           =   2745
         End
         Begin VB.CheckBox chkAsiento 
            Caption         =   "Un Asiento por recibo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   13920
            TabIndex        =   19
            Top             =   300
            Width           =   2385
         End
         Begin VB.TextBox txtDCta 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   5
            Left            =   2670
            TabIndex        =   14
            Text            =   "Text3"
            Top             =   1470
            Width           =   5745
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   1260
            TabIndex        =   0
            Text            =   "Text3"
            Top             =   330
            Width           =   1365
         End
         Begin MSComctlLib.Toolbar ToolbarAyuda 
            Height          =   390
            Left            =   16770
            TabIndex        =   33
            Top             =   30
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Ayuda"
               EndProperty
            EndProperty
         End
         Begin VB.Label Label2 
            Caption         =   "Usuario"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   11520
            TabIndex        =   37
            Top             =   60
            Width           =   2025
         End
         Begin VB.Label Label2 
            Caption         =   "Gastos banco"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   8580
            TabIndex        =   31
            Top             =   960
            Width           =   1425
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de Pago"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   8580
            TabIndex        =   28
            Top             =   60
            Width           =   2025
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   1
            Left            =   960
            Top             =   1530
            Width           =   240
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   0
            Left            =   960
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Banco"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   60
            TabIndex        =   27
            Top             =   930
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Desde Vto"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   2850
            TabIndex        =   25
            Top             =   330
            Width           =   1065
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta Vto"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   5640
            TabIndex        =   24
            Top             =   330
            Width           =   1125
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   3930
            Picture         =   "frmTESRealizarCobros.frx":6EAA
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   6780
            Picture         =   "frmTESRealizarCobros.frx":6F35
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   0
            Left            =   960
            Picture         =   "frmTESRealizarCobros.frx":6FC0
            ToolTipText     =   "Cambiar fecha contabilizacion"
            Top             =   330
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   1
            Left            =   16410
            ToolTipText     =   "AYUDA"
            Top             =   90
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   16800
            Picture         =   "frmTESRealizarCobros.frx":7302
            ToolTipText     =   "Seleccionar todos"
            Top             =   1590
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   16410
            Picture         =   "frmTESRealizarCobros.frx":744C
            ToolTipText     =   "Quitar seleccion"
            Top             =   1590
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   16
            Top             =   330
            Width           =   585
         End
         Begin VB.Label Label3 
            Caption         =   "Cliente"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   15
            Top             =   1500
            Width           =   885
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   7740
      Width           =   17085
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Contabilizar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   34
         Top             =   120
         Width           =   1425
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   5400
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   8580
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   12060
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Seleccionado"
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
         Height          =   285
         Index           =   2
         Left            =   3780
         TabIndex        =   18
         Top             =   180
         Width           =   1560
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   " PENDIENTE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Index           =   1
         Left            =   10650
         TabIndex        =   11
         Top             =   180
         Width           =   1290
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Vencido"
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
         Height          =   195
         Index           =   0
         Left            =   7530
         TabIndex        =   9
         Top             =   180
         Width           =   990
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5025
      Left            =   30
      TabIndex        =   7
      Top             =   2520
      Width           =   17475
      _ExtentX        =   30824
      _ExtentY        =   8864
      View            =   3
      LabelEdit       =   1
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
   Begin VB.Menu mnContextual 
      Caption         =   "Contextual"
      Visible         =   0   'False
      Begin VB.Menu mnNumero 
         Caption         =   "Poner numero Talón/Pagaré"
      End
      Begin VB.Menu mnbarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSelectAll 
         Caption         =   "Seleccionar todos"
      End
      Begin VB.Menu mnQUitarSel 
         Caption         =   "Quitar selección"
      End
   End
End
Attribute VB_Name = "frmTESRealizarCobros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 604



Public vSQL2 As String
Public Cobros As Boolean
Public OrdenarEfecto As Boolean
Public Regresar As Boolean
Public vTextos As String  'Dependera de donde venga
Public SegundoParametro As String
Public ContabTransfer As Boolean
Public VieneDesdeNorma57 As Boolean

    'Diversas utilidades
    '-------------------------------------------------------------------------------
    'Para las transferencias me dice que transferencia esta siendo creada/modificada
    '
    'Para mostrar un check con los efectos k se van a generar en remesa y/o pagar
 
 
 ' 13 Mayo 08
    ' Cuando contabilice el los cobros por tarjeta entonces
    ' si lleva gastos los añadire
Public ImporteGastosTarjeta_ As Currency   'Para cuando viene de recepciondocumentos pondre el importe que le falta
                                          ' y asi ofertarlo al divisonvencimiento
     '-ABRIL 2014.  Navarres. Llevara el % interes
 
 
 
 
'Agosto 2009
'Desde recepcion de talones.
'Tendra la posibilidad de desdoblar un vencimiento
Public DesdeRecepcionTalones As Boolean
 
'Febrero 2010
'Para el pago de talones y pagareses ;)
'Enviara el nº de talon/pagare
Public NumeroTalonPagere As String


'Marzo 2013
'Cuando cobro/pago un mismo clie/prov aparecera un icono para poder añadir
'cualquier cobro /pago del mismo. Se contabilizaran con los datos pendientes
Public CodmactaUnica As String

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
Private WithEvents frmBan As frmBasico2
Attribute frmBan.VB_VarHelpID = -1

Dim cad As String
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim Fecha As Date
Dim Importe As Currency
Dim Vencido As Currency
Dim impo As Currency
Dim riesgo As Currency

Dim ImpSeleccionado As Currency
Dim I As Integer
Private PrimeraVez As Boolean
Dim RiesTalPag As Currency
Private FechaAsiento As Date
Private vp As Ctipoformapago
Private SubItemVto As Integer

Private DescripcionTransferencia As String
Private GastosTransferencia As Currency



Dim CampoOrden As String
Dim Orden As Boolean
Dim Campo2 As Integer
Dim CtaAnt As String
Dim CtaAntBan As String

Dim FrasMarcadas As Collection


Private Sub chkAsiento_Click(Index As Integer)
   'Es incompatible asiento por pago y agrupar apunte bancario
   If chkAsiento(Index).Value = 1 Then
        If chkContrapar(Index).Value = 1 Then
            Incompatibilidad
            chkContrapar(Index).Value = 0
        End If
   End If
       
End Sub

Private Sub Incompatibilidad()
    If Not PrimeraVez Then MsgBox "Es incompatible agrupar apunte bancario y asiento por pago", vbExclamation
End Sub


Private Sub chkContrapar_Click(Index As Integer)
   'Es incompatible asiento por pago y agrupar apunte bancario
   If chkContrapar(Index).Value = 1 Then
        If chkAsiento(Index).Value = 1 Then
            Incompatibilidad
            chkAsiento(Index).Value = 0
        End If
   End If
End Sub

Private Sub chkGenerico_Click(Index As Integer)
    If chkGenerico(Index).Value = 0 Then
        chkGenerico(Index).FontBold = False
        chkGenerico(Index).Tag = ""
    End If
End Sub



Private Sub Generar2()
Dim Contador2 As Integer
Dim F2 As Date
Dim TipoAnt As Integer
    
    cad = ""
    For I = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
            cad = cad & "1"
            Exit For
        End If
    Next I
    If cad = "" Then
        MsgBox "Deberias selecionar algún vencimiento", vbExclamation
        Exit Sub
    End If
    
    If Combo1.ListIndex = -1 Then
        MsgBox "Deberias selecionar el tipo de pago", vbExclamation
        Exit Sub
    End If
         
    If txtCta(4).Text = "" Then
        MsgBox "Deberias introducir la cuenta de banco", vbExclamation
        Exit Sub
    End If
    
    Importe = 0
    If Combo1.ItemData(Combo1.ListIndex) = 6 Then
        If Text3(6).Text <> "" Then
            If InStr(1, Text3(6).Text, ",") > 0 Then
                Importe = ImporteFormateado(Text3(6).Text)
            Else
                Importe = CCur(TransformaPuntosComas(Text3(6).Text))
            End If
        End If
    End If
    If vParamT.IntereseCobrosTarjeta > 0 Then
        If Importe < 0 Or Importe >= 100 Then
            MsgBox "Intereses cobro tarjeta. Valor entre 0..100", vbExclamation
            PonFoco Me.Text3(6)
            Exit Sub
            
        End If
        
        'Solo dejaremos IR cliente a cliente
        If Me.txtCta(5).Text = "" And Importe > 0 Then
            MsgBox "Seleccione una cuenta cliente", vbExclamation
            PonFoco Me.txtCta(5)
            Exit Sub
        End If
    End If
    
    
    
    
    'Alguna comprobacion
    'Si es un cobro, por tarjeta y tiene gastos
    'entonces tendra que ir todo en un unico apunte
    If (Combo1.ItemData(Combo1.ListIndex) = 6 Or Combo1.ItemData(Combo1.ListIndex) = 0) And ImporteGastosTarjeta_ > 0 Then
        cad = ""
        '-----------------------------------------------------
        If Me.chkAsiento(0).Value Then
            cad = "No debe marcar la opcion de varios asientos"
        Else
            If Me.chkPorFechaVenci.Value Then
                riesgo = 0
                For I = 1 To Me.ListView1.ListItems.Count
                    If ListView1.ListItems(I).Checked Then
                        
                        Fecha = ListView1.ListItems(I).SubItems(3)
                        If riesgo = 0 Then
                            F2 = Fecha
                            riesgo = 1
                        Else
                            'Si las fechas son distintas NO dejo seguir
                            If F2 <> Fecha Then
                                cad = "Debe contabilizarlo todo en un único apunte"
                                Exit For
                            End If
                        End If
                    End If
        
                Next I
            End If
        End If
            
        If cad <> "" Then
            MsgBox cad, vbExclamation
            Exit Sub
        End If
        
        
        'Compruebo que tiene configurada la cuenta de gastos de tarjeta
        If Combo1.ItemData(Combo1.ListIndex) = 6 Then   'SOLO TARJETA
            cad = DevuelveDesdeBD("ctagastostarj", "bancos", "codmacta", txtCta(4).Text, "T")
            If cad = "" Then
                MsgBox "Falta configurar la cuenta de gastos de tarjeta", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    
    
    'Fecha dentro de ejercicios
    If CDate(Text3(0).Text) < vParam.fechaini Then
        MsgBox "Fuera de ejercios.", vbExclamation
        Exit Sub
    Else
        Fecha = DateAdd("yyyy", 1, vParam.fechafin)
        If CDate(Text3(0).Text) > Fecha Then
            If MsgBox("Fecha de ejercicio aun no abierto. ¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
            
    End If
    
    
    
    
    cad = "Desea contabilizar los vencimientos seleccionados?"
    If Combo1.ItemData(Combo1.ListIndex) = 1 Then
        I = 0
        If Not ContabTransfer And SegundoParametro <> "" Then I = 1
        If I = 1 Then
            'Estamos creando la transferencia o el pago domiciliado
            cad = RecuperaValor(Me.vTextos, 5)
            If cad = "" Then
                cad = "Desea generar la transferencia?"
            Else
                cad = "Desea generar el " & cad & "?"
            End If
        End If
    End If
    
    If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
'???
    ' para la impresion
    Dim SQL As String
    
    SQL = "delete from tmppendientes where codusu = " & vUsu.Codigo
    Conn.Execute SQL
'???
    
    
    
    Screen.MousePointer = vbHourglass
    
    'Una cosa mas.
    'Si la forma de pago es talon/pagere, y me ha escrito el numero de talon pagare...
    'Se lo tengo que pasar a la contabilizacion, con lo cual tendre que grabar
    'el nº de talon pagare en reftalonpag

        If Combo1.ItemData(Combo1.ListIndex) = vbTalon Or Combo1.ItemData(Combo1.ListIndex) = vbPagare Then
        
        
              ' llamamos a un formulario para que me introduzca la referencia de los talones o pagarés
              Dim CadInsert As String
              Dim CadValues As String
              
              SQL = "delete from tmpcobros2 where codusu = " & vUsu.Codigo
              Conn.Execute SQL

              CadInsert = "insert into tmpcobros2 (codusu,numserie,numfactu,fecfactu,numorden,fecvenci,reftalonpag,bancotalonpag) values "
              CadValues = ""

              For I = 1 To Me.ListView1.ListItems.Count
                    If ListView1.ListItems(I).Checked Then
                            CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(ListView1.ListItems(I).Text, "T") & "," & DBSet(Val(ListView1.ListItems(I).SubItems(1)), "N") & ","
                            CadValues = CadValues & DBSet(ListView1.ListItems(I).SubItems(2), "F") & "," & DBSet(ListView1.ListItems(I).SubItems(4), "N") & ","
                            CadValues = CadValues & DBSet(ListView1.ListItems(I).SubItems(3), "F") & ","
                            CadValues = CadValues & ValorNulo & "," & ValorNulo & "),"
                    End If
              Next I
              
              If CadValues <> "" Then
                  Conn.Execute CadInsert & Mid(CadValues, 1, Len(CadValues) - 1)

                  frmTESRefTalon.Show vbModal
              End If
        
              ' aqui estaba el update de cobros de la referencia talonpagare  y banco talonpagare, ahora está en lineas
        End If

    
    
    'Si el parametro dice k van todos en el mismo asiento, pues eso, todos en el mismo asiento
    'Primero leemos la forma de pago, el tipo perdon
    Set vp = New Ctipoformapago
    
    'en vtextos, en el 3 tenemos la forpa
    cad = ""
    cad = Combo1.ItemData(Combo1.ListIndex) 'RecuperaValor(vTextos, 3)
    If cad = "" Then
        I = -1
    Else
        I = Val(cad)
    End If
    If vp.Leer(I) = 1 Then
        'ERROR GRAVE LEYENDO LA FORMA DE PAGO
        Screen.MousePointer = vbDefault
        Set vp = Nothing
        End
    End If
    
    
    '--------------------------------------------------------
    'Si es realizar transferencia, crearemos la transferencia
    '--------------------------------------------------------
           
    'If Not Cobros And Tipo = 1 Then
    If Combo1.ItemData(Combo1.ListIndex) = 1 Then
        If Not ContabTransfer And SegundoParametro <> "" Then
            'Generamos la norma
            
            If Not RealizarTransferencias Then
                NumRegElim = 0
                Exit Sub
            End If
                                
            'Habra que salir
            NumRegElim = 1
            Unload Me
            Exit Sub
    
        End If
    End If
    
    '-----------------------------------------------------
    If Me.chkPorFechaVenci.Value Then
        'Contabilizaremos por fecha de vencimiento
        'Haremos una comrpobacion. Miraremos que todos los recibos marcados para
        'contabilizar , si la fecha no pertenece a actual y siguiente lo contabilizaremos con fecha
        'de cobro, es decir, la fecha con la que viene del otro form

        F2 = DateAdd("yyyy", 1, vParam.fechafin)
        Importe = 0
        riesgo = 0
        cad = ""
        SubItemVto = 3
        For I = 1 To Me.ListView1.ListItems.Count
            If ListView1.ListItems(I).Checked Then
                Fecha = ListView1.ListItems(I).SubItems(SubItemVto)
                riesgo = 0
                If Fecha < vParam.fechaini Or Fecha > F2 Then
                    riesgo = 1
                Else
                    If Fecha < vParamT.fechaAmbito Then riesgo = 1
                End If
                If riesgo = 1 Then
                    If InStr(1, cad, Format(Fecha, "dd/mm/yyyy")) = 0 Then
                        cad = cad & "    " & Format(Fecha)
                        Importe = Importe + 1
                        If Importe > 5 Then
                            cad = cad & vbCrLf
                            Importe = 0
                        End If
                    End If
                End If
            End If
        Next I
    
        If cad <> "" Then
            cad = "Las siguientes fechas están fuera de ejercicio (actual y siguiente):" & vbCrLf & vbCrLf & cad
            cad = cad & vbCrLf & vbCrLf & "Se contabilizarán con fecha: " & Text3(0).Text & vbCrLf
            cad = cad & "¿Desea continuar?"
            If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then cad = ""
                
        End If
        Importe = 0
        riesgo = 0
    End If
    
    
    DescripcionTransferencia = ""
    If ContabTransfer Then
        'Estamos contabilizando la transferencia
        cad = "stransfer"
        cad = cad & "cob"
        DescripcionTransferencia = DevNombreSQL(DevuelveDesdeBD("descripcion", cad, "codigo", SegundoParametro, "N"))
        
    End If

    
    cad = "DELETE from tmpactualizar  where codusu =" & vUsu.Codigo
    Conn.Execute cad


    Conn.BeginTrans
    
    'Si hay que generar la
    
    
    If HacerNuevaContabilizacion Then
        
        
        Conn.CommitTrans
              
        'Tenemos k borrar los listview
        For I = (ListView1.ListItems.Count) To 1 Step -1
            If ListView1.ListItems(I).Checked Then
               ListView1.ListItems.Remove I
            End If
        Next I
    Else
        TirarAtrasTransaccion
    End If


    ImpSeleccionado = 0
    Text2(2).Text = Format(ImpSeleccionado, FormatoImporte)
    If chkImprimir.Value Then Imprimir2
    
        
    Set vp = Nothing
    Screen.MousePointer = vbDefault
    
End Sub

Private Function HacerNuevaContabilizacion() As Boolean
    On Error GoTo EHacer
    HacerNuevaContabilizacion = False
    
    'Paso1. Meto todos los seleccionados en una tabla
    If Not InsertarPagosEnTemporal2 Then Exit Function
    
    
    
    'Paso 2
    'Compruebo que los vtos a cobrar no tienen ni la cuenta bloqueada, ni,
    'si contabilizo por fecha de bloqueo, alguna de los vencimienotos
    'esta fuera del de fechas
    If Not ComprobarCuentasBloquedasYFechasVencimientos Then Exit Function
    
    
    
    'Contabilizo desde la tabla. Asi puedo agrupar mejor
    ContablizaDesdeTmp
    
    HacerNuevaContabilizacion = True
    
    
    Exit Function
EHacer:
    MuestraError Err.Number, "Contabilizando"
End Function

Private Sub Imprimir2()
Dim SQL As String
Dim EsCobroTarjetaNavarres As Boolean

    EsCobroTarjetaNavarres = False
    If Combo1.ItemData(Combo1.ListIndex) = vbTarjeta Then
        'Si tiene el parametro y le ha puesto valor
        If vParamT.IntereseCobrosTarjeta > 0 And ImporteGastosTarjeta_ > 0 Then EsCobroTarjetaNavarres = True
        
        If EsCobroTarjetaNavarres Then
        
            'parrafo1
            SQL = ""
            If Combo1.ItemData(Combo1.ListIndex) = vbTarjeta Then
                If vParamT.IntereseCobrosTarjeta > 0 And ImporteGastosTarjeta_ > 0 Then
                    SQL = "1"
                    If Fecha <= vParam.fechafin Then SQL = "2"
                    SQL = DevuelveDesdeBD("contado" & SQL, "contadores", "tiporegi", ContCreditoNav) 'tarjeta credito tipo NAVARRES
                    If SQL = "" Then SQL = "1"
                    J = Val(SQL) + 1
                    SQL = Format(J, "00000")
                End If
            End If
                
        End If
        
    End If

    frmTESImpRecibo.VienedeRealizarCobro = EsCobroTarjetaNavarres
    frmTESImpRecibo.pNumFactu = SQL
    frmTESImpRecibo.pFecFactu = Text3(0).Text
    frmTESImpRecibo.Show vbModal
                                                                         
End Sub

Private Sub Refrescar()
    Screen.MousePointer = vbHourglass
    CargaList
    Screen.MousePointer = vbDefault
End Sub

Private Sub chkVerPendiente_Click()
    CargaList
End Sub

Private Sub cmdAceptar_Click()
    Generar2
    
    If VieneDesdeNorma57 Then Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
    If Combo1.ListIndex = -1 Then
        MsgBox "Debe introducir un tipo de Cobro. Revise.", vbExclamation
        Combo1.SetFocus
    
    Else
         I = 0
         If Combo1.ItemData(Combo1.ListIndex) = 1 And Me.SegundoParametro <> "" Then
             If Not ContabTransfer Then
                 I = 1
                 cad = RecuperaValor(vTextos, 5) 'Dira si es PAGO DOMICILIADO
                 If cad <> "" Then
                     If vParamT.PagosConfirmingCaixa Then
                     Else
                     End If
                 Else
                 End If
                 Image1(1).Visible = False  'No muestre la ayuda
             End If
         End If
         
         Me.chkPorFechaVenci.Visible = I = 0
         chkGenerico(0).Visible = I = 0
         Me.chkVtoCuenta(0).Visible = I = 0
         
        '++
        If Combo1.ItemData(Combo1.ListIndex) = vbTarjeta Then
            If vParamT.IntereseCobrosTarjeta > 0 Then
                ImporteGastosTarjeta_ = vParamT.IntereseCobrosTarjeta
                Text3(6).Text = Format(ImporteGastosTarjeta_, "##0.00")
            End If
        End If
        
        I = 0
        If (Combo1.ItemData(Combo1.ListIndex) = 2 Or Combo1.ItemData(Combo1.ListIndex) = 3) Then I = 1
        Me.mnBarra1.Visible = I = 1
        Me.mnNumero.Visible = I = 1
    
        CargaList
    End If
End Sub

Private Sub Combo2_Change()
    CargaList
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Me.Refresh
        espera 0.1
        
        'OCTUBRE 2014
        'PAgos por ventanilla.
        'Pondre como fecha de vencimiento la fecha que el
        'banco, en el fichero, me indica que realio el pago
        If Combo1.ListIndex <> -1 Then
        
        End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub
 

Private Sub DevuelveCadenaPorTipo(Impresion As Boolean, ByRef Cad1 As String)

    
    Cad1 = ""
    
    '++
    If Combo1.ListIndex = -1 Then Exit Sub
    
    Select Case Combo1.ItemData(Combo1.ListIndex)
    Case 0
        If Impresion Then
            Cad1 = "He recibido mediante efectivo de"
        Else
            Cad1 = "[EFECTIVO]"
        End If
        
    Case 1
        Cad1 = "[TRANSFERENCIA]"
    Case 2
        If Impresion Then
            Cad1 = "He recibido mediante TALON de"
        Else
            Cad1 = "[TALON]"
        End If
    Case 3
        If Impresion Then
            Cad1 = "He recibido mediante PAGARE de"
        Else
            Cad1 = "[PAGARE]"
        End If
    
    Case 4
        Cad1 = "[RECIBO BANCARIO]"
    
    Case 5
        Cad1 = "[CONFIRMING]"
    
    Case 6
        If Impresion Then
            Cad1 = "He recibido mediante TARJETA DE CREDITO de"
        Else
            Cad1 = "[TARJETA CREDITO]"
        End If
    
    Case Else
        
        
    End Select
End Sub

Private Sub Form_Load()

    PrimeraVez = True
    Limpiar Me
    Me.Icon = frmPpal.Icon
    For I = 0 To imgFecha.Count - 1
        Me.imgFecha(I).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next I
    For I = 0 To imgCuentas.Count - 1
        Me.imgCuentas(I).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next I
    
    
    CargaCombo
    
    CargaIconoListview Me.ListView1
    ListView1.Checkboxes = True
    imgCheck(0).Visible = True
    imgCheck(1).Visible = True
    chkPorFechaVenci.Value = 0
    
    imgFecha(2).Visible = False 'Para cambiar la fecha de contabilizacion de los pagos
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
     
     
    Text3(0).Text = Format(Now, "dd/mm/yyyy")
    Me.ImporteGastosTarjeta_ = 0
    Me.CodmactaUnica = ""
     
    DevuelveCadenaPorTipo False, cad
    LeerparametrosContabilizacion

    'Efectuar cobros
    FrameRemesar.Visible = True
    ListView1.SmallIcons = Me.ImageList1
    CargaColumnas
    
    
    
'    'Octubre 2014
'    'Norma 57 pagos ventanilla
'    'Si en el select , en el SQL, viene un
    
    
    If VieneDesdeNorma57 Then
        PosicionarCombo Combo1, 0
    End If
    
    ' Norma57 pago de ventanilla
    '         si viene de allí deshabilitamos un monton de campos y otros los damos cargados
    Text3(0).Enabled = Not VieneDesdeNorma57
    Text3(1).Enabled = Not VieneDesdeNorma57
    Text3(2).Enabled = Not VieneDesdeNorma57
    txtCta(4).Enabled = Not VieneDesdeNorma57
    txtCta(5).Enabled = Not VieneDesdeNorma57
    chkVerPendiente.Enabled = Not VieneDesdeNorma57
    Text3(6).Enabled = Not VieneDesdeNorma57
    chkImprimir.Enabled = Not VieneDesdeNorma57
    imgFecha(0).Enabled = Not VieneDesdeNorma57
    imgFecha(1).Enabled = Not VieneDesdeNorma57
    imgFecha(2).Enabled = Not VieneDesdeNorma57
    imgCuentas(0).Enabled = Not VieneDesdeNorma57
    imgCuentas(1).Enabled = Not VieneDesdeNorma57
    If VieneDesdeNorma57 Then
        Combo1.Enabled = False
        txtCta(4).Text = RecuperaValor(vTextos, 2)
        txtDCta(4).Text = RecuperaValor(vTextos, 3)
        Text3(0).Text = RecuperaValor(vTextos, 1)
    End If
    
    
    If Combo1.ListIndex <> -1 Then
        If Combo1.ItemData(Combo1.ListIndex) = 0 Then
            If InStr(1, vSQL2, "from tmpconext  WHERE codusu") > 0 Then chkPorFechaVenci.Value = 1
            
            If InStr(1, vSQL2, "from tmpconext  WHERE codusu") > 0 Then AjustarFechaVencimientoDesdeFicheroBancario
        
            If VieneDesdeNorma57 Then CargaList
        
        End If
    End If
End Sub

Private Sub Form_Resize()
Dim I As Integer
Dim H As Integer

    If Me.WindowState = 1 Then Exit Sub  'Minimizar
    If Me.Height < 2700 Then Me.Height = 2700
    If Me.Width < 2700 Then Me.Width = 2700

    'Situamos el frame y demas
    Me.frame.Width = Me.Width - 120
    Me.Frame1.Left = Me.Width - 120 - Me.Frame1.Width
    Me.Frame1.Top = Me.Height - Frame1.Height - 540 '360
    FrameRemesar.Width = Me.frame.Width - 320

    Me.ListView1.Top = Me.frame.Height + 60
    Me.ListView1.Height = Me.Frame1.Top - Me.ListView1.Top - 60
    Me.ListView1.Width = Me.frame.Width

    'Las columnas
    H = ListView1.Tag
    ListView1.Tag = ListView1.Width - ListView1.Tag - 320 'Del margen
    For I = 1 To Me.ListView1.ColumnHeaders.Count
        If InStr(1, ListView1.ColumnHeaders(I).Tag, "%") Then
            cad = (Val(ListView1.ColumnHeaders(I).Tag) * (Val(ListView1.Tag)) / 100)
        Else
            'Si no es de % es valor fijo
            cad = Val(ListView1.ColumnHeaders(I).Tag)
        End If
        Me.ListView1.ColumnHeaders(I).Width = Val(cad)
    Next I
    ListView1.Tag = H
End Sub


Private Sub CargaColumnas()
Dim ColX As ColumnHeader
Dim Columnas As String
Dim Ancho As String
Dim ALIGN As String
Dim NCols As Integer
Dim I As Integer

    ListView1.ColumnHeaders.Clear
    NCols = 13 '11
    Columnas = "Serie|Factura|F.Factura|F. VTO|Nº|CLIENTE|Tipo|Importe|Gasto|Cobrado|Pendiente|"
    Ancho = "800|10%|12%|12%|520|23%|840|12%|8%|11%|12%|0%|0%|"
    ALIGN = "LLLLLLLDDDDDD"
    
    
    ListView1.Tag = 2200  'La suma de los valores fijos. Para k ajuste los campos k pueden crecer
    
    If Combo1.ListIndex <> -1 Then
        If Combo1.ItemData(Combo1.ListIndex) = 2 Or Combo1.ItemData(Combo1.ListIndex) = 3 Then
            ''Si es un talon o pagare entonces añadire un campo mas
            NCols = NCols + 1
            Columnas = Columnas & "Nº Documento|"
            Ancho = Ancho & "2500|"
            ALIGN = ALIGN & "L"
        End If
    End If
   For I = 1 To NCols
        cad = RecuperaValor(Columnas, I)
        If cad <> "" Then
            Set ColX = ListView1.ColumnHeaders.Add()
            ColX.Text = cad
            'ANCHO
            cad = RecuperaValor(Ancho, I)
            ColX.Tag = cad
            'align
            cad = Mid(ALIGN, I, 1)
            If cad = "L" Then
                'NADA. Es valor x defecto
            Else
                If cad = "D" Then
                    ColX.Alignment = lvwColumnRight
                Else
                    'CENTER
                    ColX.Alignment = lvwColumnCenter
                End If
            End If
        End If
    Next I

End Sub

Private Sub GuardarMarcados()
Dim I As Long
    
    Set FrasMarcadas = New Collection

    For I = 1 To Me.ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
            FrasMarcadas.Add ListView1.ListItems(I).Text & "|" & ListView1.ListItems(I).SubItems(1) & "|" & ListView1.ListItems(I).SubItems(2) & "|" & ListView1.ListItems(I).SubItems(4) & "|" & ListView1.ListItems(I).Tag & "|"
        End If
    Next I

End Sub


Private Sub CargaList()
On Error GoTo ECargando

    Me.MousePointer = vbHourglass
    Screen.MousePointer = vbHourglass
    
    GuardarMarcados
    
    Set RS = New ADODB.Recordset
'    Fecha = CDate(Text1.Text)
    ListView1.ListItems.Clear
    Importe = 0
    Vencido = 0
    riesgo = 0
    ImpSeleccionado = 0
    
    CargaCobros
    
    Set FrasMarcadas = Nothing
    
    Text2(2).Text = "0,00"
    Label2(2).Caption = "Selecionado"
    Label2(2).Visible = True
    Text2(2).Visible = True
    Label2(3).Visible = True And Cobros
'        Text2(3).Visible = True And Cobros
    
ECargando:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Text2(0).Text = Format(Importe, FormatoImporte)
    Text2(1).Text = Format(Vencido, FormatoImporte)
    
    Text2(2).Text = Format(ImpSeleccionado, FormatoImporte)
    Me.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    Set RS = Nothing
End Sub

Private Sub CargaCobros()
Dim Inserta As Boolean

    RiesTalPag = 0
    cad = DevSQL
    
    'ORDENACION
    If CampoOrden = "" Then CampoOrden = "cobros.fecvenci"
    cad = cad & " ORDER BY " & CampoOrden
    If Orden Then cad = cad & " DESC"
    If CampoOrden <> "cobros.fecvenci" Then cad = cad & ", cobros.fecvenci"
    
    
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Inserta = True
        If RS!tipoformapago = vbTipoPagoRemesa Then
            
        ElseIf RS!tipoformapago = vbTalon Or RS!tipoformapago = vbPagare Then
        
        End If
        
        If Inserta Then
            InsertaItemCobro
        End If  'de insertar
        
        RS.MoveNext
    Wend
    RS.Close
End Sub


Private Sub InsertaItemCobro()
Dim vImporte As Currency
Dim DiasDif As Long
Dim ImpAux As Currency

    Set ItmX = ListView1.ListItems.Add()
    
    ItmX.Text = RS!NUmSerie
    ItmX.SubItems(1) = RS!NumFactu
    ItmX.SubItems(2) = Format(RS!FecFactu, "dd/mm/yyyy")
    ItmX.SubItems(3) = Format(RS!FecVenci, "dd/mm/yyyy")
    ItmX.SubItems(4) = RS!numorden
    ItmX.SubItems(5) = RS!Nommacta
    ItmX.SubItems(6) = RS!siglas
    
    ItmX.SubItems(7) = Format(RS!ImpVenci, FormatoImporte)
    vImporte = DBLet(RS!Gastos, "N")
    
    'Gastos
    ItmX.SubItems(8) = Format(vImporte, FormatoImporte)
    vImporte = vImporte + RS!ImpVenci
    
    If Not IsNull(RS!impcobro) Then
        ItmX.SubItems(9) = Format(RS!impcobro, FormatoImporte)
        impo = vImporte - RS!impcobro
        ItmX.SubItems(10) = Format(impo, FormatoImporte)
    Else
        impo = vImporte
        ItmX.SubItems(9) = "0.00"
        ItmX.SubItems(10) = Format(vImporte, FormatoImporte)
    End If
    If RS!tipoformapago = vbTipoPagoRemesa Then
        If Asc(Right(" " & DBLet(RS!siturem, "T"), 1)) = 81 Then
            riesgo = riesgo + vImporte
        Else
        End If
    
    ElseIf RS!tipoformapago = vbTalon Or RS!tipoformapago = vbPagare Then
            If RS!recedocu = 1 Then RiesTalPag = RiesTalPag + DBLet(RS!impcobro, "N")
    End If
    
    If RS!tipoformapago = vbTarjeta Then
        'Si tiene el parametro y le ha puesto valor
        If vParamT.IntereseCobrosTarjeta > 0 And ImporteGastosTarjeta_ > 0 Then
            DiasDif = 0
            Fecha = CDate(Text3(0).Text)
            If RS!FecVenci < Fecha Then DiasDif = DateDiff("d", RS!FecVenci, Fecha)
            If DiasDif > 0 Then
                'Si ya tenia gastos.
                If DBLet(RS!Gastos, "N") > 0 Then
                    MsgBox "Ya tenia gastos", vbExclamation
                    ItmX.ListSubItems(8).Bold = True
                    ItmX.ListSubItems(8).ForeColor = vbRed
                End If
                
                ImpAux = ((ImporteGastosTarjeta_ / 365) * DiasDif) / 100
                ImpAux = Round(ImpAux * impo, 2)
                
                impo = impo + ImpAux
                ItmX.SubItems(10) = Format(impo, FormatoImporte)
                'La de gastos
                ImpAux = DBLet(RS!Gastos, "N") + ImpAux
                ItmX.SubItems(8) = Format(ImpAux, FormatoImporte)
            End If
            
        End If
    End If
    
    If RS!FecVenci < Fecha Then
        'LO DEBE
        ItmX.SmallIcon = 1
        Vencido = Vencido + impo
    Else
'        ItmX.SmallIcon = 2
    End If
    Importe = Importe + impo
    
    ItmX.Tag = RS!codmacta
    
    If Combo1.ListIndex <> -1 Then
        If (Combo1.ItemData(Combo1.ListIndex) = 1 And SegundoParametro <> "") Or VieneDesdeNorma57 Then
            If Not IsNull(RS!transfer) Or VieneDesdeNorma57 Then
                ItmX.Checked = True
                ImpSeleccionado = ImpSeleccionado + impo
            End If
        End If
    End If
    
     ' nuevo si está marcada lo miramos
     For I = 1 To FrasMarcadas.Count
        cad = FrasMarcadas.Item(I)
        If RecuperaValor(cad, 1) = RS!NUmSerie And RecuperaValor(cad, 2) = RS!NumFactu And RecuperaValor(cad, 3) = RS!FecFactu And RecuperaValor(cad, 4) = RS!numorden And RecuperaValor(cad, 5) = RS!codmacta Then
            ItmX.Checked = True
        End If
     Next I
    
    
End Sub



Private Function DevSQL() As String
Dim cad As String
Dim vSql As String

    vSql = vSQL2
    'Llegados a este punto montaremos el sql
    
    If Text3(1).Text <> "" Then
        If vSql <> "" Then vSql = vSql & " AND "
        vSql = vSql & " cobros.fecvenci >= '" & Format(Text3(1).Text, FormatoFecha) & "'"
    End If
        
        
    If Text3(2).Text <> "" Then
        If vSql <> "" Then vSql = vSql & " AND "
        vSql = vSql & " cobros.fecvenci <= '" & Format(Text3(2).Text, FormatoFecha) & "'"
    End If

    
    'Forma de pago
    If Me.txtCta(5).Text <> "" Then
        'Los de un cliente solamente
        If vSql <> "" Then vSql = vSql & " AND "
        vSql = vSql & " cobros.codmacta = '" & txtCta(5).Text & "'"
        
        '++
        If Me.chkVerPendiente.Value = 0 Then
            If Combo1.ListIndex <> -1 Then
                If Combo1.ItemData(Combo1.ListIndex) >= 0 Then
                    If vSql <> "" Then vSql = vSql & " and "
                    vSql = vSql & " formapago.tipforpa = " & Combo1.ItemData(Combo1.ListIndex)
                End If
            End If
        End If
    Else
        If Combo1.ListIndex > 0 Then
            If Combo1.ItemData(Combo1.ListIndex) >= 0 Then
                If vSql <> "" Then vSql = vSql & " AND "
                vSql = vSql & " formapago.tipforpa = " & Combo1.ItemData(Combo1.ListIndex)    'SubTipo
            
            End If
        End If
    End If

    If vSql <> "" Then vSql = vSql & " AND "
    vSql = vSql & " ((formapago.tipforpa in (" & vbTalon & "," & vbPagare & ") and cobros.codrem is null) or not formapago.tipforpa in (" & vbTalon & "," & vbPagare & "))"

    
    ' no entran a jugar los recibos
    If vSql <> "" Then vSql = vSql & " and "
    vSql = vSql & " formapago.tipforpa <> " & vbTipoPagoRemesa
    
    ' solo los pendientes de cobro
    If vSql <> "" Then vSql = vSql & " and "
    vSql = vSql & " (coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0)) <> 0 "
    
    ' si está marcado los de solo el banco previsto
    If chkSoloBancoPrev.Value = 1 Then
        If txtCta(4).Text <> "" Then
            If vSql <> "" Then vSql = vSql & " AND "
            vSql = vSql & " cobros.ctabanc1 = " & DBSet(txtCta(4).Text, "T")
        End If
    End If
    
    ' si me han puesto usuario lo selecciono
    If Combo2.ListIndex <> -1 Then
        If Combo2.ItemData(Combo2.ListIndex) <> 1000 Then
            If vSql <> "" Then vSql = vSql & " and "
            vSql = vSql & " cobros.codusu = " & DBSet(Combo2.ItemData(Combo2.ListIndex), "N")
        End If
    End If
    
    
    'cobros
    cad = "SELECT cobros.*, formapago.nomforpa, tipofpago.descformapago, tipofpago.siglas, "
    cad = cad & " cuentas.nommacta,cuentas.codmacta,tipofpago.tipoformapago, "
    cad = cad & " coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0) imppdte "
    cad = cad & " FROM ((cobros INNER JOIN formapago ON cobros.codforpa = formapago.codforpa) INNER JOIN tipofpago ON formapago.tipforpa = tipofpago.tipoformapago) INNER JOIN cuentas ON cobros.codmacta = cuentas.codmacta"
    If vSql <> "" Then cad = cad & " WHERE " & vSql
        
    'SQL pedido
    DevSQL = cad
End Function


Private Sub CargaPagos()

    cad = DevSQL
    
    'ORDENACION
    cad = cad & " ORDER BY " & CampoOrden
    If Orden Then cad = cad & " DESC"
    If CampoOrden <> "pagos.fecefect" Then cad = cad & ", pagos.fecefect"


    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        InsertaItemPago
        RS.MoveNext
    Wend
    RS.Close

End Sub


Private Sub InsertaItemPago()
Dim J As Byte
        Set ItmX = ListView1.ListItems.Add()
        
        ItmX.Text = RS!NumFactu
        ItmX.SubItems(1) = Format(RS!FecFactu, "dd/mm/yyyy")
        ItmX.SubItems(2) = Format(RS!fecefect, "dd/mm/yyyy")
        ItmX.SubItems(3) = RS!numorden
        ItmX.SubItems(4) = RS!Nommacta
        ItmX.SubItems(5) = RS!siglas
        ItmX.SubItems(6) = Format(RS!ImpEfect, FormatoImporte)
        If Not IsNull(RS!imppagad) Then
            ItmX.SubItems(7) = Format(RS!imppagad, FormatoImporte)
            impo = RS!ImpEfect - RS!imppagad
            ItmX.SubItems(8) = Format(impo, FormatoImporte)
        Else
            impo = RS!ImpEfect
            ItmX.SubItems(7) = "0.00"
            ItmX.SubItems(8) = ItmX.SubItems(6)
        End If
        If RS!fecefect < Fecha Then
            'LO DEBE
            ItmX.SmallIcon = 1
            Vencido = Vencido + impo
        Else
            ItmX.SmallIcon = 2
        End If
        
        If Combo1.ItemData(Combo1.ListIndex) = 1 Then
            If Not IsNull(RS!transfer) Then
                ItmX.Checked = True
                ImpSeleccionado = ImpSeleccionado + impo
            End If
        End If
        'El tag lo utilizo para la cta proveedor
        ItmX.Tag = RS!ctaprove
        
        Importe = Importe + impo
        
        'Si el documento estaba emitido ya
        If Val(RS!emitdocum) = 1 Then
            'Tiene marcado DOCUMENTO EMITIDO
            ItmX.ForeColor = vbRed
            For J = 1 To ListView1.ColumnHeaders.Count - 1
                ItmX.ListSubItems(J).ForeColor = vbRed
            Next J
            If DBLet(RS!Referencia, "T") = "" Then ItmX.ListSubItems(4).ForeColor = vbMagenta
        End If
       
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Para dejar las variables bien
    ContabTransfer = False
    DesdeRecepcionTalones = False
    'Por si acaso
    NumeroTalonPagere = ""
    CodmactaUnica = ""
End Sub

Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtCta(4).Text = RecuperaValor(CadenaSeleccion, 1)
        txtDCta(4).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    cad = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtCta(5).Text = RecuperaValor(CadenaSeleccion, 1)
        txtDCta(5).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub Image1_Click(Index As Integer)
    frmTESPreguntas.Opcion = 2
    frmTESPreguntas.Show vbModal
End Sub

Private Sub imgCheck_Click(Index As Integer)
    SeleccionarTodos Index = 1 Or Index = 2
End Sub

Private Sub imgCuentas_Click(Index As Integer)
Dim SQL As String

    Select Case Index
        Case 0 ' cuenta de banco
            Set frmBan = New frmBasico2
            AyudaBanco frmBan
            Set frmBan = Nothing
            
        Case 1 ' cuenta de proveedor
            Set frmCCtas = New frmColCtas
            
            frmCCtas.DatosADevolverBusqueda = "0"
            frmCCtas.FILTRO = "1" ' clientes
            frmCCtas.Show vbModal
            
            Set frmCCtas = Nothing
    End Select
    PonFoco txtCta(Index + 4)
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Fecha = Now
    If Text3(I).Text <> "" Then
        If IsDate(Text3(I).Text) Then Fecha = CDate(Text3(I).Text)
    End If
    cad = ""
    Set frmC = New frmCal
    frmC.Fecha = Fecha
    frmC.Show vbModal
    Set frmC = Nothing
    If cad <> "" Then
        Text3(Index).Text = cad
            
        If Index = 0 Then
            'Antes de poder cambiar la fecha hay que comprobar si la fecha devuelta es OK
            '                                                'Fecha OK
            If FechaCorrecta2(CDate(cad), True) < 2 Then Text3(0).Text = cad
        End If
    End If
End Sub

Private Sub imgTraerRestoDatosCliProv_Click()

    Screen.MousePointer = vbHourglass

    'Tenemos que traer todos los vtos del client/proveedor en question
        
    'Auqi
    Set RS = New ADODB.Recordset
    cad = ""
    For I = 1 To Me.ListView1.ListItems.Count
        With Me.ListView1.ListItems(I)
                '                 RS!NUmSerie
                cad = cad & ", ('" & .Text & "'," & .SubItems(1) & ",'"
                '  fecfac,                                               RS!numorden
                cad = cad & Format(.SubItems(2), FormatoFecha) & "'," & .SubItems(4) & ")"
        
        End With
    Next
    Set RS = Nothing
        
        
    If cad <> "" Then
        cad = Mid(cad, 2)
        cad = " AND NOT (numserie,codfaccl,fecfaccl,numorden) IN (" & cad & ")"
    End If
    
    cad = " AND scobro.codmacta = '" & CodmactaUnica & "'" & cad
    cad = " AND codrem is null and situacionjuri=0 and  transfer is null  AND impvenci >0 " & cad

    cad = " AND sforpa.tipforpa = stipoformapago.tipoformapago " & cad
    cad = " FROM scobro,sforpa,cuentas,stipoformapago WHERE scobro.codmacta=cuentas.codmacta and scobro.codforpa = sforpa.codforpa " & cad
    
    
    'Hacemos un conteo
    Set RS = New ADODB.Recordset
    I = 0
    RS.Open "SELECT Count(*) " & cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        I = DBLet(RS.Fields(0), "N")
    End If
    RS.Close
    
    If I > 0 Then
        If MsgBox("Se van a añadir " & I & " vencimiento(s) a la lista.    Continuar?", vbQuestion + vbYesNo) <> vbYes Then I = 0
    Else
        MsgBox "No existen mas vencimientos para añadir", vbExclamation
    End If
    
    
    
    
    
    If I > 0 Then
        cad = " cuentas.nommacta,cuentas.codmacta,stipoformapago.tipoformapago " & cad
        cad = "SELECT scobro.*, sforpa.nomforpa, stipoformapago.descformapago, stipoformapago.siglas, " & cad

        RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        While Not RS.EOF
            InsertaItemCobro
            RS.MoveNext
        Wend
        RS.Close
    
    
        
    End If
    Set RS = Nothing
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim Campo2 As Integer

    Orden = Not Orden
'        Columnas = "Serie|Nº Factura|F.Factura|F. VTO|Nº|CLIENTE|Tipo|Importe|Gasto|Cobrado|Pendiente|"
    Select Case ColumnHeader
        Case "Serie"
            CampoOrden = "cobros.numserie"
        Case "Nº Factura"
            CampoOrden = "cobros.numfactu"
        Case "F.Factura"
            CampoOrden = "cobros.fecfactu"
        Case "F. VTO"
            CampoOrden = "cobros.fecvenci"
        Case "Nº"
            CampoOrden = "cobros.numorden"
        Case "CLIENTE"
            CampoOrden = "nommacta"
        Case "Tipo"
            CampoOrden = "siglas"
        Case "Importe"
            CampoOrden = "cobros.impvenci"
        Case "Gasto"
            CampoOrden = "cobros.gastos"
        Case "Cobrado"
            CampoOrden = "cobros.impcobro"
        Case "Pendiente"
            CampoOrden = "imppdte"
    End Select
    CargaList

End Sub


Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    I = ColD(0)
    impo = ImporteFormateado(Item.SubItems(I))
    
    If Item.Checked Then
        Set ListView1.SelectedItem = Item
        I = 1
    Else
        I = -1
    End If
    ImpSeleccionado = ImpSeleccionado + (I * impo)
    Text2(2).Text = Format(ImpSeleccionado, FormatoImporte)
    
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu Me.mnContextual
    End If
End Sub

Private Sub SeleccionarTodos(Seleccionar As Boolean)
Dim J As Integer
    J = ColD(0)
    ImpSeleccionado = 0
    For I = 1 To Me.ListView1.ListItems.Count
        ListView1.ListItems(I).Checked = Seleccionar
        impo = ImporteFormateado(ListView1.ListItems(I).SubItems(J))
        ImpSeleccionado = ImpSeleccionado + impo
    Next I
    If Not Seleccionar Then ImpSeleccionado = 0
    Text2(2).Text = Format(ImpSeleccionado, FormatoImporte)
End Sub


Private Sub mnNumero_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    CadenaDesdeOtroForm = "####"
    frmTESPreguntas.Opcion = 0
    frmTESPreguntas.vTexto = ListView1.SelectedItem.SubItems(11)
    frmTESPreguntas.Show vbModal
    If CadenaDesdeOtroForm <> "####" Then ListView1.SelectedItem.SubItems(11) = CadenaDesdeOtroForm
        
End Sub

Private Sub mnQUitarSel_Click()
    SeleccionarTodos False
End Sub

Private Sub mnSelectAll_Click()
    SeleccionarTodos True
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Function GenerarRecibos2() As Boolean
Dim SQL As String
Dim Contador As Integer
Dim J As Integer
Dim Poblacion As String



    On Error GoTo EGenerarRecibos
    GenerarRecibos2 = False
    
    
    'Limpiamos
    cad = "Delete from Usuarios.zTesoreriaComun where codusu = " & vUsu.Codigo
    Conn.Execute cad


    'Guardamos datos empresa
    cad = "Delete from Usuarios.z347carta where codusu = " & vUsu.Codigo
    Conn.Execute cad
    
    cad = "INSERT INTO Usuarios.z347carta (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir, saludos, "
    cad = cad & "parrafo1, parrafo2, parrafo3, parrafo4, parrafo5, despedida, contacto, Asunto, Referencia)"
    cad = cad & " VALUES (" & vUsu.Codigo & ", "
    
    'Estos datos ya veremos com, y cuadno los relleno
    Set miRsAux = New ADODB.Recordset
    SQL = "select nifempre,siglasvia,direccion,numero,escalera,piso,puerta,codpos,poblacion,provincia from empresa2"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "'##########','" & vEmpresa.nomempre & "','#############','######','##########','##########'"
    If Not miRsAux.EOF Then
        SQL = ""
        For J = 1 To 6
            SQL = SQL & DBLet(miRsAux.Fields(J), "T") & " "
        Next J
        SQL = Trim(SQL)
        SQL = "'" & DBLet(miRsAux!nifempre, "T") & "','" & DevNombreSQL(vEmpresa.nomempre) & "','" & DevNombreSQL(SQL) & "'"
        SQL = SQL & ",'" & DBLet(miRsAux!codpos, "T") & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "'"
        Poblacion = DevNombreSQL(DBLet(miRsAux!Poblacion, "T"))
        
    End If
    miRsAux.Close
 
    cad = cad & SQL
    'otralinea,saludos
    cad = cad & ",NULL"
    'parrafo1
    SQL = ""
    If Combo1.ItemData(Combo1.ListIndex) = vbTarjeta Then
        If vParamT.IntereseCobrosTarjeta > 0 And ImporteGastosTarjeta_ > 0 Then
            SQL = "1"
            If Fecha <= vParam.fechafin Then SQL = "2"
            SQL = DevuelveDesdeBD("contado" & SQL, "contadores", "tiporegi", "3") 'tarjeta credito tipo NAVARRES
            If SQL = "" Then SQL = "1"
            J = Val(SQL) + 1
            SQL = Format(J, "00000")
        End If
    End If
    
    cad = cad & ",'" & SQL & "'"
    
    
    '------------------------------------------------------------------------
    cad = cad & ",NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
    Conn.Execute cad

    'Empezamos
    SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2, texto3, texto4, texto5, "
    SQL = SQL & "texto6, importe1, importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion)"
    SQL = SQL & " VALUES (" & vUsu.Codigo & ","


    Contador = 0
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
            'Lo insertamos tres veces
             RellenarCadenaSQLRecibo I, Poblacion
            'Lo rellenamos por triplicado    'VER ESTO
            'For J = 1 To 3
                Contador = Contador + 1
                Conn.Execute SQL & Contador & "," & cad
            'Next J
        End If
    Next I
    GenerarRecibos2 = True
EGenerarRecibos:
    If Err.Number <> 0 Then
        MuestraError Err.Number
    End If
   
End Function


'----------------------------------
'Rellenaremos las cadenas de texto
Private Sub RellenarCadenaSQLRecibo_QUITAR(NumeroItem As Integer)
Dim Aux As String

    With ListView1.ListItems(NumeroItem)
        'texto1 , texto2, texto3, texto4, texto5,
        'texto6, importe1, importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion
        
            
        'Textos
        '---------
        '1.- Recibo nª
        cad = "'" & .Text & "/" & .SubItems(1) & "'"
        
        'Pagos: cad = "'" & .Text & "/" & .SubItems(3) & "'"
        
        
        'Lugar Vencimiento
        cad = cad & ",'VALENCIA'"
        
        'provee
        cad = cad & ",'" & DevNombreSQL(.SubItems(5)) & "',"
        'PAgos:cad = cad & ",'" & .SubItems(4) & "'"
        
        '4,5,6
        'El text 4 tendra el numero de talon, pagare SI SE HA PUESTO
        If Combo1.ItemData(Combo1.ListIndex) = 2 Or Combo1.ItemData(Combo1.ListIndex) = 3 Then
            Aux = ""
            If .SubItems(10) <> "" Then
                If Combo1.ItemData(Combo1.ListIndex) = 2 Then
                    Aux = "TALON: "
                Else
                    Aux = "PAGARE: "
                End If
                Aux = Aux & .SubItems(10)
            End If
            cad = cad & "'" & Aux & "'"
        Else
            cad = cad & "NULL"
        End If
        
        cad = cad & ",NULL,NULL"
        
        
        Importe = ImporteFormateado(.SubItems(10))
        
        'IMPORTES
        '--------------------
        cad = cad & "," & TransformaComasPuntos(CStr(Importe))
        
        'El segundo importe NULL
        cad = cad & ",NULL"
        
        'FECFAS
        '--------------
        'Libramiento o pago
        cad = cad & ",'" & Format(Text3(0).Text, FormatoFecha) & "'"
        cad = cad & ",'" & Format(.SubItems(3), FormatoFecha) & "'"
        
        '3era fecha  NULL
        cad = cad & ",NULL"
        
        'OBSERVACIONES
        '------------------
        Aux = EscribeImporteLetra(Importe)
        
        Aux = "       ** " & Aux
        cad = cad & ",'" & Aux & "**'"
        cad = cad & ",NULL"
        
        
        'OPCION
        '--------------
        cad = cad & ",NULL)"
        
        
    End With
End Sub




'TROZO COPIADO DESDE frmcobrosimprimir
Private Sub RellenarCadenaSQLRecibo(NumeroItem As Integer, Lugar As String)
Dim Aux As String
Dim EsCobroTarjetaNavarres  As Boolean
Dim QueDireccionMostrar As Byte
    '0. NO tiene
    '1. La del recibo
    '2. La de la cuenta



    'Grabara dos valores mas que no graba el resto
    'parrafo1 --> CIF
    'importe2
    EsCobroTarjetaNavarres = False
    If Combo1.ItemData(Combo1.ListIndex) = vbTarjeta Then
        If vParamT.IntereseCobrosTarjeta > 0 And ImporteGastosTarjeta_ > 0 Then EsCobroTarjetaNavarres = True
    End If

    
    With ListView1.ListItems(NumeroItem)
    
        ' IRan:   text5:  nomclien
        '         texto6: domclien
        '         observa2  cpclien  pobclien    + vbcrlf + proclien
    
        cad = "select nomclien,domclien,pobclien,cpclien,proclien,razosoci,dirdatos,codposta,despobla,desprovi"
        'MAYO 2010
        cad = cad & ",codbanco,codsucur,digcontr,scobro.cuentaba,scobro.codmacta,nifdatos "
        cad = cad & " from scobro,cuentas where scobro.codmacta =cuentas.codmacta and"
        cad = cad & " numserie ='" & .Text & "' and codfaccl=" & .SubItems(1)
        cad = cad & " and fecfaccl='" & Format(.SubItems(2), FormatoFecha) & "' and numorden=" & .SubItems(4)
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not miRsAux.EOF Then
            'El vto NO tiene datos de fiscales
            Aux = DBLet(miRsAux!nomclien, "T")
            If Aux = "" Then
                'La cuenta tampoco los tiene
                If IsNull(miRsAux!dirdatos) Then
                    QueDireccionMostrar = 0
                Else
                    QueDireccionMostrar = 2
                End If
            Else
                QueDireccionMostrar = 1
            End If
        Else
            QueDireccionMostrar = 0
        End If
        
        'texto1 , texto2, texto3, texto4, texto5,
        'texto6, importe1, importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion
        
            
        'Textos
        '---------
        '1.- Recibo nª
        cad = "'" & .Text & "/" & Format(.SubItems(1), "0000") & "'"
        
        'Pagos: cad = "'" & .Text & "/" & .SubItems(3) & "'"
        
        
        'Lugar Vencimiento
        cad = cad & ",'" & Lugar & "'"
        
        'text3 mostrare el codmacta (en pago tarjeta con intereses,NAVARRES, el NIF
        Aux = miRsAux!codmacta
        If EsCobroTarjetaNavarres Then Aux = DBLet(miRsAux!nifdatos, "T")
        cad = cad & ",'" & DevNombreSQL(Aux) & "',"
        
        
        
        'MAYO 2010.    Ahora en este campo ira el CCC del cliente si es que lo tiene
        'Cad = Cad & "'" & .SubItems(6) & "'," ANTES
        Aux = DBLet(miRsAux!codbanco, "N")
        If Aux = "" Or Aux = "0" Then
            Aux = "NULL"
        Else
            'codbanco,codsucur,digcontr,cuentaba
            Aux = Format(DBLet(miRsAux!codbanco, "N"), "0000")
            Aux = Aux & " " & Format(DBLet(miRsAux!codsucur, "N"), "0000") & " "
            Aux = Aux & Mid(DBLet(miRsAux!digcontr, "T") & "  ", 1, 2) & " "
            Aux = Aux & Right(String(10, "0") & DBLet(miRsAux!Cuentaba, "N"), 10)
            Aux = "'" & Aux & "'"
        End If
        cad = cad & Aux & ","
    
        '5 y 6.
        'text5: nomclien
        'texto6:domclien
        If QueDireccionMostrar = 0 Then
            'Cad = Cad & "NULL,NULL"
            'Siempre el nomclien
            cad = cad & "'" & DevNombreSQL(.SubItems(5)) & "',NULL"
        Else
            If QueDireccionMostrar = 1 Then
                cad = cad & "'" & DevNombreSQL(DBLet(miRsAux!nomclien, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!domclien, "T")) & "'"
            Else
                cad = cad & "'" & DevNombreSQL(DBLet(miRsAux!razosoci, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!dirdatos, "T")) & "'"
            End If
        End If
        
        Importe = ImporteFormateado(.SubItems(10))
        
        'IMPORTES
        '--------------------
        cad = cad & "," & TransformaComasPuntos(CStr(Importe))
        
        'El segundo importe NULL   Abril 2014. Tarjetas NAVARRES. Llevara los gastos
        Aux = "NULL"
        If EsCobroTarjetaNavarres Then Aux = .SubItems(8)
        cad = cad & "," & TransformaComasPuntos(CStr(Aux))
        
        'FECFAS
        '--------------
        'Libramiento o pago     Auqi pone NOW
        cad = cad & ",'" & Format(Text3(0).Text, FormatoFecha) & "'"
        cad = cad & ",'" & Format(.SubItems(3), FormatoFecha) & "'"
        
        '3era fecha  NULL
        cad = cad & ",NULL"
        
        'OBSERVACIONES
        '------------------
        Aux = EscribeImporteLetra(Importe)
        
        Aux = "       ** " & Aux
        cad = cad & ",'" & Aux & "**',"
        
        
        'Observa 2
        '         observa2:    cpclien  pobclien    + vbcrlf + proclien
        If QueDireccionMostrar = 0 Then
            Aux = "NULL"
        Else
            
            If QueDireccionMostrar = 1 Then
                Aux = DBLet(miRsAux!cpclien, "T") & "      " & DevNombreSQL(DBLet(miRsAux!pobclien, "T"))
                Aux = Trim(Aux)
                If Aux <> "" Then Aux = Aux & vbCrLf
                Aux = Aux & DevNombreSQL(DBLet(miRsAux!proclien, "T"))
            Else
                Aux = DBLet(miRsAux!codposta, "T") & "      " & DevNombreSQL(DBLet(miRsAux!desPobla, "T"))
                Aux = Trim(Aux)
                If Aux <> "" Then Aux = Aux & vbCrLf
                Aux = Aux & DevNombreSQL(DBLet(miRsAux!desProvi, "T"))
                
            End If
            Aux = "'" & Aux & "'"
        End If
        cad = cad & Aux
        
        
        
        'OPCION
        '--------------
        cad = cad & ",NULL)"
        
        
    End With
    miRsAux.Close
End Sub


Private Sub RellenarCadenaSQLReciboPagos(NumeroItem As Integer)
Dim Aux As String
Dim RT As ADODB.Recordset


    Set RT = New ADODB.Recordset
    With ListView1.ListItems(NumeroItem)
        'texto1 , texto2, texto3, texto4, texto5,
        'texto6, importe1, importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion
        
            
        'Textos
        '---------
        '1.- Recibo nª
        cad = "'" & DevNombreSQL(.Text) & "',"
        
        Aux = "nommacta,razosoci,dirdatos,codposta,despobla,desprovi"
        Aux = "Select " & Aux & " from cuentas where codmacta = '" & .Tag & "'"
        RT.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RT.EOF Then
            Aux = DBLet(RT.Fields(1), "T")
            If Aux = "" Then Aux = RT!Nommacta
            Aux = "'" & DevNombreSQL(Aux) & "'"
            For SubItemVto = 2 To 5
                Aux = Aux & ",'" & DevNombreSQL(DBLet(RT.Fields(SubItemVto), "T")) & "'"
            Next
        Else
            'VACIO. Error leyendo cuenta
            MsgBox "Error leyendo cuenta:" & .Tag, vbExclamation
            Aux = "'" & DevNombreSQL(.SubItems(4)) & "'"
            For SubItemVto = 2 To 5
                Aux = Aux & ",NULL"
            Next
        End If
        RT.Close
        cad = cad & Aux
        
        
        
        
        
        
        Importe = ImporteFormateado(.SubItems(8))
        
        'IMPORTES
        '--------------------
        cad = cad & "," & TransformaComasPuntos(CStr(Importe))
        
        'El segundo importe NULL
        cad = cad & ",NULL"
        
        'FECFAS
        '--------------
        'Libramiento o pago
        cad = cad & ",'" & Format(Text3(0).Text, FormatoFecha) & "'"
        cad = cad & ",'" & Format(.SubItems(2), FormatoFecha) & "'"
        
        '3era fecha  NULL
        cad = cad & ",NULL"
        
        'OBSERVACIONES
        '------------------
        Aux = EscribeImporteLetra(Importe)
        
        Aux = "       ** " & Aux
        cad = cad & ",'" & Aux & "**'"
        cad = cad & ",NULL"
        
        
        'OPCION
        '--------------
        cad = cad & ",NULL)"
        
        
    End With
    Set RT = Nothing
End Sub




Private Function InsertarPagosEnTemporal2() As Boolean
Dim C As String
Dim Aux As String
Dim J As Long
Dim FechaContab As Date
Dim FechaFinEjercicios As Date
Dim vGasto As Currency

    
    InsertarPagosEnTemporal2 = False
    
    C = " WHERE codusu =" & vUsu.Codigo
    Conn.Execute "DELETE FROM tmpfaclin" & C

    '++
    GastosTransferencia = CCur(ComprobarCero(Text3(6).Text))


    'Fechas fin ejercicios
    FechaFinEjercicios = DateAdd("yyyy", 1, vParam.fechafin)


    Aux = "INSERT INTO tmpfaclin (codusu, codigo, Fecha,Numfactura, cta, Cliente, NIF, Imponible,  Total) "
    Aux = Aux & "VALUES (" & vUsu.Codigo & ","
    For J = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(J).Checked Then
            C = J & ",'"
            'Si la fecha de contabilizacion esta fuera de ejercicios
            FechaContab = CDate(ListView1.ListItems(J).SubItems(3))
            

            C = C & Format(FechaContab, FormatoFecha) & "','"
            
            '-----------------------------------------------------
            'Fecha de contabilizacion
            If Me.chkPorFechaVenci.Value Then
                I = 0
                
                'Meto la fecha VTO
                If FechaContab < vParam.fechaini Then
                    I = 1
                Else
                    If FechaContab > FechaFinEjercicios Then
                        I = 1
                    Else
                        If FechaContab < vParamT.fechaAmbito Then I = 1
                    End If
                End If
                
                If I = 1 Then FechaContab = CDate(Text3(0).Text)
            Else
                'La fecha de contabilizacion es la del text
                FechaContab = CDate(Text3(0).Text)
            End If
            
            'MEto la fecha de contabilizaccion
            C = C & Format(FechaContab, FormatoFecha) & "','"
            'Cuenta contable
            C = C & ListView1.ListItems(J).Tag & "','"
            'Serie factura |FECHAfactura|
            'Neuvo febrero 2008 Serie |factura |FECHAfactura|numvto|
            C = C & ListView1.ListItems(J).Text & "|" & ListView1.ListItems(J).SubItems(1) & "|" & ListView1.ListItems(J).SubItems(2) & "|" & ListView1.ListItems(J).SubItems(4)
            C = C & "|','"
            
            'Cuenta agrupacion cobros
            If Combo1.ItemData(Combo1.ListIndex) = 1 And ContabTransfer Then
                C = C & Me.chkGenerico(1).Tag & "',"
            Else
                C = C & Me.chkGenerico(0).Tag & "',"
            End If
            'Dinerito
            'riesgo es GASTO
            I = ColD(0)
            impo = ImporteFormateado(ListView1.ListItems(J).SubItems(I))
            riesgo = ImporteFormateado(ListView1.ListItems(J).SubItems(I - 2))
            impo = impo - riesgo
            C = C & TransformaComasPuntos(CStr(impo)) & "," & TransformaComasPuntos(CStr(riesgo)) & ")"


            'Lo meto en la BD
            C = Aux & C
            Conn.Execute C
        End If
    Next J

    'Si es por tarjeta hay una opcion para meter el gasto total
    'que a partir de la cuenta de banco gasto tarjeta crear una linea mas
    If Combo1.ItemData(Combo1.ListIndex) = 6 And ImporteGastosTarjeta_ > 0 Then
        'Agosto 2014
        'Pagos credito NAVARRES  --> NO llevan esta linea
        
        If vParamT.IntereseCobrosTarjeta = 0 Then
            cad = DevuelveDesdeBD("ctagastostarj", "bancos", "codmacta", txtCta(4).Text, "T")
            
            FechaContab = CDate(Text3(0).Text)
            C = "'" & Format(FechaContab, FormatoFecha) & "'"
            C = C & "," & C
            C = J & "," & C & ",'" & cad & "','"
            'Serie factura |FECHAfactura| ----> pondre: "gastos" | fecha contab
            C = C & "GASTOS|" & FechaContab & "|','" & cad & "',"
            'Dinerito
            'riesgo es GASTO
            impo = -ImporteGastosTarjeta_
            C = C & TransformaComasPuntos(CStr(impo)) & ",0)"
            C = Aux & C
            Conn.Execute C
        End If
    
    End If
    
    'Gastos contabilizacion transferencia
'    If Combo1.ItemData(Combo1.ListIndex) = 1 And GastosTransferencia <> 0 Then
    If GastosTransferencia <> 0 Then
        'aqui ira los gastos asociados a la transferencia
        'Hay que ver los lados
        
        cad = DevuelveDesdeBD("ctagastos", "bancos", "codmacta", txtCta(4).Text, "T")
        
        FechaContab = CDate(Text3(0).Text)
        C = "'" & Format(FechaContab, FormatoFecha) & "'"
        C = C & "," & C
        C = J & "," & C & ",'" & cad & "','"
        'Serie factura |FECHAfactura| ----> pondre: "gastos" | fecha contab
        C = C & "TRA" & Format(SegundoParametro, "0000000") & "|" & FechaContab & "|','" & cad & "',"
        'Dinerito
        'riesgo es GASTO
        impo = -GastosTransferencia
        C = C & TransformaComasPuntos(CStr(impo)) & ",0)"
        C = Aux & C
        Conn.Execute C
        
    End If
    
    InsertarPagosEnTemporal2 = True
    
    

End Function

'TENGO en la tabla tmpfaclin los vtos.
'Ahora en funcion de los check haremos la contabilizacion
'agrupando de un modo o de otro
Private Sub ContablizaDesdeTmp()
Dim SQL As String
Dim ContraPartidaPorLinea As Boolean
Dim UnAsientoPorCuenta As Boolean
Dim PonerCuentaGenerica As Boolean
Dim AgrupaCuenta As Boolean
Dim RS As ADODB.Recordset
Dim MiCon As Contadores
Dim CampoCuenta As String
Dim CampoFecha As String
Dim GeneraAsiento As Boolean
Dim CierraAsiento As Boolean
Dim NumLinea As Integer
Dim ImpBanco As Currency
Dim NumVtos As Integer
Dim GastosTransDescontados As Boolean
Dim LineaUltima As Integer

    'Valores por defecto
    ContraPartidaPorLinea = True
    UnAsientoPorCuenta = False
    PonerCuentaGenerica = False
    AgrupaCuenta = False
    CampoFecha = "numfactura"
    GastosTransDescontados = False 'por lo que pueda pasar
    
    
    'Si va agrupado por cta
    If Combo1.ItemData(Combo1.ListIndex) = 1 Then '--  And ContabTransfer Then
        If Me.chkContrapar(0).Value Then ContraPartidaPorLinea = False
        If Me.chkAsiento(0).Value Then UnAsientoPorCuenta = True
        If chkGenerico(0).Value Then PonerCuentaGenerica = True
        
        'Si lleva GastosTransferencia entonce AGRUPAMOS banco
        If GastosTransferencia <> 0 Then
            
            'gastos tramtiaacion transferenca descontados importe
            SQL = DevuelveDesdeBD("GastTransDescontad", "bancos", "codmacta", txtCta(4).Text, "T")
            GastosTransDescontados = SQL = "1"
            
            AgrupaCuenta = False
        Else
            If Me.chkVtoCuenta(0).Value Then AgrupaCuenta = True
        End If
    Else
        'Si no es transferencia
        If Me.chkContrapar(0).Value Then ContraPartidaPorLinea = False
        If Me.chkAsiento(0).Value Then UnAsientoPorCuenta = True
        If chkGenerico(0).Value Then PonerCuentaGenerica = True
        If Me.chkVtoCuenta(0).Value Then AgrupaCuenta = True
        'La contabiliacion es por fecha vencimiento , no por fecha solicitada
        'YA cuando inserto en temporal miro esto
        'If chkPorFechaVenci.Value Then CampoFecha = "fecha"
    End If
    
    If PonerCuentaGenerica Then
        CampoCuenta = "NIF"
    Else
        CampoCuenta = "cta"
    End If
    'EL SQL lo empezamos aquin
    SQL = CampoCuenta & " AS cliprov,"
    'Selecciona
    SQL = "select count(*) as numvtos,codigo,numfactura,fecha,cliente," & SQL & "sum(imponible) as importe,sum(total) as gastos from tmpfaclin"
    SQL = SQL & " where codusu =" & vUsu.Codigo & " GROUP BY "
    cad = ""
    If AgrupaCuenta Then
       If PonerCuentaGenerica Then
            cad = "nif" 'La columna NIF lleva los datos de la cuenta generica
        Else
            cad = "cta"
        End If
        'Como estamos agrupando por cuenta, marcaremos tb la fecha
        'Ya que si tienen fechas distintas son apuntes distintos
        cad = cad & "," & CampoFecha
    End If
    
    'Si no agrupo por nada agrupare por codigo(es decir como si no agrupara)
    If cad = "" Then cad = "codigo"
    
    'La ordenacion
    cad = cad & " ORDER BY " & CampoFecha
    If Not PonerCuentaGenerica Then cad = cad & ",cta"
        
    
    'Tanto si agrupamos por cuenta (Generica o no)
    'el recodset tendra las lineas que habra que insertar en/los apuntes(s)
    '
    'Es decir. Que si agrupo no tengo que ir moviendome por el recodset mirando a ver si
    'las cuentas son iguales.
    'Ya que al hacer group by ya lo estaran
    cad = SQL & cad
    Set RS = New ADODB.Recordset
    RS.Open cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    'Inicializamos variables
    Fecha = CDate("01/01/1900")
    GeneraAsiento = False
    While Not RS.EOF
        'Comprobaciones iniciales
        If UnAsientoPorCuenta Then
            'Para cada linea ira su asiento
            GeneraAsiento = True
            CierraAsiento = True
            If Fecha < CDate("01/01/1950") Then CierraAsiento = False
            Fecha = CDate(RS.Fields(CampoFecha))
        Else
            'Veremos en funcion de la fecha
            GeneraAsiento = False
            If CDate(RS.Fields(CampoFecha)) = Fecha Then
                'Estamos en la misma fecha. Luego sera el mismo asiento
                'Excepto que asi no lo digan las variables
                If Not PonerCuentaGenerica Then
                    If UnAsientoPorCuenta Then
                        GeneraAsiento = True
                        If Fecha < CDate("01/01/1950") Then CierraAsiento = True
                    End If
                End If
                        
            Else
                'Fechas distintas.
                GeneraAsiento = True
                CierraAsiento = True
                If Fecha < CDate("01/01/1950") Then CierraAsiento = False
        
                Fecha = CDate(RS.Fields(CampoFecha))
            End If
        End If 'de aseinto por cuenta
        
        
        'Si tengo que cerrar el asiento anterior
        If CierraAsiento Then
            'Tirar atras el RS
            If Not ContraPartidaPorLinea Then
                RS.MovePrevious
                Fecha = CDate(RS.Fields(CampoFecha))  'Para la fecha de asiento
                impo = ImpBanco
                'Generamos las lineas de apunte que faltan
                InsertarEnAsientosDesdeTemp RS, MiCon, 2, NumLinea, NumVtos
                
                'Inserto para que actalice             3: Opcion para INSERT INTO tmpactualizar
                InsertarEnAsientosDesdeTemp RS, MiCon, 3, NumLinea, NumVtos
                
                'Reestauramos variables
                NumVtos = 0
                'Ponemos la variable
                CierraAsiento = False
                'Volvemos el RS al sitio
                RS.MoveNext
                Fecha = CDate(RS.Fields(CampoFecha))
            Else
                'Inserto para que actalice             3: Opcion para INSERT INTO tmpactualizar
                InsertarEnAsientosDesdeTemp RS, MiCon, 3, NumLinea, NumVtos
            End If
        End If
 
        
        'Si genero asiento
        If GeneraAsiento Then
            If MiCon Is Nothing Then Set MiCon = New Contadores
            MiCon.ConseguirContador "0", Fecha <= vParam.fechafin, True
                        
            'Genero la cabecera
            InsertarEnAsientosDesdeTemp RS, MiCon, 0, NumLinea, NumVtos
            
            NumLinea = 1
            ImpBanco = 0
            'Reservo la primera linea para el banco
            If GastosTransferencia <> 0 Then
                NumLinea = 2
                If Not GastosTransDescontados Then
                   ImpBanco = -GastosTransferencia
                End If
            End If
            
            riesgo = 0
        End If
        
    
        'Para el cobro /pago  que tendremos en la fila actual del recordset
        '++
        Dim VienedeGastos As Boolean
        Dim CadRes As String
        CadRes = RecuperaValor(DBLet(RS!Cliente, "T"), 1)
        VienedeGastos = (CadRes = "GASTOS" Or CadRes = "TRA")
        
        impo = RS!Importe
        InsertarEnAsientosDesdeTemp RS, MiCon, 1, NumLinea, RS!NumVtos, VienedeGastos
    
        riesgo = riesgo + RS!Gastos
        ImpBanco = ImpBanco + RS!Importe
        NumLinea = NumLinea + 1
        
        'Si tengo que generar la contrapartida
        If ContraPartidaPorLinea Then
            NumVtos = RS!NumVtos
            InsertarEnAsientosDesdeTemp RS, MiCon, 2, NumLinea, NumVtos
            NumLinea = NumLinea + 1
            ImpBanco = 0
            riesgo = 0
        Else
            NumVtos = NumVtos + RS!NumVtos
        End If
        
        'Nos movemos
        RS.MoveNext
        
        
        If RS.EOF Then
            
            If Not ContraPartidaPorLinea Then
                
                'Era la ultima linea.
                RS.MovePrevious
                
                LineaUltima = NumLinea
                
                'Cierro el apunte, del banco
                'Si fuera una transferenicia con gastos descontados, me he dejado el numlinea=1
                'si no, no hago nada
                If GastosTransferencia <> 0 Then
                    If Not GastosTransDescontados Then NumLinea = 1
                End If
                impo = ImpBanco
                InsertarEnAsientosDesdeTemp RS, MiCon, 2, NumLinea, NumVtos
    
                If GastosTransferencia <> 0 Then
                    If Not GastosTransDescontados Then
                        NumLinea = LineaUltima + 1
                
                        impo = GastosTransferencia
                        
                        InsertarEnAsientosDesdeTemp RS, MiCon, 2, NumLinea, NumVtos
                    End If
                End If
    
                'CIERRO EL APUNTE
                InsertarEnAsientosDesdeTemp RS, MiCon, 3, NumLinea, NumVtos
                
                'Y vuelvo a ponerlo ande tocaba. Para que se salga del bucle
                RS.MoveNext
                
            Else
                'Cada linea de asiento tiene su banco
                'Faltara insertarlo en tmpactualizar
                InsertarEnAsientosDesdeTemp RS, MiCon, 3, NumLinea, NumVtos
            End If
        End If
    Wend
    RS.Close
    Set RS = Nothing
    
End Sub

Private Function ColD(Colu As Integer) As Integer
    Select Case Colu
    Case 0
            'IMporte pendiente
            ColD = 10
    Case 1
    
    End Select
    If Not Cobros Then ColD = ColD - 2
End Function

Private Sub EliminarCobroPago(Indice As Integer)
    
    With ListView1.ListItems(Indice)
            
            cad = "DELETE FROM  cobros WHERE "
            cad = cad & " numserie  = '" & .Text
            cad = cad & "' and numfactu = " & .SubItems(1)
            cad = cad & " and numorden = " & .SubItems(4)
            cad = cad & " and fecfactu = '" & Format(.SubItems(2), FormatoFecha) & "'"
            
    End With
    Ejecuta cad
End Sub


Private Function RealizarTransferencias() As Boolean

On Error GoTo ERealizarTransferencias
    RealizarTransferencias = False
    
    
    impo = 0
    
    For I = 1 To ListView1.ListItems.Count
        With ListView1.ListItems(I)
            If Not Cobros Then
                'TRANSFERENCIAS A PROVEEDORES
                cad = "UPDATE spagop SET transfer= "
                If .Checked Then
                    cad = cad & SegundoParametro
                    impo = 1
                Else
                    cad = cad & "NULL"
                End If
                cad = cad & " WHERE numfactu = '" & .Text
                cad = cad & "' and fecfactu = '" & Format(.SubItems(1), FormatoFecha)
                cad = cad & "' and numorden = " & .SubItems(3)
                cad = cad & " and ctaprove = '" & .Tag & "'"
            
            Else
                'ABONOS CLIENTES
                cad = "UPDATE scobro SET transfer= "
                If .Checked Then
                    cad = cad & SegundoParametro
                    impo = 1
                Else
                    cad = cad & "NULL"
                End If
                cad = cad & " WHERE numserie = '" & .Text
                cad = cad & "' and codfaccl = " & .SubItems(1)
                cad = cad & "  and fecfaccl = '" & Format(.SubItems(2), FormatoFecha)
                cad = cad & "' and numorden = " & .SubItems(4)
                
            End If
            Conn.Execute cad
        End With
    Next I
        
    If impo > 0 Then RealizarTransferencias = True
        
    
    Exit Function
ERealizarTransferencias:
    MuestraError Err.Number
End Function

Private Sub Text3_GotFocus(Index As Integer)
    ConseguirFoco Text3(Index), 3
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index).Text)
    If Text3(Index).Text = "" Then Exit Sub

    Select Case Index
        Case 6
            PonerFormatoDecimal Text3(6), 3
            Me.ImporteGastosTarjeta_ = ComprobarCero(Text3(Index).Text)
            
        Case 0, 1, 2
        
            If Not EsFechaOK(Text3(Index)) Then
                MsgBox "Fecha incorrecta", vbExclamation
                Text3(Index).Text = ""
                PonFoco Text3(Index)
            End If
    
            If Index = 1 Or Index = 2 Then CargaList
    End Select
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtCta_GotFocus(Index As Integer)
    ConseguirFoco txtCta(Index), 3
    If Index = 5 Then CtaAnt = txtCta(Index)
    If Index = 4 Then CtaAntBan = txtCta(Index)
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtCta_LostFocus(Index As Integer)
Dim DevfrmCCtas As String
Dim SQL As String

    Select Case Index
        Case 4 ' cuenta de banco
            txtCta(Index).Text = Trim(txtCta(Index).Text)
            DevfrmCCtas = txtCta(Index).Text
            I = 0
            If DevfrmCCtas <> "" Then
                If CuentaCorrectaUltimoNivel(DevfrmCCtas, SQL) Then
                    DevfrmCCtas = DevuelveDesdeBD("codmacta", "bancos", "codmacta", DevfrmCCtas, "T")
                    If DevfrmCCtas = "" Then
                        SQL = ""
                        MsgBox "La cuenta contable no esta asociada a ninguna cuenta bancaria", vbExclamation
                    End If
                Else
                    MsgBox SQL, vbExclamation
                    DevfrmCCtas = ""
                    SQL = ""
                End If
                I = 1
            Else
                SQL = ""
            End If
            
            
            txtCta(Index).Text = DevfrmCCtas
            txtDCta(Index).Text = SQL
            If DevfrmCCtas = "" And I = 1 Then
                PonFoco txtCta(Index)
            Else
                Text3(1).Tag = txtCta(Index).Text
            End If
        
            If CtaAntBan <> txtCta(4).Text Then CargaList
        
        Case 5 ' cuenta cliente
            DevfrmCCtas = Trim(txtCta(Index).Text)
            I = 0
            If DevfrmCCtas <> "" Then
                If CuentaCorrectaUltimoNivel(DevfrmCCtas, SQL) Then
                    
                Else
                    MsgBox SQL, vbExclamation
                    If Index < 3 Or Index = 9 Or Index = 10 Or Index = 11 Then
                        DevfrmCCtas = ""
                        SQL = ""
                    End If
                End If
                I = 1
            Else
                SQL = ""
            End If
            
            txtCta(Index).Text = DevfrmCCtas
            txtDCta(Index).Text = SQL
            If DevfrmCCtas = "" And I = 1 Then
                PonFoco txtCta(Index)
            Else
                Me.CodmactaUnica = txtCta(Index)
            End If
        
            If CtaAnt <> txtCta(5).Text Then CargaList

    End Select
End Sub


Private Sub LeerparametrosContabilizacion()
Dim B As Boolean

    Me.chkAsiento(0).Value = Abs(vParamT.contapag)
    
   
    If vParamT.contapag Then
        B = False
    Else
        B = vParamT.AgrupaBancario
    End If
    Me.chkContrapar(0).Value = Abs(B)
End Sub

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'3 Opciones
'   0.- CABECERA
'   1.- LINEAS  de clientes o proeveedores
'   2.- Cierre del asiento con el BANCO, y o caja
'   3.- Para poner Boqueoactu a 0
'       Si vParam.contapag entonces cuando hago el de banco/caja lo updateo
'       pero si NO es uno por pagina entonces, la utlima vez k hago el apunte por banco/caja
'       ejecuto el update
'
'   La contrB sera la contrpartida , si lueog resulta k si, para la linea de banco o caja
'
'   FechaAsiento:  Antes estaba a "piñon" text3(0).text
'
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'ByRef m As Contadores, NumLine As Integer, Marcador As Integer, Cabecera As Byte, ByRef ContraB As String, ByRef LaUltimaAmpliacion As String, ContraParEnBanco As Boolean, CuentaDeCobroGenerica As Boolean, CodigoCtaCoborGenerica As String)
Private Function InsertarEnAsientosDesdeTemp(ByRef RS1 As ADODB.Recordset, ByRef m As Contadores, Cabecera As Byte, ByRef NumLine As Integer, NumVtos As Integer, Optional VienedeGastos As Boolean)
Dim SQL As String
Dim Ampliacion As String
Dim Debe As Boolean
Dim Conce As Integer
Dim TipoAmpliacion As Integer
Dim PonerContrPartida As Boolean
Dim Aux As String
Dim ImporteInterno As Currency
    
    
    ImporteInterno = impo
    
    'LaUltimaAmpliacion  --> Servira pq si en parametros esta marcado un apunte por movimiento, o solo metemos
    '                        un unico pagao/cobro, repetiremos numdocum, textoampliacion
    
    'El diario

    FechaAsiento = Fecha
    If Cobros Then
        Ampliacion = vp.diaricli
    Else
        Ampliacion = vp.diaripro
    End If
    
    If Cabecera = 0 Then
        'La cabecera
        SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
        SQL = SQL & Ampliacion & ",'" & Format(FechaAsiento, FormatoFecha) & "'," & m.Contador
        SQL = SQL & ",  '"
        SQL = SQL & "Generado desde Tesorería el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & vUsu.Nombre
        If Combo1.ItemData(Combo1.ListIndex) = 1 And Not Cobros Then
            'TRANSFERENCIA
            Ampliacion = DevuelveDesdeBD("descripcion", "stransfer", "codigo", SegundoParametro, "N")
            If Ampliacion <> "" Then
                Ampliacion = "Concepto: " & Ampliacion
                Ampliacion = DevNombreSQL(Ampliacion)
                Ampliacion = vbCrLf & Ampliacion
                SQL = SQL & Ampliacion
            End If
        End If
        
        SQL = SQL & "',"
        SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilizar Cobros'"

        
        SQL = SQL & ")"
        NumLine = 0
     
    Else
        If Cabecera < 3 Then
            'Lineas de apuntes o cabecera.
            'Comparten el principio
             SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
             SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
             SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada, numserie, numfaccl, fecfactu, numorden, tipforpa, reftalonpag, bancotalonpag) "
             SQL = SQL & "VALUES (" & Ampliacion & ",'" & Format(FechaAsiento, FormatoFecha) & "'," & m.Contador & "," & NumLine & ",'"
             
             '1:  Asiento para el VTO
             If Cabecera = 1 Then
                 'codmacta
                 'Si agrupa la cuenta entonces
                 SQL = SQL & RS1!cliprov & "','"
                 
                 
                 'numdocum: la factura
                 If NumVtos > 1 Then
                    Ampliacion = "Vtos: " & NumVtos
                 Else
                    Ampliacion = DevNombreSQL(RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2))
                 End If
                 SQL = SQL & Ampliacion & "',"
                
                
                 'Veamos si va al debe, al haber, si ponemos concepto debe / haber etc eyc
                 If Cobros Then
                    'CLIENTES
                    If ImporteInterno < 0 Then
                       If vParam.abononeg Then
                           Debe = False
                       Else
                           'Va al debe pero cambiado de signo
                           Debe = True
                           ImporteInterno = Abs(ImporteInterno)
                       End If
                    Else
                       Debe = False
                    End If
                    If Debe Then
                        Conce = vp.condecli
                        TipoAmpliacion = vp.ampdecli
                        PonerContrPartida = vp.ctrdecli = 1
                    Else
                        Conce = vp.conhacli
                        TipoAmpliacion = vp.amphacli
                        PonerContrPartida = vp.ctrhacli = 1
                    End If
                 
                 
                 Else
                    'PROVEEDORES
                    If ImporteInterno < 0 Then
                       If vParam.abononeg Then
                           Debe = True
                       Else
                           'Va al debe pero cambiado de signo
                           Debe = False
                           ImporteInterno = Abs(ImporteInterno)
                       End If
                    Else
                       Debe = True
                    End If
                    If Debe Then
                        Conce = vp.condepro
                        TipoAmpliacion = vp.ampdepro
                        PonerContrPartida = vp.ctrdepro = 1
                    Else
                        Conce = vp.conhapro
                        TipoAmpliacion = vp.amphapro
                        PonerContrPartida = vp.ctrhapro = 1
                    End If
                     
                 End If
                
                
                 SQL = SQL & Conce & ","
                 
                 'AMPLIACION
                 Ampliacion = ""
                


                Select Case TipoAmpliacion
                Case 0, 1
                   If TipoAmpliacion = 1 Then Ampliacion = Ampliacion & vp.siglas & " "
                   Ampliacion = Ampliacion & RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                
                Case 2
                
                   Ampliacion = Ampliacion & RecuperaValor(RS1!Cliente, 3)
                
                Case 3
                    'NUEVA AMPLIC
                    Ampliacion = DescripcionTransferencia
                Case 4
                    'Estamos en la amplicacion del cliente. Es una tonteria tener esta opcion marcada, pero bien
                    Ampliacion = RecuperaValor(vTextos, 3)
                    Ampliacion = Mid(Ampliacion, InStr(1, Ampliacion, "-") + 1)
                Case 5
                    'Si hubiera que especificar mas el documento
'                    If Tipo = vbTalon Then
'                        AUX = "TAL Nº"
'                    Else
'                        AUX = "PAG Nº"
'                    End If
'
                
                    If Cobros Then
                        'Veo la el camporefencia de ese talon
                        'Antes cogiamos numero fra
                        'ahora contrapar
                        Ampliacion = RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)  'Num tal pag
                        If False Then
                            
                            Ampliacion = "numserie = '" & RecuperaValor(RS1!Cliente, 1) & "' AND RecuperaValor(RS1!Cliente, 2)"
                            Ampliacion = Ampliacion & " AND numorden = " & RecuperaValor(RS1!Cliente, 4) & " AND fecfactu "
                            Ampliacion = DevuelveDesdeBD("reftalonpag", "cobros", Ampliacion, Format(RecuperaValor(RS1!Cliente, 3), FormatoFecha), "F")
                            
                        Else
                            'Es numero tal pag + ctrpar
                            DescripcionTransferencia = RecuperaValor(vTextos, 2)
                            DescripcionTransferencia = Mid(DescripcionTransferencia, InStr(1, DescripcionTransferencia, "-") + 1)
                            Ampliacion = Ampliacion & " " & DescripcionTransferencia
                            DescripcionTransferencia = ""
                        End If
                        If Ampliacion = "" Then
                            Ampliacion = RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                        Else
                            Ampliacion = " NºDoc: " & Ampliacion
                        End If
                    Else
                        If NumeroTalonPagere = "" Then
                            Ampliacion = ""
                        Else
                            'Cta banco
                            Ampliacion = RecuperaValor(vTextos, 2)
                            Ampliacion = Mid(Ampliacion, InStr(1, Ampliacion, "-") + 1)
                            'Numero tal/pag
                        
                            Ampliacion = NumeroTalonPagere & " " & Ampliacion
                        
                        End If
                        
                        If Ampliacion = "" Then
                            Ampliacion = RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                        Else
                            Ampliacion = "NºDoc: " & Ampliacion
                        End If
                    End If
                    
                End Select
                   
                If NumVtos > 1 Then
                    'TIENE MAS DE UN VTO. No puedo ponerlo en la ampliacion
                    Ampliacion = "Vtos: " & NumVtos
                End If
                
                 'Le concatenamos el texto del concepto para el asiento -ampliacion
                 Aux = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce)) & " "
                 'Para la ampliacion de nºtal + ctrapar NO pongo la ampliacion del concepto
                 If TipoAmpliacion = 5 Then Aux = ""
                 Ampliacion = Aux & Ampliacion
                 If Len(Ampliacion) > 30 Then Ampliacion = Mid(Ampliacion, 1, 30)
                
                 SQL = SQL & "'" & DevNombreSQL(Ampliacion) & "',"
                 
                 
                 If Debe Then
                    SQL = SQL & TransformaComasPuntos(CStr(ImporteInterno)) & ",NULL,"
                 Else
                    SQL = SQL & "NULL," & TransformaComasPuntos(CStr(ImporteInterno)) & ","
                 End If
             
                'CENTRO DE COSTE
                SQL = SQL & "NULL,"
                
                'SI pone contrapardida
                If PonerContrPartida Then
                   SQL = SQL & "'" & txtCta(4).Text & "',"
                Else
                   SQL = SQL & "NULL,"
                End If
            
             
            Else
                    '----------------------------------------------------
                    'Cierre del asiento con el total contra banco o caja
                    '----------------------------------------------------
                    'codmacta
                    SQL = SQL & txtCta(4).Text & "','"
                     
  
                    PonerContrPartida = False
                    If NumVtos = 1 Then
                        PonerContrPartida = True
                    Else
                        PonerContrPartida = False
                    End If
                       
                    If PonerContrPartida Then
                       Ampliacion = DevNombreSQL(RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2))
                    Else
                       
                       Ampliacion = ""
                    End If
                     
                    SQL = SQL & Ampliacion & "',"
                   
                    
                    If Cobros Then
                        '----------------------------------------------------------------------
                        If ImporteInterno < 0 Then
                           If vParam.abononeg Then
                               Debe = True
                           Else
                               'Va al debe pero cambiado de signo
                               Debe = False
                               ImporteInterno = Abs(ImporteInterno)
                           End If
                        Else
                           Debe = True
                        End If
                                   
                        
                        'COmo el banco o caja, siempre van al reves (Su abono es nuetro pago..)
                        If Not Debe Then
                            Conce = vp.condecli
                            TipoAmpliacion = vp.ampdecli
                        Else
                            Conce = vp.conhacli
                            TipoAmpliacion = vp.amphacli
                        End If
                        
                     Else
                        'PROVEEDORES
                        If ImporteInterno < 0 Then
                           If vParam.abononeg Then
                               Debe = False
                           Else
                               'Va al debe pero cambiado de signo
                               Debe = True
                               ImporteInterno = Abs(ImporteInterno)
                           End If
                        Else
                           Debe = False
                        End If
                        
                        If Not Debe Then
                            Conce = vp.condepro
                            TipoAmpliacion = vp.ampdepro
                        Else
                            Conce = vp.conhapro
                            TipoAmpliacion = vp.amphapro
                        End If
                     End If
                     
                        
                     
                     
                
                     SQL = SQL & Conce & ","
                     'AMPLIACION
                     'AMPLIACION
                     Ampliacion = ""
                     
                     'Si estoy contabilizando pag de UN unico proveedor entonces NumeroTalonPageretendra valor
                     If NumVtos > 1 And NumeroTalonPagere <> "" Then NumVtos = 1
                        
                     
                     If NumVtos = 1 Then
                    
                        Select Case TipoAmpliacion
                        Case 0, 1
                           If TipoAmpliacion = 1 Then Ampliacion = Ampliacion & vp.siglas & " "
                           Ampliacion = Ampliacion & RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                        
                        Case 2
                        
                           Ampliacion = Ampliacion & RecuperaValor(RS1!Cliente, 3)
                        
                        Case 3
                            'NUEVA AMPLIC
                             Ampliacion = DescripcionTransferencia
                        Case 4, 5
                            'Nombre ctrpartida
                            Ampliacion = CStr(DBLet(RS1!cliprov, "T"))
                            Ampliacion = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Ampliacion, "T")
                            DescripcionTransferencia = Ampliacion
                            If Cobros Then
                                
                                'Veo la el camporefencia de ese talon
                                Ampliacion = RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                                Ampliacion = "numserie = '" & RecuperaValor(RS1!Cliente, 1) & "' AND numfaccl = " & RecuperaValor(RS1!Cliente, 2)
                                Ampliacion = Ampliacion & " AND numorden = " & RecuperaValor(RS1!Cliente, 4) & " AND fecfactu "
                                Ampliacion = DevuelveDesdeBD("reftalonpag", "hlinapu", Ampliacion, Format(RecuperaValor(RS1!Cliente, 3), FormatoFecha), "F")
                                
                                If Ampliacion = "" Then
                                    Ampliacion = RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                                Else
                                    Ampliacion = " NºDoc: " & Ampliacion
                                End If
                                Ampliacion = Ampliacion & " " & DescripcionTransferencia
     
                            Else
                                
                                Ampliacion = NumeroTalonPagere
                                If Ampliacion = "" Then
                                    Ampliacion = RecuperaValor(RS1!Cliente, 1) & RecuperaValor(RS1!Cliente, 2)
                                Else
                                    Ampliacion = "NºDoc: " & Ampliacion
                                End If
                            End If
                          
                            Ampliacion = Ampliacion & " " & DescripcionTransferencia
                            DescripcionTransferencia = ""
                          
                          
                        End Select
                    Else
                        'Ma de un VTO.  Si no
                        If vp.tipoformapago = vbTransferencia Then
                            'SI es transferencia
                            'If TipoAmpliacion = 3 Then Ampliacion = DescripcionTransferencia
                            Ampliacion = DescripcionTransferencia
                        
                        End If
                    End If
                    
                     Aux = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce))
                     Aux = Aux & " "
                     'Para la ampliacion de nºtal + ctrapar NO pongo la ampliacion del concepto
                     If TipoAmpliacion = 5 Then Aux = ""
                     Ampliacion = Trim(Aux & Ampliacion)
                     If Len(Ampliacion) > 30 Then Ampliacion = Mid(Ampliacion, 1, 30)
                    
                     SQL = SQL & "'" & DevNombreSQL(Ampliacion) & "',"
        
                         
                     If Debe Then
                        SQL = SQL & TransformaComasPuntos(CStr(ImporteInterno)) & ",NULL,"
                     Else
                        SQL = SQL & "NULL," & TransformaComasPuntos(CStr(ImporteInterno)) & ","
                     End If
                 
                     'CENTRO DE COSTE
                     SQL = SQL & "NULL,"
                    
                     'SI pone contrapardida
                     If PonerContrPartida Then
                        SQL = SQL & "'" & RS1!cliprov & "',"
                     Else
                        SQL = SQL & "NULL,"
                     End If
                 
            End If
            
            'Trozo comun
            '------------------------
            'IdContab
            SQL = SQL & "'COBROS',"
            
            'Punteado
            SQL = SQL & "0,"
            
            If Cabecera = 1 And Not VienedeGastos Then
            
                ' nuevos campos de la factura
                'numSerie , numfacpr, FecFactu, numorden, TipForpa, reftalonpag, bancotalonpag
                SQL = SQL & DBSet(RecuperaValor(RS1!Cliente, 1), "T") & "," & DBSet(RecuperaValor(RS1!Cliente, 2), "T") & "," & DBSet(RecuperaValor(RS1!Cliente, 3), "F") & ","
                SQL = SQL & DBSet(RecuperaValor(RS1!Cliente, 4), "N") & "," & DBSet(Combo1.ItemData(Combo1.ListIndex), "N") & ","
                
                Dim SqlBanco As String
                Dim RsBanco As ADODB.Recordset
                
                SqlBanco = "select reftalonpag, bancotalonpag from tmpcobros2 where codusu = " & vUsu.Codigo
                SqlBanco = SqlBanco & " and numserie = " & DBSet(RecuperaValor(RS1!Cliente, 1), "T")
                SqlBanco = SqlBanco & " and numfactu = " & DBSet(RecuperaValor(RS1!Cliente, 2), "N")
                SqlBanco = SqlBanco & " and fecfactu = " & DBSet(RecuperaValor(RS1!Cliente, 3), "F")
                SqlBanco = SqlBanco & " and numorden = " & DBSet(RecuperaValor(RS1!Cliente, 4), "N")
                SqlBanco = SqlBanco & " and codmacta = " & DBSet(RS1!cliprov, "T")
        
                Set RsBanco = New ADODB.Recordset
                RsBanco.Open SqlBanco, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RsBanco.EOF Then
                    SQL = SQL & DBSet(RsBanco.Fields(0), "T") & "," & DBSet(RsBanco.Fields(1), "T") & ")"
                Else
                    SQL = SQL & ValorNulo & "," & ValorNulo & ")"
                End If
                Set RsBanco = Nothing
                
            Else
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
            End If
                 
             
        End If 'De cabecera menor que 3, es decir : 1y 2
    
    
    End If
    
    'Ejecutamos si:
    '   Cabecera=0 o 1
    '   Cabecera=2 y impo=0.  Esto sginifica que estamos desbloqueando el apunte e insertandolo para pasarlo a hco
    Debe = True
    If Cabecera = 3 Then Debe = False
    If Debe Then Conn.Execute SQL
    
    If Debe Then
        '++monica
        If Cobros And Cabecera = 1 And Not VienedeGastos Then
        
            Dim Situacion As Byte
            
            Situacion = 1

            SQL = "update cobros set impcobro = coalesce(impcobro,0) + " & DBSet(ImporteInterno, "N")
            SQL = SQL & " ,fecultco = " & DBSet(FechaAsiento, "F")
            SQL = SQL & ", situacion = " & DBSet(Situacion, "N")
            SQL = SQL & " where numserie = " & DBSet(RecuperaValor(RS1!Cliente, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(RS1!Cliente, 2), "N")
            SQL = SQL & " and fecfactu = " & DBSet(RecuperaValor(RS1!Cliente, 3), "F") & " and numorden = " & DBSet(RecuperaValor(RS1!Cliente, 4), "N")

            Conn.Execute SQL

        ' en tmppendientes metemos la clave primaria de cobros_recibidos y el importe en letra
                                                          'importe=nro factura,   codforpa=linea de cobros_realizados
            SQL = "insert into tmppendientes (codusu,serie_cta,importe,fecha,numorden, observa) values ("
            SQL = SQL & vUsu.Codigo & "," & DBSet(RecuperaValor(RS1!Cliente, 1), "T") & "," 'numserie
            SQL = SQL & DBSet(RecuperaValor(RS1!Cliente, 2), "N") & "," 'numfactu
            SQL = SQL & DBSet(RecuperaValor(RS1!Cliente, 3), "F") & "," 'fecfactu
            SQL = SQL & DBSet(RecuperaValor(RS1!Cliente, 4), "N") & "," 'numorden
            SQL = SQL & DBSet(EscribeImporteLetra(ImporteFormateado(CStr(ImporteInterno))), "T") & ") "
            
            Conn.Execute SQL

        End If
    
    End If
    
    
    
    
  '????????????????????????????????????
    '-------------------------------------------------------------------
    'Si es apunte de banco, y hay gastos
    If Cabecera = 2 Then
        'SOOOOLO COBROS
        If Cobros And riesgo > 0 Then
                     
             SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
             SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
             SQL = SQL & " timporteH,  ctacontr,codccost, idcontab, punteada) "
             SQL = SQL & "VALUES (" & vp.diaricli & ",'" & Format(FechaAsiento, FormatoFecha) & "'," & m.Contador & ","
             
             Ampliacion = DevuelveDesdeBD("ctaingreso", "bancos", "codmacta", txtCta(4).Text, "T")
             If Ampliacion = "" Then
                MsgBox "Cta ingreso bancario MAL configurada. Se utilizara la misma del banco", vbExclamation
                Ampliacion = txtCta(4).Text
            End If
            'linea,numdocum,codconce  amconce
            For Conce = 1 To 2
                NumLine = NumLine + 1
                Aux = NumLine & ",'"
                If Conce = 1 Then
                    Aux = Aux & txtCta(4).Text
                Else
                    Aux = Aux & Ampliacion
                End If
                Aux = Aux & "',''," & vp.condecli & ",'" & DevNombreSQL(DevuelveDesdeBD("nomconce", "conceptos", "codconce", vp.condecli)) & "',"
                If Conce = 1 Then
                    Aux = Aux & TransformaComasPuntos(CStr(riesgo)) & ",NULL"
                Else
                    Aux = Aux & "NULL," & TransformaComasPuntos(CStr(riesgo))
                End If
                If Conce = 2 Then
                    Aux = Aux & ",'" & txtCta(4).Text
                Else
                    Aux = Aux & ",'" & Ampliacion
                End If
                Aux = Aux & "',"
                'CC
                If Conce = 1 Then
                    Aux = Aux & "NULL"
                Else
                    If vParam.autocoste Then
                        Ampliacion = DevuelveDesdeBD("codccost", "bancos", "codmacta", txtCta(4).Text, "T")
                        If Ampliacion = "" Then
                            Ampliacion = "NULL"
                        Else
                            Ampliacion = "'" & Ampliacion & "'"
                        End If
                    Else
                        'NO LLEVA ANALITICA
                        Ampliacion = "NULL"
                    End If
                    Aux = Aux & Ampliacion
                End If
                Aux = Aux & ",'COBROS',0)"
                Aux = SQL & Aux
                Ejecuta Aux
            Next Conce
        End If
    End If
    
    
End Function





'----------------------------------------------------------
'   A partir de la tabla tmp
'   Se que cuentas hay y los vencimientos.Por lo tanto, comprobare
'   que si la fechas estan fuera de ejercicios o de ambito
'   y si hay cuentas bloquedas
Private Function ComprobarCuentasBloquedasYFechasVencimientos() As Boolean
    ComprobarCuentasBloquedasYFechasVencimientos = False
    On Error GoTo EComprobarCuentasBloquedasYFechasVencimientos
    Set RS = New ADODB.Recordset
    

    cad = "select codmacta,nommacta,numfactura,fecha,fecbloq,cliente from tmpfaclin,cuentas where codusu=" & vUsu.Codigo & " and cta=codmacta and not (fecbloq is null )"
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    While Not RS.EOF
        If CDate(RS!NumFactura) > RS!FecBloq Then cad = cad & RS!codmacta & "    " & RS!FecBloq & "     " & Format(RS!NumFactura, "dd/mm/yyyy") & Space(15) & RecuperaValor(RS!Cliente, 1) & RecuperaValor(RS!Cliente, 2) & vbCrLf
        RS.MoveNext
    Wend
    RS.Close


    If cad <> "" Then
        cad = vbCrLf & String(90, "-") & vbCrLf & cad
        cad = "Cta           Fec. Bloq            Fecha contab         Factura" & cad
        cad = "Cuentas bloqueadas: " & vbCrLf & vbCrLf & vbCrLf & cad
        MsgBox cad, vbExclamation
    Else
        ComprobarCuentasBloquedasYFechasVencimientos = True
    End If
EComprobarCuentasBloquedasYFechasVencimientos:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set RS = Nothing
End Function





'Busca VTO
' Para no hacer muchos seect WHERE, hacemos un unico SELECT (mirsaux)
' ahora en esta funcion buscaremos el registro correspondiente
'
Private Function BuscarVtoPago(ByRef IT As ListItem) As Boolean
Dim Fin As Boolean
    BuscarVtoPago = False
    Fin = False
    miRsAux.MoveFirst
    While Not Fin
        'numfactu fecfactu numorden
        If miRsAux!ctaprove = IT.Tag Then
            If miRsAux!NumFactu = IT.Text Then
                If miRsAux!FecFactu = IT.SubItems(1) Then
                    If miRsAux!numorden = IT.SubItems(3) Then
                        'ESTE ES
                        BuscarVtoPago = True
                        Fin = True
                    End If
                End If
            End If
        End If
        If Not Fin Then
            miRsAux.MoveNext
            Fin = miRsAux.EOF
        End If
    Wend
End Function

'CREDITO tipo navarres(Forpa 6)
Private Sub ActualizarGastosCobrosTarjetasTipoNavarres()
    
    
    cad = DevuelveDesdeBD("parrafo1", "usuarios.z347carta", "codusu", CStr(vUsu.Codigo))
    impo = Val(cad)
    DescripcionTransferencia = " NºRec:" & cad
    'update z347carta set saludos=trim(concat(coalesce(saludos,''),' ','AAe'))
    
    
    
    For I = 1 To Me.ListView1.ListItems.Count
          If ListView1.ListItems(I).Checked Then
              cad = "UPDATE cobros SET "
              cad = cad & " gastos = " & TransformaComasPuntos(ImporteFormateado(ListView1.ListItems(I).SubItems(8)))
              cad = cad & " ,obs =trim(concat(coalesce(obs,''),' ','" & DescripcionTransferencia & "')) "
              cad = cad & " WHERE numserie = '" & ListView1.ListItems(I).Text
              cad = cad & "' AND numfactu = " & Val(ListView1.ListItems(I).SubItems(1))
              cad = cad & " AND fecfactu = '" & Format(ListView1.ListItems(I).SubItems(2), FormatoFecha)
              cad = cad & "' AND numorden = " & Val(ListView1.ListItems(I).SubItems(4))
              Ejecuta cad
          End If
    Next I

    cad = "1"
    If Fecha <= vParam.fechafin Then cad = "2"
    cad = "UPDATE contadores SET contado" & cad & " =" & Val(impo) & " WHERE tiporegi = 3" 'tarjeta credito tipo NAVARRES
    Ejecuta cad
End Sub






'***********************************************************************************
'***********************************************************************************
'
'   NORMA 57  Pagos por ventanilla
'
'***********************************************************************************
'***********************************************************************************
Private Sub AjustarFechaVencimientoDesdeFicheroBancario()
Dim Fin As Boolean
    'Para cada item buscare en la tabla from tmpconext  WHERE codusu
    Set RS = New ADODB.Recordset
    '(numserie ,codfaccl,fecfaccl,numorden )
    cad = "select ccost,pos,nomdocum,numdiari,fechaent from tmpconext  WHERE codusu =" & vUsu.Codigo & " and numasien=0 "
    cad = cad & " ORDER BY 1,2,3,4"
    RS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
    If RS.EOF Then
        cad = "NINGUN VENCIMIENTO"
    Else
    For I = 1 To Me.ListView1.ListItems.Count
        Fin = False
        RS.MoveFirst
        With ListView1.ListItems(I)
            
            While Not Fin
                'Buscamos el registro... DEBERIA ESTAR
                If RS!CCost = .Text Then
                    If RS!Pos = .SubItems(1) Then
                        If Format(RS!nomdocum, "dd/mm/yyyy") = .SubItems(2) Then
                            If RS!NumDiari = .SubItems(4) Then
                                'Le pongo como fecha de vto la fecha del cobro del fichero
                                Fin = True
                                .SubItems(3) = Format(RS!FechaEnt)
                                .Checked = True
                            End If
                        End If
                    End If
                End If
                If Not Fin Then
                    RS.MoveNext
                    If RS.EOF Then
                        'Ha llegado al final, y no lo ha encotrado
                        cad = cad & "     " & .Text & .SubItems(1) & "  -  " & .SubItems(2) & vbCrLf
                        'Para que vuelva al ppio
                        Fin = True
                    End If
                End If
            Wend
        End With
    Next
    End If
    RS.Close
    
    If cad <> "" Then
        cad = cad & vbCrLf & "El programa continuara con la fecha de vencimiento"
        MsgBox "No se ha encotrado la fecha de cobro para los siguientes vencimientos:" & vbCrLf & cad, vbExclamation
    End If
    Set RS = Nothing
End Sub


Private Sub CargaCombo()
    Combo1.Clear
    'Conceptos
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from tipofpago where tipoformapago <> 4 order by descformapago", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux!descformapago
        Combo1.ItemData(Combo1.NewIndex) = miRsAux!tipoformapago
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    ' cargamos el combo de usuarios
    Combo2.Clear
    'Conceptos
    Set miRsAux = New ADODB.Recordset
    
    Combo2.AddItem ""
    Combo2.ItemData(Combo2.NewIndex) = 1000
    
    miRsAux.Open "Select * from usuarios.usuarios  order by login", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not miRsAux.EOF
        Combo2.AddItem miRsAux!Login
        Combo2.ItemData(Combo2.NewIndex) = miRsAux!codusu
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub

Private Sub HacerToolBar(ButtonIndex As Integer)
    Select Case ButtonIndex
        Case 1
            Generar2
    End Select
End Sub


Private Function EsTalonOPagare(NumSer As String, NumFact As String, FecFact As String, NumOrd As String) As Boolean
Dim SQL As String
Dim Tipo As Byte

    SQL = "select tipforpa from formapago, cobros  where cobros.codforpa = formapago.codforpa and cobros.numserie = " & DBSet(NumSer, "T")
    SQL = SQL & " and numfactu = " & DBSet(NumFact, "N") & " and fecfactu = " & DBSet(FecFact, "F") & " and numorden = " & DBSet(NumOrd, "N")
    
    Tipo = DevuelveValor(SQL)
    EsTalonOPagare = (CByte(Tipo) = 2 Or CByte(Tipo) = 3)

End Function
