VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCCCtaExplo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Ordenaci�n"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   7140
      TabIndex        =   44
      Top             =   4560
      Width           =   4455
      Begin VB.OptionButton optVarios 
         Caption         =   "Cuenta Contable"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   450
         TabIndex        =   46
         Top             =   1500
         Width           =   2415
      End
      Begin VB.OptionButton optVarios 
         Caption         =   "Centro de Coste"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   450
         TabIndex        =   45
         Top             =   810
         Width           =   2715
      End
   End
   Begin VB.Frame frameConceptoDer 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4515
      Left            =   7140
      TabIndex        =   33
      Top             =   0
      Width           =   4455
      Begin VB.Frame FrameCCComparativo 
         Height          =   1065
         Left            =   300
         TabIndex        =   39
         Top             =   1350
         Visible         =   0   'False
         Width           =   3795
         Begin VB.OptionButton optCCComparativo 
            Caption         =   "Saldo"
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
            Left            =   480
            TabIndex        =   41
            Top             =   420
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optCCComparativo 
            Caption         =   "Mes"
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
            Left            =   2340
            TabIndex        =   40
            Top             =   420
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkCtaExpCC 
         Caption         =   "Comparativo"
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
         Index           =   1
         Left            =   300
         TabIndex        =   8
         Top             =   900
         Width           =   1575
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3900
         TabIndex        =   43
         Top             =   270
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
   End
   Begin VB.Frame FrameConcepto 
      Caption         =   "Selecci�n"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4515
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtCuentas 
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
         Left            =   1230
         TabIndex        =   3
         Tag             =   "imgConcepto"
         Top             =   2520
         Width           =   1275
      End
      Begin VB.TextBox txtCuentas 
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
         Index           =   0
         Left            =   1230
         TabIndex        =   2
         Tag             =   "imgConcepto"
         Top             =   2100
         Width           =   1275
      End
      Begin VB.TextBox txtNCuentas 
         BackColor       =   &H80000018&
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
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2520
         Width           =   4215
      End
      Begin VB.TextBox txtNCuentas 
         BackColor       =   &H80000018&
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
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2100
         Width           =   4215
      End
      Begin VB.TextBox txtNCCoste 
         BackColor       =   &H80000018&
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
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   900
         Width           =   4605
      End
      Begin VB.TextBox txtNCCoste 
         BackColor       =   &H80000018&
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
         Index           =   7
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1320
         Width           =   4605
      End
      Begin VB.TextBox txtAno 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3240
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   3810
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
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
         ItemData        =   "frmCCCtaExplo.frx":0000
         Left            =   1200
         List            =   "frmCCCtaExplo.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3780
         Width           =   1935
      End
      Begin VB.TextBox txtAno 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3240
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   3330
         Width           =   855
      End
      Begin VB.ComboBox cmbFecha 
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
         ItemData        =   "frmCCCtaExplo.frx":0004
         Left            =   1200
         List            =   "frmCCCtaExplo.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3330
         Width           =   1935
      End
      Begin VB.TextBox txtCCoste 
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
         Left            =   1230
         TabIndex        =   0
         Top             =   900
         Width           =   795
      End
      Begin VB.TextBox txtCCoste 
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
         Index           =   7
         Left            =   1230
         TabIndex        =   1
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   38
         Top             =   1740
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   37
         Top             =   2130
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   36
         Top             =   2490
         Width           =   615
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   1
         Left            =   930
         Top             =   2550
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   0
         Left            =   930
         Top             =   2100
         Width           =   255
      End
      Begin VB.Image imgCCoste 
         Height          =   255
         Index           =   6
         Left            =   930
         Top             =   900
         Width           =   255
      End
      Begin VB.Image imgCCoste 
         Height          =   255
         Index           =   7
         Left            =   930
         Top             =   1350
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   29
         Top             =   1290
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   28
         Top             =   930
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   27
         Top             =   3750
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   26
         Top             =   3390
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Centro de Coste"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   25
         Top             =   540
         Width           =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "Per�odo de C�lculo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   210
         TabIndex        =   24
         Top             =   3000
         Width           =   2790
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
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
      Left            =   10320
      TabIndex        =   11
      Top             =   7410
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccion 
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
      Index           =   1
      Left            =   8730
      TabIndex        =   9
      Top             =   7410
      Width           =   1455
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Imprimir"
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
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Frame FrameTipoSalida 
      Caption         =   "Tipo de salida"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   6915
      Begin VB.CommandButton PushButtonImpr 
         Caption         =   "Propiedades"
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
         Left            =   5190
         TabIndex        =   23
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   22
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   21
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox txtTipoSalida 
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
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1680
         Width           =   4665
      End
      Begin VB.TextBox txtTipoSalida 
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
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1200
         Width           =   4665
      End
      Begin VB.TextBox txtTipoSalida 
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
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   720
         Width           =   3345
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "eMail"
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
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Width           =   975
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "PDF"
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
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "Archivo csv"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1515
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "Impresora"
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
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancelarAccion 
      Caption         =   "CANCEL"
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
      Left            =   10320
      TabIndex        =   30
      Top             =   7410
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Label15"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1890
      TabIndex        =   42
      Top             =   7350
      Visible         =   0   'False
      Width           =   5985
   End
End
Attribute VB_Name = "frmCCCtaExplo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 1003

Public opcion As Byte
' ***********************************************************************************************************
' ***********************************************************************************************************
' ***********************************************************************************************************
'
'  3 espacios
'       -Los desde hasta,
'       -las opciones / ordenacion
'       -el tipo salida
'
' ***********************************************************************************************************
' ***********************************************************************************************************
' ***********************************************************************************************************


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmDia As frmTiposDiario
Attribute frmDia.VB_VarHelpID = -1
Private WithEvents frmCCoste  As frmBasico
Attribute frmCCoste.VB_VarHelpID = -1
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1


Private Sql As String
Dim Cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim PrimeraVez As String
Dim Rs As ADODB.Recordset

Dim FechaIncioEjercicio As Date
Dim FechaFinEjercicio As Date
Dim PulsadoCancelar As Boolean

Public Legalizacion As String   'Datos para la legalizacion

Dim HanPulsadoSalir As Boolean
Dim FechaInicio As String
Dim fechafin As String

Public Sub InicializarVbles(A�adireElDeEmpresa As Boolean)
    cadFormula = ""
    cadselect = ""
    cadParam = "|"
    numParam = 0
    cadNomRPT = ""
    conSubRPT = False
    cadPDFrpt = ""
    ExportarPDF = False
    vMostrarTree = False
    
    If A�adireElDeEmpresa Then
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
    
End Sub




Private Sub chkCtaExpCC_Click(Index As Integer)
    If Index = 1 Then
         FrameCCComparativo.Visible = chkCtaExpCC(1).Value = 1
    End If
End Sub

Private Sub chkCtaExpCC_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAccion_Click(Index As Integer)
    
    If Not DatosOK Then Exit Sub
    
    PulsadoCancelar = False
    Me.cmdCancelarAccion.Visible = True
    Me.cmdCancelarAccion.Enabled = True
    
    Me.cmdCancelar.Visible = False
    Me.cmdCancelar.Enabled = False
        
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    InicializarVbles True
    
    
'    If Not PonerDesdeHasta("hlinapu.codccost", "CCO", Me.txtCCoste(6), Me.txtNCCoste(6), Me.txtCCoste(7), Me.txtNCCoste(7), "pDHCoste=""") Then Exit Sub
    
    If Not MontaSQL Then Exit Sub

    Me.cmdCancelarAccion.Visible = False
    Me.cmdCancelarAccion.Enabled = False
    
    Me.cmdCancelar.Visible = True
    Me.cmdCancelar.Enabled = True
    
    
    If Not HayRegParaInforme("tmplinccexplo", "codusu=" & vUsu.Codigo) Then Exit Sub
    
    If optTipoSal(1).Value Then
        'EXPORTAR A CSV
        AccionesCSV
    
    Else
        'Tanto a pdf,imprimiir, preevisualizar como email van COntral Crystal
        If optTipoSal(2).Value Or optTipoSal(3).Value Then
            ExportarPDF = True 'generaremos el pdf
        Else
            ExportarPDF = False
        End If
        SoloImprimir = False
        If Index = 0 Then SoloImprimir = True 'ha pulsado impirmir
        
        AccionesCrystal
    End If
    
    If Legalizacion <> "" Then
        CadenaDesdeOtroForm = "OK"
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    If Me.cmdCancelarAccion.Visible Then Exit Sub
    HanPulsadoSalir = True
    Unload Me
End Sub


Private Sub cmdCancelarAccion_Click()
    PulsadoCancelar = True
End Sub

Private Sub Form_Activate()
Dim Cont As Integer

    If PrimeraVez Then
        PrimeraVez = False
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmPpal.Icon
        
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
        
    'Otras opciones
    Me.Caption = "Cuenta de Explotaci�n Anal�tica"

    For i = 6 To 7
        Me.imgCCoste(i).Picture = frmPpal.ImageList3.ListImages(1).Picture
    Next i
    For i = 0 To 1
        Me.imgCuentas(i).Picture = frmPpal.ImageList3.ListImages(1).Picture
    Next i
    
    PrimeraVez = True
     
     
    CargarComboFecha
     
    cmbFecha(0).ListIndex = Month(vParam.fechaini) - 1
    cmbFecha(1).ListIndex = Month(vParam.fechafin) - 1

    txtAno(0).Text = Year(vParam.fechaini)
    txtAno(1).Text = Year(vParam.fechafin)
   
 
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    Me.optVarios(0).Value = True
    
    cmdCancelarAccion.Enabled = False
    cmdCancelarAccion.Visible = False
    
    
    
End Sub

Private Sub frmCCoste_DatoSeleccionado(CadenaSeleccion As String)
    txtCCoste(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNCCoste(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    txtCuentas(RC).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNCuentas(RC).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub ImgCCoste_Click(Index As Integer)

    IndCodigo = Index
    
    Set frmCCoste = New frmBasico
    AyudaCC frmCCoste
    Set frmCCoste = Nothing
    
    PonFoco txtCCoste(Index)

End Sub


Private Sub imgCuentas_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmCta = New frmColCtas
    RC = Index
    frmCta.DatosADevolverBusqueda = "0|1"
    frmCta.ConfigurarBalances = 3
    frmCta.Show vbModal
    Set frmCta = Nothing
End Sub

Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
End Sub

Private Sub PushButton2_Click(Index As Integer)
    'FILTROS
    If Index = 0 Then
        frmPpal.cd1.Filter = "*.csv|*.csv"
         
    Else
        frmPpal.cd1.Filter = "*.pdf|*.pdf"
    End If
    frmPpal.cd1.InitDir = App.Path & "\Exportar" 'PathSalida
    frmPpal.cd1.FilterIndex = 1
    frmPpal.cd1.ShowSave
    If frmPpal.cd1.FileTitle <> "" Then
        If Dir(frmPpal.cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo ya existe. Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        txtTipoSalida(Index + 1).Text = frmPpal.cd1.FileName
    End If
End Sub

Private Sub PushButtonImpr_Click()
    frmPpal.cd1.ShowPrinter
    PonerDatosPorDefectoImpresion Me, True
End Sub



Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub


Private Sub txtAno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCCoste_GotFocus(Index As Integer)
    ConseguirFoco txtCCoste(Index), 3
End Sub


Private Sub txtCCoste_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0

        LanzaFormAyuda "imgCCoste", Index
    End If
End Sub


Private Sub LanzaFormAyuda(Nombre As String, indice As Integer)
    Select Case Nombre
        Case "imgCCoste"
            ImgCCoste_Click indice
        Case "imgCuentas"
            imgCuentas_Click indice
    End Select
    
End Sub

Private Sub txtCCoste_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim RC As String
Dim Hasta As Integer

    txtCCoste(Index).Text = Trim(txtCCoste(Index).Text)
    
    Select Case Index
        Case 6, 7 'Centros de Coste
            If txtCCoste(Index).Text <> "" Then txtCCoste(Index).Text = UCase(txtCCoste(Index).Text)
            txtNCCoste(Index) = PonerNombreDeCod(txtCCoste(Index), "ccoste", "nomccost", "codccost", "T")
            
    End Select

'    PierdeFocoTiposDiario Me.txtTiposDiario(Index), Me.lblTiposDiario(Index)
End Sub



Private Sub AccionesCSV()
Dim Sql2 As String
Dim Tipo As Byte

    If chkCtaExpCC(1).Value Then 'comparativo
        If optCCComparativo(0).Value Then 'saldo
            If optVarios(0).Value Then
                Sql = "select tt.codccost CC, cc.nomccost Nombre, tt.codmacta Cuenta, cu.nommacta Descripcion,tt.antD AntDebe,tt.antH AntHaber, tt.perD PeriodoDebe,tt.perH PeriodoHaber "
                Sql = Sql & " from tmplinccexplo tt, ccoste cc, cuentas cu where tt.codusu = " & vUsu.Codigo
                Sql = Sql & " and tt.codccost = cc.codccost and tt.codmacta = cu.codmacta "
                Sql = Sql & " order by 1,2,3,4 "
            Else
                Sql = "select tt.codmacta Cuenta, cu.nommacta Descripcion, tt.codccost CC, cc.nomccost Nombre, tt.antD AntDebe,tt.antH AntHaber, tt.perD PeriodoDebe,tt.perH PeriodoHaber "
                Sql = Sql & " from tmplinccexplo tt, ccoste cc, cuentas cu where tt.codusu = " & vUsu.Codigo
                Sql = Sql & " and tt.codccost = cc.codccost and tt.codmacta = cu.codmacta "
                Sql = Sql & " order by 1,2,3,4 "
            End If
        Else 'mes
            If optVarios(0).Value Then
                Sql = "select tt.codccost CC, cc.nomccost Nombre, tt.codmacta Cuenta, cu.nommacta Descripcion,tt.mes,tt.anyo, tt.antD AntDebe,tt.antH AntHaber, tt.perD PeriodoDebe,tt.perH PeriodoHaber "
                Sql = Sql & " from tmplinccexplo tt, ccoste cc, cuentas cu where tt.codusu = " & vUsu.Codigo
                Sql = Sql & " and tt.codccost = cc.codccost and tt.codmacta = cu.codmacta "
                Sql = Sql & " order by 1,2,3,4,5,6 "
            Else
                Sql = "select tt.codmacta Cuenta, cu.nommacta Descripcion, tt.codccost CC, cc.nomccost Nombre,tt.mes,tt.anyo, tt.antD AntDebe,tt.antH AntHaber, tt.perD PeriodoDebe,tt.perH PeriodoHaber "
                Sql = Sql & " from tmplinccexplo tt, ccoste cc, cuentas cu where tt.codusu = " & vUsu.Codigo
                Sql = Sql & " and tt.codccost = cc.codccost and tt.codmacta = cu.codmacta "
                Sql = Sql & " order by 1,2,3,4,5,6 "
            End If
        End If
    Else
        If optVarios(0).Value Then
            Sql = "select tt.codccost CC, cc.nomccost Nombre, tt.codmacta Cuenta, cu.nommacta Descripcion, tt.perD PeriodoDebe,tt.perH PeriodoHaber, tt.antD AntDebe,tt.antH AntHaber,tt.mes,tt.anyo "
            Sql = Sql & " from tmplinccexplo tt, ccoste cc, cuentas cu where tt.codusu = " & vUsu.Codigo
            Sql = Sql & " and tt.codccost = cc.codccost and tt.codmacta = cu.codmacta "
            Sql = Sql & " order by 1,2,3,4 "
        Else
            Sql = "select tt.codmacta Cuenta, cu.nommacta Descripcion, tt.codccost CC, cc.nomccost Nombre, tt.perD PeriodoDebe,tt.perH PeriodoHaber, tt.antD AntDebe,tt.antH AntHaber,tt.mes,tt.anyo "
            Sql = Sql & " from tmplinccexplo tt, ccoste cc, cuentas cu where tt.codusu = " & vUsu.Codigo
            Sql = Sql & " and tt.codccost = cc.codccost and tt.codmacta = cu.codmacta "
            Sql = Sql & " order by 1,2,3,4 "
        End If
    End If
        
            
    'LLamos a la funcion
    GeneraFicheroCSV Sql, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim Tipo As Byte
Dim UltimoNivel As Integer
Dim indRPT As String
Dim nomDocu As String

    
    cadParam = cadParam & "pDHFecha=""" & cmbFecha(0).Text & " " & txtAno(0).Text & " a " & cmbFecha(1).Text & " " & txtAno(1).Text & """|"
    numParam = numParam + 1
    
    
    vMostrarTree = False
    conSubRPT = False
        
    indRPT = "1003-00"
    
    If chkCtaExpCC(1).Value = 1 Then
        indRPT = "1003-01" ' comparativo
        If optCCComparativo(0).Value Then cadParam = cadParam & "pPorMeses=0|"
        If optCCComparativo(1).Value Then cadParam = cadParam & "pPorMeses=1|"
        numParam = numParam + 1
    End If
    
    If optVarios(0).Value Then
        cadParam = cadParam & "pGrupo1={tmplinccexplo.codccost}|"
        cadParam = cadParam & "pGrupo2={tmplinccexplo.codmacta}|"
    Else
        cadParam = cadParam & "pGrupo1={tmplinccexplo.codmacta}|"
        cadParam = cadParam & "pGrupo2={tmplinccexplo.codccost}|"
    End If
    numParam = numParam + 4
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu '"SumasySaldos.rpt"

    cadFormula = "{tmplinccexplo.codusu}=" & vUsu.Codigo

    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, (Legalizacion <> "")
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 2
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Function MontaSQL() As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RC As String
Dim RC2 As String

    MontaSQL = GeneraCtaExplotacionCC
    
           
End Function


Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    If txtCCoste(6).Text <> "" And txtCCoste(7).Text <> "" Then
        If txtCCoste(6).Text > txtCCoste(7).Text Then
            MsgBox "Centro de coste inicio mayor que centro de coste fin", vbExclamation
            Exit Function
        End If
    End If
    
    If txtAno(0).Text = "" Or txtAno(1).Text = "" Then
        MsgBox "Introduce las fechas(a�os) de consulta", vbExclamation
        Exit Function
    End If
    
    If Not ComparaFechasCombos(0, 1, 0, 1) Then Exit Function
     
    
    'Comprobamos que el total de meses no supera el a�o
    i = Val(txtAno(0).Text)
    Cont = Val(txtAno(1).Text)
    Cont = Cont - i
    i = 0
    If Cont > 1 Then
       i = 1  'Ponemos a uno para luego salir del bucle
    Else
        If Cont = 1 Then
            'Se diferencian un a�o, luego el mes fin tienes k ser menor a mes inicio
            If Me.cmbFecha(1).ListIndex >= Me.cmbFecha(0).ListIndex Then i = 1
        End If
    End If
    If i <> 0 Then
        MsgBox "El intervalo tiene que ser de un a�o como m�ximo", vbExclamation
        Exit Function
    End If

    DatosOK = True

End Function


Private Sub CargarComboFecha()
Dim J As Integer

QueCombosFechaCargar "0|1|"

End Sub


Private Sub QueCombosFechaCargar(Lista As String)
Dim L As Integer

L = 1
Do
    Cad = RecuperaValor(Lista, L)
    If Cad <> "" Then
        i = Val(Cad)
        With cmbFecha(i)
            .Clear
            For Cont = 1 To 12
                RC = "25/" & Cont & "/2002"
                RC = Format(RC, "mmmm") 'Devuelve el mes
                .AddItem RC
            Next Cont
        End With
    End If
    L = L + 1
Loop Until Cad = ""
End Sub


Private Function ComprobarCuentas(Indice1 As Integer, Indice2 As Integer) As Boolean
Dim L1 As Integer
Dim L2 As Integer
    ComprobarCuentas = False
    If txtCCoste(Indice1).Text <> "" And txtCCoste(Indice2).Text <> "" Then
        L1 = Len(txtCCoste(Indice1).Text)
        L2 = Len(txtCCoste(Indice2).Text)
        If L1 > L2 Then
            L2 = L1
        Else
            L1 = L2
        End If
        If Val(Mid(txtCCoste(Indice1).Text & "000000000", 1, L1)) > Val(Mid(txtCCoste(Indice2).Text & "0000000000", 1, L1)) Then
            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
            Exit Function
        End If
    End If
    ComprobarCuentas = True
End Function

Private Function ComparaFechasCombos(Indice1 As Integer, Indice2 As Integer, InCombo1 As Integer, InCombo2 As Integer) As Boolean
    ComparaFechasCombos = False
    If txtAno(Indice1).Text <> "" And txtAno(Indice2).Text <> "" Then
        If Val(txtAno(Indice1).Text) > Val(txtAno(Indice2).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Function
        Else
            If Val(txtAno(Indice1).Text) = Val(txtAno(Indice2).Text) Then
                If Me.cmbFecha(InCombo1).ListIndex > Me.cmbFecha(InCombo2).ListIndex Then
                    MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
                    Exit Function
                End If
            End If
        End If
    End If
    ComparaFechasCombos = True
End Function


'Siempre k la fecha no este en fecha siguiente
Private Function HayAsientoCierre(Mes As Byte, Anyo As Integer, Optional Contabilidad As String) As Boolean
Dim C As String
    HayAsientoCierre = False
    'C = "01/" & CStr(Me.cmbFecha(1).ListIndex + 1) & "/" & txtAno(1).Text
    C = "01/" & CStr(Mes) & "/" & Anyo
    'Si la fecha es menor k la fecha de inicio de ejercicio entonces SI k hay asiento de cierre
    If CDate(C) < vParam.fechaini Then
        HayAsientoCierre = True
    Else
        If CDate(C) > vParam.fechafin Then
            'Seguro k no hay
            Exit Function
        Else
            C = "Select count(*) from " & Contabilidad
            C = C & " hlinapu where (codconce=960 or codconce = 980) and fechaent>='" & Format(vParam.fechaini, FormatoFecha)
            C = C & "' AND fechaent <='" & Format(vParam.fechafin, FormatoFecha) & "'"
            Rs.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                If Not IsNull(Rs.Fields(0)) Then
                    If Rs.Fields(0) > 0 Then HayAsientoCierre = True
                End If
            End If
            Rs.Close
        End If
    End If
End Function



Private Sub txtCuentas_GotFocus(Index As Integer)
    ConseguirFoco txtCuentas(Index), 3
End Sub


Private Sub txtCuentas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0

        LanzaFormAyuda "imgCuentas", Index
    End If
End Sub





Private Sub txtCuentas_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim RC As String
Dim Hasta As Integer

    txtCuentas(Index).Text = Trim(txtCuentas(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    txtCuentas(Index).Text = Trim(txtCuentas(Index).Text)
    If txtCuentas(Index).Text = "" Then
        txtNCuentas(Index).Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtCuentas(Index).Text) Then
        If InStr(1, txtCuentas(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser num�rica: " & txtCuentas(Index).Text, vbExclamation
        txtCuentas(Index).Text = ""
        txtNCuentas(Index).Text = ""
        Exit Sub
    End If



    Select Case Index
        Case 0, 1 'Cuentas
            'lblCuentas(Index).Caption = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", txtCuentas(Index), "T")
            
            RC = txtCuentas(Index).Text
            If CuentaCorrectaUltimoNivelSIN(RC, Sql) Then
                txtCuentas(Index) = RC
                txtNCuentas(Index).Text = Sql
            Else
                MsgBox Sql, vbExclamation
                txtCuentas(Index).Text = ""
                txtNCuentas(Index).Text = ""
                PonFoco txtCuentas(Index)
            End If
            
            If Index = 0 Then Hasta = 1
            If Hasta >= 1 Then
                txtCuentas(Hasta).Text = txtCuentas(Index).Text
                txtNCuentas(Hasta).Text = txtNCuentas(Index).Text
            End If
    End Select

'    PierdeFocoTiposDiario Me.txtTiposDiario(Index), Me.lblTiposDiario(Index)
End Sub


Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtTitulo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub


Private Function GeneraCtaExplotacionCC() As Boolean
Dim RC As Byte
Dim FIniP As Date
Dim FFinP As Date
Dim FIniPAnt As Date
Dim FFinPAnt As Date
Dim CadInsert As String

    On Error GoTo EGeneraCtaExplotacionCC

    GeneraCtaExplotacionCC = False
    
    'Borramos datos
    Sql = "Delete from tmplinccexplo where codusu = " & vUsu.Codigo
    Conn.Execute Sql
    
    FIniP = "01/" & Format(cmbFecha(0).ListIndex + 1, "00") & "/" & txtAno(0).Text
    FFinP = DateAdd("d", -1, DateAdd("m", 1, "01/" & Format(cmbFecha(1).ListIndex + 1, "00") & "/" & txtAno(1).Text))
    FIniPAnt = "01/01/1900"
    FFinPAnt = "01/01/1900"
    If chkCtaExpCC(1).Value Then
        FIniPAnt = DateAdd("yyyy", -1, FIniP)
        FFinPAnt = DateAdd("yyyy", -1, FFinP)
    Else
        If FIniP <> vParam.fechaini Then
            If cmbFecha(0).ListIndex + 1 < Month(vParam.fechaini) Then
                FIniPAnt = "01/" & Format(Month(vParam.fechaini), "00") & "/" & Format(CLng(txtAno(0).Text) - 1, "0000")
            Else
                FIniPAnt = "01/" & Format(Month(vParam.fechaini), "00") & "/" & Format(CLng(txtAno(0).Text), "0000")
            End If
            FFinPAnt = DateAdd("d", -1, FIniP)
        End If
    End If

    Label15.Visible = True


    If chkCtaExpCC(1).Value = 1 Then
        If optCCComparativo(1).Value Then ' por meses
        
            Sql = "insert into tmplinccexplo (codusu,codccost,codmacta, mes, anyo, perD,perH) "
            Sql = Sql & " select " & vUsu.Codigo & ", codccost, codmacta, month(fechaent) mes, year(fechaent) anyo, sum(coalesce(timported,0)), sum(coalesce(timporteh,0))  "
            Sql = Sql & " FROM hlinapu  "
            Sql = Sql & " where mid(hlinapu.codmacta,1,1) IN (" & DBSet(vParam.grupogto, "T") & "," & DBSet(vParam.grupovta, "T") & ")"
            Sql = Sql & " and fechaent between " & DBSet(FIniP, "F") & " and " & DBSet(FFinP, "F")
            Sql = Sql & " and not codccost is null and codccost <> '' "
            If cadselect <> "" Then Sql = Sql & " and " & cadselect
            Sql = Sql & " group by 1,2,3,4,5 "
            Sql = Sql & " ORDER BY 1,2,3,4,5 "
        
            Conn.Execute Sql
        
            Label15.Caption = "Insertando periodo por meses comparativo"
            Me.Refresh
        
            CadInsert = "insert into tmplinccexplo (codusu,codccost,codmacta,mes,anyo,AntD,AntH) values ("
            Sql = "select " & vUsu.Codigo & ", codccost, codmacta, month(fechaent) mes, year(fechaent) anyo, sum(coalesce(timported,0)) impd, sum(coalesce(timporteh,0)) imph  "
            Sql = Sql & " FROM hlinapu  "
            Sql = Sql & " where mid(hlinapu.codmacta,1,1) IN (" & DBSet(vParam.grupogto, "T") & "," & DBSet(vParam.grupovta, "T") & ")"
            Sql = Sql & " and fechaent between " & DBSet(FIniPAnt, "F") & " and " & DBSet(FFinPAnt, "F")
            Sql = Sql & " and not codccost is null and codccost <> '' "
            If cadselect <> "" Then Sql = Sql & " and " & cadselect
            Sql = Sql & " group by 1,2,3,4,5 "
            Sql = Sql & " ORDER BY 1,2,3,4,5 "
            
            Label15.Caption = "Insertando periodo anterior por meses comparativo"
            Me.Refresh
            
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                Sql = "select count(*) from tmplinccexplo where codusu = " & vUsu.Codigo & " and codccost = " & DBSet(Rs!codccost, "T") & " and codmacta = " & DBSet(Rs!codmacta, "T")
                Sql = Sql & " and mes = " & DBSet(Rs!Mes, "N")
                Sql = Sql & " and anyo = " & DBSet(Rs!Anyo, "N")
                If TotalRegistros(Sql) = 0 Then
                    Sql = CadInsert & DBSet(vUsu.Codigo, "N") & "," & DBSet(Rs!codccost, "T") & "," & DBSet(Rs!codmacta, "T") & "," & DBSet(Rs!Mes, "N") & "," & DBSet(Rs!Anyo, "N") & "," & DBSet(Rs!ImpD, "N") & "," & DBSet(Rs!ImpH, "N") & ")"
                Else
                    Sql = "update tmplinccexplo set antd = " & DBSet(Rs!ImpD, "N") & ", anth = " & DBSet(Rs!ImpH, "N")
                    Sql = Sql & " where codusu = " & vUsu.Codigo & " and codccost =  " & DBSet(Rs!codccost, "T") & " and codmacta = " & DBSet(Rs!codmacta, "T")
                    Sql = Sql & " and mes = " & DBSet(Rs!Mes, "N")
                    Sql = Sql & " and anyo = " & DBSet(Rs!Anyo, "N")
                End If
                
                Conn.Execute Sql
                
                Rs.MoveNext
            Wend
            Set Rs = Nothing
            
            B = HacerRepartoSubcentrosCoste
            
            GeneraCtaExplotacionCC = True
            Label15.Visible = False
            Exit Function
        End If
    End If
    
    Sql = "insert into tmplinccexplo (codusu,codccost,codmacta,perD,perH) "
    Sql = Sql & " select " & vUsu.Codigo & ", codccost, codmacta, sum(coalesce(timported,0)), sum(coalesce(timporteh,0))  "
    Sql = Sql & " FROM hlinapu  "
    Sql = Sql & " where mid(hlinapu.codmacta,1,1) IN (" & DBSet(vParam.grupogto, "T") & "," & DBSet(vParam.grupovta, "T") & ")"
    Sql = Sql & " and fechaent between " & DBSet(FIniP, "F") & " and " & DBSet(FFinP, "F")
    Sql = Sql & " and not codccost is null and codccost <> '' "
    If cadselect <> "" Then Sql = Sql & " and " & cadselect
    Sql = Sql & " group by 1,2,3 "
    Sql = Sql & " ORDER BY 1,2,3 "

    Conn.Execute Sql

    Label15.Caption = "Insertando periodo"
    Me.Refresh


    ' si el periodo no coincide con el inicio de ejercicio, grabamos el acumulado anterior
    If FIniP <> vParam.fechaini Or chkCtaExpCC(1).Value = 1 Then
        CadInsert = "insert into tmplinccexplo (codusu,codccost,codmacta,AntD,AntH) values ("
        Sql = "select " & vUsu.Codigo & ", codccost, codmacta, sum(coalesce(timported,0)) impd, sum(coalesce(timporteh,0)) imph  "
        Sql = Sql & " FROM hlinapu  "
        Sql = Sql & " where mid(hlinapu.codmacta,1,1) IN (" & DBSet(vParam.grupogto, "T") & "," & DBSet(vParam.grupovta, "T") & ")"
        Sql = Sql & " and fechaent between " & DBSet(FIniPAnt, "F") & " and " & DBSet(FFinPAnt, "F")
        Sql = Sql & " and not codccost is null and codccost <> '' "
        If cadselect <> "" Then Sql = Sql & " and " & cadselect
        Sql = Sql & " group by 1,2,3 "
        Sql = Sql & " ORDER BY 1,2,3 "
        
        Label15.Caption = "Insertando periodo anterior"
        Me.Refresh
        
        
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            Sql = "select count(*) from tmplinccexplo where codusu = " & vUsu.Codigo & " and codccost = " & DBSet(Rs!codccost, "T") & " and codmacta = " & DBSet(Rs!codmacta, "T")
            If TotalRegistros(Sql) = 0 Then
                Sql = CadInsert & DBSet(vUsu.Codigo, "N") & "," & DBSet(Rs!codccost, "T") & "," & DBSet(Rs!codmacta, "T") & "," & DBSet(Rs!ImpD, "N") & "," & DBSet(Rs!ImpH, "N") & ")"
            Else
                Sql = "update tmplinccexplo set antd = " & DBSet(Rs!ImpD, "N") & ", anth = " & DBSet(Rs!ImpH, "N")
                Sql = Sql & " where codusu = " & vUsu.Codigo & " and codccost =  " & DBSet(Rs!codccost, "T") & " and codmacta = " & DBSet(Rs!codmacta, "T")
            End If
            
            Conn.Execute Sql
            
            Rs.MoveNext
        Wend
        Set Rs = Nothing
    End If
    
    
    B = HacerRepartoSubcentrosCoste
    
    
    GeneraCtaExplotacionCC = True
    Label15.Visible = False
    Exit Function

EGeneraCtaExplotacionCC:
    Label15.Visible = False
    MuestraError Err.Number, "Genera Cuenta Explotacion", Err.Description

End Function

