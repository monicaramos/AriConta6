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
      Height          =   7215
      Left            =   7140
      TabIndex        =   35
      Top             =   0
      Width           =   4455
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
         Index           =   7
         ItemData        =   "frmCCCtaExplo.frx":0000
         Left            =   1980
         List            =   "frmCCCtaExplo.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   1050
         Width           =   2055
      End
      Begin VB.Frame FrameCCComparativo 
         Height          =   1065
         Left            =   270
         TabIndex        =   42
         Top             =   3240
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
            TabIndex        =   44
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
            TabIndex        =   43
            Top             =   420
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkCtaExpCC 
         Caption         =   "Ver movimientos posteriores"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1740
         Width           =   3795
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
         Left            =   240
         TabIndex        =   10
         Top             =   2790
         Width           =   1575
      End
      Begin VB.CheckBox chkCtaExpCC 
         Caption         =   "Solo mostrar subcentros de reparto"
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
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   2250
         Width           =   4005
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3900
         TabIndex        =   47
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
      Begin VB.Label Label3 
         Caption         =   "Mes de cálculo"
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
         Index           =   10
         Left            =   270
         TabIndex        =   36
         Top             =   1110
         Width           =   1410
      End
   End
   Begin VB.Frame FrameConcepto 
      Caption         =   "Selección"
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
      TabIndex        =   22
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         ItemData        =   "frmCCCtaExplo.frx":0004
         Left            =   1200
         List            =   "frmCCCtaExplo.frx":0006
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
         ItemData        =   "frmCCCtaExplo.frx":0008
         Left            =   1200
         List            =   "frmCCCtaExplo.frx":000A
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   540
         Width           =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "Mes / Año"
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
         TabIndex        =   26
         Top             =   3030
         Width           =   1410
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
      TabIndex        =   13
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
      TabIndex        =   11
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
      TabIndex        =   12
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
      TabIndex        =   14
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
         TabIndex        =   25
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   24
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   23
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
      TabIndex        =   32
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
      TabIndex        =   45
      Top             =   7350
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


Private SQL As String
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

Public Sub InicializarVbles(AñadireElDeEmpresa As Boolean)
    cadFormula = ""
    cadselect = ""
    cadParam = "|"
    numParam = 0
    cadNomRPT = ""
    conSubRPT = False
    cadPDFrpt = ""
    ExportarPDF = False
    vMostrarTree = False
    
    If AñadireElDeEmpresa Then
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
    
    
    If Not HayRegParaInforme("tmpsaldoscc", "codusu=" & vUsu.Codigo) Then Exit Sub
    
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
    Me.Caption = "Cuenta de Explotación Analítica"

    For i = 6 To 7
        Me.imgCCoste(i).Picture = frmPpal.ImageList3.ListImages(1).Picture
    Next i
    
    PrimeraVez = True
     
     
    CargarComboFecha
     
    cmbFecha(0).ListIndex = Month(vParam.fechaini) - 1
    cmbFecha(1).ListIndex = Month(vParam.fechafin) - 1

    txtAno(0).Text = Year(vParam.fechaini)
    txtAno(1).Text = Year(vParam.fechafin)
   
 
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
    
    cmdCancelarAccion.Enabled = False
    cmdCancelarAccion.Visible = False
    
    
    
End Sub

Private Sub frmCCoste_DatoSeleccionado(CadenaSeleccion As String)
    txtCCoste(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNCCoste(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
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
            txtNCCoste(Index) = PonerNombreDeCod(txtCCoste(Index), "ccoste", "nomccost", "codccost", "T")
            
    End Select

'    PierdeFocoTiposDiario Me.txtTiposDiario(Index), Me.lblTiposDiario(Index)
End Sub



Private Sub AccionesCSV()
Dim SQL2 As String
Dim Tipo As Byte

    SQL = "select CCoste Cuenta , nomCCoste Titulo, aperturad, aperturah, case when coalesce(aperturad,0) - coalesce(aperturah,0) > 0 then concat(coalesce(aperturad,0) - coalesce(aperturah,0),'D') when coalesce(aperturad,0) - coalesce(aperturah,0) < 0 then concat(coalesce(aperturah,0) - coalesce(aperturad,0),'H') when coalesce(aperturad,0) - coalesce(aperturah,0) = 0 then 0 end Apertura, "
    SQL = SQL & " acumantd AcumAnt_deudor, acumanth AcumAnt_acreedor, acumperd AcumPer_deudor, acumperh AcumPer_acreedor, "
    SQL = SQL & " totald Saldo_deudor, totalh Saldo_acreedor, case when coalesce(totald,0) - coalesce(totalh,0) > 0 then concat(coalesce(totald,0) - coalesce(totalh,0),'D') when coalesce(totald,0) - coalesce(totalh,0) < 0 then concat(coalesce(totalh,0) - coalesce(totald,0),'H') when coalesce(totald,0) - coalesce(totalh,0) = 0 then 0 end Saldo"
    SQL = SQL & " from tmpbalancesumas where codusu = " & vUsu.Codigo
    SQL = SQL & " order by 1 "
        
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim Tipo As Byte
Dim UltimoNivel As Integer
Dim indRPT As String
Dim nomDocu As String


    cadParam = cadParam & "pTipo=" & Tipo & "|"
    numParam = numParam + 1
    
    
    cadParam = cadParam & "pDHFecha=""" & cmbFecha(0).Text & " " & txtAno(0).Text & " a " & cmbFecha(1).Text & " " & txtAno(1).Text & """|"
    numParam = numParam + 1
    
    
    vMostrarTree = False
    conSubRPT = False
        
    indRPT = "1002-00"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu '"SumasySaldos.rpt"

    cadFormula = "{tmpsaldoscc.codusu}=" & vUsu.Codigo

    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, (Legalizacion <> "")
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 2
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Function MontaSQL() As Boolean
Dim SQL As String
Dim SQL2 As String
Dim RC As String
Dim RC2 As String

    MontaSQL = GeneraCtaExplotacionCC
    
           
End Function

Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    If txtCCost(2).Text <> "" And txtCCost(3).Text <> "" Then
        If txtCCost(2).Text > txtCCost(3).Text Then
            MsgBox "Centro de coste inicio mayor que centro de coste fin", vbExclamation
            Exit Function
        End If
    End If
    
    If txtAno(0).Text = "" Or txtAno(1).Text = "" Then
        MsgBox "Introduce las fechas(años) de consulta", vbExclamation
        Exit Function
    End If
    
    If Me.cmbFecha(7).ListIndex < 0 Then
        MsgBox "Seleccione un mes de cálculo", vbExclamation
        Exit Function
    End If
    
    If Not ComparaFechasCombos(0, 1, 0, 1) Then Exit Sub
     
    
    'Comprobamos que el total de meses no supera el año
    i = Val(txtAno(0).Text)
    Cont = Val(txtAno(1).Text)
    Cont = Cont - i
    i = 0
    If Cont > 1 Then
       i = 1  'Ponemos a uno para luego salir del bucle
    Else
        If Cont = 1 Then
            'Se diferencian un año, luego el mes fin tienes k ser menor a mes inicio
            If Me.cmbFecha(1).ListIndex >= Me.cmbFecha(0).ListIndex Then i = 1
        End If
    End If
    If i <> 0 Then
        MsgBox "El intervalo tiene que ser de un año como máximo", vbExclamation
        Exit Function
    End If


    'No puede pedir movimientos posteriores y comparativo
    If chkCtaExpCC(0).Value = 1 And chkCtaExpCC(1).Value = 1 Then
        MsgBox "No puede pedir comparativo y movimientos posteriores", vbExclamation
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
        If InStr(1, txtCuentas(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser numérica: " & txtCuentas(Index).Text, vbExclamation
        txtCuentas(Index).Text = ""
        txtNCuentas(Index).Text = ""
        Exit Sub
    End If



    Select Case Index
        Case 0, 1 'Cuentas
            'lblCuentas(Index).Caption = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", txtCuentas(Index), "T")
            
            RC = txtCuentas(Index).Text
            If CuentaCorrectaUltimoNivelSIN(RC, SQL) Then
                txtCuentas(Index) = RC
                txtNCuentas(Index).Text = SQL
            Else
                MsgBox SQL, vbExclamation
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

    GeneraCtaExplotacionCC = False
    
    
    'Borramos datos
    SQL = "Delete from tmpctaexpcc where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
    

    If chkCtaExpCC(1).Value = 1 Then
        'Hacemos primero el periodo anterior
        RC = HacerCtaExploxCC(CInt(txtAno(7).Text) - 1, CInt(txtAno(8).Text) - 1)
        If RC = 0 Then Exit Function  'ha habido algun error
        
        'Si ha Generado datos los paso un momento para que el seguiente proceso no los borre
        If RC = 2 Then
            If Not ProcesoCtaExplotacionCC(0) Then Exit Function
        End If
        
    End If
    
    'Este bloque lo hace siempre
    RC = HacerCtaExploxCC(CInt(txtAno(7).Text), CInt(txtAno(8).Text))
    If RC > 0 Then
        RC = 0
        'Si era compartativo
        'Mezclamos los datos de ahora con los de antes
        If chkCtaExpCC(1).Value = 1 Then
            If Not ProcesoCtaExplotacionCC(1) Then RC = 1
        End If
        If RC = 0 Then GeneraCtaExplotacionCC = True
    End If
    

    
    'Eliminamos datos temporales
    If chkCtaExpCC(1).Value = 1 Then
        ProcesoCtaExplotacionCC 2
    End If
    

End Function

'0: Error    1: No hay datos       2: OK
Private Function HacerCtaExploxCC(Anyo1 As Integer, Anyo2 As Integer) As Byte
Dim A1 As Integer, M1 As Integer
Dim Post As Boolean

    On Error GoTo EGeneraCtaExplotacionCC
    HacerCtaExploxCC = 0
    
    UltimoMesAnyoAnal1 M1, A1
    
    'Si años consulta iguales
    If txtAno(0).Text = txtAno(1).Text Then
         Cad = " anoccost=" & Anyo1 & " AND mesccost>=" & Me.cmbFecha(0).ListIndex + 1
         Cad = Cad & " AND mesccost<=" & Me.cmbFecha(1).ListIndex + 1
         
    Else
        'Años disitintos
        'Inicio
        Cad = "( anoccost=" & Anyo1 & " AND mesccost>=" & Me.cmbFecha(0).ListIndex + 1 & ")"
        Cad = Cad & " OR ( anoccost=" & Anyo2 & " AND mesccost<=" & Me.cmbFecha(1).ListIndex + 1 & ")"
        'Por si la diferencia es mas de un año
        If Val(txtAno(1).Text) - Val(txtAno(0).Text) > 1 Then
            Cad = Cad & " OR (anoccost >" & Anyo1 & " AND anoccost < " & Anyo2 & ")"
        End If
    End If
    Cad = " (" & Cad & ")"
    
    RC = ""
    If txtCCoste(6).Text <> "" Then RC = " codccost >='" & txtCCoste(6).Text & "'"
    If txtCCoste(7).Text <> "" Then
        If RC <> "" Then RC = RC & " AND "
        RC = RC & " codccost <='" & txtCCoste(7).Text & "'"
    End If
    
    
    'Si han puesto desde hasta cuenta
    If txtCuentas(0).Text <> "" Then
        If RC <> "" Then RC = RC & " AND "
        RC = RC & " codmacta >='" & txtCuentas(0).Text & "'"
    End If
    
    If txtCuentas(1).Text <> "" Then
        If RC <> "" Then RC = RC & " AND "
        RC = RC & " codmacta <='" & txtCuentas(1).Text & "'"
    End If
    
    
    'Cogemos prestada la tabla tmpCierre cargando las cuentas k
    'tengan en hpsaldanal y hpsaldana1 si asi lo recuieren las fechas
    SQL = "Delete  from tmpctaexpCC"
    Conn.Execute SQL
    
    
    
    'Insertamos las cuentas desde hpsald1 si hicera o hiciese falta
    Tablas = ""
    If Anyo1 < A1 Then
        Tablas = "SI"
    Else
        If Anyo1 = A1 Then
            'Dependera del mes
            If M1 > (Me.cmbFecha(0).ListIndex + 1) Then Tablas = "OK"
        End If
    End If

    If RC <> "" Then Cad = RC & " AND " & Cad
    SQL = "INSERT INTO tmpctaexpCC (codusu,cta,codccost) SELECT "
    SQL = SQL & vUsu.Codigo & ",codmacta,codccost from hsaldosanal"
    'Si es de hco
    If Tablas <> "" Then SQL = SQL & "1"
    SQL = SQL & " Where "
    SQL = SQL & Cad
    SQL = SQL & " GROUP BY codccost,codmacta"
    Conn.Execute SQL
    
    
    
    'Diciembre 2012
    'Si ha marcado "solo" centros de reparto elimino aquellos eque n
    'Borro todos aquellos cc que no sean de reparto
    If chkCtaExpCC(2).Value = 1 Then
        SQL = "DELETE FROm tmpctaexpCC WHERE codusu = " & vUsu.Codigo & " AND "
        SQL = SQL & " NOT codccost IN (select distinct(subccost) from linccost) "
        Conn.Execute SQL
    End If
    
    
    'AHora en  tenemos todas las cuentas a tratar
    'Para ello cogeremos
    SQL = "Select count(*) from tmpctaexpCC where codusu = " & vUsu.Codigo
    
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cont = 0
    If Not Rs.EOF Then
        Cont = DBLet(Rs.Fields(0), "N")
    End If
    Rs.Close

    If Cont = 0 Then
        HacerCtaExploxCC = 1
        'MsgBox "Ningun registro a mostrar", vbExclamation
    Else
        'mes ini, ano ini, mes pedido, ano pedido, mes fin, ano fin
        'Cad = cmbFecha(5).ListIndex + 1 & "|" & txtAno(7).Text & "|"
        'Cad = Cad & cmbFecha(7).ListIndex + 1 & "|"
        Cad = cmbFecha(5).ListIndex + 1 & "|" & Anyo1 & "|"
        Cad = Cad & cmbFecha(7).ListIndex + 1 & "|"
        
        
        
        'El año del mes de calculo tiene k estar entre los años pedidos
        If cmbFecha(7).ListIndex >= cmbFecha(0).ListIndex Then
            Cad = Cad & Anyo1
        Else
            Cad = Cad & Anyo2
        End If
        Cad = Cad & "|"
        Cad = Cad & cmbFecha(1).ListIndex + 1 & "|" & Anyo2 & "|"
        
        'Ajusta los valores en modulo
        AjustaValoresCtaExpCC Cad
        
        'Si ha pediod los movimientos posteriores
        Post = (chkCtaExpCC(0).Value = 1)
        
        SQL = "Select cta,tmpctaexpCC.codccost,nommacta,nomccost from tmpctaexpCC,cuentas,cabccost where cuentas.codmacta=tmpctaexpCC.cta and cabccost.codccost=tmpctaexpCC.codccost and codusu = " & vUsu.Codigo
        'Vemos hasta donde hay de fechas en hco
        FechaFinEjercicio = CDate("01/" & M1 & "/" & A1)
        Rs.Open SQL, Conn, adOpenStatic, adLockPessimistic, adCmdText
        While Not Rs.EOF
            
            Tablas = Rs.Fields(0) & "|" & Rs.Fields(1) & "|" & DevNombreSQL(Rs.Fields(2))
            Tablas = Tablas & "|" & DevNombreSQL(Rs.Fields(3)) & "|"
        
            'Tb ponemos la pb
            Label15.Caption = Rs.Fields(0)
            Label15.Refresh
        
            CtaExploCentroCoste Tablas, Post, FechaFinEjercicio
    
            'Siguiente
            Rs.MoveNext
        Wend
        Rs.Close
        
        
        
        
        
        A1 = 0
        SQL = "Select count(*) from tmpctaexpcc where codusu =" & vUsu.Codigo
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            A1 = DBLet(Rs.Fields(0), "N")
        End If
        Rs.Close
        If A1 = 0 Then
            'MsgBox "Ningun registro a mostrar", vbExclamation
            'GeneraCtaExplotacionCC = False
            HacerCtaExploxCC = 1
        Else
            HacerCtaExploxCC = 2
            'GeneraCtaExplotacionCC = True
        End If
        
    End If
    
    Exit Function
EGeneraCtaExplotacionCC:
    MuestraError Err.Number, "Genera Cta Explotacion CC"
End Function


Private Sub UltimoMesAnyoAnal1(ByRef Mes As Integer, ByRef Anyo As Integer)
    Anyo = 1900
    Mes = 13
    Set miRsAux = New ADODB.Recordset
    SQL = "select max(anoccost) from hsaldosanal1"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        Anyo = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    
    If Anyo > 1900 Then
        SQL = "select max(mesccost) from hsaldosanal1 where anoccost =" & Anyo
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Mes = 1
        If Not miRsAux.EOF Then
            Mes = DBLet(miRsAux.Fields(0), "N")
        End If
        miRsAux.Close
    End If
    Set miRsAux = Nothing
End Sub

'Proceso:    1.- Updatear codusu para que no borre los datos
'            2.- Mezclar los datos de otros actual /siguiente
'            3.- Borramos
Private Function ProcesoCtaExplotacionCC(Proceso As Byte) As Boolean

    ProcesoCtaExplotacionCC = False
    Select Case Proceso
    Case 0
        Cont = 0
        i = 0
        SQL = "Select min(codusu) from Usuarios.zctaexpcc"
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            If Not IsNull(Rs.Fields(0)) Then
                i = Rs.Fields(0)
                Cont = 1 'para indicar que no es null
            End If
        End If
        Rs.Close
        If i >= 10 Then
            i = 9
        Else
            If i = 0 Then
                'Si NO ERA NULL es que estan ocupados(cosa rara) desde el 0 al 9
                If Cont = 1 Then
                    MsgBox "Error inesperado. Descripcion: codusu entre 0..9", vbExclamation
                    Exit Function
                End If
                i = 9
            Else
                i = i - 1
            End If
        End If
        NumRegElim = i
        SQL = "UPDATE  Usuarios.zctaexpcc set codusu = " & NumRegElim & " WHERE codusu = " & vUsu.Codigo
        Conn.Execute SQL
    
    Case 1
        'UPDATEO los valores de codusu=vusu
        SQL = "UPDATE Usuarios.zctaexpcc  SET acumd=0,acumh=0,postd=0,posth=0 where codusu = " & vUsu.Codigo
        Conn.Execute SQL
        
        'Para los valores comarativos
        SQL = "UPDATE Usuarios.zctaexpcc  SET acumd=perid,acumh=perih,postd=saldod,posth=saldoh where codusu = " & NumRegElim
        Conn.Execute SQL
        SQL = "UPDATE Usuarios.zctaexpcc  SET perid=0,perih=0,saldod=0,saldoh=0 where codusu = " & NumRegElim
        Conn.Execute SQL
        
        'En RS cargo todas las referencias de codusu= vusu
        SQL = "Select * from Usuarios.zctaexpcc WHERE codusu = "
        Rs.Open SQL & vUsu.Codigo, Conn, adOpenKeyset, adLockOptimistic, adCmdText
        
        
        'Cojere Todas las referencias de la tabla zctaexpcc para numregelim
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open SQL & NumRegElim, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not miRsAux.EOF
            'Busco si tiene referncia en ACTUAL
            SQL = "codmacta = '" & miRsAux!codmacta & "' AND codccost ='" & DevNombreSQL(miRsAux!codccost) & "'"
            
            
            If Not EncontrarEn_zctaexpcc(miRsAux!codmacta, UCase(miRsAux!codccost)) Then
                'Updateo solo el codusu
                Cad = "UPDATE Usuarios.zctaexpcc  SET codusu = " & vUsu.Codigo & " WHERE codusu = " & NumRegElim & " AND " & SQL
            Else
                'UPDATEO
                Cad = "UPDATE Usuarios.zctaexpcc  SET acumd=" & TransformaComasPuntos(CStr(miRsAux!acumd))
                Cad = Cad & ", acumH=" & TransformaComasPuntos(CStr(miRsAux!acumh))
                Cad = Cad & ", postd=" & TransformaComasPuntos(CStr(miRsAux!postd))
                Cad = Cad & ", posth=" & TransformaComasPuntos(CStr(miRsAux!posth))
                Cad = Cad & " WHERE codusu = " & vUsu.Codigo & " AND " & SQL
            End If
            Conn.Execute Cad
            'Siguiente
            miRsAux.MoveNext
            
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        Rs.Close
        
    Case 2
            'Finalmente borramos los datos de codusu=numregeleim
            SQL = "DELETE FROM Usuarios.zctaexpcc WHERE codusu = " & NumRegElim
            Conn.Execute SQL
        
            'Ahora, para que """SOLO""" aparezcan  los que tienen valor
            ' Los importes seran
            '                   MES                 SALDO
            '   actual      perid   perdih     saldod  saldoh
            '   anterior    acumd   acumh      postd   posth
        
            If optCCComparativo(1).Value Then
                'QUiero ver movimientos del periodo con lo cual me cargare aquellos
                ' que mov periodo en actual y anterior sea 0
                Cad = "acumd =0 and acumh=0 and perid=0 and perih=0"
                SQL = "DELETE FROM Usuarios.zctaexpcc WHERE codusu = " & vUsu.Codigo & " AND " & Cad
                Conn.Execute SQL
            End If
    End Select
    ProcesoCtaExplotacionCC = True
End Function

