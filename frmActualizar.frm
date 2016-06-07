VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmActualizar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizar diario"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmActualizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFacturas 
      BorderStyle     =   0  'None
      Height          =   4755
      Left            =   0
      TabIndex        =   31
      Top             =   -30
      Width           =   5055
      Begin VB.TextBox Text6 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   65
         Text            =   "Text6"
         Top             =   3000
         Width           =   3735
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   240
         TabIndex        =   39
         Text            =   "Text5"
         Top             =   3000
         Width           =   615
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   3480
         Visible         =   0   'False
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Frame tapa 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   270
         TabIndex        =   32
         Top             =   1920
         Width           =   4335
      End
      Begin VB.CommandButton cmdFacturas 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   3720
         TabIndex        =   41
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton cmdFacturas 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   40
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   33
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   3
         Left            =   3360
         TabIndex        =   34
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtNumFac 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   35
         Text            =   "Text5"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtNumFac 
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   36
         Text            =   "Text5"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Index           =   0
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Index           =   1
         Left            =   3360
         MaxLength       =   1
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   840
         Picture         =   "frmActualizar.frx":000C
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Diario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   64
         Top             =   2760
         Width           =   555
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         Height          =   315
         Left            =   120
         TabIndex        =   63
         Top             =   4260
         Width           =   1575
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmActualizar.frx":685E
         Top             =   720
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   3
         Left            =   3000
         Picture         =   "frmActualizar.frx":68E9
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   51
         Top             =   765
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   50
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   2520
         TabIndex        =   49
         Top             =   765
         Width           =   420
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   48
         Top             =   1485
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Nº Factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   47
         Top             =   1200
         Width           =   1000
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   2520
         TabIndex        =   46
         Top             =   1485
         Width           =   420
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   45
         Top             =   2205
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Serie"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   44
         Top             =   1920
         Width           =   435
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   7
         Left            =   2520
         TabIndex        =   43
         Top             =   2205
         Width           =   420
      End
      Begin VB.Label lblFac 
         Caption         =   "Facturas "
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
         Left            =   240
         TabIndex        =   42
         Top             =   120
         Width           =   4575
      End
   End
   Begin VB.Frame FrameRecalculo 
      Height          =   4635
      Left            =   120
      TabIndex        =   53
      Top             =   0
      Width           =   4875
      Begin MSComctlLib.ProgressBar pb3 
         Height          =   315
         Left            =   300
         TabIndex        =   60
         Top             =   2400
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.CommandButton cmdRecalCANCEL 
         Caption         =   "Salir"
         Height          =   435
         Left            =   3540
         TabIndex        =   55
         Top             =   4080
         Width           =   1155
      End
      Begin VB.CommandButton cmdRecalcula 
         Caption         =   "Aceptar"
         Height          =   435
         Left            =   2280
         TabIndex        =   54
         Top             =   4080
         Width           =   1095
      End
      Begin ComCtl2.Animation Animation2 
         Height          =   735
         Left            =   360
         TabIndex        =   59
         Top             =   2820
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   1296
         _Version        =   327681
         FullWidth       =   285
         FullHeight      =   49
      End
      Begin VB.Label Label10 
         Caption         =   "Es conveniente hacer una copia de seguridad."
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
         Index           =   3
         Left            =   300
         TabIndex        =   62
         Top             =   1440
         Width           =   4395
      End
      Begin VB.Label Label11 
         Height          =   255
         Left            =   300
         TabIndex        =   61
         Top             =   2040
         Width           =   3255
      End
      Begin VB.Label Label10 
         Caption         =   "Este proceso puede costar algunos minutos."
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
         Index           =   2
         Left            =   300
         TabIndex        =   58
         Top             =   1080
         Width           =   4395
      End
      Begin VB.Label Label10 
         Caption         =   "de la empresa: "
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
         Left            =   300
         TabIndex        =   57
         Top             =   660
         Width           =   4440
      End
      Begin VB.Label Label10 
         Caption         =   "No debe haber nadie trabajando en la contabilidad"
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
         Index           =   0
         Left            =   300
         TabIndex        =   56
         Top             =   360
         Width           =   4395
      End
   End
   Begin VB.Frame FrameListaContabilizar 
      Height          =   5535
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton cmdActuList 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   70
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton cmdActuList 
         Caption         =   "Actualizar"
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   69
         Top             =   5040
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4575
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   8070
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Diario"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Asiento"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Observaciones"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmActualizar.frx":6974
         ToolTipText     =   "Selecciona todos"
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmActualizar.frx":6ABE
         ToolTipText     =   "Quita seleccion"
         Top             =   5160
         Width           =   240
      End
   End
   Begin VB.Frame frameResultados 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   0
      TabIndex        =   25
      Top             =   -120
      Visible         =   0   'False
      Width           =   5115
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   3840
         TabIndex        =   26
         Top             =   4080
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3615
         Left            =   60
         TabIndex        =   30
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   6376
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nº Asien"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha Entrada"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Obteniendo resultados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   28
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label Label8 
         Caption         =   "Errores:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   27
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   14
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   13
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   1
      Left            =   3360
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CheckBox chkMostrarListview 
      Caption         =   "Mostrar lista asientos"
      Height          =   255
      Left            =   240
      TabIndex        =   66
      Top             =   4000
      Width           =   2535
   End
   Begin VB.Frame frame1Asiento 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4815
      Begin ComCtl2.Animation Animation1 
         Height          =   735
         Left            =   600
         TabIndex        =   12
         Top             =   1800
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1296
         _Version        =   327681
         FullWidth       =   241
         FullHeight      =   49
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label9 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   120
         Width           =   4335
      End
      Begin VB.Label lblAsiento 
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Asiento :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   3855
      End
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   1
      Left            =   3000
      Picture         =   "frmActualizar.frx":6C08
      Top             =   3240
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   840
      Picture         =   "frmActualizar.frx":6C93
      Top             =   2400
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   0
      Left            =   840
      Picture         =   "frmActualizar.frx":D4E5
      Top             =   3240
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   840
      Picture         =   "frmActualizar.frx":D570
      Top             =   2000
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Actualización de asientos"
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
      Left            =   720
      TabIndex        =   24
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Desde"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Asiento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   22
      Top             =   720
      Width           =   645
   End
   Begin VB.Label Label5 
      Caption         =   "Desde"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   1980
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Desde"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   3285
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Hasta"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   19
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Diario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   18
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   17
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Hasta"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   2430
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Hasta"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   15
      Top             =   3285
      Width           =   615
   End
End
Attribute VB_Name = "frmActualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OpcionActualizar As Byte
    '1.- Actualizar 1 asiento
    '2.- Desactualiza pero NO insertes en apuntes
    '3.- Desactualizar asiento desde hco
    
    'Si el asiento es de una factura entonces NUMSERIE tendra "FRACLI" o "FRAPRO"
    ' con lo cual habra que poner su factura asociada a NULL
    
    '4.- Si es para enviar datos a impresora
    '5.- Actualiza mas de 1 asiento
    
    '6.- Integra 1 factura
    '7.- Elimina factura integrada . DesINTEGRA   . C L I E N T E S
    '8.- Integra 1 factura PROVEEDORES
    '9.- Elimina factura integrada . Desintegra. P R O V E E D O R E S
    
    '10 .- Integracion masiva facturas clientes
    '11 .- Integracion masiva facturas Proveedores
    
    
    '12 .- Recalcular saldos desde hlinapu

    '13 .- IMPRIMIR asientos errores
    
    
    
    '15.-   RECALCULAR SALDOS desde otro proceso. No pregunta. Sigue adelante
    
        
Public NumAsiento As Long
Public FechaAsiento As Date
Public NumDiari As Integer
Public NUmSerie As String
Public NumFac As Long
Public FechaAnterior As Date
Public Proveedor As String
Public FACTURA As String
Public FechaFactura As Date

Public DentroBeginTrans As Boolean

'Nuevo. 17 Cotubre de 2005
'-------------------------
'  Los clientes que facturan con mas de un diario, las facturas SIEMPRE
'  van al diaro de parametros, con lo cual ES una cagada
Public DiarioFacturas As Integer
Private WithEvents frmD As frmTiposDiario
Attribute frmD.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private Cuenta As String
Private ImporteD As Currency
Private ImporteH As Currency
Private CCost As String
'Y estas son privadas
Private Mes As Integer
Private Anyo As Integer
Dim Fecha As String  'TENDRA la fecha ya formateada en yyy-mm-dd
Dim PrimeraVez As Boolean
Dim SQL As String
Dim RS As Recordset

Dim INC As Long

Dim NE As Integer
Dim ErroresAbiertos As Boolean
Dim NumErrores As Long

Dim ItmX As ListItem  'Para mostra errores masivos

Private Sub AñadeError(ByRef Mensaje As String)
On Error Resume Next
'Escribimos en el fichero
If Not ErroresAbiertos Then
    NE = FreeFile
    ErroresAbiertos = True
    Open App.Path & "\ErrActua.txt" For Output As NE
    If Err.Number <> 0 Then
        MsgBox " Error abriendo fichero errores", vbExclamation
        Err.Clear
    End If
End If
Print #NE, Mensaje
If Err.Number <> 0 Then
    Err.Clear
    NumErrores = -20000
Else
    NumErrores = NumErrores + 1
End If
End Sub



Private Function CadenaImporte(VaAlDebe As Boolean, ByRef Importe As Currency, ElImporteEsCero As Boolean) As String
Dim CadImporte As String

'Si va al debe, pero el importe es negativo entonces va al haber a no ser que la contabilidad admita importes negativos
    If Importe < 0 Then
        If Not vParam.abononeg Then
            VaAlDebe = Not VaAlDebe
            Importe = Abs(Importe)
        End If
    End If
    ElImporteEsCero = (Importe = 0)
    CadImporte = TransformaComasPuntos(CStr(Importe))
    If VaAlDebe Then
        CadenaImporte = CadImporte & ",NULL"
    Else
        CadenaImporte = "NULL," & CadImporte
    End If
End Function

Private Sub CargaProgres(Valor As Integer)
Me.ProgressBar1.Max = Valor
Me.ProgressBar1.Value = 0
End Sub


Private Function ComprobarFactura() As Boolean
Dim RT As Recordset
Dim B As Boolean

On Error GoTo EComprobarFactura
            'numfac     --> CODIGO FACTURA
            'NumDiari       --> AÑO FACTURA
            'NUmSerie       --> SERIE DE LA FACTURA
            'FechaAsiento   --> Fecha factura
    ComprobarFactura = False

    'Compruebo primero la fecha
    varFecOk = FechaCorrecta2(FechaAsiento)
    If varFecOk > 1 Then
        If varFecOk = 2 Then
            InsertaError varTxtFec
        Else
            InsertaError "Fuera ejercicios"
        End If
        ComprobarFactura = False
        Exit Function
    End If
    
    
    If OpcionActualizar = 10 Then
        SQL = "Select totfaccl as importetotafaccl, ba1faccl as importeba1faccl"
        SQL = SQL & " ,ba2faccl as importeba2faccl, ba3faccl as importeba3faccl"
        SQL = SQL & " ,ti1faccl as importeti1faccl, ti2faccl as importeti2faccl"
        SQL = SQL & " ,ti3faccl as importeti3faccl, tr1faccl as importetr1faccl"
        SQL = SQL & " ,tr2faccl as importetr2faccl, tr3faccl as importetr3faccl"
        SQL = SQL & " ,trefaccl as importetrefaccl,numasien FROM cabfact"
        Fecha = "  WHERE numserie ='" & Me.NUmSerie & "'"
        Fecha = Fecha & " AND codfaccl = " & NumFac
        Fecha = Fecha & " AND anofaccl =" & NumDiari
    Else
        SQL = "Select totfacpr as importetotafaccl, ba1facpr as importeba1faccl"
        SQL = SQL & " ,ba2facpr as importeba2faccl, ba3facpr as importeba3faccl"
        SQL = SQL & " ,ti1facpr as importeti1faccl, ti2facpr as importeti2faccl"
        SQL = SQL & " ,ti3facpr as importeti3faccl, tr1facpr as importetr1faccl"
        SQL = SQL & " ,tr2facpr as importetr2faccl, tr3facpr as importetr3faccl"
        SQL = SQL & " ,trefacpr as importetrefaccl,numasien FROM cabfactprov"
        Fecha = "  WHERE numregis = " & NumFac
        Fecha = Fecha & " AND anofacpr =" & NumDiari
    End If
    SQL = SQL & Fecha
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If RT.EOF Then
        'Insertamos error factura no encontrada
        B = False
    Else
        'Sumamos importes
        B = True
        
        
        If Not IsNull(RT!NumAsien) Then
            If RT!NumAsien = 0 Then
                InsertaError "NO se puede contabilizar. (nºasien=0)"
            Else
                InsertaError "Ya contabilizada: " & RT!NumAsien
            End If
            RT.Close
            Exit Function
        End If
        
        If IsNull(RT!Importetotafaccl) Then
            'Insertremos el error
            InsertaError "Sin importe"
            RT.Close
            Exit Function
        End If
        'Sumamos las bases
        ImporteD = 0
        If Not IsNull(RT!Importeba1faccl) Then ImporteD = ImporteD + RT!Importeba1faccl
        If Not IsNull(RT!Importeba2faccl) Then ImporteD = ImporteD + RT!Importeba2faccl
        If Not IsNull(RT!Importeba3faccl) Then ImporteD = ImporteD + RT!Importeba3faccl
        ImporteH = ImporteD  'En importe D guardamos las bases imponibles
        
        'Le sumamos los IVAS
        If Not IsNull(RT!Importeti1faccl) Then ImporteD = ImporteD + RT!Importeti1faccl
        If Not IsNull(RT!Importeti2faccl) Then ImporteD = ImporteD + RT!Importeti2faccl
        If Not IsNull(RT!Importeti3faccl) Then ImporteD = ImporteD + RT!Importeti3faccl
        
        'Los recargos
        If Not IsNull(RT!Importetr1faccl) Then ImporteD = ImporteD + RT!Importetr1faccl
        If Not IsNull(RT!Importetr2faccl) Then ImporteD = ImporteD + RT!Importetr2faccl
        If Not IsNull(RT!Importetr3faccl) Then ImporteD = ImporteD + RT!Importetr3faccl
        
        'La retencion( es en negativo)
        If Not IsNull(RT!Importetrefaccl) Then ImporteD = ImporteD - RT!Importetrefaccl
        
        If ImporteD <> RT!Importetotafaccl Then
            InsertaError "Suma de importes distinto total factura"
            RT.Close
            Exit Function
        End If
        RT.Close
        
        
        'Calculamos las lineas
        If OpcionActualizar = 10 Then
            SQL = "SELECT sum(impbascl) FROM linfact"
        Else
            SQL = "SELECT sum(impbaspr) FROM linfactprov"
        End If
        SQL = SQL & Fecha
        RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        ImporteD = 0
        If Not RT.EOF Then
            If Not IsNull(RT.Fields(0)) Then ImporteD = RT.Fields(0)
        End If
        RT.Close
        If ImporteD <> ImporteH Then
            'La suma de lineas y la de bases no es igual
            'Insertaremos error
            InsertaError "Suma lineas distinto de suma de bases"
            Exit Function
        End If
    End If
    
    
    ComprobarFactura = True
EComprobarFactura:
    If Err.Number <> 0 Then
        InsertaError "Error en SQL: " & SQL & Fecha
    End If
    Set RT = Nothing
End Function

Private Sub IncrementaP2(v As Integer)
On Error Resume Next
ProgressBar2.Value = ProgressBar2 + v
If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub IncrementaProgres(Veces As Integer)
On Error Resume Next
Me.ProgressBar1.Value = Me.ProgressBar1.Value + (Veces * INC)
If Err.Number <> 0 Then
    Err.Clear
    ProgressBar1.Value = 0
End If

End Sub


Private Sub cmdAceptar_Click()

    Screen.MousePointer = vbHourglass
    'Obtendremos el sql
    SQL = ObtenerSQL
    
    If Me.OpcionActualizar <> 5 Then
        'IMPRESION de asientos ... normales / errores
        If frmActualizar.OpcionActualizar = 4 Then
            '
            '
            '  NORMALES
            '
            '
            If IDiariosPendientes(SQL) Then   'Prepara datos impresion
                'Mandaremos el formulario de impresion con los datos
                frmImprimir.Opcion = 4
                frmImprimir.NumeroParametros = 0
                frmImprimir.FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                frmImprimir.Show vbModal
            End If
            
        Else
            '
            '
            '  Asientos con errores
            '
            '
            If IAsientosErrores(SQL) Then   'Prepara datos impresion
                'Mandaremos el formulario de impresion con los datos
                frmImprimir.Opcion = 66
                frmImprimir.NumeroParametros = 0
                frmImprimir.FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                frmImprimir.Show vbModal
            End If
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Label9.Caption = "Actualizando asientos"
    Label9.Refresh
    If SQL = "" Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    
    If Me.chkMostrarListview.Value = 1 Then
        CargaAsientosPorActualizar
        Screen.MousePointer = vbDefault
        Me.Refresh
        DoEvents
        If ListView2.ListItems.Count = 0 Then
            MsgBox "No hay asientos pendientes de actualizar", vbExclamation
            Exit Sub
        End If
        'Me.Height = 4865
        Me.Height = 6075
        FrameListaContabilizar.Visible = True
        Exit Sub
    End If
    
    
    If MsgBox("Va a actualizar los asientos entre las fechas. ¿Desea continuar?", vbQuestion + vbYesNo) <> vbYes Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'Borramos el tmp
    BorrarArchivoTemporal
    
    'Si tenemos que imprimir luego los asientos actualizados entonces
    If vParam.emitedia Then
        Conn.Execute "Delete FROM Usuarios.zhistoapu  where codusu =" & vUsu.Codigo
    End If
        
    'Vamos a ver cuales de estos registros en cabapu cunmplen el sql ademas de no estar bloqueados
    If ObtenerRegistrosParaActualizar Then    'Y bloquearlos
       
    
        'ACtualizarRegistros
        ActualizaASientosDesdeTMP
    End If
    
    
    
    'Ahora si todo ha ido bien mostraremos datos de las actualizaciones
    'Set Rs = Nothing
    'Set Rs = New ADODB.Recordset
    'SQL = "Select count(*) from tmpactualizar where codusu=" & vUsu.Codigo
    'Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Me.Height = 4965
    frame1Asiento.Visible = False
    
    Me.frameResultados.Visible = True
    Me.Refresh
    Screen.MousePointer = vbHourglass
    If NumErrores > 0 Then
        Close #NE
        Label7.Caption = "Se han producido errores."
        CargaListAsiento
    Else
       Label7.Caption = "NO se han producido errores."
       Me.Refresh
    End If
    
    'Ahora comprobamos, si emite diario al actualizar, que tiene datos
    If vParam.emitedia Then
        INC = 0
        SQL = "Select count(*) from  Usuarios.zhistoapu  where codusu =" & vUsu.Codigo
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            INC = DBLet(RS.Fields(0), "N")
        End If
        RS.Close
        Set RS = Nothing
        If INC > 0 Then
            'Si k ha actualizado apuntes, con lo cual hay k imprimir
             With frmImprimir
                SQL = "Actualización del dia " & Format(Now, "dd/mm/yyyy") & " a las " & Format(Now, "hh:mm") & "."
                .OtrosParametros = "Fechas= """ & SQL & """|Cuenta= """"|"
                .NumeroParametros = 2
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = True
                'Opcion dependera del combo
                .Opcion = 12
                .Show vbModal
            End With
        End If
    End If

    If NumErrores = 0 Then
        Me.Refresh
        DoEvents
        espera 0.5
        Unload Me
    End If
    'Fin
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaListAsiento()

NE = FreeFile
If Dir(App.Path & "\ErrActua.txt") = "" Then
    MsgBox "Los errores han sido eliminados. Imposible ver errores. Modulo: CargaLisAsiento", vbExclamation
    Exit Sub
End If

Me.frameResultados.Visible = True
'Los encabezados
ListView1.ColumnHeaders.Clear
ListView1.ColumnHeaders.Add , , "Diario", 750
ListView1.ColumnHeaders.Add , , "Fecha", 1000
ListView1.ColumnHeaders.Add , , "Nº Asie.", 1000
ListView1.ColumnHeaders.Add , , "Error", 3000


Open App.Path & "\ErrActua.txt" For Input As #NE
While Not EOF(NE)
    Line Input #NE, Cuenta
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = RecuperaValor(Cuenta, 1)
    ItmX.SubItems(1) = RecuperaValor(Cuenta, 2)
    ItmX.SubItems(2) = RecuperaValor(Cuenta, 3)
    ItmX.SubItems(3) = RecuperaValor(Cuenta, 4)
Wend
Close #NE
End Sub




Private Sub CargaListFacturas()

NE = FreeFile
If Dir(App.Path & "\ErrActua.txt") = "" Then
    MsgBox "Los errores han sido eliminados. Imposible ver errores. Modulo: CargaLisAsiento"
    Exit Sub
End If

Me.frameResultados.Visible = True
'Los encabezados
ListView1.ColumnHeaders.Clear
ListView1.ColumnHeaders.Add , , "Factura", 1200
ListView1.ColumnHeaders.Add , , "Fecha ", 1200
ListView1.ColumnHeaders.Add , , "Error ", 3000

Open App.Path & "\ErrActua.txt" For Input As #NE
While Not EOF(NE)
    Line Input #NE, Cuenta
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = RecuperaValor(Cuenta, 1)
    ItmX.SubItems(1) = RecuperaValor(Cuenta, 2)
    ItmX.SubItems(2) = RecuperaValor(Cuenta, 3)
Wend
Close #NE
End Sub



'Eliminar factura con asiento
Private Function EliminaFacturaConAsiento()
Dim Donde As String
Dim bol As Boolean
Dim LEtra As String
Dim Mc As Contadores
Dim Contabilizada As String

    On Error GoTo EEliminaFacturaConAsiento
    'Sabemos que
    'numasiento     --> Nº aseinto
    'numfac         --> CODIGO FACTURA
    'NumDiari       --> ATENCION -> Nº de diario, no como al integrar
    'FechaAsiento   --> Fecha asiento
    'NUmSerie       --> SERIE DE LA FACTURA  y el año (sep. con pipes)

    'Obtenemos el mes y el año
    Mes = Month(FechaAsiento)
    Anyo = Year(FechaAsiento)
    Fecha = Format(FechaAsiento, FormatoFecha)
    
    'Aqui bloquearemos
    Conn.BeginTrans
    
    'Eliminamos factura
    LEtra = RecuperaValor(NUmSerie, 1)
    If Me.OpcionActualizar = 7 Then
        '-------------------------------------------------------------
        '               C L I E N T E S
        '-------------------------------------------------------------
        SQL = " WHERE numserie = '" & LEtra & "'"
        SQL = SQL & " AND numfactu = " & NumFac
        SQL = SQL & " AND anofactu= " & RecuperaValor(NUmSerie, 2)
        'Las lineas
        Donde = "Linea factura"
        Cuenta = "DELETE from factcli_lineas " & SQL
        Conn.Execute Cuenta
        
        'totales de factura
        Donde = "Totales factura"
        Cuenta = "DELETE from factcli_totales " & SQL
        Conn.Execute Cuenta
        
        
        Contabilizada = "select count(*) from cobros where numserie = " & DBSet(LEtra, "T") & " and numfactu = " & NumFac & " and fecfactu = " & DBSet(FechaAsiento, "F") & " and impcobro <> 0 and not impcobro is null "
        
        If TotalRegistros(Contabilizada) <> 0 Then
            MsgBox "Hay cobros que ya se han efectuado. Revise cartera y contabilidad.", vbExclamation
        Else
            ' cobro de la factura
            Donde = "Cobro factura"
            
            Cuenta = "DELETE from cobros_realizados where numserie = " & DBSet(LEtra, "T") & " and numfactu = " & NumFac & " and fecfactu = " & DBSet(FechaAsiento, "F")
            Conn.Execute Cuenta
            
            
            Cuenta = "DELETE from cobros where numserie = " & DBSet(LEtra, "T") & " and numfactu = " & NumFac & " and fecfactu = " & DBSet(FechaAsiento, "F")
            Conn.Execute Cuenta
        End If
        
        'La factura
        Donde = "Cabecera factura"
        Cuenta = "DELETE from factcli " & SQL
        Conn.Execute Cuenta

    Else
        '-------------------------------------------------------------
        '       P R O V E E D O R E S
        '-------------------------------------------------------------
        SQL = " WHERE numserie = '" & LEtra & "'"
        SQL = SQL & " AND numregis = " & NumFac
        SQL = SQL & " AND anofactu= " & RecuperaValor(NUmSerie, 2)
        'Las lineas
        Donde = "Linea factura"
        Cuenta = "DELETE from factpro_lineas " & SQL
        Conn.Execute Cuenta
        
        'totales de factura
        Donde = "Totales factura"
        Cuenta = "DELETE from factpro_totales " & SQL
        Conn.Execute Cuenta
        
        Contabilizada = "select count(*) from pagos where numserie = " & DBSet(LEtra, "T") & " and codmacta = " & DBSet(Proveedor, "T") & " and numfactu = " & DBSet(FACTURA, "T") & " and fecfactu = " & DBSet(FechaFactura, "F") & " and imppagad <> 0 and not imppagad is null "
        
        If TotalRegistros(Contabilizada) <> 0 Then
            MsgBox "Hay pagos que ya se han efectuado. Revise cartera y contabilidad.", vbExclamation
        Else
            ' cobro de la factura
            Donde = "Pago factura"
            
            Cuenta = "DELETE from pagos_realizados where numserie = " & DBSet(LEtra, "T") & " and codmacta = " & DBSet(Proveedor, "T") & " and numfactu = " & DBSet(FACTURA, "T") & " and fecfactu = " & DBSet(FechaFactura, "F")
            Conn.Execute Cuenta
            
            
            Cuenta = "DELETE from pagos where numserie = " & DBSet(LEtra, "T") & " and codmacta = " & DBSet(Proveedor, "T") & " and numfactu = " & DBSet(FACTURA, "T") & " and fecfactu = " & DBSet(FechaFactura, "F")
            Conn.Execute Cuenta
        End If
        
        'La factura
        Donde = "Cabecera factura"
        Cuenta = "DELETE from factpro " & SQL
        Conn.Execute Cuenta
        LEtra = RecuperaValor(NUmSerie, 1) '"1"
    End If

    bol = DesActualizaElASiento(Donde)

EEliminaFacturaConAsiento:
        If Err.Number <> 0 Then
            SQL = "Actualiza Asiento." & vbCrLf & "----------------------------" & vbCrLf
            SQL = SQL & Donde
            MuestraError Err.Number, SQL, Err.Description
            bol = False
        End If
        If bol Then
            Conn.CommitTrans
            
            'Intentamos devolver el contador
            If FechaAsiento >= vParam.fechaini Then
                Set Mc = New Contadores
                Mc.DevolverContador LEtra, (FechaAsiento <= vParam.fechafin), NumFac
                Set Mc = Nothing
            End If
            
            
            'INSERTO EN LOG
            Mes = 6
            If Me.OpcionActualizar <> 7 Then
                Mes = 9   'FRARPO
                LEtra = ""
            End If
            
            vLog.Insertar CByte(Mes), vUsu, LEtra & Format(NumFac, "000000")
            
            
            EliminaFacturaConAsiento = True
            AlgunAsientoActualizado = True
        Else
            Conn.RollbackTrans
        End If
    
End Function

Private Sub cmdActuList_Click(Index As Integer)
Dim N As Integer
Dim cad As String

    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    SQL = ""
    For N = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(N).Checked Then SQL = SQL & "K"
    Next N
    If SQL = "" Then
        MsgBox "Seleccione algun asiento para actualizar", vbExclamation
        Exit Sub
    Else
        cad = Len(SQL)
        SQL = "Va a actualizar " & cad & " asiento(s). ¿Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    'Borramos el tmp
    BorrarArchivoTemporal
    
    'Si tenemos que imprimir luego los asientos actualizados entonces
    If vParam.emitedia Then
        Conn.Execute "Delete FROM Usuarios.zhistoapu  where codusu =" & vUsu.Codigo
    End If
        
    
    'Para el LOG
    vLog.InicializarDatosDesc
    
    
    'Nos ponemos manos a la obra
    NumErrores = 0
    For N = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(N).Checked Then
            NumDiari = Val(ListView2.ListItems(N).Text)
            FechaAsiento = CDate(ListView2.ListItems(N).SubItems(1))
            Fecha = Format(FechaAsiento, FormatoFecha)
            NumAsiento = ListView2.ListItems(N).SubItems(2)
            
            
            cad = RegistroCuadrado
            If cad = "" Then
                SQL = ""
                If BloquearAsiento(CStr(NumAsiento), CStr(NumDiari), Fecha, SQL) Then
            
                
                    'Actualiza el asiento
                    If ActualizaAsiento = False Then
                         DesbloquearAsiento CStr(NumAsiento), CStr(NumDiari), Format(FechaAsiento, FormatoFecha)
                    Else
                        vLog.AnyadeTextoDatosDes CStr(NumAsiento)
                        
                        'Si tiene k imprimir al finalizar entonces
                        If vParam.emitedia Then
                            SQL = NumAsiento & "|" & Format(FechaAsiento, FormatoFecha) & "|" & NumDiari & "|"
                            IHcoApuntesAlActualizarModificar (SQL)
                        End If
                    End If
                    
                
                Else
                    'HA habido error bloqueando el archivo
                    '-------------------------------------
                    If SQL = "" Then SQL = "Asiento bloqueado"
                    InsertaError SQL
                    
                End If
                
            
            Else
                'Registro NO cuadrado o sin lineas ...
                 InsertaError cad
            End If
        End If  'Del checked
    Next N
    
    If vLog.DatosDescripcion <> "" Then vLog.Insertar 10, vUsu, vLog.DatosDescripcion
    
    Me.Height = 5000
    frame1Asiento.Visible = False
    Me.FrameListaContabilizar.Visible = False
    Me.frameResultados.Visible = True
    Me.Refresh
    Screen.MousePointer = vbHourglass
    If NumErrores > 0 Then
        Close #NE
        Label7.Caption = "Se han producido errores."
        CargaListAsiento
    Else
       Label7.Caption = "NO se han producido errores."
       Me.Refresh
    End If
    
    'Ahora comprobamos, si emite diario al actualizar, que tiene datos
    If vParam.emitedia Then
        INC = 0
        SQL = "Select count(*) from  Usuarios.zhistoapu  where codusu =" & vUsu.Codigo
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            INC = DBLet(RS.Fields(0), "N")
        End If
        RS.Close
        Set RS = Nothing
        If INC > 0 Then
            'Si k ha actualizado apuntes, con lo cual hay k imprimir
             With frmImprimir
                SQL = "Actualización del dia " & Format(Now, "dd/mm/yyyy") & " a las " & Format(Now, "hh:mm") & "."
                .OtrosParametros = "Fechas= """ & SQL & """|Cuenta= """"|"
                .NumeroParametros = 2
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = True
                'Opcion dependera del combo
                .Opcion = 12
                .Show vbModal
            End With
        End If
    End If
    
    If NumErrores = 0 Then
        Me.Refresh
        DoEvents
        espera 0.5
        Unload Me
    End If
    'Fin
    Screen.MousePointer = vbDefault

    
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub


Private Sub cmdFacturas_Click(Index As Integer)
Dim F As Date
    If Index = 1 Then
        'Cancelar
        Unload Me
    Else
        NumErrores = 0
        'Aceptar
        If Text3(2).Text = "" Or Text3(3).Text = "" Then
            MsgBox "Ponga el desde / hasta fecha", vbExclamation
            Exit Sub
        End If
        
        F = CDate(Text3(2).Text)
        If F < vParam.fechaini Or F > DateAdd("yyyy", 1, vParam.fechafin) Then
            MsgBox "Fecha Desde fuera de ejercicios", vbExclamation
            Exit Sub
        End If
        
        F = CDate(Text3(3).Text)
        If F < vParam.fechaini Or F > DateAdd("yyyy", 1, vParam.fechafin) Then
            MsgBox "Fecha Hasta fuera de ejercicios", vbExclamation
            Exit Sub
        End If
        
        
        
        If Text3(2).Text <> "" And Text3(3).Text <> "" Then
            If CDate(Text3(2).Text) > CDate(Text3(3).Text) Then
                MsgBox "Fecha Desde mayor que hasta", vbExclamation
                Exit Sub
            End If
        End If
         
        If txtnumfac(0).Text <> "" And txtnumfac(1).Text <> "" Then
            If Val(txtnumfac(0).Text) > Val(txtnumfac(1).Text) Then
                MsgBox "Nº factura Desde mayor que hasta", vbExclamation
                Exit Sub
            End If
        End If
        
        
        
        If Text5.Text = "" Then
            MsgBox "Numero de diario obligatorio", vbExclamation
            Exit Sub
        End If
        SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text5.Text, "N")
        If SQL = "" Then
            MsgBox "Diario incorrecto o incexistente", vbExclamation
            Exit Sub
        End If
        
        If Me.OpcionActualizar = 10 Then
            'Cliente
            SQL = " FACTCLI"
        Else
            SQL = " FACTPRO"
        End If
        
        If Not BloqueoManual(True, "CONTABILIZA", SQL) Then
             MsgBox "Proceso bloqueado realizandose por otro usuario", vbExclamation
        Else
            PrepararIntegrarFacturas
            
            
            'Borro los bloqueos de regiostro
            SQL = "cabfact"
            If Me.OpcionActualizar <> 10 Then SQL = SQL & "prov"
            BloqueoManual False, SQL, ""
            
            
            
            BloqueoManual False, "CONTABILIZA", ""
        End If
    End If
End Sub

    
Private Sub PrepararIntegrarFacturas()
Dim F As Date
Dim RF As Recordset

'JUNIO 2010
'BLoquearemos todas las fras que se vayan a integrar
'Si alguna esta bloqueada por otro usuario avisamos
Dim CadenaBloqueoRegistros As String
Dim MiC As String

    SQL = ""
    'Generamos el SQL
    If Me.OpcionActualizar = 10 Then
        'Cliente
        Cuenta = " fecfaccl"
    Else
        Cuenta = " fecrecpr"
    End If
    If Text3(2).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & Cuenta & " >= '" & Format(Text3(2).Text, FormatoFecha) & "'"
    End If
    If Text3(3).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & Cuenta & " <= '" & Format(Text3(3).Text, FormatoFecha) & "'"
    End If
    'Codigo de factura
    If Me.OpcionActualizar = 10 Then
        'Cliente
        Cuenta = "codfaccl"
    Else
        Cuenta = "numregis"
    End If
    If txtnumfac(0).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & Cuenta & " >= " & txtnumfac(0).Text
    End If
    If txtnumfac(1).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & Cuenta & " <= " & txtnumfac(1).Text
    End If
    
    'Solo para CLIENTES
    If Me.OpcionActualizar = 10 Then
        If txtSerie(0).Text <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & " numserie >= '" & txtSerie(0).Text & "'"
        End If
        If txtSerie(1).Text <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & " numserie <= '" & txtSerie(1).Text & "'"
        End If
    End If
    
    'Ahora con el SQL vemos las facturas a integrar
    BorrarArchivoTemporal
    
    'Abrimos el recordset
    If Me.OpcionActualizar = 10 Then
        Cuenta = "Select codfaccl,anofaccl,fecfaccl,numserie from cabfact "
    Else
        Cuenta = "Select numregis,anofacpr,fecrecpr from cabfactprov "
        
        'Antes de 20 Enero 2004
        'Cuenta = "Select numregis,anofacpr,fecfacpr from cabfactprov "
    End If
    Cuenta = Cuenta & " WHERE (numasien IS NULL) "
    If SQL <> "" Then Cuenta = Cuenta & " AND " & SQL
    Cuenta = Cuenta & " ORDER BY "
    If Me.OpcionActualizar = 10 Then
        Cuenta = Cuenta & "fecfaccl,numserie,codfaccl"
    Else
        Cuenta = Cuenta & "fecrecpr,numregis"
    End If
    
    
    Set RF = New ADODB.Recordset
    
    'Contrloamos PUNTUALMENTE el error del openrecordset
    'Por si acaso hay bloqueada alguna factura
    On Error Resume Next
    RF.Open Cuenta, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Err.Number <> 0 Then
        MsgBox "Compruebe que no se estan modificando facturas." & vbCrLf & _
            "Si el error persiste consulte con soporte técnico", vbExclamation
        Set RF = Nothing
        Exit Sub
    End If
    
    '-------------------------------------------------
    'Desactivamos en este modulo el control de errores
    On Error GoTo 0
    
    
    
    If Not RF.EOF Then
    
        vLog.InicializarDatosDesc   'Para el LOG
    
    
        RF.MoveFirst
        Mes = 1
        'Contamos las facturas k hay para integrar

        'JUNIO 2010
        'ademas de contar las bloqueo en zbloqueos
        SQL = ""
        While Not RF.EOF
        
            '*********************************************************************************
            'BLoqueamos a ver si nos deja
            MiC = "Insert into zBloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'cabfact"
            If Me.OpcionActualizar = 10 Then
                'FRACLI
                '24002,'cabfact',"2009|'V'|32|")
                MiC = MiC & "',""" & RF!anofaccl & "|'" & RF!NUmSerie & "'|" & RF!codfaccl & "|"")"
            Else
                'FRAPRO
                '2009|11478|)
                MiC = MiC & "prov','" & RF!anofacpr & "|" & RF!NumRegis & "|')"
            End If
            If Not EjecutaSQL(MiC) Then
                If Me.OpcionActualizar = 10 Then
                    'CLIENTES
                    CadenaBloqueoRegistros = CadenaBloqueoRegistros & RF!NUmSerie & Format(RF!codfaccl, "000000") & "  " & RF!anofaccl & vbCrLf
                    
                Else
                    'FRAPRO
                    CadenaBloqueoRegistros = CadenaBloqueoRegistros & Format(RF!NumRegis, "000000") & "  " & RF!anofacpr & vbCrLf
                End If
                
            Else
                Mes = Mes + 1
            End If
            RF.MoveNext
        Wend
        RF.MoveFirst
        
        If CadenaBloqueoRegistros <> "" Then
            MiC = "Facturas  bloqueadas: " & vbCrLf & CadenaBloqueoRegistros & vbCrLf & "¿Continuar?"
            If MsgBox(MiC, vbQuestion + vbYesNoCancel) <> vbYes Then
                RF.Close
                Set RF = Nothing
                Exit Sub
            End If
        End If
        
        'Datos para el progress
        ProgressBar2.Visible = (Mes > 4)
        ProgressBar2.Value = 0
        If Mes < 15000 Then
            INC = 1
            ProgressBar2.Max = Mes * 2
        Else
            ProgressBar2.Max = Mes
            INC = 0
        End If
        
        
        '---------------------------------------
        'numfac     --> CODIGO FACTURA
        'NumDiari       --> AÑO FACTURA
        'NUmSerie       --> SERIE DE LA FACTURA
        'FechaAsiento   --> Fecha factura
        '---------------------------------------
        
        'Nuevo en OCTUBRE 2005
        DiarioFacturas = Val(Text5.Text)
        If CadenaBloqueoRegistros <> "" Then CadenaBloqueoRegistros = Replace(CadenaBloqueoRegistros, vbCrLf, "|")
        While Not RF.EOF
            'Caption = Now
            NumFac = RF.Fields(0)
            NumDiari = RF.Fields(1)
            FechaAsiento = RF.Fields(2)
            Label12.Caption = NumFac & " - " & RF.Fields(2)
            Me.Refresh
            
            
            'Si bloqueamos el registro
                        
                If OpcionActualizar = 10 Then
                    NUmSerie = RF.Fields(3)
                    'Para comprobar k no esta entre las facturas bloquedas DOS ESPACIOS EN BLANCO
                    MiC = RF!NUmSerie & Format(RF!codfaccl, "000000") & "  " & RF!anofaccl
                Else
                    MiC = Format(RF!NumRegis, "000000") & "  " & RF!anofacpr
                    NUmSerie = ""
                End If
                NumAsiento = -1
                
                
                If InStr(1, CadenaBloqueoRegistros, MiC) = 0 Then
                
                    'Comprobaremos la factura
                    If ComprobarFactura Then
                        If OpcionActualizar = 10 Then NUmSerie = RF.Fields(3)
                        'Mandamos a integrar la factura
                        IntegraFactura
                    End If
                    IncrementaP2 1
                    IncrementaP2 CInt(INC)
                    
                Else
                    'Stop  'bloqueada
                End If
            RF.MoveNext
        Wend
        RF.Close
        Set RF = Nothing
        Label12.Caption = ""
        If vLog.DatosDescripcion <> "" Then
            If OpcionActualizar = 10 Then
                NumAsiento = 11
            Else
                NumAsiento = 12
            End If
            vLog.Insertar CByte(NumAsiento), vUsu, vLog.DatosDescripcion
        End If
        
        'Ahora si numero de errores >0 entonces mostramos los errores
        If NumErrores > 0 Then
            Close #NE
            CargaListFacturas
        Else
            Unload Me
        End If
    Else
        MsgBox "Ninguna factura a contabilizar", vbExclamation
    End If

End Sub






Private Sub cmdRecalCANCEL_Click()
    'Salir de RECLACULO
    Unload Me
End Sub

Private Sub cmdRecalcula_Click()
    SQL = "¿Seguro que desea continuar con el recálculo de saldos?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
       
       
       
    'Comprobacion si hay alguien trabajando
    
    If UsuariosConectados("Recalculando saldos", True) Then Exit Sub
    
    
    SQL = "Este proceso puede durar mucho tiempo" & vbCrLf & vbCrLf & vbCrLf & "¿Continuar?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    
   'BLOQUEAMOS LA BD
   If Not Bloquear_DesbloquearBD(True) Then
        MsgBox "No se ha podido bloquea a nivel de BD.", vbExclamation
        Exit Sub
    End If
    

    Screen.MousePointer = vbHourglass
    PonerAVI 2
    espera 1
    Label11.Visible = True
    pb3.Visible = False
    NE = 0
    'El recalculo va aqui
    pb3.Visible = False
    Label11.Visible = False
    Animation2.Stop
    Animation2.Visible = False

    
    
    Bloquear_DesbloquearBD False
    Screen.MousePointer = vbDefault
    If NE = 0 Then
        MsgBox "El proceso ha finalizado correctamente.", vbInformation
        cmdRecalcula.Enabled = False
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub






Private Sub Form_Activate()
Dim bol As Boolean
If PrimeraVez Then
    PrimeraVez = False
    Me.Refresh
    bol = False
    Select Case OpcionActualizar
    Case 1
        ActualizaAsiento
        bol = True
    Case 2, 3
        DesActualizaAsiento
        bol = True
    Case 4, 13
        Text1(0).Text = NUmSerie
        Text1(1).Text = NUmSerie
        NUmSerie = ""
    Case 6, 8
        'Integramos la factura (Dependera del opcion si es de clientes o de proveedores
        IntegraFactura
        bol = True
    Case 7, 9
         'Integramos la factura (Dependera del opcion si es de clientes o de proveedores
        EliminaFacturaConAsiento
        bol = True
    Case 10, 11
        txtnumfac(0).SetFocus
        
    Case 15
        'Hacemos el recalculo
        RecalculoAutomatico
        bol = True
        
    Case 16
        'Insertar Asiento en el hco
        
    End Select
    If bol Then Unload Me
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_DblClick()
    CargaListAsiento
End Sub

Private Sub Form_Load()
Dim B As Boolean

    Me.Icon = frmPpal.Icon


    ErroresAbiertos = False
    Limpiar Me
    Label12.Caption = ""
    PrimeraVez = True
    Me.frameResultados.Visible = False
    Me.FrameFacturas.Visible = False
    Me.FrameRecalculo.Visible = False
    FrameListaContabilizar.Visible = False
    
    ListView1.ListItems.Clear
    Select Case OpcionActualizar
    Case 1, 2, 3
        Label1.Caption = "Nº Asiento"
        Me.lblAsiento.Caption = NumAsiento
        INC = 10  'Incremento para el proggress
        If OpcionActualizar = 1 Then
            Label9.Caption = "Actualizar"
        Else
            Label9.Caption = "Modi/Eliminar"
        End If
        'Tamaño
        Me.Height = 3000
        B = True
    Case 4, 5, 13
        If OpcionActualizar <> 5 Then
            Me.chkMostrarListview.Visible = False
        Else
            Me.chkMostrarListview.Visible = True
        End If
        Me.Height = 4865
        Text3(0).Text = Format(vParam.fechaini, "dd/mm/yyyy")
        If Now > vParam.fechafin Then
            Text3(1).Text = Format(Now, "dd/mm/yyyy")
        Else
            Text3(1).Text = Format(vParam.fechafin, "dd/mm/yyyy")
        End If
        If OpcionActualizar <> 5 Then
            Label2.Caption = "Seleccion para imprimir"
            Me.Caption = "Impresión"
            If OpcionActualizar = 13 Then
                Label2.Caption = "Impresion ERRORES"
                Me.Caption = Me.Caption & " ERRORES"
            End If
        Else
            'La opcion 5: Actualizar
            Label2.Caption = "Asientos para actualizar"
            Me.Caption = "Actualizar asientos"
        End If
        B = False
        
        If OpcionActualizar = 5 Then ChkListaAsientos True
            
    Case 6, 7, 8, 9
        '// Estamos en Facturas
        Label1.Caption = "Nº factura"
        If OpcionActualizar < 8 Then
            Label1.Caption = Label1.Caption & " Cliente"
        Else
            Label1.Caption = Label1.Caption & " Proveedor"
        End If
        Me.lblAsiento.Caption = NUmSerie & NumAsiento
        INC = 10  'Incremento para el proggress
        If OpcionActualizar = 6 Or OpcionActualizar = 8 Then
            Label9.Caption = "Integrar Factura"
        Else
            Label9.Caption = "Eliminar Factura"
        End If
        Me.Caption = "Actualizar facturas"
        'Tamaño
        Me.Height = 3315
        B = True
    Case 10, 11
        Me.Caption = "Contabilizar facturas"
        Me.FrameFacturas.Visible = True
        lblFac = "Facturas"
        ProgressBar2.Visible = False
        If Me.OpcionActualizar = 10 Then
            lblFac.Caption = lblFac.Caption & " clientes"
            Me.Label4(4).Caption = "Nº Factura"
        Else
            lblFac.Caption = lblFac.Caption & " proveedores"
            Me.Label4(4).Caption = "Nº registro"
        End If
        tapa.Visible = Me.OpcionActualizar = 11
        'Para k no cojan foco
        txtSerie(0).Visible = Not tapa.Visible
        
        txtSerie(1).Visible = Not tapa.Visible
        cmdFacturas(1).Cancel = True
        
        'Ofertamos las fechas del ejercicio
        Text3(2).Text = Format(vParam.fechaini, "dd/mm/yyyy")
        Text3(3).Text = Format(Now, "dd/mm/yyyy")
        
        'Ofertamos el diario
        If OpcionActualizar = 10 Then
            Text5.Text = vParam.numdiacl
        Else
            Text5.Text = vParam.numdiapr
        End If
        
        SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text5.Text, "N")
        Text6.Text = SQL
        Me.Height = FrameFacturas.Height + 210
        
    Case 12, 15
        Me.Caption = "Recálculo de saldos"
        pb3.Visible = False
        FrameRecalculo.Visible = True
        cmdRecalCANCEL.Cancel = OpcionActualizar = 12
        cmdRecalcula.Visible = OpcionActualizar = 12
        cmdRecalCANCEL.Visible = OpcionActualizar = 12
        Label11.Caption = ""
        Label10(1).Caption = "de la empresa " & UCase(vEmpresa.nomempre)
        Animation2.Visible = False
        Me.Height = 5300
    End Select
    Me.frame1Asiento.Visible = B
    Me.Animation1.Visible = B
End Sub



Private Function IntegraFactura() As Boolean
Dim B As Boolean
Dim Donde As String
Dim vConta As Contadores

Dim TipoConce As String
On Error GoTo EIntegraFactura
    
    IntegraFactura = False
    
    If Not DentroBeginTrans Then Conn.BeginTrans
    Fecha = Format(FechaAsiento, FormatoFecha)
    
    
    'Vemos si estamos intentato forzar numero de asiento
    If NumAsiento > 0 Then
        'Primero que nada obtendremos el contador
        If AsientoExiste Then
            MsgBox "Ya existe el asiento con la numeración: " & NumAsiento & " " & FechaAsiento & " " & NumDiari, vbExclamation
            'Vamoa al final del proceso de esta factura
            GoTo EIntegraFactura
        End If
    Else
        Donde = "Conseguir contador"
        Set vConta = New Contadores
        If vConta.ConseguirContador("0", (FechaAsiento <= vParam.fechafin), True) = 1 Then
            MsgBox "Error consiguiendo contador asiento", vbExclamation
            'Vamoa al final del proceso de esta factura
            GoTo EIntegraFactura
        End If
        
        If Not vConta.YaExisteContador((FechaAsiento <= vParam.fechafin), vParam.fechafin, (OpcionActualizar < 10)) Then
            If OpcionActualizar > 9 Then InsertaError "Error contadores asiento: " & vConta.Contador
            GoTo EIntegraFactura
        End If
        NumAsiento = vConta.Contador
        Set vConta = Nothing
    End If
    
    'Actualizamos los datos
    If OpcionActualizar = 6 Or OpcionActualizar = 10 Then
        'B = IntegraLaFactura(Donde, ConcFac)
        B = IntegraLaFactura(Donde)
    Else
        'B = IntegraLaFacturaProv(Donde, ConcFac)
        B = IntegraLaFacturaProv(Donde)
    End If
    
EIntegraFactura:
    If Err.Number <> 0 Then
        If OpcionActualizar > 9 Then
            'Esta actualizando varias a la vez
            InsertaError Donde & " - " & Err.Description
        Else
            MuestraError Err.Number, "Integra factura(I)" & vbCrLf & Donde
        End If
        Err.Clear
        B = False
    End If
    If B Then
        If OpcionActualizar > 9 Then
            'Actualizando desde/hasta y ha ido bien. La meto al LOG
            vLog.AnyadeTextoDatosDes NUmSerie & Format(NumFac, "000000")
            'If OpcionActualizar = 10 Then
            '    'FRACLI
        End If
    End If
    IntegraFactura = B
    AlgunAsientoActualizado = B
    
    If Not DentroBeginTrans Then
        If B Then
            Conn.CommitTrans
        Else
            Conn.RollbackTrans
        End If
    End If
End Function

Private Function IntegraLaFactura(ByRef A_Donde As String) As Boolean
Dim cad As String
Dim Cad2 As String
Dim Cad3 As String
Dim Amplia2 As String
Dim DocConcAmp As String
Dim RF As Recordset
Dim ImporteNegativo As Boolean
Dim Importe0 As Boolean
Dim PrimeraContrapartida As String
    
    Dim SqlIva As String
    Dim RsIvas As ADODB.Recordset

    IntegraLaFactura = False
    'Sabemos que
    'numfac     --> CODIGO FACTURA
    'NumDiari       --> AÑO FACTURA
    'NUmSerie       --> SERIE DE LA FACTURA
    'FechaAsiento   --> Fecha factura
    'FecFactuAnt    --> FecFactura Anterior
    
    'Obtenemos los datos de la factura
    A_Donde = "Leyendo datos factura"
    Set RF = New ADODB.Recordset
    SQL = "SELECT * FROM factcli"
    SQL = SQL & " WHERE numserie='" & NUmSerie
    SQL = SQL & "' AND numfactu= " & NumFac
    SQL = SQL & " AND anofactu=" & NumDiari
    RF.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RF.EOF Then
        MsgBox "No se encuentra la factura: " & vbCrLf & SQL, vbExclamation
        RF.Close
        Exit Function
    End If
    
 
    SQL = "select count(*) from hcabapu where numdiari = " & DBSet(DiarioFacturas, "N") & " and fechaent = " & DBSet(FechaAnterior, "F") & " and numasien = " & DBSet(NumAsiento, "N")
    If TotalRegistros(SQL) > 0 Then
        A_Donde = "Actualiza cabecera hco apuntes"
        
        SQL = "UPDATE hcabapu SET "
        SQL = SQL & " fechaent = " & DBSet(Fecha, "F")
        SQL = SQL & ", obsdiari = " & DBSet(RF!observa, "T", "N")
        SQL = SQL & " where numdiari = " & DBSet(DiarioFacturas, "N")
        SQL = SQL & " and fechaent = " & DBSet(FechaAnterior, "F")
        SQL = SQL & " and numasien = " & DBSet(NumAsiento, "N")
    
        Conn.Execute SQL
    Else
        'Cabecera del hco de apuntes
        A_Donde = "Inserta cabecera hco apuntes"
        SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES ("
        'Ant 17/OCT/2005
        'SQL = SQL & vParam.numdiacl & ",'" & Fecha & "'," & NumAsiento
        SQL = SQL & DiarioFacturas & ",'" & Fecha & "'," & NumAsiento
        SQL = SQL & ","
        'Marzo 2010
        'Si tiene observaciones las llevo al apunte
        cad = DBLet(RF!observa, "T")
        If cad = "" Then
            cad = "NULL,"
        Else
            cad = "'" & DevNombreSQL(cad) & "',"
        End If
        cad = cad & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilización Factura de Cliente " & NUmSerie & Format(NumFac, "0000000") & " " & Fecha & "')"
        
        
        SQL = SQL & cad
        Conn.Execute SQL
    End If
    
    'Lineas fijas, es decir la linea de cliente, importes y tal y tal
    'Para el sql
    cad = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, "
    cad = cad & "codconce,ampconce, timporteD, timporteH,codccost, ctacontr, idcontab, punteada)"
    'Ant 17/oct/05
    'Cad = Cad & " VALUES (" & vParam.numdiacl & ",'" & Fecha & "'," & NumAsiento & ","
    cad = cad & " VALUES (" & DiarioFacturas & ",'" & Fecha & "'," & NumAsiento & ","
    Mes = 1 'Contador de lineas
    
    
    A_Donde = "Linea cliente"
    '-------------------------------------------------------------------
    'LINEA Cliente
    SQL = Mes & ",'" & RF!codmacta & "',"
    
    'AQUI ESTABA EL NUMERO DE SERIE y formateaba a 10 ceros. 22 sept 03
    'DocConcAmp = "'" & NUmSerie & Format(NumFac, "000000000") & "'," & vParam.concefcl & ",'"
    '[Monica]15/05/2015: en el numdocum ponemos serie y factura con formato 0000000
    DocConcAmp = "'" & NUmSerie & Format(NumFac, "0000000") & "'," & vParam.concefcl & ",'"
    
    
    'Ampliacion segun parametros
    Select Case vParam.nctafact
    Case 1
        If RF!totfaccl < 0 Then
            Cad2 = RecuperaValor(vParam.AmpliacionFacurasCli, 2)
        Else
            Cad2 = RecuperaValor(vParam.AmpliacionFacurasCli, 1)
        End If
        '28/02/2007.
        'Añado numerie
        Cad2 = Cad2 & " " & NUmSerie & Format(NumFac, "0000000")
    Case 2
        Cad2 = DevNombreSQL(DBLet(RF!Nommacta))
    Case Else
        Cad2 = DBLet(RF!confaccl)
    End Select
    
    '   Modificacion para k aparezca en la ampliacio el CC en la ampliacion de codmacta
    '
    Amplia2 = Cad2
    If vParam.CCenFacturas Then
        A_Donde = "CC en Facturas."
        Cad3 = DevuelveCentroCosteFactura(True, PrimeraContrapartida)
        If Cad3 <> "" Then
            If Len(Amplia2) > 21 Then Amplia2 = Mid(Amplia2, 1, 21)
            'Opcion1
            'Amplia2 = Amplia2 & " .CC:" & Cad3
            'Opcion2
            Amplia2 = Amplia2 & " [" & Cad3 & "]"
        End If
    End If
    A_Donde = "Linea cliente"
    
    
    SQL = SQL & DocConcAmp & Amplia2 & "'"
    DocConcAmp = DocConcAmp & Cad2 & "'"   'DocConcAmp Sirve para el IVA
    
    'Esta variable sirve para las demas
    ImporteNegativo = (DBLet(RF!totfaccl, "N") < 0)
    
    'Importes, atencion importes negativos
    '  antes --> Cad2 = CadenaImporte(ImporteNegativo, True, RF!totfaccl)
    Cad2 = CadenaImporte(True, DBLet(RF!totfaccl, "N"), Importe0)
    SQL = SQL & "," & Cad2 & ",NULL,"
    
    'Contrpartida. 28 Marzo 2006
    If PrimeraContrapartida <> "" Then
        SQL = SQL & "'" & PrimeraContrapartida & "'"
    Else
        SQL = SQL & "NULL"
    End If
    SQL = SQL & ",'FRACLI',0)"
    
    
    Conn.Execute cad & SQL
    Mes = Mes + 1 'Es el contador de lineaapunteshco
    
    ' cuentas de iva ahora se sacan de las tablas de totales
    SqlIva = "select * from factcli_totales "
    SqlIva = SqlIva & " WHERE numserie='" & NUmSerie
    SqlIva = SqlIva & "' AND numfactu= " & NumFac
    SqlIva = SqlIva & " AND anofactu=" & NumDiari
    SqlIva = SqlIva & " order by numlinea "
    
    Set RsIvas = New ADODB.Recordset
    RsIvas.Open SqlIva, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RsIvas.EOF
        Cad3 = "cuentarr"
        Cad2 = DevuelveDesdeBD("cuentare", "tiposiva", "codigiva", RsIvas!codigiva, "N", Cad3)
        If Cad2 <> "" Then
            SQL = Mes & ",'" & Cad2 & "'," & DocConcAmp
            Cad2 = CadenaImporte(False, RsIvas!Impoiva, Importe0)
            SQL = SQL & "," & Cad2 & ","
            SQL = SQL & "NULL,'" & RF!codmacta & "','FRACLI',0)"
'            If Not Importe0 Then
                Conn.Execute cad & SQL
                Mes = Mes + 1
'            End If
            
            'La de recargo  1-----------------
            If Not IsNull(RsIvas!ImpoRec) Then
                     SQL = Mes & "," & Cad3 & "," & DocConcAmp
                    'Importes, atencion importes negativos
                    Cad2 = CadenaImporte(False, RsIvas!ImpoRec, Importe0)
                    SQL = SQL & "," & Cad2 & ","
                    SQL = SQL & "NULL,'" & RF!codmacta & "','FRACLI',0)"
                    If Not Importe0 Then
                        Conn.Execute cad & SQL
                        Mes = Mes + 1
                    End If
            End If
        Else
            MsgBox "Error leyendo TIPO de IVA: " & RsIvas!codigiva, vbExclamation
            RF.Close
            Exit Function
        End If
    
        RsIvas.MoveNext
    Wend
    Set RsIvas = Nothing
    
    '-------------------------------------
    ' RETENCION
    A_Donde = "Retencion"
    If Not IsNull(RF!cuereten) Then
        SQL = Mes & ",'" & RF!cuereten & "'," & DocConcAmp
        'Importes, atencion importes negativos
        Cad2 = CadenaImporte(True, RF!trefaccl, Importe0)
        SQL = SQL & "," & Cad2 & ","
        SQL = SQL & "NULL,NULL,'FRACLI',0)"
       
        Conn.Execute cad & SQL
        Mes = Mes + 1 'Es el contador de lineaapunteshco
    End If
    
    
    IncrementaProgres 2
    
    '------------------------------------------------------------
    'Las lineas de la factura. Para ello guardaremos algunos datos
    Cad2 = RF!codmacta
    ImporteD = DBLet(RF!totfaccl, "N")
    
    
    'Cerramos el RF
    Cuenta = RF!codmacta
    RF.Close
    
    
    
    A_Donde = "Leyendo lineas factura"
    SQL = "Select factcli_lineas.* , cuentas.codmacta FROM factcli_lineas,Cuentas "
    SQL = SQL & " WHERE numserie='" & NUmSerie
    SQL = SQL & "' AND numfactu= " & NumFac
    SQL = SQL & " AND anofactu=" & NumDiari
    SQL = SQL & " AND factcli_lineas.codmacta = Cuentas.codmacta"
    RF.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Para cada linea insertamos
    Cad2 = ""
    A_Donde = "Procesando lineas"
    While Not RF.EOF
        'Importes, atencion importes negativos
        If Cad2 = "" Then PrimeraContrapartida = RF!codmacta
        SQL = Mes & ",'" & RF!codmacta & "'," & DocConcAmp
        Cad2 = CadenaImporte(False, RF!Baseimpo, Importe0)
        SQL = SQL & "," & Cad2 & ","
        If IsNull(RF!codccost) Then
            Cad2 = "NULL"
        Else
            Cad2 = "'" & RF!codccost & "'"
        End If
        
        SQL = SQL & Cad2 & ",'" & Cuenta & "','FRACLI',0)"
    
        Conn.Execute cad & SQL
        Mes = Mes + 1 'Es el contador de lineaapunteshco
        
        'Siguiente
        IncrementaProgres 1
        RF.MoveNext
        If Not RF.EOF Then PrimeraContrapartida = ""
    Wend
    RF.Close
    
    
    
    
    'AHora viene lo bueno.  MARZO 2006
    'Si el valor fuera true YA lo habria insertado en la cabcera
    If Not vParam.CCenFacturas Then
        If PrimeraContrapartida <> "" Then
            SQL = "UPDATE factcli_lineas SET codmacta ='" & PrimeraContrapartida & "'"
            SQL = SQL & " WHERE numdiari = " & DiarioFacturas & " AND fechaent ='" & Fecha & "' and numasien = " & NumAsiento
            SQL = SQL & " AND numlinea =1 " 'LA PRIMERA LINEA SIEMPRE ES LA DE LA CUENTA
            EjecutaSQL SQL  'Lo hacemos aqui para controlar el error y que no explote
        End If
    End If
        
    
    
    
    'Actualimos en factura, el nº de asiento
    SQL = "UPDATE factcli SET numdiari = " & DiarioFacturas & ", fechaent = '" & Fecha & "', numasien =" & NumAsiento
    SQL = SQL & " WHERE numserie='" & NUmSerie
    SQL = SQL & "' AND numfactu= " & NumFac
    SQL = SQL & " AND anofactu= " & NumDiari
    Conn.Execute SQL
    
    'Para los saldos ponemos el numero de asiento donde toca
    '
    A_Donde = "Saldos factura"
    NumDiari = vParam.numdiacl
    NumDiari = DiarioFacturas
    Mes = Month(FechaAsiento)
    Anyo = Year(FechaAsiento)
    'Actualizaremos los saldos
'    If Not CalcularLineasYSaldosFacturas Then Exit Function
    
    IntegraLaFactura = True
End Function



'////////////////////////////////////////////////////////////////////
'
'           Facturas proveedores
Private Function IntegraLaFacturaProv(ByRef A_Donde As String) As Boolean
Dim cad As String
Dim Cad2 As String
Dim Cad3 As String
Dim DocConcAmp As String
Dim Amplia2 As String
Dim RF As Recordset
Dim ImporteNegativo As Boolean
Dim Importe0 As Boolean 'Para saber si el importe es 0
Dim PrimeraContrapartida As String  'Si hay solo una linea entonces la pondremos como contrapartida de la primera base


'Modificacion de 31 Enero 2005
'-------------------------------------
'-------------------------------------
Dim ColumnaIVA As String
Dim TipoDIva As Byte
    
    Dim SqlIva As String
    Dim RsIvas As ADODB.Recordset

    IntegraLaFacturaProv = False
    
    'Sabemos que
    'numfac     --> CODIGO FACTURA
    'NumDiari       --> AÑO FACTURA
    'FechaAsiento   --> Fecha factura
    
    
    'Obtenemos los datos de la factura
    A_Donde = "Leyendo datos factura"
    Set RF = New ADODB.Recordset
    SQL = "SELECT * FROM factpro"
    SQL = SQL & " WHERE numregis = " & NumFac
    SQL = SQL & " AND anofactu=" & NumDiari
    RF.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RF.EOF Then
        MsgBox "No se encuentra la factura: " & vbCrLf & SQL, vbExclamation
        RF.Close
        Exit Function
    End If
    
    SQL = "select count(*) from hcabapu where numdiari = " & DBSet(DiarioFacturas, "N") & " and fechaent = " & DBSet(FechaAnterior, "F") & " and numasien = " & DBSet(NumAsiento, "N")
    If TotalRegistros(SQL) > 0 Then
        A_Donde = "Actualiza cabecera hco apuntes"
        
        SQL = "UPDATE hcabapu SET "
        SQL = SQL & " fechaent = " & DBSet(Fecha, "F")
        SQL = SQL & ", obsdiari = " & DBSet(RF!observa, "T", "N")
        SQL = SQL & " where numdiari = " & DBSet(DiarioFacturas, "N")
        SQL = SQL & " and fechaent = " & DBSet(FechaAnterior, "F")
        SQL = SQL & " and numasien = " & DBSet(NumAsiento, "N")
    
        Conn.Execute SQL
    Else
        'Cabecera del hco de apuntes
        A_Donde = "Inserta cabecera hco apuntes"
        SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES ("
        'SQL = SQL & vParam.numdiapr & ",'" & Fecha & "'," & NumAsiento
        SQL = SQL & DiarioFacturas & ",'" & Fecha & "'," & NumAsiento
        
        'Marzo 2010
        'Si tiene observaciones las llevo al apunte
        cad = DBLet(RF!observa, "T")
        If cad = "" Then
            cad = "NULL,"
        Else
            cad = "'" & DevNombreSQL(cad) & "',"
        End If
        
        cad = cad & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilización Factura Proveedor Registro " & Format(NumFac, "0000000") & " " & Fecha & "')"
        
        SQL = SQL & "," & cad
        
        Conn.Execute SQL
        
    End If
    
    
    
    'Lineas fijas, es decir la linea de cliente, importes y tal y tal
    'Para el sql
    cad = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, "
    cad = cad & "codconce,ampconce, timporteD, timporteH,codccost, ctacontr, idcontab, punteada)"
    'Cad = Cad & " VALUES (" & vParam.numdiapr & ",'" & Fecha & "'," & NumAsiento & ","
    cad = cad & " VALUES (" & DiarioFacturas & ",'" & Fecha & "'," & NumAsiento & ","
    Mes = 1 'Contador de lineas
    PrimeraContrapartida = ""
    
    'Esta variable sirve para las demas
    ImporteNegativo = (RF!totfacpr < 0)
    A_Donde = "Linea proveedor"
    '-------------------------------------------------------------------
    'LINEA Proveedor
    SQL = Mes & ",'" & RF!codmacta & "',"
    
    'Documento "numdocum"
    If vParam.CodiNume = 1 Then
        Cad2 = Format(NumFac, "0000000000")
    Else
        Cad2 = DBLet(RF!NumFactu)
    End If
    

    DocConcAmp = "'" & Cad2 & "'," & vParam.concefpr & ",'"
    
    
    'Ampliacion segun parametros
    Select Case vParam.nctafact
    Case 1
        If RF!totfacpr < 0 Then
            Cad2 = RecuperaValor(vParam.AmpliacionFacurasPro, 2)
        Else
            Cad2 = RecuperaValor(vParam.AmpliacionFacurasPro, 1)
        End If
        Cad2 = Cad2 & " " & DevNombreSQL(RF!NumFactu)
        
        Cad2 = Cad2 & " (" & Format(RF!FecFactu, "ddmmyy") & ")"
    Case 2
        Cad2 = DevNombreSQL(DBLet(RF!Nommacta))
    Case Else
        Cad2 = DBLet(RF!confacpr)
    End Select
    
        
    
    'Modificacion para k aparezca en la ampliacio el CC en la ampliacion de codmacta
    '
    Amplia2 = Cad2
    If vParam.CCenFacturas Then
        A_Donde = "CC en Facturas."
        Cad3 = DevuelveCentroCosteFactura(False, PrimeraContrapartida)
        If Cad3 <> "" Then
            If Len(Amplia2) > 26 Then Amplia2 = Mid(Amplia2, 1, 26)
            'Opcion1
            'Amplia2 = Amplia2 & " .CC:" & Cad3
            'Opcion2
            Amplia2 = Amplia2 & "[" & Cad3 & "]"
        End If
    End If
    A_Donde = "Linea cliente"
    
    
    SQL = SQL & DocConcAmp & Amplia2 & "'"
    DocConcAmp = DocConcAmp & Cad2 & "'"   'DocConcAmp Sirve para el IVA
    
    
    'Importes, atencion importes negativos
    Cad2 = CadenaImporte(False, RF!totfacpr, Importe0)
    SQL = SQL & "," & Cad2 & ",NULL,"
    
    'Contrpartida. 28 Marzo 2006
    If PrimeraContrapartida <> "" Then
        SQL = SQL & "'" & PrimeraContrapartida & "'"
    Else
        SQL = SQL & "NULL"
    End If
    SQL = SQL & ",'FRAPRO',0)"
    
    Conn.Execute cad & SQL
    Mes = Mes + 1 'Es el contador de lineaapunteshco
    
    ' cuentas de iva ahora se sacan de las tablas de totales
    SqlIva = "select * from factpro_totales "
    SqlIva = SqlIva & " WHERE numserie='" & NUmSerie
    SqlIva = SqlIva & "' AND numregis= " & NumFac
    SqlIva = SqlIva & " AND anofactu=" & NumDiari
    SqlIva = SqlIva & " order by numlinea "
    
    
    Dim EsSujetoPasivo As Boolean
    Dim EsImportacion As Boolean
    
    EsImportacion = (DBLet(RF!codopera, "N") = 2)
    EsSujetoPasivo = ((DBLet(RF!codopera, "N") = 1) Or (DBLet(RF!codopera, "N") = 4))
    
    Set RsIvas = New ADODB.Recordset
    RsIvas.Open SqlIva, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RsIvas.EOF
        TipoDIva = DevuelveValor("select tipodiva from tiposiva where codigiva = " & DBSet(RsIvas!codigiva, "N"))
        If TipoDIva = 1 Then
            'Es iva NO deducible
            ColumnaIVA = "cuentasn"
        Else
            ColumnaIVA = "cuentaso"   'La normal
        End If
        
        Cad3 = "cuentasr"
        Cad2 = DevuelveDesdeBD(ColumnaIVA, "tiposiva", "codigiva", RsIvas!codigiva, "N", Cad3)
        If Cad2 <> "" Then
            SQL = Mes & ",'" & Cad2 & "'," & DocConcAmp
            Cad2 = CadenaImporte(True, RsIvas!Impoiva, Importe0)
            SQL = SQL & "," & Cad2 & ","
            SQL = SQL & "NULL,'" & RF!codmacta & "','FRAPRO',0)"
'            If Not Importe0 Then
            If Not EsImportacion Then
                Conn.Execute cad & SQL
                Mes = Mes + 1
            End If
            
            'La de recargo  1-----------------
            If Not IsNull(RsIvas!ImpoRec) Then
                SQL = Mes & "," & Cad3 & "," & DocConcAmp
                'Importes, atencion importes negativos
                Cad2 = CadenaImporte(True, RsIvas!ImpoRec, Importe0)
                SQL = SQL & "," & Cad2 & ","
                SQL = SQL & "NULL,'" & RF!codmacta & "','FRAPRO',0)"
                If Not Importe0 Then
                    Conn.Execute cad & SQL
                    Mes = Mes + 1
                End If
            End If
            
            If EsSujetoPasivo Then
                Cad3 = "cuentarr"
                Cad2 = DevuelveDesdeBD("cuentare", "tiposiva", "codigiva", RsIvas!codigiva, "N", Cad3)
                
                Cad3 = Cad2 & "|" & Cad3 & "|"
                
                
                SQL = Mes & ",'" & RecuperaValor(Cad3, 1) & "'," & DocConcAmp
                Cad2 = CadenaImporte(False, RsIvas!Impoiva, Importe0)
                SQL = SQL & "," & Cad2 & ","
                SQL = SQL & "NULL,'" & RF!codmacta & "','FRAPRO',0)"
                'If Not Importe0 Then
                    Conn.Execute cad & SQL
                    Mes = Mes + 1
                'End If
               
                If Not IsNull(RsIvas!ImpoRec) Then
                     SQL = Mes & "," & RecuperaValor(Cad3, 2) & "," & DocConcAmp
                    'Importes, atencion importes negativos
                    Cad2 = CadenaImporte(False, RsIvas!ImpoRec, Importe0)
                    SQL = SQL & "," & Cad2 & ","
                    SQL = SQL & "NULL,'" & RF!codmacta & "','FRAPRO',0)"
                    If Not Importe0 Then
                        Conn.Execute cad & SQL
                        Mes = Mes + 1
                    End If
                End If
            End If
            
        Else
            MsgBox "Error leyendo TIPO de IVA: " & RsIvas!codigiva, vbExclamation
            RF.Close
            Exit Function
        End If
    
        RsIvas.MoveNext
    Wend
    Set RsIvas = Nothing
    
    '-------------------------------------
    
    '-------------------------------------
    ' RETENCION
    A_Donde = "Retencion"
    If Not IsNull(RF!cuereten) Then
        SQL = Mes & ",'" & RF!cuereten & "'," & DocConcAmp
        'Importes, atencion importes negativos
        Cad2 = CadenaImporte(False, RF!trefacpr, Importe0)
        SQL = SQL & "," & Cad2 & ","
        SQL = SQL & "NULL,NULL,'FRAPRO',0)"
       
        Conn.Execute cad & SQL
        Mes = Mes + 1 'Es el contador de lineaapunteshco
    End If
    
    
    IncrementaProgres 2
    
    '------------------------------------------------------------
    'Las lineas de la factura. Para ello guardaremos algunos datos
    Cad2 = RF!codmacta
    ImporteD = RF!totfacpr
    
    
    
    'Cerramos el RF
    Cuenta = RF!codmacta
    RF.Close
    
    
    
    A_Donde = "Leyendo lineas factura"
    SQL = "Select factpro_lineas.*  FROM factpro_lineas "
    SQL = SQL & " WHERE numregis= " & NumFac
    SQL = SQL & " AND anofactu=" & NumDiari
    RF.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Para cada linea insertamos
    A_Donde = "Procesando lineas"
    Cad2 = ""
    While Not RF.EOF
        'Importes, atencion importes negativos
        If Cad2 = "" Then PrimeraContrapartida = RF!codmacta
        SQL = Mes & ",'" & RF!codmacta & "'," & DocConcAmp
        Cad2 = CadenaImporte(True, RF!Baseimpo, Importe0)
        SQL = SQL & "," & Cad2 & ","
        If IsNull(RF!codccost) Then
            Cad2 = "NULL"
        Else
            Cad2 = "'" & RF!codccost & "'"
        End If
        
        SQL = SQL & Cad2 & ",'" & Cuenta & "','FRAPRO',0)"
    
        Conn.Execute cad & SQL
        Mes = Mes + 1 'Es el contador de lineaapunteshco
        
        'Siguiente
        IncrementaProgres 1
        RF.MoveNext
        If Not RF.EOF Then PrimeraContrapartida = ""
    Wend
    RF.Close
    
    
    'AHora viene lo bueno.  MARZO 2006
    'Si el valor fuera true YA lo habria insertado en la cabcera
    If Not vParam.CCenFacturas Then
        If PrimeraContrapartida <> "" Then
            SQL = "UPDATE hlinapu SET ctacontr ='" & PrimeraContrapartida & "'"
            SQL = SQL & " WHERE numdiari = " & DiarioFacturas & " AND fechaent ='" & Fecha & "' and numasien = " & NumAsiento
            SQL = SQL & " AND linliapu =1 " 'LA PRIMERA LINEA SIEMPRE ES LA DE LA CUENTA
            EjecutaSQL SQL  'Lo hacemos aqui para controlar el error y que no explote
        End If
    End If
    
    'Actualimos en factura, el nº de asiento
    SQL = "UPDATE factpro SET numdiari = " & DiarioFacturas & ", fechaent = '" & Fecha & "', numasien =" & NumAsiento
    SQL = SQL & " WHERE  numregis = " & NumFac
    SQL = SQL & " AND anofactu=" & NumDiari
    Conn.Execute SQL
    
    'Para los saldos ponemos el numero de asiento donde toca
    '
    A_Donde = "Saldos factura"
    NumDiari = DiarioFacturas
    Mes = Month(FechaAsiento)
    Anyo = Year(FechaAsiento)
    'Actualizaremos los saldos
'    If Not CalcularLineasYSaldosFacturas Then Exit Function
    
    IntegraLaFacturaProv = True
End Function










Private Function ActualizaAsiento() As Boolean
    Dim bol As Boolean
    Dim Donde As String
    On Error GoTo EActualizaAsiento
    
    'Obtenemos el mes y el año
    Mes = Month(FechaAsiento)
    Anyo = Year(FechaAsiento)
    Fecha = Format(FechaAsiento, FormatoFecha)
    
    'Comprobamos que no existe en historico
    If AsientoExiste Then
        If OpcionActualizar = 1 Then
            MsgBox "El asiento ya existe. Fecha: " & Fecha & "     Nº: " & NumAsiento, vbExclamation
            Exit Function
        Else
            SQL = "Comprobar  -> El asiento ya existe. Fecha: " & Fecha & "     Nº: " & NumAsiento
            InsertaError SQL
        End If
    End If
    
    'Aqui bloquearemos
    
    Conn.BeginTrans
    bol = ActualizaElASiento(Donde)
    
EActualizaAsiento:
        If Err.Number <> 0 Then
            SQL = "Actualiza Asiento." & vbCrLf & "----------------------------" & vbCrLf
            SQL = SQL & Donde
            If OpcionActualizar = 1 Then
                MuestraError Err.Number, SQL, Err.Description
            Else
                SQL = Donde & " -> " & Err.Description
                SQL = Mid(SQL, 1, 200)
                InsertaError SQL
            End If
            bol = False
        End If
        If bol Then
            Conn.CommitTrans
            ActualizaAsiento = True
            AlgunAsientoActualizado = True
        Else
            If OpcionActualizar = 1 Then
                MsgBox "Error: " & Donde, vbExclamation
            Else
                'FALTA###
            End If
            Conn.RollbackTrans
        End If
End Function


Private Function ActualizaElASiento(ByRef A_Donde As String) As Boolean



    ActualizaElASiento = False
    
    'Insertamos en cabeceras
    A_Donde = "Insertando datos en historico cabeceras asiento"
    If Not InsertarCabecera Then Exit Function
    IncrementaProgres 1
    
    'Insertamos en lineas
    A_Donde = "Insertando datos en historico lineas asiento"
    If Not InsertarLineas Then Exit Function
    IncrementaProgres 2
    
    
    
    'Modificar saldos
    A_Donde = "Calculando Lineas y saldos "
    If Not CalcularLineasYSaldos(False) Then Exit Function
    
    
    'Borramos cabeceras y lineas del asiento
    A_Donde = "Borrar cabeceras y lineas en asientos"
    If Not BorrarASiento(True) Then Exit Function
    IncrementaProgres 2
    ActualizaElASiento = True
End Function


Private Function InsertarCabecera() As Boolean
On Error Resume Next

    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari) SELECT numdiari,fechaent,numasien,obsdiari from cabapu where "
    SQL = SQL & " numdiari =" & NumDiari
    SQL = SQL & " AND fechaent='" & Fecha & "'"
    SQL = SQL & " AND numasien=" & NumAsiento

    Conn.Execute SQL

    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarCabecera = False
    Else
        InsertarCabecera = True
    End If
End Function


Private Function InsertarCabeceraApuntes() As Boolean
On Error Resume Next

    SQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, obsdiari) SELECT numdiari,fechaent,numasien,obsdiari from hcabapu where "
    SQL = SQL & " numdiari =" & NumDiari
    SQL = SQL & " AND fechaent='" & Fecha & "'"
    SQL = SQL & " AND numasien=" & NumAsiento

    Conn.Execute SQL

    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarCabeceraApuntes = False
    Else
        InsertarCabeceraApuntes = True
    End If
End Function



Private Function AsientoExiste() As Boolean
    AsientoExiste = True
    SQL = "SELECT numdiari from hcabapu"
    SQL = SQL & " WHERE numdiari =" & NumDiari
    SQL = SQL & " AND fechaent='" & Fecha & "'"
    SQL = SQL & " AND numasien=" & NumAsiento
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If RS.EOF Then AsientoExiste = False
    RS.Close
    Set RS = Nothing
End Function


Private Function CalcularLineasYSaldos(EsDesdeRecalcular As Boolean) As Boolean
Dim Reparto As Boolean
Dim T As String



    On Error GoTo ECalcularLineasYSaldos

    Dim RL As Recordset
    Set RL = New ADODB.Recordset
    
    '----------------------------------------------------------------------
    ' Este hace el group by por cuenta pero ya cuando el select tiene
    ' el asiento que quiero.   19 Septiembre 2006
    '----------------------------------------------------------------------
    'AQUI###
    SQL = "SELECT sum(timporteD) AS SD, sum(timporteH) AS SH, codmacta"
    SQL = SQL & "  FROM"
    If EsDesdeRecalcular Then
        SQL = SQL & " hlinapu"
    Else
        SQL = SQL & " linapu"
    End If
    'SQL = SQL & " GROUP BY codmacta, numdiari, fechaent, numasien"
    SQL = SQL & " WHERE (((numdiari)= " & NumDiari
    SQL = SQL & ") AND ((fechaent)='" & Fecha & "'"
    SQL = SQL & ") AND ((numasien)=" & NumAsiento
    SQL = SQL & ")) group by codmacta"
    
        
    
   
    Set RL = New ADODB.Recordset
    RL.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        Cuenta = RL!codmacta
        If IsNull(RL!sD) Then
            ImporteD = 0
        Else
            'ImporteD = RL!tImporteD
            ImporteD = RL!sD
        End If
        If IsNull(RL!sH) Then
            ImporteH = 0
        Else
            'ImporteH = RL!tImporteH
            ImporteH = RL!sH
        End If
        
        
        'Sig
        RL.MoveNext
    Wend
    RL.Close
    IncrementaProgres 3
    If Not vParam.autocoste Then
        'NO tiene analitica
        CalcularLineasYSaldos = True
        Exit Function
    End If
    
    
    '------------------------------------------
    '       ANALITICA     -> Modificado para 2 de Julio, para subcentros de reparto
    
    If EsDesdeRecalcular Then
        T = "h"
    Else
        T = ""
    End If
    

    
    SQL = "SELECT timporteD AS SD, timporteH AS SH, codmacta,idsubcos," & T & "linapu.codccost"
    SQL = SQL & " FROM " & T & "linapu,ccoste WHERE ccoste.codccost=" & T & "linapu.codccost"
    'SQL = SQL & " GROUP BY codmacta, fechaent, numdiari, numasien, codccost"
    SQL = SQL & " AND numdiari=" & NumDiari
    SQL = SQL & " AND fechaent='" & Fecha & "'"
    SQL = SQL & " AND numasien=" & NumAsiento
    SQL = SQL & " AND " & T & "linapu.codccost Is Not Null;"
    
    
    RL.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        Cuenta = RL!codmacta
        CCost = RL!codccost
        ImporteD = DBLet(RL!sD, "N")
        ImporteH = DBLet(RL!sH, "N")
        Reparto = (RL!idsubcos = 1)
        If Not CalcularSaldosAnal Then
            RL.Close
            Exit Function
        End If
        If Reparto Then
            If Not HacerReparto(True) Then
                RL.Close
                Exit Function
            End If
        End If
        'Sig
        RL.MoveNext
    Wend
    RL.Close
    IncrementaProgres 2
    CalcularLineasYSaldos = True
    Exit Function
ECalcularLineasYSaldos:
    Err.Clear
End Function




Private Function HacerReparto(Actualizar As Boolean) As Boolean
Dim RR As ADODB.Recordset
Dim AD As Currency
Dim AH As Currency
Dim TD As Currency
Dim TH As Currency
Dim B As Boolean

    HacerReparto = False
    TD = ImporteD
    TH = ImporteH
    AD = 0
    AH = 0
    Set RR = New ADODB.Recordset
    SQL = "Select * from ccoste_lineas WHERE codccost = '" & CCost & "'"
    RR.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RR.EOF
        'Cargamos los porcentajes
        CCost = RR!subccost
        ImporteD = (RR!porccost) / 100
        ImporteH = ImporteD
        'Importe porcentajeado
        ImporteD = Round(ImporteD * TD, 2)
        ImporteH = Round(ImporteH * TH, 2)
        'Movemos al sguiente
        RR.MoveNext
        'Por si acaso los decimales quedan sueltos entonces
        'Los valores para el ultimo subcentro de reaparto se obtienen por diferencias
        'con el acumulado
        If RR.EOF Then
            ImporteD = TD - AD
            ImporteH = TH - AH
        Else
            'Acumulo
            AD = AD + ImporteD
            AH = AH + ImporteH
        End If
        If Actualizar Then
            B = CalcularSaldosAnal
        Else
            B = CalcularSaldosAnalDesactualizar
        End If
        If Not B Then
            RR.Close
            Exit Function
        End If
    Wend
    RR.Close
    HacerReparto = True
End Function


'/////////////////////////////////////////////////
'//
'//
'//     Calcula los saldos del asiento desde las facturas
'//     Estoes, el asiento esta ya en hco, con lo cual las tablas son de hco
Private Function CalcularLineasYSaldosFacturas() As Boolean
    Dim Reparto As Boolean
    Dim RL As Recordset
    Set RL = New ADODB.Recordset
    
    CalcularLineasYSaldosFacturas = False
   
    'Abril 2004. Objetivo : QUITAR GROUP BY
    SQL = "SELECT hlinapu.timporteD AS SD, hlinapu.timporteH AS SH, hlinapu.codmacta"
    'SQL = SQL & " , hlinapu.numdiari, hlinapu.fechaent, hlinapu.numasien"
    SQL = SQL & " From hlinapu"
    SQL = SQL & " WHERE (((hlinapu.numdiari)= " & NumDiari
    SQL = SQL & ") AND ((hlinapu.fechaent)='" & Fecha & "'"
    SQL = SQL & ") AND ((hlinapu.numasien)=" & NumAsiento
    SQL = SQL & "));"
    
   
    Set RL = New ADODB.Recordset
    RL.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        Cuenta = RL!codmacta
        If IsNull(RL!sD) Then
            ImporteD = 0
        Else
            'ImporteD = RL!tImporteD
            ImporteD = RL!sD
        End If
        If IsNull(RL!sH) Then
            ImporteH = 0
        Else
            'ImporteH = RL!tImporteH
            ImporteH = RL!sH
        End If
        
        
        'Sig
        RL.MoveNext
    Wend
    RL.Close
    IncrementaProgres 3
    If Not vParam.autocoste Then
        'NO tiene analitica
        CalcularLineasYSaldosFacturas = True
        Exit Function
    End If
    
    
    '------------------------------------------
    '       ANALITICA
    SQL = "SELECT hlinapu.timporteD AS SD, hlinapu.timporteH AS SH, hlinapu.codmacta,"
    SQL = SQL & " hlinapu.fechaent, hlinapu.numdiari, hlinapu.numasien, hlinapu.codccost,idsubcos"
    SQL = SQL & " From hlinapu,ccoste WHERE ccoste.codccost=hlinapu.codccost"
    SQL = SQL & " AND hlinapu.numdiari =" & NumDiari
    SQL = SQL & " AND hlinapu.fechaent='" & Fecha & "'"
    SQL = SQL & " AND hlinapu.numasien=" & NumAsiento
    SQL = SQL & " AND hlinapu.codccost Is Not Null;"
    RL.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        Cuenta = RL!codmacta
        CCost = RL!codccost
        ImporteD = DBLet(RL!sD, "N")
        ImporteH = DBLet(RL!sH, "N")
        Reparto = (RL!idsubcos = 1)
        If Not CalcularSaldosAnal Then
            RL.Close
            Exit Function
        End If
        'Sig
        
        If Reparto Then
            If Not HacerReparto(True) Then
                RL.Close
                Exit Function
            End If
        End If
        RL.MoveNext
    Wend
    RL.Close
    IncrementaProgres 2
    CalcularLineasYSaldosFacturas = True
End Function




Private Function InsertarLineas() As Boolean
On Error Resume Next
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab,punteada,traspasado)"
    SQL = SQL & " SELECT numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab,punteada,traspasado From linapu"
    SQL = SQL & " WHERE numasien = " & NumAsiento
    SQL = SQL & " AND numdiari = " & NumDiari
    SQL = SQL & " AND fechaent='" & Fecha & "'"
    Conn.Execute SQL
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        
        InsertarLineas = False
    Else
        InsertarLineas = True
    End If
End Function


Private Function InsertarLineasApuntes() As Boolean
On Error Resume Next
    SQL = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab,punteada,traspasado)"
    SQL = SQL & " SELECT numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab,punteada,traspasado From hlinapu"
    SQL = SQL & " WHERE numasien = " & NumAsiento
    SQL = SQL & " AND numdiari = " & NumDiari
    SQL = SQL & " AND fechaent='" & Fecha & "'"
    Conn.Execute SQL
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarLineasApuntes = False
    Else
        InsertarLineasApuntes = True
    End If
End Function





'-------------------------------------------------------
'-------------------------------------------------------
'ANALITICA
'-------------------------------------------------------
'-------------------------------------------------------

Private Function CalcularSaldosAnal() As Boolean
    
    CalcularSaldosAnal = CalcularSaldos1NivelAnal(vEmpresa.numnivel)

End Function

Private Function CalcularSaldosAnalDesactualizar() As Boolean
    'Dim i As Integer
    'CalcularSaldosAnalDesactualizar = False
    'For i = vEmpresa.numnivel To 1 Step -1
    CalcularSaldosAnalDesactualizar = CalcularSaldos1NivelAnalDesactualizar(vEmpresa.numnivel)

End Function

Private Function CalcularSaldos1NivelAnal(Nivel As Integer) As Boolean
    Dim ImpD As Currency
    Dim ImpH As Currency
    Dim TD As String
    Dim TH As String
    Dim Cta As String
    Dim I As Integer
    
    
    CalcularSaldos1NivelAnal = False
    I = DigitosNivel(Nivel)
    If I < 0 Then Exit Function
    
    Cta = Mid(Cuenta, 1, I)
    SQL = "Select debccost,habccost from hsaldosanal where "
    SQL = SQL & " codccost='" & CCost & "' AND"
    SQL = SQL & " Codmacta = '" & Cta & "' AND anoccost = " & Anyo & " AND mesccost = " & Mes
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        I = 0   'Nuevo
        ImpD = 0
        ImpH = 0
    Else
        I = 1
        ImpD = RS.Fields(0)
        ImpH = RS.Fields(1)
    End If
    RS.Close
    'Acumulamos
    ImpD = ImpD + ImporteD
    ImpH = ImpH + ImporteH
    TD = TransformaComasPuntos(CStr(ImpD))
    TH = TransformaComasPuntos(CStr(ImpH))
    If I = 0 Then
        'Nueva insercion
        SQL = "INSERT INTO hsaldosanal(codccost,codmacta,anoccost,mesccost,debccost,habccost)"
        SQL = SQL & " VALUES('" & CCost & "','" & Cta & "'," & Anyo & "," & Mes & "," & TD & "," & TH & ")"
        Else
        SQL = "UPDATE hsaldosanal SET debccost=" & TD & ", habccost = " & TH
        SQL = SQL & " WHERE Codmacta = '" & Cta & "' AND Anoccost = " & Anyo & " AND mesccost = " & Mes
        SQL = SQL & " AND codccost = '" & CCost & "';"
    End If
    Conn.Execute SQL
    CalcularSaldos1NivelAnal = True
End Function



Private Function CalcularSaldos1NivelAnalDesactualizar(Nivel As Integer) As Boolean
    Dim ImpD As Currency
    Dim ImpH As Currency
    Dim TD As String
    Dim TH As String
    Dim Cta As String
    Dim I As Integer
    Dim NoHaySaldoContinuar As Boolean
    
    CalcularSaldos1NivelAnalDesactualizar = False
    I = DigitosNivel(Nivel)
    If I < 0 Then Exit Function
    
    Cta = Mid(Cuenta, 1, I)
    SQL = "Select debccost,habccost from hsaldosanal where "
    SQL = SQL & " codccost='" & CCost & "' AND"
    SQL = SQL & " Codmacta = '" & Cta & "' AND anoccost = " & Anyo & " AND mesccost = " & Mes
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        I = 0
        'If vUsu.Nivel = 0 Then
            SQL = "Error grave. No habia saldos en analitica: " & vbCrLf
            SQL = SQL & "Cuenta:    " & Cta & "      " & CCost & vbCrLf
            SQL = SQL & "Mes-año:     " & Mes & " / " & Anyo & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then
                NoHaySaldoContinuar = False
            Else
                NoHaySaldoContinuar = True
            End If
            ImpD = 0
            ImpH = 0
        'SEPT 2010
        'Para que no les moleste, cualquier usuario puede corregir el error
        'Else
        '    MsgBox "Error grave. No habia saldos en analitica para la cuenta: " & Cta & " " & CCost, vbCritical
        '    NoHaySaldoContinuar = False
        'End If
        
        
        If Not NoHaySaldoContinuar Then
            
                
            RS.Close
            Exit Function
        End If
    Else
        I = 1
        ImpD = RS.Fields(0)
        ImpH = RS.Fields(1)
    End If
    RS.Close
    'Acumulamos
    ImpD = ImpD - ImporteD 'Con respecto a ACTUALIZAR CAMBIA EL SIGNO
    ImpH = ImpH - ImporteH
    TD = TransformaComasPuntos(CStr(ImpD))
    TH = TransformaComasPuntos(CStr(ImpH))
    If I > 0 Then
        SQL = "UPDATE hsaldosanal SET debccost=" & TD & ", habccost = " & TH
        SQL = SQL & " WHERE Codmacta = '" & Cta & "' AND Anoccost = " & Anyo & " AND mesccost = " & Mes
        SQL = SQL & " AND codccost = '" & CCost & "';"
        Conn.Execute SQL
    Else
        SQL = "INSERT INTO hsaldosanal (codmacta, anoccost, mesccost, debccost, habccost,codccost) VALUES "
        SQL = SQL & "('" & Cta & "'," & Anyo & "," & Mes & ","
        SQL = SQL & TD & "," & TH & ",'" & CCost & "')"
        EjecutaSQL SQL   'Para que si da error nos deje tranquilos
    End If
    CalcularSaldos1NivelAnalDesactualizar = True
End Function




Private Function BorrarASiento(BorrarCabecera As Boolean) As Boolean

On Error GoTo EBorrarASiento
    BorrarASiento = False
    
    'Borramos las lineas
    SQL = "Delete from hlinapu"
    SQL = SQL & " WHERE numasien = " & NumAsiento
    SQL = SQL & " AND numdiari = " & NumDiari
    SQL = SQL & " AND fechaent=" & DBSet(FechaAnterior, "F")
    Conn.Execute SQL
    
    If BorrarCabecera Then
        'La cabecera
        SQL = "Delete from hcabapu"
        SQL = SQL & " WHERE numdiari =" & NumDiari
        SQL = SQL & " AND fechaent=" & DBSet(FechaAnterior, "F")
        SQL = SQL & " AND numasien=" & NumAsiento
        
        Conn.Execute SQL
    Else
        'Actualizamos la fecha de la cabecera
        SQL = "Update hcabapu"
        SQL = SQL & " set fechaent = " & DBSet(Fecha, "F")
        SQL = SQL & " WHERE numdiari =" & NumDiari
        SQL = SQL & " AND fechaent=" & DBSet(FechaAnterior, "F")
        SQL = SQL & " AND numasien=" & NumAsiento
    
        Conn.Execute SQL
    End If
    
    BorrarASiento = True
    Exit Function
EBorrarASiento:
    Err.Clear
    
End Function

Private Sub ObtenFoco(ByRef T As TextBox)
T.SelStart = 0
T.SelLength = Len(T.Text)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If NumErrores > 0 Then CerrarFichero
    If OpcionActualizar = 5 Then ChkListaAsientos False
End Sub

Private Sub CerrarFichero()
On Error Resume Next
If NE = 0 Then Exit Sub
Close #NE
If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub frmC_Selec(vFecha As Date)
Text3(INC).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmD_DatoSeleccionado(CadenaSeleccion As String)
    If INC > -1 Then
        Text2(INC).Text = RecuperaValor(CadenaSeleccion, 1)
        Text4(INC).Text = RecuperaValor(CadenaSeleccion, 2)
    Else
        Text5.Text = RecuperaValor(CadenaSeleccion, 1)
        Text6.Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub Image1_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmD = New frmTiposDiario
    INC = Index
    frmD.DatosADevolverBusqueda = "0"
    frmD.Show vbModal
    Set frmD = Nothing
End Sub

Private Sub Image2_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text3(Index).Text <> "" Then frmC.Fecha = CDate(Text3(Index).Text)
    INC = Index
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub Image3_Click()
'    Screen.MousePointer = vbHourglass
'    Set frmD = New frmTiposDiario
'    INC = -1
'    frmD.DatosADevolverBusqueda = "0"
'    frmD.Show vbModal
'    Set frmD = Nothing
    Image1_Click -1
End Sub

Private Sub imgCheck_Click(Index As Integer)
    For INC = 1 To ListView2.ListItems.Count
        ListView2.ListItems(INC).Checked = (Index = 1)
    Next INC
End Sub

Private Sub Text1_GotFocus(Index As Integer)
ObtenFoco Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).Text = Trim(Text1(Index).Text)
If Text1(Index).Text = "" Then Exit Sub

If Not IsNumeric(Text1(Index).Text) Then
    MsgBox "El Nº asiento debe ser numérico.", vbExclamation
    Text1(Index).Text = ""
    Text1(Index).SetFocus
    Exit Sub
End If

End Sub

Private Sub Text2_GotFocus(Index As Integer)
ObtenFoco Text2(Index)
End Sub

'++
Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYBusqueda KeyAscii, 0
            Case 1:  KEYBusqueda KeyAscii, 1
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    Image1_Click (Indice)
End Sub
'++




Private Sub Text2_LostFocus(Index As Integer)
Text2(Index).Text = Trim(Text2(Index).Text)
Text4(Index).Text = ""
If Text2(Index).Text = "" Then Exit Sub

If Not IsNumeric(Text2(Index).Text) Then
    MsgBox "El código diario debe ser numérico.", vbExclamation
    Text2(Index).Text = ""
    Text2(Index).SetFocus
    Exit Sub
End If

SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text2(Index).Text, "N")

If SQL = "" Then
    MsgBox "Diario NO encontrado: " & Text2(Index).Text, vbExclamation
    Text2(Index).Text = ""
End If
Text4(Index).Text = SQL
End Sub

Private Sub Text3_GotFocus(Index As Integer)
ObtenFoco Text3(Index)
End Sub

'++
Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0:  KEYFecha KeyAscii, 0
            Case 1:  KEYFecha KeyAscii, 1
            Case 2:  KEYFecha KeyAscii, 2
            Case 3:  KEYFecha KeyAscii, 3
            
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    Image2_Click (Indice)
End Sub
'++



Private Sub Text3_LostFocus(Index As Integer)
Text3(Index).Text = Trim(Text3(Index).Text)

If Text3(Index).Text = "" Then Exit Sub

If Not EsFechaOK(Text3(Index)) Then
    MsgBox "Fecha incorrecta: " & Text3(Index).Text, vbExclamation
    Text3(Index).Text = ""
    Text3(Index).SetFocus
    Exit Sub
End If
Text3(Index).Text = Format(Text3(Index).Text, "dd/mm/yyyy")
End Sub



Private Function ObtenerSQL() As String
Dim cad As String
Dim Aux As String
Dim CADENA As String

ObtenerSQL = ""
cad = ""

'Comprobacioines
If Text1(0).Text <> "" And Text1(1).Text <> "" Then
    If Val(Text1(0).Text) > Val(Text1(1).Text) Then
        MsgBox "Nº asiento hasta mayor que desde", vbExclamation
        Exit Function
    End If
End If
If Text2(0).Text <> "" And Text2(1).Text <> "" Then
    If Val(Text2(0).Text) > Val(Text2(1).Text) Then
        MsgBox "Diario desde mayor que hasta", vbExclamation
        Exit Function
    End If
End If

If Text3(0).Text <> "" And Text3(1).Text <> "" Then
    If CDate(Text3(0).Text) > CDate(Text3(1).Text) Then
    MsgBox "Fecha Desde mayor que hasta", vbExclamation
    Exit Function
    End If
End If




If Me.OpcionActualizar = 4 Then
    'Cadena = "{ado.numasien}"
    CADENA = "cabapu_0.numasien"
Else
   CADENA = "numasien"
End If

'Nº asiento
Aux = ""
If Text1(0).Text <> "" Then Aux = " " & CADENA & "  >= " & Text1(0).Text
If Text1(1).Text <> "" Then
    If Aux <> "" Then Aux = Aux & " AND "
    Aux = Aux & " " & CADENA & "  <= " & Text1(1).Text
End If

If Aux <> "" Then cad = "( " & Aux & ")"

'Nº diario
If Me.OpcionActualizar = 4 Then
    CADENA = "cabapu_0.numdiari"
Else
    CADENA = "numdiari"
End If

Aux = ""
If Text2(0).Text <> "" Then Aux = " " & CADENA & "  >= " & Text2(0).Text
If Text2(1).Text <> "" Then
    If Aux <> "" Then Aux = Aux & " AND "
    Aux = Aux & " " & CADENA & "  <= " & Text2(1).Text
End If

If Aux <> "" Then
    If cad <> "" Then cad = cad & " AND "
    cad = cad & "(" & Aux & ")"
End If

'Fecha
Aux = ""
If Me.OpcionActualizar = 4 Then
    CADENA = "cabapu_0.fechaent"
    Else
    CADENA = "fechaent"
End If

    If Text3(0).Text <> "" Then Aux = CADENA & " >= '" & Format(Text3(0).Text, FormatoFecha) & "'"
    If Text3(1).Text <> "" Then
        If Aux <> "" Then Aux = Aux & " AND "
        Aux = Aux & CADENA & " <= '" & Format(Text3(1).Text, FormatoFecha) & "'"
    End If


If Aux <> "" Then
    If cad <> "" Then cad = cad & " AND "
    cad = cad & "(" & Aux & ")"
End If


If cad = "" Then cad = CADENA & " >= '2000-01-01'"
    

ObtenerSQL = cad
End Function


Private Function ObtenerRegistrosParaActualizar() As Boolean
Dim cad As String

ObtenerRegistrosParaActualizar = False
'Borramos temporal
Conn.Execute "Delete From tmpActualizar where codusu = " & vUsu.Codigo
Conn.Execute "Delete From tmpactualizarError where codusu = " & vUsu.Codigo

Set RS = New ADODB.Recordset
RS.Open "Select count(*) from Cabapu WHERE " & SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If RS.EOF Then
    'NINGUN REGISTTRO A ACTUALIZAR
    NumAsiento = 0
Else
    NumAsiento = RS.Fields(0)
End If
RS.Close
If NumAsiento = 0 Then
    MsgBox "Ningún asiento para actualizar entre estos valores.", vbExclamation
    Exit Function
End If

'Cargamos valores
If NumAsiento < 32000 Then
    CargaProgres CInt(NumAsiento)
    INC = 1
End If

'Ponemos en marcha la peli
If NumAsiento > 4 Then
    PonerAVI 1
    Me.Refresh
    DoEvents
    espera 1
End If



'Ponemos el form como toca
Label1.Caption = "Obtener registros actualización."
lblAsiento.Caption = ""
Me.Height = 3315
Me.frame1Asiento.Visible = True
Me.Refresh
Me.Height = 3315
Me.Refresh

RS.Open "Select * from Cabapu WHERE " & SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
While Not RS.EOF
    IncrementaProgres 1
    'Para poder acceder a ellos desde cualquier sitio
    NumAsiento = RS!NumAsien
    FechaAsiento = RS!FechaEnt
    Fecha = Format(RS!FechaEnt, FormatoFecha)
    NumDiari = RS!NumDiari
    If RS!bloqactu <> 0 Then
        'INSERTAERROR
        InsertaError "Asiento bloqueado"
        
    Else
        'No esta bloqueado
        'Comprobamos que esta cuadrado
        cad = RegistroCuadrado
        If cad = "" Then
            cad = BloqAsien
        End If
        If cad <> "" Then InsertaError cad
    End If
    'Siguiente
    RS.MoveNext
    
    'Si esta visible el avi
    If Animation1.Visible Then espera 0.5
Wend
RS.Close
Set RS = Nothing
ObtenerRegistrosParaActualizar = True
End Function

Private Function BloqAsien() As String
Dim C As String
On Error Resume Next
'Bloqueamos e insertamos
BloqAsien = ""
C = ""
If BloquearAsiento(CStr(NumAsiento), CStr(NumDiari), Fecha, C) Then
    'Utilizamos una variable existente
    Cuenta = "INSERT INTO tmpactualizar (numdiari, fechaent, numasien, codusu) VALUES ("
    Cuenta = Cuenta & NumDiari & ",'"
    Cuenta = Cuenta & Fecha & "',"
    Cuenta = Cuenta & NumAsiento & ","
    Cuenta = Cuenta & vUsu.Codigo & ")"
    Conn.Execute Cuenta
    If Err.Number <> 0 Then
        Err.Clear
        BloqAsien = "Error al insertar temporal"
        DesbloquearAsiento CStr(NumAsiento), CStr(NumDiari), Fecha
    End If
Else
    If C <> "" Then
        BloqAsien = C
    Else
        BloqAsien = "Error al bloquear el asiento."
    End If
End If
End Function

Private Sub PonerAVI(NumAVI As Integer)
On Error GoTo EPonerAVI
If NumAVI = 1 Then
    Me.Animation1.Open App.Path & "\actua.avi"
    Me.Animation1.Play
    Me.Animation1.Visible = True
Else
    Me.Animation2.Open App.Path & "\actua.avi"
    Me.Animation2.Visible = True
    Me.Animation2.Play
End If
Exit Sub
EPonerAVI:
    MuestraError Err.Number, "Poner Video"
End Sub


Private Function RegistroCuadrado() As String
    Dim Deb As Currency
    Dim hab As Currency
    Dim RSUM As ADODB.Recordset

    'Trabajamos con RS que es global
    RegistroCuadrado = "" 'Todo bien
    
    
    
    'Primero compruebo el ambito de fechas
    varFecOk = FechaCorrecta2(FechaAsiento)
    If varFecOk > 1 Then
        If varFecOk = 2 Then
            RegistroCuadrado = varTxtFec
        Else
            'Fecha fuera de ejerecicios
            RegistroCuadrado = "Fecha fuera de ejercicios"
        End If
        Exit Function
    End If
    
    Set RSUM = New ADODB.Recordset
    SQL = "SELECT Sum(linapu.timporteD) AS SumaDetimporteD, Sum(linapu.timporteH) AS SumaDetimporteH"
    SQL = SQL & " ,linapu.numdiari,linapu.fechaent,linapu.numasien"
    SQL = SQL & " From linapu GROUP BY linapu.numdiari, linapu.fechaent, linapu.numasien "
    SQL = SQL & " HAVING (((linapu.numdiari)=" & NumDiari
    SQL = SQL & ") AND ((linapu.fechaent)='" & Fecha
    SQL = SQL & "') AND ((linapu.numasien)=" & NumAsiento
    SQL = SQL & "));"
    
    
    
    
    
    
    RSUM.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RSUM.EOF Then
        Deb = DBLet(RSUM.Fields(0), "N")
        'Deb = Round(Deb, 2)
        hab = DBLet(RSUM.Fields(1), "N")
        'Hab = Round(Hab, 2)
        CCost = ""
    Else
        Deb = 0
        hab = -1
        CCost = "Asiento sin lineas"
    End If
    
    RSUM.Close
    Set RSUM = Nothing
    If Deb <> hab Then
        If CCost = "" Then CCost = "Asiento descuadrado"
        RegistroCuadrado = CCost
    End If
    
    
    
    
    
End Function

Private Function InsertaError(ByRef CADENA As String)
Dim vS As String
    'Insertamos en errores
    'Esta lo tratamos con error especifico
    
    On Error Resume Next

    If OpcionActualizar < 10 Then
        'Insertamos error para ASIENTOS
        vS = NumDiari & "|"
        vS = vS & Fecha & "|"
        vS = vS & NumAsiento & "|"
        vS = vS & CADENA & "|"
    
    Else
        vS = NUmSerie & " " & NumFac & "|"
        vS = vS & FechaAsiento & "|"
        vS = vS & CADENA & "|"
    End If
    'Modificacion del 10 de marzo
    'Conn.Execute vS
    AñadeError vS
    
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error." & vbCrLf & Err.Description & vbCrLf & vS
        Err.Clear
    End If
End Function


Private Function ActualizaASientosDesdeTMP()
Dim RT As Recordset


'Para el progress
NumAsiento = ProgressBar1.Max
Me.lblAsiento.Caption = "Nº asiento:"
If NumAsiento < 3000 Then
    CargaProgres NumAsiento * 10
    Else
    CargaProgres 32000
End If
INC = 1

vLog.InicializarDatosDesc

SQL = "Select * from tmpactualizar where codusu=" & vUsu.Codigo
Set RT = New ADODB.Recordset
RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
While Not RT.EOF
    NumAsiento = RT!NumAsien
    FechaAsiento = RT!FechaEnt
    NumDiari = RT!NumDiari
    'Actualiza el asiento
    If ActualizaAsiento = False Then
         DesbloquearAsiento CStr(NumAsiento), CStr(NumDiari), Fecha
    Else
        vLog.AnyadeTextoDatosDes CStr(NumAsiento)
        'Si tiene k imprimir al finalizar entonces
        If vParam.emitedia Then
            SQL = NumAsiento & "|" & Format(FechaAsiento, FormatoFecha) & "|" & NumDiari & "|"
            IHcoApuntesAlActualizarModificar (SQL)
        End If
    End If
    
    'Siguiente
    RT.MoveNext
Wend
RT.Close
Set RT = Nothing


vLog.Insertar 10, vUsu, vLog.DatosDescripcion


End Function




Private Function DesActualizaAsiento() As Boolean
    Dim bol As Boolean
    Dim Donde As String
    On Error GoTo EDesActualizaAsiento
    
    
    '2.- Desactualiza pero NO insertes en apuntes
    '3.- Desactualizar asiento desde hco
    
    'Obtenemos el mes y el año
    Mes = Month(FechaAsiento)
    Anyo = Year(FechaAsiento)
    Fecha = Format(FechaAsiento, FormatoFecha)
    
    'Comprobamos que no existe en APUNTES
    'Obviamente solo comprobamos si vamos a insertar
    'en apuntes
    If Me.OpcionActualizar = 3 Then
        If AsientoExiste Then Exit Function
    End If
    'Aqui bloquearemos
    
    Conn.BeginTrans
    
    bol = DesActualizaElASiento(Donde)
    
EDesActualizaAsiento:
        If Err.Number <> 0 Then
            SQL = "Actualiza Asiento." & vbCrLf & "----------------------------" & vbCrLf
            SQL = SQL & Donde
            MuestraError Err.Number, SQL, Err.Description
            bol = False
        End If
        If bol Then
            Conn.CommitTrans
            espera 0.2
            DesActualizaAsiento = True
            AlgunAsientoActualizado = True
        Else
            Conn.RollbackTrans
        End If
End Function


Private Function DesActualizaElASiento(ByRef A_Donde As String) As Boolean

    '2  .- Desactualiza pero NO insertes en apuntes
    '      Si viene FRACLI o FRAPROV habrá que volver
    '3  .- Desactualizar asiento desde hco
        


    DesActualizaElASiento = False
    
    Select Case Me.OpcionActualizar
    Case 3
        'Insertamos en cabeceras
        A_Donde = "Insertando datos en cabeceras de apuntes"
        If Not InsertarCabeceraApuntes Then Exit Function
        IncrementaProgres 1
        
        'Insertamos en lineas
        A_Donde = "Insertando datos en lineas asiento"
        If Not InsertarLineasApuntes Then Exit Function
        IncrementaProgres 2
    
    Case 2
        If NUmSerie = "FRACLI" Or NUmSerie = "FRAPRO" Then
            A_Donde = "Desvinculando facturas"
            If Not DesvincularFactura(NUmSerie = "FRACLI") Then Exit Function
            IncrementaProgres 1
        End If
    End Select
    
    'Modificar saldos
'    A_Donde = "Recalculando lineas y saldos"
'    If Not CalcularLineasYSaldosDesactualizar Then Exit Function
    
    'Borramos cabeceras y lineas del asiento
    A_Donde = "Borrar cabeceras y lineas en historico"
    
    If OpcionActualizar = 2 Then
        If Not BorrarASiento(False) Then Exit Function
    Else
        If Not BorrarASiento(True) Then Exit Function
    End If
    
    IncrementaProgres 2
    DesActualizaElASiento = True
End Function

Private Function DesvincularFactura(Clientes As Boolean) As Boolean
On Error Resume Next
    Set RS = New ADODB.Recordset
    If Clientes Then
        CCost = "factcli"
    Else
        CCost = "factpro"
    End If
    SQL = "Select * From " & CCost
    SQL = SQL & " WHERE numasien=" & NumAsiento
    SQL = SQL & " AND numdiari = " & NumDiari
    SQL = SQL & " AND fechaent = '" & Fecha & "'"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        
        SQL = "UPDATE " & CCost & " SET numasien=NULL, fechaent=NULL, numdiari=NULL"
        If Clientes Then
            SQL = SQL & " WHERE numfactu = " & RS!codfaccl
            SQL = SQL & " AND anofaccl =" & RS!anofaccl
            SQL = SQL & " AND numserie = '" & RS!NUmSerie & "'"
        Else
            'proveedores
            SQL = SQL & " WHERE numregis = " & RS!NumRegis
            SQL = SQL & " AND anofactu =" & RS!anofactu
        End If
        Conn.Execute SQL
    End If
    If Err.Number <> 0 Then
        DesvincularFactura = False
        MuestraError Err.Number, "Desvincular factura"
    Else
        DesvincularFactura = True
    End If
End Function


Private Function CalcularLineasYSaldosDesactualizar() As Boolean
    Dim RL As Recordset
    Dim Reparto As Boolean
    Set RL = New ADODB.Recordset
    
    
    '------------------------------------------
    'SALDOS
    'Calculamos sumas importes asiento en hco
    CalcularLineasYSaldosDesactualizar = False
    
    SQL = "SELECT sum(timporteD) AS SD, sum(timporteH) AS SH, codmacta"
    SQL = SQL & "  FROM  hlinapu"
    SQL = SQL & " WHERE (((numdiari)= " & NumDiari
    SQL = SQL & ") AND ((fechaent)='" & Fecha & "'"
    SQL = SQL & ") AND ((numasien)=" & NumAsiento
    SQL = SQL & ")) group by codmacta"
    
    
    
    Set RL = New ADODB.Recordset
    RL.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        Cuenta = RL!codmacta
        ImporteD = DBLet(RL!sD, "N")
        ImporteH = DBLet(RL!sH, "N")
        'Sig
        RL.MoveNext
    Wend
    RL.Close
    IncrementaProgres 3
    
    If Not vParam.autocoste Then
        'NO tiene analitica
        CalcularLineasYSaldosDesactualizar = True
        Exit Function
    End If
    
    
    
    '       ANALITICA
    SQL = "SELECT hlinapu.timporteD AS SD, hlinapu.timporteH AS SH, hlinapu.codmacta,"
    SQL = SQL & " hlinapu.fechaent, hlinapu.numdiari, hlinapu.numasien, hlinapu.codccost,ccoste.idsubcos"
    SQL = SQL & " From hlinapu,ccoste WHERE hlinapu.codccost=ccoste.codccost"
    SQL = SQL & " AND hlinapu.numdiari=" & NumDiari
    SQL = SQL & " AND hlinapu.fechaent='" & Fecha & "'"
    SQL = SQL & " AND hlinapu.numasien=" & NumAsiento
    SQL = SQL & " AND hlinapu.codccost Is Not Null;"
    RL.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        Cuenta = RL!codmacta
        CCost = RL!codccost
        ImporteD = DBLet(RL!sD, "N")
        ImporteH = DBLet(RL!sH, "N")
        Reparto = (RL!idsubcos = 1)
        If Not CalcularSaldosAnalDesactualizar Then
            RL.Close
            Exit Function
        End If
        If Reparto Then
            If Not HacerReparto(False) Then
                RL.Close
                Exit Function
            End If
        End If
        'Sig
        RL.MoveNext
    Wend
    RL.Close
    IncrementaProgres 2
    CalcularLineasYSaldosDesactualizar = True
End Function





Private Sub Text5_GotFocus()
    ObtenFoco Text5
End Sub

'++
Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        KEYBusqueda2 KeyAscii, 0
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda2(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    Image3_Click
End Sub
'++

Private Sub Text5_LostFocus()
    Text5.Text = Trim(Text5.Text)
    SQL = ""
    If Text5.Text <> "" Then
        If Not IsNumeric(Text5.Text) Then
            MsgBox "Diario debe ser numérico", vbExclamation
            Text5.Text = ""
        Else
            SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", Text5.Text, "N")
            If SQL = "" Then Text5.Text = ""
        End If
    End If
    Text6.Text = SQL
End Sub

Private Sub txtNumFac_GotFocus(Index As Integer)
    ObtenFoco txtnumfac(Index)
End Sub

Private Sub txtNumFac_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtNumFac_LostFocus(Index As Integer)
With txtnumfac(Index)
    .Text = Trim(.Text)
    If .Text = "" Then Exit Sub
    If Not IsNumeric(.Text) Then
        MsgBox "Campo numero factura debe de ser numérico", vbExclamation
        .Text = ""
        Exit Sub
    End If
End With
End Sub

Private Sub txtSerie_GotFocus(Index As Integer)
    ObtenFoco txtSerie(Index)
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub txtSerie_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub BorrarArchivoTemporal()
On Error Resume Next
If Dir(App.Path & "\ErrActua.txt") <> "" Then Kill App.Path & "\ErrActua.txt"
If Err.Number <> 0 Then MuestraError Err.Number, "Borrar fichero temporal"
End Sub

Private Sub txtSerie_LostFocus(Index As Integer)
If txtSerie(Index).Text <> "" Then txtSerie(Index).Text = UCase(txtSerie(Index).Text)
End Sub



Private Sub InsertandoEnRecalculodeSaldos()
        Label11.Caption = "Insertando ......."
        Label11.Refresh
        DoEvents
        Cuenta = SQL & Cuenta & ";"
        Conn.Execute Cuenta
        Cuenta = ""
        Me.Refresh
End Sub

Private Sub REcalculoDesdeAsiento()
    On Error GoTo ERec
    CalcularLineasYSaldos True
    Exit Sub
ERec:
    NE = 1
    MuestraError Err.Number
End Sub



Private Function DevuelveCentroCosteFactura(Cliente As Boolean, LaPrimeraContrapartida As String) As String
Dim R As ADODB.Recordset
Dim SQL As String
    DevuelveCentroCosteFactura = ""
    If Cliente Then
        
        SQL = "SELECT codccost,numlinea,codtbase FROM linfact"
        SQL = SQL & " WHERE numserie='" & NUmSerie
        SQL = SQL & "' AND codfaccl= " & NumFac
        SQL = SQL & " AND anofaccl=" & NumDiari
        SQL = SQL & " AND not (codccost is null)"   'El primero k devuelva
        SQL = SQL & " ORDER BY numlinea"
    Else
        SQL = "SELECT codccost,numlinea,codtbase FROM linfactprov"
        SQL = SQL & " WHERE numregis = " & NumFac
        SQL = SQL & " AND anofacpr=" & NumDiari
        SQL = SQL & " AND not (codccost is null)"   'El primero k devuelva
        SQL = SQL & " ORDER BY numlinea"
    End If
    
    
    Set R = New ADODB.Recordset
    R.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not R.EOF Then
        If Not IsNull(R.Fields(0)) Then DevuelveCentroCosteFactura = R.Fields(0)
        LaPrimeraContrapartida = R!codtbase
        R.MoveNext
        If Not R.EOF Then LaPrimeraContrapartida = ""
    End If
    R.Close
    Set R = Nothing
End Function



Private Sub ChkListaAsientos(leer As Boolean)

    On Error GoTo EChkListaAsientos
    SQL = App.Path & "\chkListato.xdf"
    If leer Then
        Me.chkMostrarListview.Value = 0
        If Dir(SQL, vbArchive) <> "" Then Me.chkMostrarListview.Value = 1
        
    Else
        If Me.chkMostrarListview.Value = 0 Then
            If Dir(SQL, vbArchive) <> "" Then Kill SQL
            
        Else
            If Dir(SQL, vbArchive) = "" Then
                Mes = FreeFile
                Open SQL For Output As #Mes
                Print vUsu.Nombre & " " & Now
                Close #Mes
            End If
        End If
        
    End If
    Exit Sub
EChkListaAsientos:
    MuestraError Err.Number, "Sub :  ChkListaAsientos"
End Sub


Private Sub CargaAsientosPorActualizar()

    ListView2.ListItems.Clear
    Set RS = New ADODB.Recordset
    
    If SQL <> "" Then SQL = " AND " & SQL
    SQL = "Select * from cabapu where bloqactu=0 " & SQL
    
    SQL = SQL & " ORDER by numdiari,fechaent,numasien"
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set ItmX = ListView2.ListItems.Add(, , RS!NumDiari)
        ItmX.SubItems(1) = Format(RS!FechaEnt, "dd/mm/yyyy")
        ItmX.SubItems(2) = RS!NumAsien
        ItmX.SubItems(3) = DBLet(RS!obsdiari, "T")
        ItmX.Checked = False
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub


Private Sub RecalculoAutomatico()
    Screen.MousePointer = vbHourglass
    DoEvents
    Bloquear_DesbloquearBD (True)
    PonerAVI 2
    espera 1
    Label11.Visible = True
    pb3.Visible = False
    NE = 0
    'El recalculo va aqui
    pb3.Visible = False
    Label11.Visible = False
    Animation2.Stop
    Animation2.Visible = False

    
    
    Bloquear_DesbloquearBD False
    Screen.MousePointer = vbDefault
    If NE = 0 Then
        MsgBox "El proceso ha finalizado correctamente.", vbInformation
        cmdRecalcula.Enabled = False
    End If
End Sub
