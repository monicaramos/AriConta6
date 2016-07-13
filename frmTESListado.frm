VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESListado 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listados"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12060
   Icon            =   "frmTESListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCompensaciones 
      Height          =   6045
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8235
      Begin VB.CheckBox chkCompensa 
         Caption         =   "Dejar sólo importe compensacion"
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
         Left            =   960
         TabIndex        =   13
         Top             =   5370
         Width           =   4005
      End
      Begin VB.Frame FrameCambioFPCompensa 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   120
         TabIndex        =   27
         Top             =   2400
         Width           =   7785
         Begin VB.TextBox txtDescFPago 
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
            Index           =   8
            Left            =   3360
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   240
            Width           =   4335
         End
         Begin VB.TextBox txtFPago 
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
            Index           =   8
            Left            =   2220
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Forma Pago vto"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   49
            Left            =   90
            TabIndex        =   29
            Top             =   240
            Width           =   1590
         End
         Begin VB.Image imgFP 
            Height          =   240
            Index           =   8
            Left            =   1920
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.ComboBox cboCompensaVto 
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
         Left            =   2370
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1440
         Width           =   4245
      End
      Begin VB.TextBox txtConcpto 
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
         Left            =   2340
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   4440
         Width           =   645
      End
      Begin VB.TextBox txtDescConcepto 
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
         Index           =   1
         Left            =   3030
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   4440
         Width           =   4785
      End
      Begin VB.CommandButton cmdContabCompensaciones 
         Caption         =   "&Aceptar"
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
         Left            =   5700
         TabIndex        =   14
         Top             =   5370
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
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
         Index           =   22
         Left            =   6780
         TabIndex        =   15
         Top             =   5370
         Width           =   975
      End
      Begin VB.TextBox txtDescConcepto 
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
         Index           =   0
         Left            =   3030
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   3960
         Width           =   4785
      End
      Begin VB.TextBox txtConcpto 
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
         Left            =   2340
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   3960
         Width           =   645
      End
      Begin VB.TextBox txtDescDiario 
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
         Index           =   0
         Left            =   3030
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   3240
         Width           =   4785
      End
      Begin VB.TextBox txtDiario 
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
         Left            =   2340
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   3240
         Width           =   645
      End
      Begin VB.TextBox txtCtaBanc 
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
         Left            =   2370
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2040
         Width           =   1275
      End
      Begin VB.TextBox txtDescBanc 
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
         Index           =   2
         Left            =   3660
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2040
         Width           =   4125
      End
      Begin VB.TextBox Text3 
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
         Index           =   23
         Left            =   2370
         TabIndex        =   6
         Top             =   840
         Width           =   1125
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   0
         Left            =   480
         Top             =   5370
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Compensa sobre Vto."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   47
         Left            =   210
         TabIndex        =   26
         Top             =   1440
         Width           =   2160
      End
      Begin VB.Label Label6 
         Caption         =   "Pagos"
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
         Index           =   21
         Left            =   960
         TabIndex        =   25
         Top             =   4440
         Width           =   705
      End
      Begin VB.Label Label6 
         Caption         =   "Cobros"
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
         Index           =   20
         Left            =   960
         TabIndex        =   24
         Top             =   3960
         Width           =   765
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   1
         Left            =   2040
         Picture         =   "frmTESListado.frx":000C
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Conceptos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   46
         Left            =   210
         TabIndex        =   22
         Top             =   3600
         Width           =   1050
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   0
         Left            =   2040
         Picture         =   "frmTESListado.frx":685E
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgDiario 
         Height          =   240
         Index           =   0
         Left            =   2040
         Picture         =   "frmTESListado.frx":D0B0
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Diario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   45
         Left            =   210
         TabIndex        =   20
         Top             =   3240
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta bancaria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   44
         Left            =   210
         TabIndex        =   18
         Top             =   2040
         Width           =   1635
      End
      Begin VB.Image imgCtaBanc 
         Height          =   240
         Index           =   2
         Left            =   2040
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   23
         Left            =   2040
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha contab."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   43
         Left            =   210
         TabIndex        =   16
         Top             =   840
         Width           =   1440
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Contabilización compensaciones"
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
         Height          =   405
         Index           =   12
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   5370
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameProgreso 
      Height          =   1935
      Left            =   3360
      TabIndex        =   0
      Top             =   2280
      Width           =   4095
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Max             =   1000
      End
      Begin VB.Label lbl2 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label lblPPAL 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame FrameDividVto 
      Height          =   2415
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   33
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdDivVto 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3000
         TabIndex        =   34
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   27
         Left            =   4200
         TabIndex        =   35
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "euros"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   62
         Left            =   3240
         TabIndex        =   37
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Datos vto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   57
         Left            =   240
         TabIndex        =   36
         Top             =   960
         Width           =   5040
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Dividir vencimiento "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   16
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   4890
      End
      Begin VB.Label Label4 
         Caption         =   "Datos vto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   56
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   5040
      End
   End
   Begin VB.Frame FrameOperAsegComunica 
      Height          =   5655
      Left            =   120
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Frame FrameFraPendOpAseg 
         Height          =   1455
         Left            =   120
         TabIndex        =   50
         Top             =   2520
         Width           =   4815
         Begin VB.CheckBox chkVarios 
            Caption         =   "Resumen"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   52
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkVarios 
            Caption         =   "Solo asegurados"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   51
            Top             =   720
            Value           =   1  'Checked
            Width           =   1815
         End
      End
      Begin VB.Frame FrameSelEmpre1 
         Height          =   3015
         Left            =   120
         TabIndex        =   47
         Top             =   1920
         Width           =   4815
         Begin MSComctlLib.ListView ListView3 
            Height          =   2055
            Left            =   840
            TabIndex        =   48
            Top             =   720
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   3625
            View            =   3
            Arrange         =   2
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
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   5292
            EndProperty
         End
         Begin VB.Label Label2 
            Caption         =   "Empresas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   41
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   825
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   0
            Left            =   960
            Picture         =   "frmTESListado.frx":13902
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgCheck 
            Height          =   240
            Index           =   1
            Left            =   1320
            Picture         =   "frmTESListado.frx":13A4C
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdOperAsegComunica 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2640
         TabIndex        =   46
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   35
         Left            =   3600
         TabIndex        =   44
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   34
         Left            =   1200
         TabIndex        =   43
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   39
         Left            =   3840
         TabIndex        =   39
         Top             =   5040
         Width           =   975
      End
      Begin VB.Image ImageAyudaImpcta 
         Height          =   240
         Index           =   1
         Left            =   120
         Top             =   5160
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   35
         Left            =   3360
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   39
         Left            =   2880
         TabIndex        =   45
         Top             =   1605
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha factura"
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
         Index           =   75
         Left            =   240
         TabIndex        =   42
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   34
         Left            =   960
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   195
         Index           =   38
         Left            =   360
         TabIndex        =   41
         Top             =   1605
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "XX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   22
         Left            =   2310
         TabIndex        =   40
         Top             =   480
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmTESListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SaltoLinea = """ + chr(13) + """

Public Opcion As Byte
    '1.- Cobros pendientes por cliente
    
    '3.- Reclamaciones por mail
    
    '4.- lISTADO agentes
    '5.- Departamentos
    
    '6.- Listado remesas
    
    '8.- Listado caja
    
    '9-  Devol remesas
    
    '10.- Listado formas de pago

    
    '11.- Transferencias PRovee   (o confirmings (domicilados o caixaconfirming)
    
    '12.- Listado previsional de gstos/pagos
    
    '13.- Transferencias ABONOS
    
    
    'Operaciones aseguradas
    '----------------------------
    '15.-  datos basicos
    '16.-  listado facturacion
    '17.-  Impagados asegurados
    
    
    '20.- Pregunta cuenta COBRO GENERICO
    '       La pongo aqui pq tengo implemntado todo
    
    
    '22.- Datos para la contabilizacion de las compensaciones
        
    '23.- Datos para la contbailiacion de la recpcion de documentos
    
    
    '24.-  Listado de documento(tal/pag) recibidos
    
    '25.-  Listado de pagos ordenados por banco  **** AHORA NO DEBERIA ENTRAR AQUI
    
    '26.-  Cancel remesa TAL/PAG.  Cando los importe no coinden. Solicitud cta y cc
    '27.-  Divide el vencimiento en dos vtos a partir del importe introducido en el text
        
        
    '30.-  Historico RECLAMACIONES
    '31.-   Gastos fijos
        
    '33.-  ASEGURADOS.  Listados avisos falta pago, avisos prorroga, aviso siniestro
        
    '34.-  Eliminar una recepcion de documentos, que ya ha sido contb con la puente
        
    '35.-  Gastos transferencias
        
    '36.-  Compensar ABONOS cobros
            
    '38.-  Recaudacion ejecutiva
        
    '39.-   Informe de comunicacion al seguro
    '40.-    Fras pendientes operaciones aseguradas
    
    '42.-   IMportar fichero norma 57 (recibos al cobro en ventanilla)
    
    '43.-   Confirmings
    '44.-   Caixaconfirming   igual que el de arriba
    
    
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
'--monica
'Private WithEvents frmB As frmBuscaGrid
Private WithEvents frmBa As frmBanco
Attribute frmBa.VB_VarHelpID = -1
Private WithEvents frmA As frmAgentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmP As frmFormaPago
Attribute frmP.VB_VarHelpID = -1
Private WithEvents frmD As frmBasico '--monica frmDepartamentos
Attribute frmD.VB_VarHelpID = -1
Private WithEvents frmS As frmBasico '--monica frmSerie
Attribute frmS.VB_VarHelpID = -1

Dim SQL As String
Dim RC As String
Dim RS As Recordset
Dim PrimeraVez As Boolean

Dim cad As String
Dim CONT As Long
Dim I As Integer
Dim TotalRegistros As Long

Dim Importe As Currency
Dim MostrarFrame As Boolean
Dim Fecha As Date

Dim DevfrmCCtas As String


Private Function ComprobarObjeto(ByRef T As TextBox) As Boolean
    Set miTag = New CTag
    ComprobarObjeto = False
    If miTag.Cargar(T) Then
        If miTag.Cargado Then
            If miTag.Comprobar(T) Then ComprobarObjeto = True
        End If
    End If

    Set miTag = Nothing
End Function










Private Sub cboCobro_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboCompensaVto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub






Private Sub chkCompensa_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkDesglosaGastoFijo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkEfectosContabilizados_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub chkFormaPago_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkLstTalPag_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkMarcarUtlRecla_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkNOremesar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub ChkObserva_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub chkPrevision_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 20 Or Index = 23 Or Index >= 26 Then
        CadenaDesdeOtroForm = "" 'Por si acaso. Tiene que devolve "" para que no haga nada
    End If
    Unload Me
End Sub


Private Function PonerTipoPagoCobro_(ParaSelect As Boolean, EsReclamacion As Boolean) As String
Dim I As Integer
Dim Sele As Integer
Dim Aux As String
Dim Visibles As Byte

    Aux = ""
    Sele = 0
    Visibles = 0
    If Not EsReclamacion Then
        For I = 0 To Me.chkTipPago.Count - 1
            If Me.chkTipPago(I).Visible Then
                Visibles = Visibles + 1
                If Me.chkTipPago(I).Value = 1 Then
                    Sele = Sele + 1
                    If ParaSelect Then
                        Aux = Aux & ", " & I
                    Else
                        Aux = Aux & "-" & Me.chkTipPago(I).Caption
                    End If
                End If
            End If
        Next
        'No ha marcado ninguno o los ha nmarcado todos. NO pongo nada
        If Sele = 0 Or Sele = Visibles Then Aux = ""
        
    Else
        '************************
        'Reclamaciones
        
        For I = 0 To Me.chkTipPagoRec.Count - 1
            If Me.chkTipPagoRec(I).Visible Then
                Visibles = Visibles + 1
                If Me.chkTipPagoRec(I).Value = 1 Then
                    Sele = Sele + 1
                    If ParaSelect Then
                        Aux = Aux & ", " & I
                    Else
                        Aux = Aux & "-" & Me.chkTipPagoRec(I).Caption
                    End If
                End If
            End If
        Next
        'No ha marcado ninguno o los ha nmarcado todos. NO pongo nada
        If Sele = 0 Or Sele = Visibles Then Aux = ""
    End If
    If Aux <> "" Then
        Aux = Mid(Aux, 2)
        Aux = "(" & Aux & ")"
    End If
    PonerTipoPagoCobro_ = Aux
End Function



Private Sub cmdCompensar_Click()
    
    cad = DevuelveDesdeBD("informe", "scryst", "codigo", 10) 'Orden de pago a bancos
    If cad = "" Then
        MsgBox "No esta configurada la aplicación. Falta informe(10)", vbCritical
        Exit Sub
    End If
    Me.Tag = cad
    
    cad = ""
    RC = ""
    CONT = 0
    TotalRegistros = 0
    NumRegElim = 0
    For I = 1 To Me.lwCompenCli.ListItems.Count
        If Me.lwCompenCli.ListItems(I).Checked Then
            If Trim(lwCompenCli.ListItems(I).SubItems(6)) = "" Then
                'Es un abono
                TotalRegistros = TotalRegistros + 1
            Else
                NumRegElim = NumRegElim + 1
            End If
        End If
        If Me.lwCompenCli.ListItems(I).Bold Then
            cad = cad & "A"
            If CONT = 0 Then CONT = I
            
            
            
        End If
    Next
    
    I = 0
    SQL = ""
    If Len(cad) <> 1 Then
        'Ha seleccionado o cero o mas de uno
        If txtimpNoEdit(0).Text <> txtimpNoEdit(1).Text Then
            'importes distintos. Solo puede seleccionar UNO
            SQL = "Debe selecionar uno(y solo uno) como vencimiento destino"
            
        End If
    Else
        'Comprobaremos si el selecionado esta tb checked
        If Not lwCompenCli.ListItems(CONT).Checked Then
            SQL = "El vencimiento seleccionado no esta marcado"
        
        Else
            'Si el importe Cobro es mayor que abono, deberia estar
            Importe = CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag)  'txtimpNoEdit(1).Tag es negativo
            If Importe <> 0 Then
                If Importe > 0 Then
                    'Es un abono
                    If Trim(lwCompenCli.ListItems(CONT).SubItems(6)) = "" Then SQL = "cobro"
                Else
                    If Trim(lwCompenCli.ListItems(CONT).SubItems(6)) <> "" Then SQL = "abono"
                End If
                If SQL <> "" Then SQL = "Debe marcar un " & SQL
            End If
            
        End If
    End If
    If TotalRegistros = 0 Or NumRegElim = 0 Then SQL = "Debe selecionar cobro(s) y abono(s)" & vbCrLf & SQL
        
    'Sep 2012
    'NO se pueden borrar las observaciones que ya estuvieran
    'RecuperaValor("text41csb|text42csb|text43csb|text51csb|text52csb|text53csb|text61csb|text62csb|text63csb|text71csb|text72csb|text73csb|text81csb|text82csb|text83csb|", J)
    If CONT > 0 Then
        'Hay seleccionado uno vto
        Set miRsAux = New ADODB.Recordset
        cad = "text41csb,text42csb,text43csb,text51csb,text52csb,text53csb,text61csb,text62csb,text63csb,text71csb,text72csb,text73csb,text81csb,text82csb,text83csb"
        cad = "Select " & cad & " FROM scobro WHERE numserie ='" & lwCompenCli.ListItems(CONT).Text & "' AND codfaccl="
        cad = cad & lwCompenCli.ListItems(CONT).SubItems(1) & " AND fecfaccl='" & Format(lwCompenCli.ListItems(CONT).SubItems(2), FormatoFecha)
        cad = cad & "' AND numorden = " & lwCompenCli.ListItems(CONT).SubItems(3)
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If miRsAux.EOF Then
            SQL = SQL & vbCrLf & " NO se ha encontrado el veto. destino"
        Else
            'Vamos a ver cuantos registros son
            CadenaDesdeOtroForm = ""
            RC = "0"
            For I = 0 To 14
                If DBLet(miRsAux.Fields(I), "T") = "" Then
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & miRsAux.Fields(I).Name & "|"
                    RC = Val(RC) + 1
                End If
            Next I
            
                
                
            'If TotalRegistros + NumRegElim > 15 Then SQL = SQL & vbCrLf & "No caben los textos de los vencimientos"
            If TotalRegistros + NumRegElim > Val(RC) Then SQL = SQL & vbCrLf & "No caben los textos de los vencimientos"
        End If
        miRsAux.Close
        Set miRsAux = Nothing
    End If
    
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    Else
        If MsgBox("Seguro que desea realizar la compensación?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    
    
    
    
    Me.FrameCompensaAbonosCliente.Enabled = False
    Me.Refresh
    Screen.MousePointer = vbHourglass
    
    RealizarCompensacionAbonosClientes
    Me.FrameCompensaAbonosCliente.Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub





Private Sub cmdContabCompensaciones_Click()

    'COmprobaciones y leches
    If Me.txtConcpto(0).Text = "" Or txtDiario(0).Text = "" Or Text3(23).Text = "" Or _
        Me.txtConcpto(1).Text = "" Then
        MsgBox "Todos los campos de contabilizacion  son obligatorios", vbExclamation
        Exit Sub
    End If

    If Me.cboCompensaVto.ListIndex = 0 Then
        If Me.txtCtaBanc(2).Text = "" Then
            MsgBox "Campo banco no puede estar vacio", vbExclamation
            Exit Sub
        End If
    Else
        If Me.txtFPago(8).Text <> "" Then
            RC = DevuelveDesdeBD("codforpa", "formapago", "codforpa", txtFPago(8).Text, "N")
            If RC = "" Then
                MsgBox "No existe la forma de pago", vbExclamation
                Exit Sub
            End If
        End If
    End If

    If FechaCorrecta2(CDate(Text3(23).Text), True) > 1 Then
        PonFoco Text3(23)
        Exit Sub
    End If

    If Me.cboCompensaVto.ListIndex = 0 Then
        'No compensa sobre ningun vencimiento.
        'No puede marcar la opcion del importe
        If chkCompensa.Value = 1 Then
            MsgBox "'Dejar sólo importe compensación' disponible cuando compense sobre un vencimiento", vbExclamation
            Exit Sub
        End If
    End If

    'Cargamos la cadena y cerramos
    CadenaDesdeOtroForm = Me.txtConcpto(0).Text & "|" & Me.txtConcpto(1).Text & "|" & txtDiario(0).Text & "|" & Text3(23).Text & "|" & Me.txtCtaBanc(2).Text & "|" & DevNombreSQL(txtDescBanc(2).Text) & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Me.txtFPago(8).Text & "|" & Me.cboCompensaVto.ItemData(Me.cboCompensaVto.ListIndex) & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & Me.chkCompensa.Value & "|"
    Unload Me
End Sub

Private Sub cmdContabilizarNorma57_Click()
    SQL = ""
    If Me.lwNorma57Importar(0).ListItems.Count = 0 Then SQL = SQL & "-Ningun vencimiento desde el fichero" & vbCrLf
    If Me.txtCtaBanc(5).Text = "" Then SQL = SQL & "-Cuenta bancaria" & vbCrLf
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    
    'La madre de las batallas
    'El sql que mando
    SQL = "(numserie ,codfaccl,fecfaccl,numorden ) IN (select ccost,pos,nomdocum,numdiari from tmpconext "
    SQL = SQL & " WHERE codusu =" & vUsu.Codigo & " and numasien=0 ) "
    'CUIDADO. El trozo 'from tmpconext  WHERE codusu' tiene que estar extamente ASI
    '  ya que en ver cobros, si encuentro esto, pong la fecha de vencimiento la del PAGO por
    ' ventanilla que devuelve el banco y contabilizamos en funcion de esa fecha
            
            
    cad = Format(Now, "dd/mm/yyyy") & "|" & Me.txtCtaBanc(5).Text & " - " & Me.txtDescBanc(5).Text & "|0|"  'efectivo
    With frmTESVerCobrosPagos
        .ImporteGastosTarjeta_ = 0
        .OrdenacionEfectos = 3
        .vSQL = SQL
        .OrdenarEfecto = True
        .Regresar = False
        .ContabTransfer = False
        .Cobros = True
        .Tipo = 0
        .SegundoParametro = ""
        'Los textos
        .vTextos = cad
        .CodmactaUnica = ""

        .Show vbModal
    End With

    
    'Borro haya cancelado o no
    LimpiarDelProceso
End Sub

Private Sub cmdDepto_Click()

    RC = ""
    cad = ""
    If txtCta(7).Text <> "" Then
        cad = " AND departamentos.codmacta >='" & txtCta(7).Text & "'"
        RC = "Desde " & txtCta(7).Text & " - " & DtxtCta(7).Text
    End If
    
    If txtCta(8).Text <> "" Then
        cad = cad & " AND departamentos.codmacta <='" & txtCta(8).Text & "'"
        RC = RC & "  hasta " & txtCta(8).Text & " - " & DtxtCta(8).Text
    End If


    
    SQL = "select departamentos.codmacta, nommacta,dpto,descripcion from departamentos,cuentas where cuentas.codmacta=departamentos.codmacta" & cad
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQL = "DELETE from Usuarios.zpendientes where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    CONT = 0
    SQL = "INSERT INTO Usuarios.zpendientes (codusu,  numorden,codforpa,  nomforpa, codmacta, nombre) VALUES (" & vUsu.Codigo & ","
    While Not RS.EOF
        CONT = CONT + 1
        cad = CONT & "," & RS!Dpto & ",'" & DevNombreSQL(RS!Descripcion) & "','" & RS!codmacta & "','" & DevNombreSQL(RS!Nommacta) & "')"
        cad = SQL & cad
        Conn.Execute cad
        
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    If CONT = 0 Then
        MsgBox "Ningún dato con esos valores", vbExclamation
        Exit Sub
    End If
    
    cad = "DesdeHasta= """ & Trim(RC) & """|"
    
    With frmImprimir
            .OtrosParametros = cad
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 10
            .Show vbModal
        End With
    
End Sub

Private Sub cmdDivVto_Click()
Dim Im As Currency

    'Dividira el vto en dos. En uno dejara el importe que solicita y en el otro el resto
    'Los gastos s quedarian en uno asi como el cobrado si diera lugar
    SQL = ""
    If txtImporte(1).Text = "" Then SQL = "Ponga el importe" & vbCrLf
    
    RC = RecuperaValor(CadenaDesdeOtroForm, 3)
    Importe = CCur(RC)
    Im = ImporteFormateado(txtImporte(1).Text)
    If Im = 0 Then
        SQL = "Importe no puede ser cero"
    Else
        If Importe > 0 Then
            'Vencimiento normal
            If Im > Importe Then SQL = "Importe superior al máximo permitido(" & Importe & ")"
            
        Else
            'ABONO
            If Im > 0 Then
                SQL = "Es un abono. Importes negativos"
            Else
                If Im < Importe Then SQL = "Importe superior al máximo permitido(" & Importe & ")"
            End If
        End If
        
    End If
    
    
    If SQL = "" Then
        Set RS = New ADODB.Recordset
        
        
        'CadenaDesdeOtroForm. Pipes
        '           1.- cadenaSQL numfac,numsere,fecfac
        '           2.- Numero vto
        '           3.- Importe maximo
        I = -1
        RC = "Select max(numorden) from scobro WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
        RS.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RS.EOF Then
            SQL = "Error. Vencimiento NO encontrado: " & CadenaDesdeOtroForm
        Else
            I = RS.Fields(0) + 1
        End If
        RS.Close
        Set RS = Nothing
        
    End If
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        PonFoco txtImporte(1)
        Exit Sub
        
    Else
        SQL = "¿Desea desdoblar el vencimiento con uno de : " & Im & " euros?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    'OK.  a desdoblar
    SQL = "INSERT INTO scobro (`numorden`,`gastos`,impvenci,`fecultco`,`impcobro`,`recedocu`,"
    SQL = SQL & "`tiporem`,`codrem`,`anyorem`,`siturem`,reftalonpag,"
    SQL = SQL & "`numserie`,`codfaccl`,`fecfaccl`,`codmacta`,`codforpa`,`fecvenci`,`ctabanc1`,`codbanco`,`codsucur`,`digcontr`,`cuentaba`,`ctabanc2`,`text33csb`,`text41csb`,`text42csb`,`text43csb`,`text51csb`,`text52csb`,`text53csb`,`text61csb`,`text62csb`,`text63csb`,`text71csb`,`text72csb`,`text73csb`,`text81csb`,`text82csb`,`text83csb`,`ultimareclamacion`,`agente`,`departamento`,`Devuelto`,`situacionjuri`,`noremesar`,`obs`,`nomclien`,`domclien`,`pobclien`,`cpclien`,`proclien`,iban) "
    'Valores
    SQL = SQL & " SELECT " & I & ",NULL," & TransformaComasPuntos(CStr(Im)) & ",NULL,NULL,0,"
    SQL = SQL & "NULL,NULL,NULL,NULL,NULL,"
    SQL = SQL & "`numserie`,`codfaccl`,`fecfaccl`,`codmacta`,`codforpa`,`fecvenci`,`ctabanc1`,`codbanco`,`codsucur`,`digcontr`,`cuentaba`,`ctabanc2`,`text33csb`,`text41csb`,`text42csb`,`text43csb`,`text51csb`,`text52csb`,`text53csb`,`text61csb`,`text62csb`,`text63csb`,`text71csb`,`text72csb`,`text73csb`,`text81csb`,`text82csb`,"
    'text83csb`,
    SQL = SQL & "'Div vto." & Format(Now, "dd/mm/yyyy hh:nn") & "'"
    SQL = SQL & ",`ultimareclamacion`,`agente`,`departamento`,`Devuelto`,`situacionjuri`,`noremesar`,`obs`,`nomclien`,`domclien`,`pobclien`,`cpclien`,`proclien`,iban FROM "
    SQL = SQL & " scobro WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
    SQL = SQL & " AND numorden = " & RecuperaValor(CadenaDesdeOtroForm, 2)
    Conn.BeginTrans
    
    'Hacemos
    CONT = 1
    If Ejecuta(SQL) Then
        'Hemos insertado. AHora updateamos el impvenci del que se queda
        If Im < 0 Then
            'Abonos
            SQL = "UPDATE scobro SET impvenci= impvenci + " & TransformaComasPuntos(CStr(Abs(Im)))
        Else
            'normal
            SQL = "UPDATE scobro SET impvenci= impvenci - " & TransformaComasPuntos(CStr(Im))
        End If
        
        SQL = SQL & " WHERE " & RecuperaValor(CadenaDesdeOtroForm, 1)
        SQL = SQL & " AND numorden = " & RecuperaValor(CadenaDesdeOtroForm, 2)
        If Ejecuta(SQL) Then CONT = 0 'TODO BIEN ******
    End If
    'Si mal, volvemos
    If CONT = 1 Then
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
        CadenaDesdeOtroForm = I
        Unload Me
    End If
    
    
End Sub

Private Sub cmdEfecDev_Click()
    'Listado de efectos devueltos
    SQL = ""
    RC = CampoABD(Text3(13), "F", "fechadev", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(14), "F", "fechadev", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    Set RS = New ADODB.Recordset
    
    RC = "SELECT count(*) from sefecdev where numorden>=0" & SQL
    RS.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If Not RS.EOF Then
        If DBLet(RS.Fields(0), "N") > 0 Then I = 1
    End If
    RS.Close
    Set RS = Nothing
    
    If I = 0 Then
        RC = "Ningun dato para mostrar"
        If SQL <> "" Then RC = RC & " con esos valores"
        MsgBox RC, vbExclamation
        Exit Sub
    End If
        
    Screen.MousePointer = vbHourglass
    If ListadoEfectosDevueltos(SQL) Then
        
        cad = DesdeHasta("F", 13, 14)
        If cad <> "" Then cad = "Fecha devolución: " & cad
        cad = "Desde= """ & Trim(cad) & """|"
        If Me.optImpago(0).Value Then
            I = 13
        Else
            I = 14
        End If
        
        With frmImprimir
            .OtrosParametros = cad
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = I
            .Show vbModal
        End With
    
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdFormaPago_Click()
    cad = ""
    RC = CampoABD(txtFPago(4), "N", "codforpa", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(txtFPago(5), "N", "codforpa", False)
    If RC <> "" Then cad = cad & " AND " & RC
    I = 0
    If cad <> "" Then
        I = 1
        'Forma pago
        SQL = ""
        RC = DesdeHasta("FP", 4, 5)
        SQL = "Cuenta= """ & Trim(RC) & """|"
    
    Else
        I = 0
        SQL = ""
    End If
    
        
    If ListadoFormaPago(cad) Then
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = I
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 27
            .Show vbModal
        End With
    End If
End Sub

Private Sub cmdGastosFijos_Click()
    If ListadoGastosFijos() Then
        With frmImprimir
            SQL = "Detalla= " & Abs(Me.chkDesglosaGastoFijo.Value) & "|DH= """ & cad & """|"
            
            
            .OtrosParametros = SQL
            .NumeroParametros = 2
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 89
            .Show vbModal
        End With
    End If
End Sub

Private Sub cmdGastosTransfer_Click()

       If Me.txtImporte(2).Text = "" Then
            CadenaDesdeOtroForm = 0
       Else
            CadenaDesdeOtroForm = Me.txtImporte(2).Text
       End If
       Unload Me
End Sub

Private Sub cmdListadoPagosBanco_Click()
    If ListadoOrdenPago Then
    
        'Orden de pagos. Habran dos. El que devuelve la funcion de abajo y
        'el acabado en F que irá ordenado por fecha dentro del grupo del banco
    
        CadenaDesdeOtroForm = DevuelveDesdeBD("informe", "scryst", "codigo", 7) 'Orden de pago a bancos
        
        If Me.chkPagBanco(0).Value = 1 Then
            If CadenaDesdeOtroForm = "" Then
                MsgBox "Falta registro 7 scryst", vbExclamation
                Exit Sub
            End If
            SQL = Mid(CadenaDesdeOtroForm, 1, Len(CadenaDesdeOtroForm) - 4)
            SQL = SQL & "F.rpt"
            RC = App.Path & "\InformesT\" & SQL
            If Dir(RC, vbArchive) = "" Then
                MsgBox "No existe el listado ordenado por fecha. Consulte soporte técnico" & vbCrLf & "El programa continuará", vbExclamation
            Else
                CadenaDesdeOtroForm = SQL
            End If
        End If
        With frmImprimir
            .NumeroParametros = 1
            .FormulaSeleccion = "{zlistadopagos.codusu}=" & vUsu.Codigo
            
            .SoloImprimir = False
            .Opcion = 62
            .Show vbModal
        End With
    End If
End Sub

Private Sub cmdListaRecpDocum_Click()
Dim NomFile As String


    'Si marca la opcion de imprimir el justifacante de recepcion, el desglose tiene que estar marcado
    If chkLstTalPag(2).Value = 1 Then
        chkLstTalPag(1).Value = 1
        NomFile = DevuelveNombreInformeSCRYST(8, "Confir. recepcion talón")
        If NomFile = "" Then Exit Sub  'El msgbox ya lo da la funcion
        
    Else
        NomFile = ""
    End If
    
    If GeneraDatosTalPag Then
        
        RC = "FechaIMP= " & Format(Now, "dd/mm/yyyy") & "|Cuenta= "
    
        SQL = DesdeHasta("F", 24, 25, "F. Recep")
        If (chkLstTalPag(0).Value = 1) Xor (chkLstTalPag(1).Value = 1) Then
            'Solo uno seleccionado
            cad = "Talón"
            If (chkLstTalPag(0).Value = 1) Then cad = "Pagaré"
            SQL = Trim(SQL & Space(15) & "F. pago: " & cad)
        End If
        
        
        cad = DesdeHasta("NF", 2, 3, "Id. ")
        If cad <> "" Then
            SQL = Trim(SQL & Space(15) & cad)
        End If
        
        
        
        If cboListPagare.ListIndex >= 1 Then
            If cboListPagare.ListIndex = 1 Then
                cad = "Llevadas a "
            Else
                cad = "Pendientes de llevar"
            End If
            cad = cad & " banco"
            SQL = Trim(SQL & Space(15) & cad)
        End If
        SQL = RC & """" & SQL & """|"
        

        
        CadenaDesdeOtroForm = NomFile   'Por si es el ersonalizable
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = 2
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            If chkLstTalPag(3).Value = 1 Then
                'Si esta marcado la confirmacion recepcion
                If chkLstTalPag(2).Value = 1 Then
                    .Opcion = 87
                Else
                    .Opcion = 61
                End If
            Else
                .Opcion = 63
            End If
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
    
    
End Sub

Private Sub cmdListRem_Click()
Dim B As Boolean
    '-------------------------------------
    'LISTADO REMESAS
    'Utilizaremos las tablas de informes
    ' ztesoreriacomun, ztmplibrodiario
    '------------------------------------

    'Comprobaciones iniciales
    If Me.chkTipoRemesa(0).Value = 0 And Me.chkTipoRemesa(1).Value = 0 And Me.chkTipoRemesa(2).Value = 0 Then
        MsgBox "Seleccione algún tipo de remesa", vbExclamation
        Exit Sub
    End If
    
    If Me.chkTipoRemesa(0).Value = 1 And Me.chkRem(1).Value = 1 Then
        MsgBox "Formato banco para talones / pagarés", vbExclamation
        Exit Sub
    End If
    
    If chkRem(1).Value = 1 Then
        If Me.chkRem(0).Value = 1 Then MsgBox "Listado banco NO detalla vencimientos", vbExclamation
        chkRem(0).Value = 0
    End If
    
    Screen.MousePointer = vbHourglass
    '------------------------------
    If Me.chkRem(1).Value Then
        'FORMATO BANCO
        B = ListadoRemesasBanco
    Else
        'El de siempre
        B = ListadoRemesas
    End If
    If B Then
        With frmImprimir
            RC = "0"
            If Me.chkRem(1).Value = 1 Then
                RC = "1"
                I = 88
            Else
                I = 11
                If Me.chkRem(0).Value = 1 Then RC = "1"
            End If
            .OtrosParametros = "Mostrar= " & RC & "|"
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = I
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdNoram57Fich_Click()

    If Me.lwNorma57Importar(0).ListItems.Count > 0 Or lwNorma57Importar(1).ListItems.Count > 0 Then
        SQL = "Ya hay un proceso . ¿ Desea importar otro archivo?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    Me.cmdContabilizarNorma57.Visible = False
    
    cd1.FileName = ""
    cd1.ShowOpen
    If cd1.FileName = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    LimpiarDelProceso
    Me.Refresh
   
    If procesarficheronorma57 Then
        
        'El fichero que ha entrado es correcto.
        'Ahora vamos a buscar los vencimientos
        If BuscarVtosNorma57 Then
            
            'AHORA cargamos los listviews
            CargaLWNorma57 True   'los correctos 'Si es que hay
            
            'Los errores
            CargaLWNorma57 False
    
    
    
            Me.cmdContabilizarNorma57.Visible = Me.lwNorma57Importar(0).ListItems.Count > 0
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOperAsegComunica_Click()
Dim B As Boolean


    'Fecha hasta
    SQL = ""
    If Text3(35).Text = "" Then SQL = SQL & "-Fecha hasta obligatoria" & vbCrLf

    If Opcion = 39 Then
            
            RC = ""
            For I = 1 To Me.ListView3.ListItems.Count
                If Me.ListView3.ListItems(I).Checked Then RC = RC & "1"
            Next
            If RC = "" Then SQL = SQL & "-Seleccione alguna empresa" & vbCrLf
            
            If SQL <> "" Then
                SQL = "Campos obligatorios: " & vbCrLf & vbCrLf & SQL
                MsgBox SQL, vbExclamation
                Exit Sub
            End If
    Else
    
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Comun para los dos
    SQL = "DELETE FROM Usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo
    Conn.Execute SQL

    
    If Opcion = 39 Then
        B = ComunicaDatosSeguro_
        I = 92
        CONT = 0
        SQL = ""
    Else
        B = GeneraDatosFrasAsegurados
        If Me.chkVarios(1).Value = 1 Then
            'Resumido
            I = 94
        Else
            I = 93
        End If
    End If
    If B Then
            SQL = ""
            CONT = 0
            RC = ""
            If Opcion <> 39 Then If Me.chkVarios(0).Value = 1 Then SQL = "SOLO asegurados"
                

            If Me.Text3(34).Text <> "" Then RC = RC & "desde " & Text3(34).Text
            If Me.Text3(35).Text <> "" Then RC = RC & "     hasta " & Text3(35).Text
            If RC <> "" Then
                RC = Trim(RC)
                RC = "Fechas : " & RC
                SQL = Trim(SQL & "       " & RC)
            End If
            
            SQL = "pDH= """ & SQL & """|"
            CONT = CONT + 1
            
            If Me.Opcion = 40 Then
                '   True: De factura ALZIRA
                '   False: vto      HERBELCA
                
                '//En el rpt DeFactura : Alzira es 1 (fra)    y herbelca es 0 (vto)
                RC = Abs(vParamT.FechaSeguroEsFra)
                SQL = SQL & "DeFactura= " & RC & "|"
                CONT = CONT + 1
            End If
    
            'Declaracion seguro
            'Cominicacion datos grupo
            If Opcion = 39 Then
                RC = ""
                For NumRegElim = 1 To Me.ListView3.ListItems.Count
                    If Me.ListView3.ListItems(NumRegElim).Checked Then
                        If RC <> "" Then RC = RC & SaltoLinea
                        RC = RC & Me.ListView3.ListItems(NumRegElim).Text
                    End If
                Next
        
                SQL = SQL & "Empresas= """ & RC & """|"
                CONT = CONT + 1
                
                
                RC = DevuelveDesdeBD("informe", "scryst", "codigo", 11) '
                If RC = "" Then
                    MsgBox "No esta configurada la aplicación. Falta informe(11)", vbCritical
                    Exit Sub
                End If
                CadenaDesdeOtroForm = RC
                
            End If
        
        
    End If
    Screen.MousePointer = vbDefault
        
    
    If B Then
        With frmImprimir
            .OtrosParametros = SQL
            .NumeroParametros = CONT
            .FormulaSeleccion = "{ztesoreriacomun.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = I
            .Show vbModal
        End With
    
    
        'rComunicaSeguro.rpt
    Else
        MsgBox "No se ha generado ningún dato", vbExclamation
    End If
    
End Sub

Private Sub cmdPagosprov_Click()
    'Hago las comprobaciones
    If Text3(5).Text = "" Then
        MsgBox "Fecha calculo no puede estar en blanco", vbExclamation
        Exit Sub
    End If
    
    
    
    
   'QUIEREN DETALLAR LAS CUENTAS
    CadenaDesdeOtroForm = ""
    If Me.cmbCuentas(1).ListIndex = 1 Then
        
        frmVarios.Opcion = 21
        CadenaDesdeOtroForm = Me.cmbCuentas(1).Tag
        frmVarios.Show vbModal
        If CadenaDesdeOtroForm = "" Then
            Me.cmbCuentas(1).ListIndex = 0
            Exit Sub
        Else
            
            Me.cmbCuentas(1).Tag = CadenaDesdeOtroForm
            GeneraComboCuentas
            Me.cmbCuentas(1).ListIndex = 2
        End If
    Else
        If Me.cmbCuentas(1).ListIndex = 2 Then CadenaDesdeOtroForm = Me.cmbCuentas(1).Tag
    End If
    
    
    
    Screen.MousePointer = vbHourglass
    If PagosPendienteProv(CadenaDesdeOtroForm) Then
        'Tesxto que iran
        SQL = "FECHA CALCULO: " & Text3(5).Text & "  "
        
        'Fechas
        cad = DesdeHasta("F", 3, 4)
        SQL = SQL & cad
        
        'Cuenta
        cad = DesdeHasta("C", 2, 3)
        If cad <> "" Then cad = SaltoLinea & Trim(cad)
        SQL = SQL & cad
        
        
        'Si lleva la cuentas seleccionadas una a una, las pondremos en el encabezado
        If Me.cmbCuentas(1).ListIndex = 2 Then
            If Me.cmbCuentas(1).Tag <> "" Then
                RC = Me.cmbCuentas(1).Tag
                cad = ""
                Do
                    I = InStr(1, RC, "|")
                    If I > 0 Then
                        If cad <> "" Then cad = cad & ","
                        cad = cad & "  " & Mid(RC, 1, I - 1)
                        RC = Mid(RC, I + 1)
                    End If
                Loop Until I = 0
                If cad <> "" Then
                    cad = SaltoLinea & "Cuentas: " & cad
                    SQL = SQL & cad
                End If
            End If
        End If
        
        
        
        
        
        
        'Desde hasta FP
        cad = DesdeHasta("FP", 6, 7)
        If cad <> "" Then cad = SaltoLinea & Trim(cad)
        SQL = SQL & cad
        
        
        'Formulas
        cad = "Cuenta= """ & SQL & """|"
        
        'Fecha imp
        cad = cad & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
        
        
        
        'Totaliza
        cad = cad & "Totalizar= " & Abs(chkProv.Value) & "|"
        'marzo 2014
        cad = cad & "EsPorTipo= " & Abs(Me.optMostraFP(0).Value) & "|"
        
        
        With frmImprimir
            .OtrosParametros = cad
            .NumeroParametros = 4
            .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
            .SoloImprimir = False
            If Me.optProv(0).Value Then
                If chkProv2.Value Then
                    .Opcion = 4
                Else
                    .Opcion = 6
                End If
            Else
                .Opcion = 5
            End If

            
            .Show vbModal
        End With

    
    End If
    Me.FrameProgreso.Visible = False
    Screen.MousePointer = vbDefault
    
    
    

End Sub



Private Sub cmdPrevisionGastosCobros_Click()


    'Borramos las lineas en usuarios
    lblPrevInd.Caption = "Preparando ..."
    lblPrevInd.Refresh
    Conn.Execute "DELETE FROM Usuarios.ztmpconext WHERE codusu =" & vUsu.Codigo
    Conn.Execute "DELETE FROM Usuarios.ztmpconextcab WHERE codusu =" & vUsu.Codigo
    Set miRsAux = New ADODB.Recordset



    'Hacemos el selecet
    SQL = "select cuentas.codmacta,nommacta from ctabancaria,cuentas where cuentas.codmacta=ctabancaria.codmacta"
    RC = CampoABD(txtCtaBanc(0), "T", "ctabancaria.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCtaBanc(1), "T", "ctabancaria.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    TotalRegistros = 0
    While Not RS.EOF
        '---
        If Not HacerPrevisionCuenta(RS!codmacta, RS!Nommacta) Then
        '---
            SQL = "DELETE FROM Usuarios.ztmpconextcab WHERE codusu =" & vUsu.Codigo
            SQL = SQL & " AND cta ='" & RS!codmacta & "'"
            Conn.Execute SQL
        Else
            TotalRegistros = TotalRegistros + 1
        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    lblPrevInd.Caption = ""
    Me.Refresh
    
    
    If TotalRegistros = 0 Then
        MsgBox "Ningun dato generado", vbExclamation
        Exit Sub
    End If
    
    If Me.optPrevision(0).Value Then
        SQL = "Fecha"
    Else
        SQL = "Tipo"
    End If
    'txtCtaBanc  txtDescBanc
    
    
    
    SQL = "Titulo= ""Informe tesorería (" & SQL & ")""|"
    'Fechas intervalor
    SQL = SQL & "Fechas= ""Fecha hasta " & Text3(18).Text & """|"
    'Cuentas
    RC = DesdeHasta("BANCO", 0, 1)
    SQL = SQL & "Cuenta= """ & RC & """|"
    SQL = SQL & "FechaIMP= """ & Format(Now, "dd/mm/yyyy") & """|"
    SQL = SQL & "NumPag= 0|"
    SQL = SQL & "Salto= 2|"

    'SQL = SQL & "MostrarAnterior= " & MostrarAnterior & "|"
    
    Screen.MousePointer = vbDefault
    With frmImprimir
        .OtrosParametros = SQL
        .NumeroParametros = 6
        .FormulaSeleccion = "{ado_lineas.codusu}=" & vUsu.Codigo
        '.SoloImprimir = True
        'Opcion dependera del combo
        .Opcion = 29
        .Show vbModal
    End With
    

    
    
    
    
End Sub

Private Function HacerPrevisionCuenta(Cta As String, Nommacta As String) As Boolean
Dim SaldoArrastrado As Currency
Dim Id As Currency
Dim IH As Currency


    HacerPrevisionCuenta = False
    
    lblPrevInd.Caption = Cta & " - " & Nommacta
    lblPrevInd.Refresh
    ' Las fechas son del periodo, luego me importa una mierda las fechas desde hasta
    '
    '
    CargaDatosConExt Cta, Now, Now, " 1 = 1", Nommacta
    
    Conn.Execute "insert into Usuarios.ztmpconextcab select * from tmpconextcab where codusu =" & vUsu.Codigo
    
    Conn.Execute "DELETE FROM tmpfaclin where codusu =" & vUsu.Codigo
    
    RC = "INSERT INTO tmpfaclin (codusu, IVA,codigo, Fecha, Cliente, cta,"
    RC = RC & " ImpIVA, Total) VALUES (" & vUsu.Codigo & ","
    
    'PARA CADA CUENTA
    'mETEREMOS TODOS LOS REGISTROS EN LA TABLA
    '
    '           TMPFACLIN
    '
    'TANTO COBROS COMO PAGOS I GASTOS
    '
    'Luego, en funcion del orden(TIPO o fecha) los iremos insertando en la tabla, para que
    'el saldo que va arrastrando sea el correcto
    
    
       
        
    CONT = 0
    
    
    '--------------------
    'DETALLAR COBROS
    lblPrevInd.Caption = Cta & " - Cobros"
    lblPrevInd.Refresh
    SQL = " WHERE fecvenci<='" & Format(Text3(18).Text, FormatoFecha) & "'"
    SQL = SQL & " AND ctabanc1 ='" & Cta & "'"
    If chkPrevision(0).Value = 0 Then
        SQL = "select sum(impvenci),sum(impcobro),fecvenci from scobro " & SQL
        SQL = SQL & " GROUP BY fecvenci"
        
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        While Not miRsAux.EOF
        
            Id = DBLet(miRsAux.Fields(0), "N")
            IH = DBLet(miRsAux.Fields(1), "N")
            Importe = Id - IH

            If Importe <> 0 Then
                CONT = CONT + 1
                cad = "'COBROS'," & CONT & ",'" & Format(miRsAux!FecVenci, FormatoFecha) & "','COBROS PENDIENTES',NULL,"
                'HAY COBROS
                If Importe < 0 Then
                    cad = cad & "NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
                Else
                    cad = cad & TransformaComasPuntos(CStr(Importe)) & ",NULL"
                End If
                cad = RC & cad & ")"
                Conn.Execute cad
                
            End If
                miRsAux.MoveNext
        Wend
        miRsAux.Close
                
    Else
         'DETALLAR PAGOS COBROS
            '(codusu, cta, ccost,Pos, fechaent, nomdocum, ampconce,"
            'timporteD,timporteH, saldo
            
        'SQL = "select scobro.*,nommacta from scobro,cuentas where scobro.codmacta=cuentas.codmacta"
        'SQL = SQL & " AND fecvenci<='2006-01-01'"
         
        SQL = "select scobro.* from scobro " & SQL
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            CONT = CONT + 1
            cad = "'COBROS'," & CONT & ",'" & Format(miRsAux!FecVenci, FormatoFecha) & "','"
            'NUmero factura
            cad = cad & miRsAux!NUmSerie & miRsAux!codfaccl & "/" & miRsAux!numorden & "',"
            
            cad = cad & "'" & miRsAux!codmacta & "',"
            Importe = miRsAux!ImpVenci - DBLet(miRsAux!impcobro, "N")
            If Importe <> 0 Then
                If Importe < 0 Then
                    cad = cad & "NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
                Else
                    cad = cad & TransformaComasPuntos(CStr(Importe)) & ",NULL"
                End If
                cad = cad & ")"
                cad = RC & cad
                Conn.Execute cad
            End If
            miRsAux.MoveNext
            
        Wend
        miRsAux.Close
        
    End If
    
    
    '--------------------
    '--------------------
    '--------------------
    'DETALLAR PAGOS
    '--------------------
    '--------------------
    lblPrevInd.Caption = Cta & " - pagos"
    lblPrevInd.Refresh
    SQL = " WHERE fecefect<='" & Format(Text3(18).Text, FormatoFecha) & "'"
    SQL = SQL & " AND ctabanc1 ='" & Cta & "'"
    
    If chkPrevision(1).Value = 0 Then
        SQL = "select sum(impefect),sum(imppagad),fecefect from spagop " & SQL & " GROUP BY fecefect"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Importe = 0
        While Not miRsAux.EOF

                Id = DBLet(miRsAux.Fields(0), "N")
                IH = DBLet(miRsAux.Fields(1), "N")
                Importe = Id - IH
            
                If Importe <> 0 Then
                    CONT = CONT + 1
                    cad = "'PAGOS'," & CONT & ",'" & Format(miRsAux!fecefect, FormatoFecha) & "','PAGOS PENDIENTES',NULL,"
                    'HAY COBROS
                    If Importe > 0 Then
                        cad = cad & "NULL," & TransformaComasPuntos(CStr(Importe))
                    Else
                        cad = cad & TransformaComasPuntos(CStr(Abs(Importe))) & ",NULL"
                    End If
                    cad = RC & cad & ")"
                    Conn.Execute cad
                End If
                miRsAux.MoveNext
        Wend
        miRsAux.Close
    Else
         'DETALLAR PAGOS COBROS
        'codusu, IVA,codigo, Fecha, Cliente, cta,"
        ' ImpIVA, Total) VALUES (" & vUsu.Codigo & ","
        
        'SQL = "select spagop.*,nommacta from spagop,cuentas where ctaprove=codmacta"
        'SQL = SQL & " AND fecefect<='2006-01-01'"
        
        SQL = "select spagop.* from spagop " & SQL
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            CONT = CONT + 1
            cad = "'PAGOS'," & CONT & ",'" & Format(miRsAux!fecefect, FormatoFecha) & "','"
            'NUmero factura
            cad = cad & DevNombreSQL(miRsAux!NumFactu) & "/" & miRsAux!numorden & "',"
            
            cad = cad & "'" & miRsAux!ctaprove & "',"
            Importe = miRsAux!ImpEfect - DBLet(miRsAux!imppagad, "N")
            If Importe <> 0 Then
                If Importe > 0 Then
                    cad = cad & "NULL," & TransformaComasPuntos(CStr(Importe))
                Else
                    cad = cad & TransformaComasPuntos(CStr(Abs(Importe))) & ",NULL"
                End If
                cad = cad & ")"
                cad = RC & cad
                Conn.Execute cad
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
    End If
    
    
    
    
    
    
    '--------------------
    '--------------------
    '--------------------
    'DETALLAR GASTOS GASTOS
    '--------------------
    '--------------------
    
    SQL = " from sgastfij,sgastfijd where sgastfij.codigo= sgastfijd.codigo"
    SQL = SQL & " and fecha >='" & Format(Now, FormatoFecha)
    SQL = SQL & "' AND fecha <='" & Format(Format(Text3(18).Text, FormatoFecha), FormatoFecha) & "'"
    SQL = SQL & " and ctaprevista='" & Cta & "'"
    
    'Desde 5 Abril 2006
    '------------------
    ' Si el gasto esta contbilizado desde la tesoreria, tiene la marca "contabilizado"
    SQL = SQL & " and contabilizado=0"
    
        ' ImpIVA, Total) VALUES (" & vUsu.Codigo & ","
        
        'SQL = "select spagop.*,nommacta from spagop,cuentas where ctaprove=codmacta"
        'SQL = SQL & " AND fecefect<='2006-01-01'"
     
     
    'ABro el recodset aqui.
    'Si es EOF entonces no necesito abrir la pantalla, puesto
    ' que no habran gastos para seleccionar
    'Si NO es EOF entonces abro el form y entonces alli(en frmvarios)
    'recorro el recodset
    SQL = " select sgastfij.codigo,descripcion,fecha,importe " & SQL
    
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If miRsAux.EOF Then
        miRsAux.Close
    Else
        NumRegElim = CONT
        CadenaDesdeOtroForm = "Gastos cuenta: " & Nommacta & "|" & Cta & "|" & Val(chkPrevision(2).Value) & "|"
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & RC & "|"
        frmVarios.Opcion = 18
        frmVarios.Show vbModal
        Set miRsAux = New ADODB.Recordset
        CONT = NumRegElim
        Me.Refresh
    End If
    
    
    If CONT = 0 Then Exit Function
    
    lblPrevInd.Caption = Cta & " - Informe"
    lblPrevInd.Refresh
    'Cargo INFORME
    '------------------------------------------------------------------------------------------
    'Leo el  saldo inicial
    RC = "Select * from tmpconextcab where codusu =" & vUsu.Codigo & " AND cta ='" & Cta & "'"
    miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SaldoArrastrado = 0
    If Not miRsAux.EOF Then SaldoArrastrado = DBLet(miRsAux!acumtotT, "N")
    miRsAux.Close
    
    'Si desgloso cobros, los detallo, si no hago el acumu
    RC = "INSERT INTO Usuarios.ztmpconext (codusu, cta, ccost,Pos, fechaent, nomdocum, ampconce,"
    RC = RC & "timporteD,timporteH, saldo) VALUES (" & vUsu.Codigo & ",'" & Cta & "','"
        
    
    
    'Ahora cogere todos los registros que estan cargados en tmpfaclin y los metere ya
    'en la tabla con los importes, ordenado como dice el option y
    'arrastrando saldo
    SQL = "select tmpfaclin.*,nommacta from tmpfaclin left join cuentas on cta=codmacta where codusu =" & vUsu.Codigo & " ORDER BY "
    'EL ORDEN
    If optPrevision(0).Value Then
        SQL = SQL & "fecha,cta"
    Else
        SQL = SQL & "cta,fecha"
    End If
    CONT = 1
    Id = 0
    IH = 0
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        cad = Mid(miRsAux!iva, 1, 4) & "'," & CONT & ",'" & Format(miRsAux!Fecha, FormatoFecha) & "','"
        
        
        
        If IsNull(miRsAux!Cta) Then
            'Stop
            cad = cad & "','" & DevNombreSQL(miRsAux!Cliente) & "'"
        Else
            cad = cad & Mid(DevNombreSQL(miRsAux!Cliente), 1, 10) & "',"
            If IsNull(miRsAux!Nommacta) Then
                cad = cad & "NULL"
            Else
                cad = cad & "'" & DevNombreSQL(miRsAux!Nommacta) & "'"
            End If
        End If
        If IsNull(miRsAux!Total) Then
            'VA AL DEBE
            Importe = miRsAux!ImpIva
            cad = cad & "," & TransformaComasPuntos(CStr(miRsAux!ImpIva)) & ",NULL,"
            Id = Id + Importe
        Else
            'HABER
            Importe = miRsAux!Total * -1
            cad = cad & ",NULL," & TransformaComasPuntos(CStr(miRsAux!Total)) & ","
            IH = IH + miRsAux!Total
        End If
        SaldoArrastrado = SaldoArrastrado + Importe
        cad = cad & TransformaComasPuntos(CStr(SaldoArrastrado)) & ")"
        cad = RC & cad
        Conn.Execute cad
        
        
        CONT = CONT + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Ajusto los importes de la tabla tmpconextcab
    SQL = "UPDATE Usuarios.ztmpconextcab SET acumantD=acumtotD,acumantH=acumtotH,acumantT=acumtotT"
    SQL = SQL & " where codusu =" & vUsu.Codigo & " AND cta ='" & Cta & "'"
    Conn.Execute SQL
    SQL = "UPDATE Usuarios.ztmpconextcab SET acumperD=" & TransformaComasPuntos(CStr(Id))
    SQL = SQL & ", acumperH=" & TransformaComasPuntos(CStr(IH))
    SQL = SQL & ", acumperT=" & TransformaComasPuntos(CStr(Id - IH))
    SQL = SQL & ", acumtott=" & TransformaComasPuntos(CStr(SaldoArrastrado))
    
    SQL = SQL & " where codusu =" & vUsu.Codigo & " AND cta ='" & Cta & "'"
    Conn.Execute SQL
    
    HacerPrevisionCuenta = True
    
End Function

Private Sub cmdRecaudaEjecutiva_Click()
'--monica
'
'    SQL = " scobro.codmacta=cuentas.codmacta AND"
'    SQL = SQL & " fecejecutiva is null and impvenci+coalesce(gastos)-coalesce(impcobro,0)>0"
'    'Si fechvto
'    RC = CampoABD(Text3(32), "F", "fecvenci", True)
'    If RC <> "" Then SQL = SQL & " AND " & RC
'    RC = CampoABD(Text3(33), "F", "fecvenci", False)
'    If RC <> "" Then SQL = SQL & " AND " & RC
'    'Codmacta
'    RC = CampoABD(txtCta(18), "T", "scobro.codmacta", True)
'    If RC <> "" Then SQL = SQL & " AND " & RC
'    RC = CampoABD(txtCta(18), "T", "scobro.codmacta", False)
'    If RC <> "" Then SQL = SQL & " AND " & RC
'
'
'
'    'hacemos un COUNT
'    RC = DevuelveDesdeBD("count(*)", "scobro,cuentas", SQL & " AND 1", "1")
'    If RC = "" Then RC = "0"
'    If Val(RC) = 0 Then
'        MsgBox "No existen registros", vbExclamation
'        Exit Sub
'    End If
'
'    SQL = " FROM scobro,cuentas WHERE " & SQL
'
'    frmTESVarios.NumeroDocumento = SQL
'    frmTESVarios.Opcion = 29
'    frmTESVarios.Show vbModal
    
End Sub

Private Sub cmdRecepDocu_Click()
    If txtDiario(1).Text = "" Or Me.txtConcpto(2).Text = "" Or txtConcpto(3).Text = "" Then
        MsgBox "Campos obligatorios", vbExclamation
        Exit Sub
    End If
    
    If Me.Label4(55).Visible Then
        If Me.txtCta(14).Text = "" Then
            MsgBox "Cuentas " & Label4(55).Caption & " requerida", vbExclamation
            Exit Sub
        End If
        SQL = ""
        If vParam.autocoste Then
            RC = Mid(txtCta(14).Text, 1, 1)
            If RC = 6 Or RC = 7 Then
                If txtCCost(0).Text = "" Then
                    MsgBox "Centro de coste requerido", vbExclamation
                    Exit Sub
                Else
                    SQL = txtCCost(0).Text
                End If
            End If
            
                
        End If
        txtCCost(0).Text = SQL
        
    Else
        txtCCost(0).Text = ""
        Me.txtCta(14).Text = ""
    End If
    
    
    
    
    I = 0
    If Me.chkAgruparCtaPuente(0).Visible Then
        If Me.chkAgruparCtaPuente(0).Value Then I = 1
    End If
    CadenaDesdeOtroForm = txtDiario(1).Text & "|" & Me.txtConcpto(2).Text & "|" & txtConcpto(3).Text & "|" & I & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & txtCta(14).Text & "|" & txtCCost(0).Text & "|"
    
    Unload Me
End Sub








Private Sub cmdReclamas_Click()
    
    
    Screen.MousePointer = vbHourglass
    '------------------------------
    If ListadoReclamas Then
        With frmImprimir
            cad = "Cadena= """ & cad & """|"
            .OtrosParametros = cad
            .NumeroParametros = 1
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            .Opcion = 86
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdTransfer_Click()
    
    
    Screen.MousePointer = vbHourglass
    '------------------------------
    If ListadoTransferencias Then
        With frmImprimir
            cad = "Mostrar= 1|tipot= """
            I = 28
            SQL = "Listado transferencias"
            If Opcion = 11 Then
                cad = cad & "(Pagos)"
            ElseIf Opcion = 13 Then
                cad = cad & "(Abonos)"
                If Me.chkCartaAbonos.Value Then
                    CadenaDesdeOtroForm = DevuelveNombreInformeSCRYST(12, "Carta abono0")
                    I = 95
                End If
            Else
                If Opcion = 44 Then
                    SQL = "Caixa confirming"
                Else
                    SQL = "Pagos domiciliados"
                End If
            End If
            
            cad = cad & """|ErTitulo= """ & SQL & """|"
            
            .OtrosParametros = cad
            .NumeroParametros = 3
            .FormulaSeleccion = "{ado.codusu}= " & vUsu.Codigo
            .SoloImprimir = False
            
            .Opcion = I
            .Show vbModal
        End With
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdVtoDestino_Click(Index As Integer)
    
    If Index = 0 Then
        TotalRegistros = 0
        If Not Me.lwCompenCli.SelectedItem Is Nothing Then TotalRegistros = Me.lwCompenCli.SelectedItem.Index
    
    
        For I = 1 To Me.lwCompenCli.ListItems.Count
            If Me.lwCompenCli.ListItems(I).Bold Then
                Me.lwCompenCli.ListItems(I).Bold = False
                Me.lwCompenCli.ListItems(I).ForeColor = vbBlack
                For CONT = 1 To Me.lwCompenCli.ColumnHeaders.Count - 1
                    Me.lwCompenCli.ListItems(I).ListSubItems(CONT).ForeColor = vbBlack
                    Me.lwCompenCli.ListItems(I).ListSubItems(CONT).Bold = False
                Next
            End If
        Next
        Me.Refresh
        
        If TotalRegistros > 0 Then
            I = TotalRegistros
            Me.lwCompenCli.ListItems(I).Bold = True
            Me.lwCompenCli.ListItems(I).ForeColor = vbRed
            For CONT = 1 To Me.lwCompenCli.ColumnHeaders.Count - 1
                Me.lwCompenCli.ListItems(I).ListSubItems(CONT).ForeColor = vbRed
                Me.lwCompenCli.ListItems(I).ListSubItems(CONT).Bold = True
            Next
        End If
        lwCompenCli.Refresh
        
        PonerFocoLw Me.lwCompenCli

    Else
    
        cad = DevuelveDesdeBD("informe", "scryst", "codigo", 10) 'Orden de pago a bancos
        If cad = "" Then
            MsgBox "No esta configurada la aplicación. Falta informe(10)", vbCritical
            Exit Sub
        End If
        CadenaDesdeOtroForm = cad
    
        LanzaBuscaGrid 1, 4


    End If
End Sub

Private Sub Command1_Click()
    If txtCta(13).Text = "" Then
        MsgBox "Ponga la cuenta", vbExclamation
        Exit Sub
    End If
    CadenaDesdeOtroForm = txtCta(13).Text
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion

        Case 22
            'Contabi efectos
            If CONT > 0 Then
                For I = 1 To Me.cboCompensaVto.ListCount
                    If Me.cboCompensaVto.ItemData(I) = CONT Then
                        CONT = I
                        Exit For
                    End If
                Next
            End If
            Me.cboCompensaVto.ListIndex = CONT
            PonFoco Text3(23)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub



    
Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim Img As Image


    Limpiar Me
    Me.Icon = frmPpal.Icon
    CargaImagenesAyudas Me.imgCtaBanc, 1, "Cuenta contable bancaria"
    CargaImagenesAyudas Image2, 2
    CargaImagenesAyudas Me.imgFP, 1, "Forma de pago"
    CargaImagenesAyudas Me.Image3, 1, "Cuenta contable"
    CargaImagenesAyudas Me.Imagente, 1, "Seleccionar agente"
    CargaImagenesAyudas imgCCoste, 1, "Centro de coste"
    CargaImagenesAyudas Me.ImageAyudaImpcta, 3
    For Each Img In Me.imgConcepto
        Img.ToolTipText = "Concepto"
    Next
    For Each Img In Me.imgDiario
        Img.ToolTipText = "Diario"
    Next
    
    
    
    For Each Img In Me.imgDpto
        Img.ToolTipText = "Departamento"
    Next
    
    
    'Limpiamos el tag
    txtCta(6).Tag = ""
    PrimeraVez = True
    FrCobrosPendientesCli.Visible = False
    frpagosPendientes.Visible = False
    FramereclaMail.Visible = False
    FrameAgentes.Visible = False
    FrameDpto.Visible = False
    FrameListRem.Visible = False
    FrameListadoCaja.Visible = False
    FrameDevEfec.Visible = False
    Me.FrameFormaPago.Visible = False
    FrameTransferencias.Visible = False
    Me.FramePrevision.Visible = False
    FrameAseg_Bas.Visible = False
    FrameCobroGenerico.Visible = False
    FrameCompensaciones.Visible = False
    FrameRecepcionDocumentos.Visible = False
    FrameListaRecep.Visible = False
    frameListadoPagosBanco.Visible = False
    FrameDividVto.Visible = False
    FrameReclama.Visible = False
    FrameGastosFijos.Visible = False
    FrameGastosTranasferencia.Visible = False
    FrameCompensaAbonosCliente.Visible = False
    FrameRecaudaEjec.Visible = False
    FrameOperAsegComunica.Visible = False
    FrameNorma57Importar.Visible = False
    CommitConexion
    
    Select Case Opcion
    Case 22
        
        
        For H = 0 To 1
            
            txtConcpto(H).Text = RecuperaValor(CadenaDesdeOtroForm, (H * 2) + 1)
            txtDescConcepto(H).Text = RecuperaValor(CadenaDesdeOtroForm, (H * 2) + 2)
        Next H
        Me.cboCompensaVto.Clear
        InsertaItemComboCompensaVto "No compensa sobre ningún vencimiento", 0
        
        'Veremos si puede sobre un Vto o no
        H = RecuperaValor(CadenaDesdeOtroForm, 5)
        CONT = 0
        If H = 1 Then CONT = RecuperaValor(CadenaDesdeOtroForm, 6)
        FrameCambioFPCompensa.Visible = CONT > 0
        'chkCompensaVto.Value = 0
        'chkCompensaVto.Enabled = h = 1
        'chkCompensaVto.Caption = RecuperaValor(CadenaDesdeOtroForm, 6)
        CadenaDesdeOtroForm = ""
        H = FrameCompensaciones.Height + 120
        W = FrameCompensaciones.Width
        FrameCompensaciones.Visible = True
        Caption = "Compensacion efectos"
        Text3(23).Text = Format(Now, "dd/mm/yyyy")
        
        
    Case 27
                'CadenaDesdeOtroForm. Pipes
        '           1.- cadenaSQL numfac,numsere,fecfac
        '           2.- Numero vto
        '           3.- Importe maximo
        H = FrameDividVto.Height + 120
        W = FrameDividVto.Width
        FrameDividVto.Visible = True
        
        
    End Select
    
    Me.Width = W + 300
    Me.Height = H + 400
    
    I = Opcion
    If Opcion = 13 Or I = 43 Or I = 44 Then I = 11
    
    'Aseguradas
    If Opcion >= 15 And Opcion <= 18 Then I = 15  'aseguradoas
    If Opcion = 33 Then I = 15 'aseguradoas
    If Opcion = 34 Then I = 23 'Eliminar recepcion documento
    If Opcion = 40 Then I = 39
    Me.cmdCancelar(I).Cancel = True
    
    PonerFrameProgreso

End Sub


Private Sub PonerFrameProgreso()
Dim I As Integer

    'Ponemos el frame al pricnipio de todo
    FrameProgreso.Visible = False
    FrameProgreso.ZOrder 0
    
    'lo ubicamos
    'Posicion horizintal WIDTH
    I = Me.Width - FrameProgreso.Width
    If I > 100 Then
        I = I \ 2
    Else
        I = 0
    End If
    FrameProgreso.Left = I
    'Posicion  VERTICAL HEIGHT
    I = Me.Height - FrameProgreso.Height
    If I > 100 Then
        I = I \ 2
    Else
        I = 0
    End If
    FrameProgreso.Top = I
End Sub





Private Sub Form_Unload(Cancel As Integer)
    If Opcion = 1 Then
        CheckValueGuardar "Listcta", CByte(Abs(Me.optCuenta(0).Value))
        CheckValueGuardar "Infapa", chkApaisado(0)
    End If
    If Opcion = 23 Then CheckValueGuardar "Agrup0", Me.chkAgruparCtaPuente(0)
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtAgente(RC).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescAgente(RC).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    DevfrmCCtas = CadenaDevuelta
End Sub

Private Sub frmBa_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CInt(RC)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub



Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 1)
    DtxtCta(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub







Private Sub frmD_DatoSeleccionado(CadenaSeleccion As String)
    'Si no habia cuenta
    
        txtCta(0).Text = RecuperaValor(CadenaSeleccion, 1)
        DtxtCta(0).Text = RecuperaValor(CadenaSeleccion, 2)
        txtCta(1).Text = RecuperaValor(CadenaSeleccion, 1)
        DtxtCta(1).Text = RecuperaValor(CadenaSeleccion, 2)
    
    'El dpto
    txtDpto(RC).Text = RecuperaValor(CadenaSeleccion, 3)
    txtDescDpto(RC).Text = RecuperaValor(CadenaSeleccion, 4)
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtFPago(RC).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescFPago(RC).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
    txtSerie(RC).Text = RecuperaValor(CadenaSeleccion, 1)
    
End Sub

Private Sub Image2_Click(Index As Integer)
    
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text3(Index).Text <> "" Then frmC.Fecha = CDate(Text3(Index).Text)
    RC = Index
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub Image3_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmCta = New frmColCtas
    RC = Index
    frmCta.DatosADevolverBusqueda = "0|1"
    frmCta.ConfigurarBalances = 3
    frmCta.Show vbModal
    Set frmCta = Nothing
    If Index = 17 Then PonerVtosCompensacionCliente
End Sub










Private Sub ImageAyudaImpcta_Click(Index As Integer)
Dim C As String
    Select Case Index
    Case 0
            C = "Compensaciones" & vbCrLf & String(60, "-") & vbCrLf
            C = C & "Cuando compense sobre un vencimiento al marcar la opción " & vbCrLf
            C = C & Space(10) & Me.chkCompensa.Caption & vbCrLf
            C = C & "se modificará el importe vencimiento poniendo el total a compensar  y en importe cobrado un cero"
    Case 1
            C = "Asegurados" & vbCrLf & String(60, "-") & vbCrLf
            C = C & "Fecha 'hasta' es campo obligado para considerar la fecha de baja de los asegurados." & vbCrLf
            C = C & "En los listados saldrán aquellos que si tienen fecha de baja , es superior al hasta solicitado "
            
            'ALZIRA
            C = C & vbCrLf & vbCrLf & "Comunicación datos seguro" & vbCrLf & "Salen TODAS las facturas entre el periodo seleccionado para los"
            C = C & " clientes asegurados"
            
    End Select
    MsgBox C, vbInformation

End Sub

Private Sub Imagente_Click(Index As Integer)
    Set frmA = New frmAgentes
    RC = Index
    frmA.DatosADevolverBusqueda = "0|1|"
    frmA.Show vbModal
    Set frmA = Nothing
End Sub

Private Sub ImageSe_Click(Index As Integer)
    RC = Index
'    Set frmS = New frmSerie
'    frmS.DatosADevolverBusqueda = "0"
'    frmS.Show vbModal
    Set frmS = New frmBasico
    AyudaContadores frmS, txtSerie(RC).Text, "tiporegi REGEXP '^[0-9]+$' = 0"
    Set frmS = Nothing


    Set frmS = Nothing
End Sub

Private Sub imgCarta_Click()
    Screen.MousePointer = vbHourglass
'--monica
'    Set frmB = New frmBuscaGrid
'    DevfrmCCtas = ""
'    frmB.vSQL = ""
'
'    '###A mano
'    frmB.vDevuelve = "0|1|"   'Siempre el 0
'
'    frmB.vSelElem = 1
'
'    cad = "Codigo|codcarta|N|15·"
'    cad = cad & "Descripcion|descarta|T|65·"
'    frmB.vCampos = cad
'    frmB.vTabla = "scartas"
'    frmB.vTitulo = "Cartas reclamación"
'    frmB.Show vbModal
'    Set frmB = Nothing
'    If DevfrmCCtas <> "" Then
'        Me.txtDescCarta.Text = RecuperaValor(DevfrmCCtas, 2)
'        txtCarta.Text = RecuperaValor(DevfrmCCtas, 1)
'    End If
End Sub

Private Sub ImgCCoste_Click(Index As Integer)
    LanzaBuscaGrid Index, 2
End Sub

Private Sub imgCheck_Click(Index As Integer)
    For I = 1 To Me.ListView3.ListItems.Count
        Me.ListView3.ListItems(I).Checked = (Index = 1)
    Next
        
End Sub

Private Sub imgConcepto_Click(Index As Integer)
    LanzaBuscaGrid Index, 1
End Sub

Private Sub imgCtaBanc_Click(Index As Integer)
    SQL = ""
    Set frmBa = New frmBanco
    frmBa.DatosADevolverBusqueda = "OK"
    frmBa.Show vbModal
    Set frmBa = Nothing
    If SQL <> "" Then
        txtCtaBanc(Index).Text = RecuperaValor(SQL, 1)
        Me.txtDescBanc(Index).Text = RecuperaValor(SQL, 2)
    End If
End Sub

Private Sub imgDiario_Click(Index As Integer)
    LanzaBuscaGrid Index, 0
End Sub

Private Sub imgDpto_Click(Index As Integer)
    SQL = "NO"
    If txtCta(1).Text <> "" And txtCta(0).Text <> "" Then
        
        If txtCta(1).Text <> txtCta(0).Text Then
            MsgBox "Debe seleccionar un mismo cliente", vbExclamation
            txtDpto(Index).Text = ""
            SQL = ""
        End If
    End If
    If SQL = "" Then Exit Sub
'--monica
'    Set frmD = New frmDepartamentos
'    RC = Index
'    frmD.DatosADevolverBusqueda = "1|2|"
'    frmD.vCuenta = txtCta(0).Text
'    frmD.Show vbModal
    
    RC = Index
    
    Set frmD = New frmBasico
    AyudaDepartamentos frmD, txtDpto(RC).Text, "codmacta = " & DBSet(txtCta(1).Text, "T")
    Set frmD = Nothing
    PonFoco txtDpto(RC)
    
    
    
    
    Set frmD = Nothing
End Sub


Private Sub imgFP_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    'Set frmCta = New frmColCtas
    Set frmP = New frmFormaPago
    RC = Index
    frmP.DatosADevolverBusqueda = "0|1"
    frmP.Show vbModal
    Set frmP = Nothing
End Sub






Private Sub imgGastoFijo_Click(Index As Integer)
     LanzaBuscaGrid Index, 3
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lwCompenCli_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim C As Currency
Dim Cobro As Boolean

    Cobro = True
    C = Item.Tag
    If Trim(Item.SubItems(6)) = "" Then
        'Es un abono
        Cobro = False
        C = -C
    
    End If
    
    'Si no es checkear cambiamos los signos
    If Not Item.Checked Then C = -C
    
    I = 0
    If Not Cobro Then I = 1
    
    Me.txtimpNoEdit(I).Tag = Me.txtimpNoEdit(I).Tag + C
    txtimpNoEdit(I).Text = Format(Abs(txtimpNoEdit(I).Tag))
    txtimpNoEdit(2).Text = Format(CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag), FormatoImporte)
            
End Sub

Private Sub optAsegAvisos_Click(Index As Integer)
    If Index = 0 Then
        Label4(39).Caption = "Fecha aviso falta pago"
    ElseIf Index = 1 Then
        Label4(39).Caption = "Fecha aviso prorroga"
    Else
        Label4(39).Caption = "Fecha aviso siniestro"
    End If
    
End Sub

Private Sub optAsegAvisos_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub optAsegBasic_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optCuenta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub optImpago_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optLCobros_Click(Index As Integer)
    Me.Check1.Enabled = Me.optLCobros(1).Value
    Me.Check2.Enabled = Not Me.Check1.Enabled
End Sub

Private Sub optLCobros_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optPrevision_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optProv_Click(Index As Integer)
    Me.chkProv.Enabled = Me.optProv(1).Value
    Me.chkProv2.Enabled = Not Me.chkProv.Enabled
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    PonFoco Text3(Index)
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index))
    If Text3(Index) = "" Then Exit Sub
    If Not EsFechaOK(Text3(Index)) Then
        MsgBox "Fecha incorrecta: " & Text3(Index), vbExclamation
        Text3(Index).Text = ""
        Text3(Index).SetFocus
    End If
End Sub
















Private Sub txtConcpto_GotFocus(Index As Integer)
     PonFoco txtConcpto(Index)
End Sub

Private Sub txtConcpto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtConcpto_LostFocus(Index As Integer)
    SQL = ""
    txtConcpto(Index).Text = Trim(txtConcpto(Index).Text)
    If txtConcpto(Index).Text <> "" Then
        
        If Not IsNumeric(txtConcpto(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            txtConcpto(Index).Text = ""
        Else
            txtConcpto(Index).Text = Val(txtConcpto(Index).Text)
            SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtConcpto(Index).Text, "N")
            If SQL = "" Then
                MsgBox "No existe el concepto: " & Me.txtConcpto(Index).Text, vbExclamation
                Me.txtConcpto(Index).Text = ""
            End If
        End If
        If txtConcpto(Index).Text = "" Then SubSetFocus txtConcpto(Index)
    End If
    Me.txtDescConcepto(Index).Text = SQL
    
End Sub

Private Sub txtDiario_GotFocus(Index As Integer)
    PonFoco txtDiario(Index)
End Sub

Private Sub txtDiario_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtDiario_LostFocus(Index As Integer)
    
    SQL = ""
    txtDiario(Index).Text = Trim(txtDiario(Index).Text)
    If txtDiario(Index).Text <> "" Then
        
        If Not IsNumeric(txtDiario(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            txtDiario(Index).Text = ""
            SubSetFocus txtDiario(Index)
        Else
            txtDiario(Index).Text = Val(txtDiario(Index).Text)
            SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtDiario(Index).Text, "N")
            
            If SQL = "" Then
                MsgBox "No existe el diario: " & Me.txtDiario(Index).Text, vbExclamation
                Me.txtDiario(Index).Text = ""
                PonFoco txtDiario(Index)
            End If
        End If
    End If
    Me.txtDescDiario(Index).Text = SQL
     
End Sub


Private Sub txtImporte_GotFocus(Index As Integer)
    ConseguirFoco txtImporte(Index), 3
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
Dim Mal As Boolean
    txtImporte(Index).Text = Trim(txtImporte(Index).Text)
    If txtImporte(Index).Text = "" Then Exit Sub
    Mal = False
    If Not EsNumerico(txtImporte(Index).Text) Then Mal = True

    If Not Mal Then Mal = Not CadenaCurrency(txtImporte(Index).Text, Importe)

    If Mal Then
        txtImporte(Index).Text = ""
        txtImporte(Index).SetFocus

    Else
        txtImporte(Index).Text = Format(Importe, FormatoImporte)
    End If
End Sub



Private Function ComprobarCuentas(Indice1 As Integer, Indice2 As Integer) As Boolean
Dim L1 As Integer
Dim L2 As Integer
    ComprobarCuentas = False
    If txtCta(Indice1).Text <> "" And txtCta(Indice2).Text <> "" Then
        L1 = Len(txtCta(Indice1).Text)
        L2 = Len(txtCta(Indice2).Text)
        If L1 > L2 Then
            L2 = L1
        Else
            L1 = L2
        End If
        If Val(Mid(txtCta(Indice1).Text & "000000000", 1, L1)) > Val(Mid(txtCta(Indice2).Text & "0000000000", 1, L1)) Then
            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
            Exit Function
        End If
    End If
    ComprobarCuentas = True
End Function


Private Function ComprobarFechas(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarFechas = False
    If Text3(Indice1).Text <> "" And Text3(Indice2).Text <> "" Then
        If CDate(Text3(Indice1).Text) > CDate(Text3(Indice2).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Function
        End If
    End If
    ComprobarFechas = True
End Function





Private Sub txtCtaBanc_GotFocus(Index As Integer)
    PonFoco txtCtaBanc(Index)
End Sub

Private Sub txtCtaBanc_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCtaBanc_LostFocus(Index As Integer)
    txtCtaBanc(Index).Text = Trim(txtCtaBanc(Index).Text)
    If txtCtaBanc(Index).Text = "" Then
        txtDescBanc(Index).Text = ""
        Exit Sub
    End If
    
    cad = txtCtaBanc(Index).Text
    I = CuentaCorrectaUltimoNivelSIN(cad, SQL)
    If I = 0 Then
        MsgBox "NO existe la cuenta: " & txtCtaBanc(Index).Text, vbExclamation
        SQL = ""
        cad = ""
    Else
        cad = DevuelveDesdeBD("codmacta", "bancos", "codmacta", cad, "T")
        If cad = "" Then
            MsgBox "Cuenta no asoaciada a ningun banco", vbExclamation
            SQL = ""
            I = 0
        End If
    End If
    
    txtCtaBanc(Index).Text = cad
    Me.txtDescBanc(Index).Text = SQL
    If I = 0 Then PonFoco txtCtaBanc(Index)
    
End Sub

Private Sub txtDias_GotFocus()
    PonFoco txtDias
End Sub

Private Sub txtDias_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtDias_LostFocus()
    txtDias.Text = Trim(txtDias.Text)
    If txtDias.Text <> "" Then
        If Not IsNumeric(txtDias.Text) Then
            MsgBox "Numero de dias debe ser numérico", vbExclamation
            txtDias.Text = ""
            SubSetFocus txtDias
        End If
    End If
End Sub



Private Sub txtDpto_GotFocus(Index As Integer)
    PonFoco txtDpto(Index)
End Sub

Private Sub txtDpto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtDpto_LostFocus(Index As Integer)
    
    'Pierde foco
    txtDpto(Index).Text = Trim(txtDpto(Index).Text)
    If txtDpto(Index).Text = "" Then
        Me.txtDescDpto(Index).Text = ""
        Exit Sub
    End If
    
    SQL = "NO"
    If txtCta(1).Text = "" Or txtCta(0).Text = "" Then
        MsgBox "Debe seleccionar un unico cliente", vbExclamation
        txtDpto(Index).Text = ""
        SQL = ""
    Else
        If txtCta(1).Text <> txtCta(0).Text Then
            MsgBox "Debe seleccionar un mismo cliente", vbExclamation
            txtDpto(Index).Text = ""
            SQL = ""
        End If
    End If
    
    If SQL <> "" Then
        SQL = ""
        If txtCta(1).Text <> "" Then
            If txtDpto(Index).Text <> "" Then
                If Not IsNumeric(txtDpto(Index).Text) Then
                      MsgBox "Codigo departamento debe ser numerico: " & txtDpto(Index).Text
                      txtDpto(Index).Text = ""
                Else
                      'Comproamos en la BD
                       Set RS = New ADODB.Recordset
                       cad = "Select descripcion from departamentos where codmacta='" & txtCta(0).Text
                       cad = cad & "' AND Dpto = " & txtDpto(Index).Text
                       RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                       If Not RS.EOF Then SQL = DBLet(RS.Fields(0), "T")
                       RS.Close
                       Set RS = Nothing
                End If
            End If
        Else
            If txtDpto(Index).Text <> "" Then
                MsgBox "Seleccione un cliente", vbExclamation
                txtDpto(Index).Text = ""
            End If
        End If
    End If
    Me.txtDescDpto(Index).Text = SQL
End Sub

Private Sub txtFPago_GotFocus(Index As Integer)
    PonFoco txtFPago(Index)
End Sub

Private Sub txtFPago_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtFPago_LostFocus(Index As Integer)
    If ComprobarCampoENlazado(txtFPago(Index), txtDescFPago(Index), "N") > 0 Then
        If txtFPago(Index).Text <> "" Then
            'Tiene valor.
            SQL = DevuelveDesdeBD("nomforpa", "formapago", "codforpa", txtFPago(Index).Text, "N")
            If SQL = "" Then SQL = "Codigo no encontrado"
            txtDescFPago(Index).Text = SQL
        Else
            'Era un error
            SubSetFocus txtFPago(Index)
        End If
    End If
End Sub




Private Sub SubSetFocus(Obje As Object)
    On Error Resume Next
    Obje.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


'Si tiene valor el campo fecha, entonces lo ponemos con el BD
Private Function CampoABD(ByRef T As TextBox, Tipo As String, CampoEnLaBD, Mayor_o_Igual As Boolean) As String

    CampoABD = ""
    If T.Text <> "" Then
        If Mayor_o_Igual Then
            CampoABD = " >= "
        Else
            CampoABD = " <= "
        End If
        Select Case Tipo
        Case "F"
            CampoABD = CampoEnLaBD & CampoABD & "'" & Format(T.Text, FormatoFecha) & "'"
        Case "T"
            CampoABD = CampoEnLaBD & CampoABD & "'" & T.Text & "'"
        Case "N"
            CampoABD = CampoEnLaBD & CampoABD & T.Text
        End Select
    End If
End Function



Private Function CampoBD_A_SQL(ByRef C As ADODB.Field, Tipo As String, Nulo As Boolean) As String

    If IsNull(C) Then
        If Nulo Then
            CampoBD_A_SQL = "NULL"
        Else
            If Tipo = "T" Then
                CampoBD_A_SQL = "''"
            Else
                CampoBD_A_SQL = "0"
            End If
        End If

    Else
    
        Select Case Tipo
        Case "F"
            CampoBD_A_SQL = "'" & Format(C.Value, FormatoFecha) & "'"
        Case "T"
            CampoBD_A_SQL = "'" & DevNombreSQL(C.Value) & "'"
        Case "N"
            CampoBD_A_SQL = TransformaComasPuntos(CStr(C.Value))
        End Select
    End If
End Function



Private Function DesdeHasta(Tipo As String, Desde As Integer, Hasta As Integer, Optional Des As String)
Dim C As String
    C = ""
    Select Case Tipo
    Case "F"
        'Son los text3(desde)....
        If Text3(Desde).Text <> "" Then C = C & "Desde " & Text3(Desde).Text
        
        If Text3(Hasta).Text <> "" Then
            If C <> "" Then C = C & "  "
            C = C & "Hasta " & Text3(Hasta).Text
        End If
        
    Case "C"
        'Cuenta
        If txtCta(Desde).Text <> "" Then C = "Desde " & txtCta(Desde).Text & "-" & DtxtCta(Desde).Text
        
        
        If txtCta(Hasta).Text <> "" Then
            If C <> "" Then C = C & "   "
            C = C & "Hasta " & txtCta(Hasta).Text & "-" & DtxtCta(Hasta).Text
        End If
        
        
    Case "FP"
        'FORMA DE PAGO
        If txtFPago(Desde).Text <> "" Then C = "Desde " & txtFPago(Desde).Text & "-" & txtDescFPago(Desde).Text
        
        
        If txtFPago(Hasta).Text <> "" Then
            If C <> "" Then C = C & "   "
            C = C & "Hasta " & txtFPago(Hasta).Text & "-" & txtDescFPago(Hasta).Text
        End If
    
    Case "BANCO"
        '    'txtCtaBanc  txtDescBanc
        If txtCtaBanc(Desde).Text <> "" Then C = "Desde " & txtCtaBanc(Desde).Text & "-" & txtDescBanc(Desde).Text
        
        If txtCtaBanc(Hasta).Text <> "" Then
            If C <> "" Then C = C & "   "
            C = C & "Hasta " & txtCtaBanc(Hasta).Text & "-" & txtDescBanc(Hasta).Text
        End If
    
    
    Case "S"
        'Serie
        If txtSerie(Desde).Text <> "" Then C = C & "Desde " & txtSerie(Desde).Text
        
        If txtSerie(Hasta).Text <> "" Then
            If C <> "" Then C = C & "  "
            C = C & "Hasta " & txtSerie(Hasta).Text
        End If
    
    Case "NF"
        'Numero factura
        If txtnumfac(Desde).Text <> "" Then C = C & "Desde " & txtnumfac(Desde).Text
        
        If txtnumfac(Hasta).Text <> "" Then
            If C <> "" Then C = C & "  "
            C = C & "Hasta " & txtnumfac(Hasta).Text
        End If
    
    Case "GF"
        'Gasto fijo
        
        If txtGastoFijo(Desde).Text <> "" Then C = C & "Desde " & txtGastoFijo(Desde).Text & " - " & Me.txtDescGastoFijo(Desde).Text
        
        If txtGastoFijo(Hasta).Text <> "" Then
            If C <> "" Then C = C & "  "
            C = C & "Hasta " & txtGastoFijo(Hasta).Text & " - " & Me.txtDescGastoFijo(Hasta).Text
        End If
    
    
    End Select
    If C <> "" Then C = "  " & Des & " " & C
    DesdeHasta = C
End Function


Private Sub PonerFrameProgressVisible(Optional TEXTO As String)
        If TEXTO = "" Then TEXTO = "Generando datos"
        Me.lblPPAL.Caption = TEXTO
        Me.lbl2.Caption = ""
        Me.ProgressBar1.Value = 0
        Me.FrameProgreso.Visible = True
        Me.Refresh
End Sub





'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------

Private Function CobrosPendientesCliente(ByVal ListadoCuentas As String) As Boolean
Dim TieneRemesa As Boolean
Dim RemesaTalones As Boolean
Dim RemesaPagares As Boolean
Dim RemesaEfectos As Boolean
Dim SePuedeRemesar As Boolean
Dim InsertarLinea As Boolean


Dim CadenaInsert As String

    On Error GoTo ECobrosPendientesCliente
    CobrosPendientesCliente = False

    
    'De parametros contapagarepte contatalonpte
    cad = DevuelveDesdeBD("contatalonpte", "paramtesor", "codigo", "1")
    If cad = "" Then cad = "0"
    RemesaTalones = (Val(cad) = 1)
    
    cad = DevuelveDesdeBD("contapagarepte", "paramtesor", "codigo", "1")
    If cad = "" Then cad = "0"
    RemesaPagares = (Val(cad) = 1)
    
    cad = DevuelveDesdeBD("contaefecpte", "paramtesor", "codigo", "1")
    If cad = "" Then cad = "0"
    RemesaEfectos = (Val(cad) = 1)
    

    
    
    
    
    'Trozo basico
    cad = " FROM scobro ,cuentas,sforpa ,stipoformapago"
    cad = cad & " WHERE  scobro.codmacta = cuentas.codmacta"
    cad = cad & " AND scobro.codforpa = sforpa.codforpa"
    cad = cad & " AND sforpa.tipforpa = stipoformapago.tipoformapago"

    
    
    'Fecha factura
    RC = CampoABD(Text3(1), "F", "fecfaccl", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(Text3(2), "F", "fecfaccl", False)
    If RC <> "" Then cad = cad & " AND " & RC



    'Se me habia olvidado
    'A G E N T E
    RC = CampoABD(txtAgente(0), "N", "agente", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(txtAgente(1), "N", "agente", False)
    If RC <> "" Then cad = cad & " AND " & RC




    'Fecha vencimiento
    RC = CampoABD(Text3(19), "F", "fecvenci", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(Text3(20), "F", "fecvenci", False)
    If RC <> "" Then cad = cad & " AND " & RC

    'SERIE
    RC = CampoABD(txtSerie(0), "T", "numserie", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(txtSerie(1), "T", "numserie", False)
    If RC <> "" Then cad = cad & " AND " & RC
    
    'Numero factura
    RC = CampoABD(txtnumfac(0), "T", "codfaccl", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(txtnumfac(1), "T", "codfaccl", False)
    If RC <> "" Then cad = cad & " AND " & RC
    


    'Cliente
    RC = CampoABD(txtCta(1), "T", "scobro.codmacta", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(txtCta(0), "T", "scobro.codmacta", False)
    If RC <> "" Then cad = cad & " AND " & RC
    
    'Forma PAGO
    RC = CampoABD(txtFPago(0), "T", "scobro.codforpa", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(txtFPago(1), "T", "scobro.codforpa", False)
    If RC <> "" Then cad = cad & " AND " & RC
    
    
    'Cliente con departamento
    'If txtCta(0).Text <> "" Then
    '    If cad <> "" Then cad = cad & " AND "
    '    cad = cad & " scobro.codmacta = '" & txtCta(6).Text & "'"
    'End If
    
    'Departamento
    RC = CampoABD(Me.txtDpto(0), "N", "departamento", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(Me.txtDpto(1), "N", "departamento", False)
    If RC <> "" Then cad = cad & " AND " & RC
    
    
    'Solo los NO remesar
    If Me.chkNOremesar.Value = 1 Then
        cad = cad & " AND noremesar = 1 "
    End If
    
    'Docuemtno recibido y devuelto. Los combos  recedocu  Devuelto
    If Me.cboCobro(0).ListIndex > 0 Then cad = cad & " AND recedocu = " & cboCobro(0).ItemData(cboCobro(0).ListIndex)
    If Me.cboCobro(1).ListIndex > 0 Then cad = cad & " AND Devuelto = " & cboCobro(1).ItemData(cboCobro(1).ListIndex)
    
    
    'Y lista de cuentas

    If ListadoCuentas <> "" Then
        NumRegElim = 1
        SQL = ""
        Do
            TotalRegistros = InStr(NumRegElim, ListadoCuentas, "|")
            If TotalRegistros > 0 Then
                If NumRegElim > 1 Then SQL = SQL & ","
                NumRegElim = 2
                SQL = SQL & "'" & Mid(ListadoCuentas, 1, TotalRegistros - 1) & "'"
                ListadoCuentas = Mid(ListadoCuentas, TotalRegistros + 1)
            End If
           
            
        Loop Until TotalRegistros = 0
        NumRegElim = 0
        cad = cad & " AND scobro.codmacta IN (" & SQL & ")"
    End If
    
    
    
    'Si ha marcado alguna forma de pago
    RC = PonerTipoPagoCobro_(True, False)
    If RC <> "" Then cad = cad & " AND tipoformapago IN " & RC
    RC = ""
    
    'Contador
    SQL = "Select count(*) "
    SQL = SQL & cad
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalRegistros = 0
    If Not RS.EOF Then
        'Total registros
        TotalRegistros = RS.Fields(0)
    End If
    RS.Close
    
    If TotalRegistros = 0 Then
        'NO hay registros
        MsgBox "Ningun dato con esos valores", vbExclamation
        Exit Function
    End If
    
    'Llegados aqui, si k hay registros.
    '100 POR EJEMPLO
    If TotalRegistros > 100 Then
        'Ponemos visible el frame
        MostrarFrame = True
        PonerFrameProgressVisible
    Else
        MostrarFrame = False 'NO lo mostramos
    End If
    
    
    
    'Borramos tmp
    Conn.Execute "DELETE FROM Usuarios.zpendientes where codusu = " & vUsu.Codigo
    
    'Marzo 2015
    'Si agrupamos por forma de pago, agruparemos antes por Tipo de pago
    If chkFormaPago.Value = 1 Then Conn.Execute "DELETE FROM Usuarios.zsimulainm where codusu = " & vUsu.Codigo
    
    
    
    
    
    SQL = "SELECT scobro.*, cuentas.nommacta, nifdatos,stipoformapago.descformapago ,stipoformapago.tipoformapago,nomforpa " & cad
    ' ----------------
    '  20 Diciembre 2005
    '  Remesados o no remesados
    '
    CONT = 1
    If Me.ChkAgruparSituacion.Value = 1 Then
        '
        CONT = 0
    Else
        If Me.chkEfectosContabilizados.Value = 1 Then CONT = 0
    End If
    If CONT = 1 Then SQL = SQL & " AND codrem is null"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    TieneRemesa = False
    SQL = "INSERT INTO Usuarios.zpendientes (codusu, serie_cta, factura, fecha, numorden,"
    SQL = SQL & "codforpa, nomforpa, codmacta,nombre, fecVto, importe, pag_cob,gastos,vencido,Situacion"
    'Nuevo Enero 2009
    'Si esta apaisado ponemos los departamentos
    If Me.chkApaisado(0).Value = 1 Then
        SQL = SQL & ",coddirec,nomdirec"
    Else
        'Metemos el NIF para futors listados. Pej. El de cobors por cliente lo pondra
        SQL = SQL & ",nomdirec"
    End If
    SQL = SQL & ",devuelto,recibido"
    'SQL = SQL & ",observa) VALUES (" & vUsu.Codigo & ",'"
    'Dic 2013 . Acelerar proceso
    SQL = SQL & ",observa) VALUES "
    
    
    CadenaInsert = "" 'acelerar carga datos
    Fecha = CDate(Text3(0).Text)
    While Not RS.EOF
        CONT = CONT + 1
        If MostrarFrame Then
            lbl2.Caption = "Registro: " & CONT
            lbl2.Refresh
        End If
        
        
        'If Rs!codmacta = "4300019" Then Stop
        
        cad = RS!NUmSerie & "','" & Format(RS!codfaccl, "0000") & "','" & Format(RS!fecfaccl, FormatoFecha) & "'," & RS!numorden
        
        'Modificacion. Enero 2010. Tiene k aparacer la forma de pago, no el tipo
        'Cad = Cad & "," & Rs!codforpa & ",'" & DevNombreSQL(Rs!descformapago) & "','"
        cad = cad & "," & RS!codforpa & ",'" & DevNombreSQL(RS!nomforpa) & "','"
        
        cad = cad & RS!codmacta & "','" & DevNombreSQL(RS!Nommacta) & "','"
        cad = cad & Format(RS!FecVenci, FormatoFecha) & "',"
        cad = cad & TransformaComasPuntos(CStr(RS!ImpVenci)) & ","
        'Cobrado, si no es nulo
        If Not IsNull(RS!impcobro) Then
            cad = cad & TransformaComasPuntos(CStr(RS!impcobro))
        Else
            cad = cad & "0"
        End If
        
        'Gastos
        cad = cad & "," & TransformaComasPuntos(DBLet(RS!Gastos, "N"))
        
        If Fecha > RS!FecVenci Then
            cad = cad & ",1"
        Else
            cad = cad & ",0"
        End If

        'Hay que añadir la situacion. Bien sea juridica....
        ' Si NO agrupa por situacion, en ese campo metere la referencia del cobro (rs!referencia)
         'vbTalon = 2 vbPagare = 3
        InsertarLinea = True
        
        If Me.ChkAgruparSituacion.Value = 0 Then
            cad = cad & ",'" & DevNombreSQL(DBLet(RS!referencia, "T")) & "'"
            
            'Si no agrupa por situacion de vto y no tiene el riesgo marcado
            'Talones pagares
            'Si se ha recepcionado documento, NO deben salir
            
            
            'Enero 2010
            'Comentamos esto ya que tiene la marca de recibidos si/no
'            If Me.chkEfectosContabilizados.Value = 0 Then
'                If Val(Rs!tipoformapago) = vbTalon Or Val(Rs!tipoformapago) = vbPagare Then
'                    If DBLet(Rs!recedocu, "N") = 1 Then InsertarLinea = False
'                End If
'            End If

            
        Else
            'La situacion.
            'Si esta en situacion juridica.
            ' Si no, si esta devuelto y no es una remesa
            ' y luego si es una remesa, sitaucion o devuelto
            If RS!situacionjuri = 1 Then
                cad = cad & ",'SITUACION JURIDICA'"
            Else
                'Cambio Marzo 2009
                ' Ahora tb se remesan los pagares y talones
                
                If Not IsNull(RS!siturem) Then
                    TieneRemesa = True
                    cad = cad & ",'R" & Format(RS!AnyoRem, "0000") & Format(RS!CodRem, "0000000000") & "'"
                    
                Else
                    
                    If RS!Devuelto = 1 Then
                        cad = cad & ",'DEVUELTO'"
                    Else
                            
                        SePuedeRemesar = False
                        If RemesaEfectos Then SePuedeRemesar = RS!tipoformapago = vbTipoPagoRemesa
                        If RemesaPagares Then SePuedeRemesar = RS!tipoformapago = vbPagare
                        If RemesaTalones Then SePuedeRemesar = RS!tipoformapago = vbTalon
                        
                    
                        If Not SePuedeRemesar Then
                            cad = cad & ",'PENDIENTE COBRO'"
                        Else
                            cad = cad & ",'PENDIENTE REMESAR'" '& Rs!anyorem
                        End If
                        
                        
                        
                        'Talones pagares
                        'Si se ha recepcionado documento, NO deben salir
                        'ENERO 2010
                        'Tiene la marca de recibidos
                        
                        'If Val(Rs!tipoformapago) = vbTalon Or Val(Rs!tipoformapago) = vbPagare Then
                        '    If Me.chkEfectosContabilizados.Value = 0 Then
                        '        If DBLet(Rs!recedocu, "N") = 1 Then InsertarLinea = False
                        '    End If
                        'End If
                        
                    
                    End If  'De devuelto
               End If  'de SITUREM null
            End If ' de situacion juridica
        End If  'de agrupa por sitacuib
        cad = cad & ","
        If Me.chkApaisado(0).Value = 1 Then
            'SI carga departamentos. Esto podriamos mejorar la velocidad si
            'pregarmos un rs o en la select linkamos con departamento...
            If IsNull(RS!departamento) Then
                cad = cad & "NULL,NULL,"
            Else
                cad = cad & "'" & RS!departamento & "','"
                cad = cad & DevNombreSQL(DevuelveDesdeBD("Descripcion", "departamentos", "codmacta = '" & RS!codmacta & "' AND dpto", RS!departamento, "N")) & "',"
            End If
            
        Else
            'Nif datos
            'Stop
             cad = cad & "'" & DevNombreSQL(DBLet(RS!nifdatos, "T")) & "',"
        End If
        
        If DBLet(RS!Devuelto, "N") = 0 Then
            cad = cad & "'',"
        Else
            cad = cad & "'S',"
        End If
        If DBLet(RS!recedocu, "N") = 0 Then
            cad = cad & "''"
        Else
            cad = cad & "'S'"
        End If
            
        cad = cad & ",'"
        If Me.ChkObserva.Value Then
            cad = cad & DevNombreSQL(DBLet(RS!Obs, "T"))
'        Else
'            Cad = Cad & "''"
        End If
        cad = cad & "')"
        
        If InsertarLinea Then
        
            CadenaInsert = CadenaInsert & ", (" & vUsu.Codigo & ",'" & cad
        
            If Len(CadenaInsert) > 20000 Then
                cad = SQL & Mid(CadenaInsert, 2)
                Conn.Execute cad
                CadenaInsert = ""
            End If
            'Cad = SQL & Cad
            'Conn.Execute Cad
        Else
            'Tiramos para atras el contador total
            CONT = CONT - 1
        End If
        'Siguiente
        RS.MoveNext
    Wend
    RS.Close
    
    If Len(CadenaInsert) > 0 Then
        cad = SQL & Mid(CadenaInsert, 2)
        Conn.Execute cad
        CadenaInsert = ""
    End If

    
    'Si esta seleccacona SITIACUIN VENCIMIENTO
    ' y tenia remesas , entonces updateo la tabla poniendo
    ' la situacion de la remesa
    If TieneRemesa Then
        cad = "Select codigo,anyo,  descsituacion"
        cad = cad & " from remesas left join tiposituacionrem on situacion=situacio"
        RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            Debug.Print RS!Codigo
            If Not IsNull(RS!descsituacion) Then
                cad = "R" & Format(RS!Anyo, "0000") & Format(RS!Codigo, "0000000000")
                cad = " WHERE situacion='" & cad & "'"
                cad = "UPDATE Usuarios.zpendientes set Situacion='Remesados: " & RS!descsituacion & "' " & cad
                Conn.Execute cad
            End If
            RS.MoveNext
        Wend
        RS.Close
    End If
    
    'Marzo 2015.
    'Nivel de anidacion para los agrupados por forma de pago
    ' que es TIPO DE PAGO
    If chkFormaPago.Value = 1 Then
    
        cad = "select codforpa from Usuarios.zpendientes where codusu =" & vUsu.Codigo & " group by 1"
        RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cad = ""
        While Not RS.EOF
            cad = cad & ", " & RS!codforpa
            RS.MoveNext
        Wend
        RS.Close
        
        If cad <> "" Then
            cad = Mid(cad, 2)
            cad = " and codforpa IN (" & cad & ")"
            cad = " FROM sforpa , stipoformapago WHERE sforpa.tipforpa=stipoformapago.tipoformapago" & cad
            cad = "SELECT " & vUsu.Codigo & ",codforpa,tipoformapago,descformapago " & cad
            cad = "INSERT INTO Usuarios.zsimulainm(codusu,codigo,conconam,nomconam) " & cad
        
            Conn.Execute cad
        End If
    End If
    
    
    
    'Voy a comprobar si ha metido algun registo
    espera 0.3
    SQL = "Select count(*) FROM  Usuarios.zpendientes where codusu = " & vUsu.Codigo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    If Not RS.EOF Then CONT = DBLet(RS.Fields(0), "N")
    RS.Close
    If CONT = 0 Then
        MsgBox "No se ha generado ningun valor(2)", vbExclamation
    Else
        CobrosPendientesCliente = True 'Para imprimir
    End If
    Exit Function
ECobrosPendientesCliente:
    MuestraError Err.Number, Err.Description
End Function



Private Function PagosPendienteProv(ByVal ListadoCuentas As String) As Boolean
'Dim MismaClavePrimaria As String


    On Error GoTo EPagosPendienteProv
    PagosPendienteProv = False
    
    'Trozo basico de PAGOS
    cad = "FROM spagop ,cuentas ,sforpa,stipoformapago"
    cad = cad & " WHERE spagop.ctaprove = cuentas.codmacta"
    cad = cad & " AND spagop.codforpa = sforpa.codforpa"
    cad = cad & " AND sforpa.tipforpa = stipoformapago.tipoformapago"

    
    'Fecha
    RC = CampoABD(Text3(3), "F", "fecefect", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(Text3(4), "F", "fecefect", False)
    If RC <> "" Then cad = cad & " AND " & RC

    'Cliente
    RC = CampoABD(txtCta(2), "T", "ctaprove", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(txtCta(3), "T", "ctaprove", False)
    If RC <> "" Then cad = cad & " AND " & RC
    
    
    'FORMA PAGO
    RC = CampoABD(txtFPago(6), "N", "spagop.codforpa", True)
    If RC <> "" Then cad = cad & " AND " & RC
    RC = CampoABD(txtFPago(7), "N", "spagop.codforpa", False)
    If RC <> "" Then cad = cad & " AND " & RC
    
    
    
    
    
    
    'Y lista de cuentas

    If ListadoCuentas <> "" Then
        NumRegElim = 1
        SQL = ""
        Do
            TotalRegistros = InStr(NumRegElim, ListadoCuentas, "|")
            If TotalRegistros > 0 Then
                If NumRegElim > 1 Then SQL = SQL & ","
                NumRegElim = 2
                SQL = SQL & "'" & Mid(ListadoCuentas, 1, TotalRegistros - 1) & "'"
                ListadoCuentas = Mid(ListadoCuentas, TotalRegistros + 1)
            End If
           
            
        Loop Until TotalRegistros = 0
        NumRegElim = 0
        cad = cad & " AND spagop.ctaprove IN (" & SQL & ")"
        
    End If
    
    
    'ORDEN
    cad = cad & " ORDER BY numfactu"
   
    
    
    
    'Contador
    SQL = "Select count(*) "
    SQL = SQL & cad
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalRegistros = 0
    If Not RS.EOF Then
        'Total registros
        TotalRegistros = RS.Fields(0)
    End If
    RS.Close
    
    If TotalRegistros = 0 Then
        'NO hay registros
        MsgBox "Ningun dato con esos valores", vbExclamation
        Exit Function
    End If
    
    'Llegados aqui, si k hay registros.
    '100 POR EJEMPLO
    If TotalRegistros > 100 Then
        'Ponemos visible el frame
        MostrarFrame = True
        PonerFrameProgressVisible
    Else
        MostrarFrame = False 'NO lo mostramos
    End If
    
    
    
    'Borramos tmp
    Conn.Execute "DELETE FROM Usuarios.zpendientes where codusu = " & vUsu.Codigo
    
    SQL = "SELECT spagop.*, cuentas.nommacta, stipoformapago.descformapago, stipoformapago.siglas,nomforpa " & cad
    
    'Cambiamos
''''''    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''''''    'Compruebo si hay repetidos en fecfactu|numfactu|siglas
''''''    SQL = ""
''''''    MismaClavePrimaria = "|"
''''''    While Not RS.EOF
''''''        RC = Format(RS!fecfactu, "yymmdd") & RS!siglas & RS!numfactu
''''''        If RC = SQL Then
''''''            RC = RC & "|"
''''''            If InStr(1, MismaClavePrimaria, "|" & RC) = 0 Then MismaClavePrimaria = MismaClavePrimaria & RC
''''''        Else
''''''            SQL = RC
''''''        End If
''''''        RS.MoveNext
''''''    Wend
''''''    SQL = RS.Source
''''''    RS.Close
    
    'Vuelvo a abrirlo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Agosto 2013
    'Añadimos en campo SITUACION donde pondra si esta emitido o no (emitdocum)
    
    'Mayo 2014
    'La factura la metemos en nomdirec. Asi NO da error duplicados
    
    CONT = 0
    SQL = "INSERT INTO Usuarios.zpendientes (codusu, serie_cta, factura, fecha, numorden,nomdirec,"
    SQL = SQL & "codforpa, nomforpa, codmacta,nombre, fecVto, importe, pag_cob,vencido,situacion) VALUES (" & vUsu.Codigo & ",'"
    Fecha = CDate(Text3(5).Text)
    DevfrmCCtas = ""
    While Not RS.EOF
        CONT = CONT + 1
        If MostrarFrame Then
            lbl2.Caption = "Registro: " & CONT
            lbl2.Refresh
        End If
        
'        'Por si se repiten misma fecfactura, mismo numero factura, mismo tipo de pago
'        If MismaClavePrimaria <> "" Then
'            'Hay claves repetidas no tiene en cuenta el vto
'            RC = "|" & Format(RS!fecfactu, "yymmdd") & RS!siglas & RS!numfactu & "|"
'            'Enero 2011
'            RC = "|" & Format(RS!fecfactu, "yymmdd") & RS!siglas & RS!numfactu & "#" & RS!numorden & "|"
'
'
'            If InStr(1, MismaClavePrimaria, RC) > 0 Then
'                RC = DevNombreSQL(RS!numfactu)
'                RC = FijaNumeroFacturaRepetido(RC)
'                Cad = RS!siglas & "','" & RC & "','" & Format(RS!fecfactu, FormatoFecha) & "'," & RS!numorden
'            Else
'                'El mismo de abajo
'                Cad = RS!siglas & "','" & DevNombreSQL(RS!numfactu) & "','" & Format(RS!fecfactu, FormatoFecha) & "'," & RS!numorden
'            End If
'        Else
'            'El procedimiento de antes
'            Cad = RS!siglas & "','" & DevNombreSQL(RS!numfactu) & "','" & Format(RS!fecfactu, FormatoFecha) & "'," & RS!numorden
'        End If
'
        
        'mayo 2014
        cad = RS!siglas & "','" & Format(CONT, "00000") & "','" & Format(RS!FecFactu, FormatoFecha) & "'," & RS!numorden & ",'" & DevNombreSQL(RS!NumFactu) & "'"
        
        
        'optMostraFP
        cad = cad & "," & RS!codforpa & ",'"
        If Me.optMostraFP(0).Value Then
            cad = cad & DevNombreSQL(RS!descformapago)
        Else
            cad = cad & DevNombreSQL(RS!nomforpa)
        End If
        cad = cad & "','" & RS!ctaprove & "','" & DevNombreSQL(RS!Nommacta) & "','"
        cad = cad & Format(RS!fecefect, FormatoFecha) & "',"
        cad = cad & TransformaComasPuntos(CStr(RS!ImpEfect)) & ","
        'Cobrado, si no es nulo
        If Not IsNull(RS!imppagad) Then
            cad = cad & TransformaComasPuntos(CStr(RS!imppagad))
        Else
            cad = cad & "0"
        End If
        If Fecha > RS!fecefect Then
            cad = cad & ",1"
        Else
            cad = cad & ",0"
        End If
        
        'Agosto 2013
        'Si esta en un tal-pag
        cad = cad & ",'"
        If DBLet(RS!emitdocum, "N") > 0 Then cad = cad & "*"
        
        cad = cad & "')"  'lleva el apostrofe
        cad = SQL & cad
        Conn.Execute cad
        'Siguiente
        RS.MoveNext
    Wend
    RS.Close
     
    PagosPendienteProv = True 'Para imprimir
    Exit Function
EPagosPendienteProv:
    MuestraError Err.Number, Err.Description
End Function



Private Function FijaNumeroFacturaRepetido(Numerofactura) As String
Dim I As Integer
Dim Aux As String
        If Len(Numerofactura) >= 10 Then
            MsgBox "Clave duplicada. Imposible insertar. " & RS!NumFactu & ": " & RS!FecFactu, vbExclamation
            FijaNumeroFacturaRepetido = Numerofactura
            Exit Function
        End If
        
        'Añadiremos guienos por detras
        For I = Len(Numerofactura) To 10
            'Añadirenos espacios en blanco al final
            Aux = RS!NumFactu & String(I - Len(Numerofactura), "_")
            If InStr(1, DevfrmCCtas, "|" & Aux & "|") = 0 Then
                'Devolvemos este numero de factura
                FijaNumeroFacturaRepetido = Aux
                If DevfrmCCtas = "" Then DevfrmCCtas = "|"
                DevfrmCCtas = DevfrmCCtas & Aux & "|"
                Exit Function
            End If
            
        Next I
        
        'Si llega aqui probaremos con los -
        For I = Len(Numerofactura) + 1 To 10
            'Añadirenos espacios en blanco al final
            Aux = String(I - Len(Numerofactura), "_") & RS!NumFactu
            If InStr(1, DevfrmCCtas, "|" & Aux & "|") = 0 Then
                'Devolvemos este numero de factura
                FijaNumeroFacturaRepetido = Aux
                DevfrmCCtas = DevfrmCCtas & Aux & "|"
                Exit Function
            End If
            
        Next I
End Function


Private Sub txtNumero_GotFocus(Index As Integer)
    PonFoco txtNumero(Index)
End Sub



Private Sub txtNumero_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtNumFac_GotFocus(Index As Integer)
    PonFoco txtnumfac(Index)
End Sub

Private Sub txtNumFac_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtNumFac_LostFocus(Index As Integer)
    txtnumfac(Index).Text = Trim(txtnumfac(Index).Text)
    If txtnumfac(Index).Text = "" Then Exit Sub
    
    If Not IsNumeric(txtnumfac(Index).Text) Then
        MsgBox "Campo debe ser numerico.", vbExclamation
        txtnumfac(Index).Text = ""
        PonFoco txtnumfac(Index)
    End If
End Sub

Private Sub txtRem_GotFocus(Index As Integer)
    PonFoco txtRem(Index)
End Sub

Private Sub txtRem_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtSerie_GotFocus(Index As Integer)
    PonFoco txtSerie(Index)
End Sub

Private Sub txtSerie_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtSerie_LostFocus(Index As Integer)
    txtSerie(Index).Text = UCase(txtSerie(Index))
End Sub

Private Sub txtVarios_GotFocus(Index As Integer)
    PonFoco txtVarios(Index)
End Sub

Private Sub txtVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Function ListadoRemesas() As Boolean
Dim Aux As String
    On Error GoTo EListadoRemesas
    ListadoRemesas = False
    
    SQL = ""
    RC = CampoABD(txtRem(0), "N", "codigo", True)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(1), "N", "codigo", False)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(2), "N", "anyo", True)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(3), "N", "anyo", False)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    'Tipo remesa
    RC = RemesaSeleccionTipoRemesa(chkTipoRemesa(0).Value = 1, chkTipoRemesa(1).Value = 1, chkTipoRemesa(2).Value = 1)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    Set RS = New ADODB.Recordset
    
    'ANTES
    RC = "SELECT remesas.*,nommacta from remesas,cuentas "
    RC = RC & " WHERE remesas.codmacta = cuentas.codmacta"
    If SQL <> "" Then RC = RC & " AND " & SQL
    
    
    'AHORA
    RC = "Select codigo,anyo, fecremesa,tiporemesa.descripcion as d1,descsituacion,remesas.codmacta,nommacta,"
    RC = RC & " Importe , remesas.descripcion, remesas.Tipo,situacion,tiporem"
    RC = RC & " from cuentas,tiposituacionrem,remesas left join tiporemesa on remesas.tipo=tiporemesa.tipo where remesas.codmacta=cuentas.codmacta"
    RC = RC & " and situacio=situacion"
    If SQL <> "" Then RC = RC & " AND " & SQL
    
    
    RS.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        MsgBox "Ninguna remesa para listar", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Function
    End If
    
    SQL = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "Delete from Usuarios.ztmplibrodiario where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    Set miRsAux = New ADODB.Recordset
    
    
    'Para insertar en una                       codigo /año desc                tipo   situacion
    SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2,texto3,texto4, texto5,importe1,  fecha1) VALUES ("
    'Para insertar en la otra
    cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien, fechaent,linliapu, codmacta, nommacta, numdocum, ampconce, debe,haber) VALUES ("
    
    
    
    TotalRegistros = 0
    
    While Not RS.EOF
        'Insertamos la cabecera de la remesas
        'Para ello el codigo sera: 200500001   es decir remesa 1 del año 2005 ...
        CONT = RS!Anyo * 100000 + RS!Codigo
        
        
        RC = vUsu.Codigo & "," & CONT & ",'" & DevNombreSQL(DBLet(RS!Descripcion, "T")) & "','" & DevNombreSQL(RS!Nommacta) & "',"
        'TIPO   situacion
        
        RC = RC & "'" & DBLet(RS!d1, "t") & "','" & DBLet(RS!descsituacion, "T") & "','"
        
        'Tipo remesa
        If RS!Tiporem = 2 Then
            RC = RC & "PAG"
        ElseIf RS!Tiporem = 3 Then
            RC = RC & "TAL"
        Else
            RC = RC & "EFE"
        End If
        RC = RC & "'," & TransformaComasPuntos(CStr(RS!Importe)) & ",'" & Format(RS!fecremesa, FormatoFecha) & "')"
    
        RC = SQL & RC
        Conn.Execute RC
       
        I = 1
        If Me.chkRem(0).Value = 1 Then
            'fecfaccl,scobro.codmacta,siturem,impcobro,impvenci,gastos,codfaccl,numserie,codbanco,codsucur,digcontr,cuentaba
            RC = "SELECT fecfaccl,scobro.codmacta,siturem,impcobro,impvenci,gastos,codfaccl,numserie,codbanco,codsucur,digcontr,scobro.cuentaba,nommacta"
            RC = RC & " ,fecvenci,scobro.iban from scobro,cuentas where codrem=" & RS!Codigo & " AND anyorem =" & RS!Anyo
            RC = RC & " AND cuentas.codmacta = scobro.codmacta  ORDER BY "
            If Me.optOrdenRem(1).Value Then
                'Codmacta
                RC = RC & "scobro.codmacta,numserie,codfaccl,fecfaccl"
            ElseIf Me.optOrdenRem(2).Value Then
                'Nommacta
                RC = RC & "nommacta,numserie,codfaccl,fecfaccl"
            ElseIf Me.optOrdenRem(0).Value Then
                'Numero factura
                RC = RC & "numserie,codfaccl,fecfaccl"
            Else
                'fcto
                RC = RC & "fecvenci,numserie,codfaccl,fecfaccl"
            
            End If
            miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                'If CONT = 200900004 Then Stop
                'cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien,
                'fechaent,linliapu, codmacta, nommacta, numdocum, ampconce, debe
                RC = vUsu.Codigo & "," & CONT & ",'" & Format(miRsAux!fecfaccl, FormatoFecha) & "',"
                RC = RC & I & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
                'Importe = miRsAux!impvenci - DBLet(miRsAux!impcobro, "N") + DBLet(miRsAux!Gastos, "N")
                If miRsAux!siturem > "B" Then
                    'No deberia ser NULL
                    Importe = DBLet(miRsAux!impcobro, "N")
                Else
                    Importe = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N")
                End If
                RC = RC & Format(miRsAux!codfaccl, "000000000") & "','"
                
                'Aqui pondre el CCC para los efectos
                '---------------------------------------------------
                'rc=rc & codbanco,codsucur,digcontr,scobro.cuentaba
                Aux = ""
                If RS!Tiporem = 1 Then
                        If IsNull(miRsAux!codbanco) Then
                            Aux = "0000"
                        Else
                            Aux = Format(miRsAux!codbanco, "0000")
                        End If
                        'AUX = AUX & " "
                        If IsNull(miRsAux!codsucur) Then
                            Aux = Aux & "0000"
                        Else
                            Aux = Aux & Format(miRsAux!codsucur, "0000")
                        End If
                        'AUX = AUX & " "
                        If IsNull(miRsAux!digcontr) Then
                            Aux = Aux & "**"
                        Else
                            Aux = Aux & Format(miRsAux!digcontr, "00")
                        End If
                        'AUX = AUX & " "
                        If IsNull(miRsAux!Cuentaba) Then
                            Aux = Aux & "0000"
                        Else
                            Aux = Aux & Format(miRsAux!Cuentaba, "0000000000")
                        End If
                Else
                    'Talon / Pagare. Si tiene numero puesto lo pondre
                 
                End If
                
                'Nuevo ENERO 2010
                'Fecha vto
                Aux = DBLet(miRsAux!IBAN, "T") & Aux
                If Len(Aux) > 24 Then Aux = Mid(Aux, 1, 24)
                Aux = Aux & "|" & Format(miRsAux!FecVenci, "dd/mm/yy")
                
                RC = RC & Aux & "'," & TransformaComasPuntos(CStr(Importe))
                
                'En el haber pongo el ascii de la serie
                '--------------------------------------
                RC = RC & "," & Asc(Left(DBLet(miRsAux!NUmSerie, "T") & " ", 1)) & ")"
                RC = cad & RC
            
                Conn.Execute RC
            
                'Sig
                I = I + 1
                miRsAux.MoveNext
            Wend
            miRsAux.Close
       
        Else
            'Tenemos k insertar una unica linea a blancos
            RC = vUsu.Codigo & "," & CONT & ",'1999-12-31'," & I & ",'','','','',0,0)"
            RC = cad & RC
            
            Conn.Execute RC
        End If
        TotalRegistros = TotalRegistros + 1
        RS.MoveNext
    Wend
    
    RS.Close
    
    
    
    
    Set RS = Nothing
    Set miRsAux = Nothing
    
    If Me.chkRem(0).Value = 1 Then
        If TotalRegistros = 0 Then
            MsgBox "No hay vencimientos asociados a las remesas", vbExclamation
            Exit Function
        End If
    End If
    ListadoRemesas = True
    Exit Function
EListadoRemesas:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
    Set miRsAux = Nothing

End Function









Private Function ListadoRemesasBanco() As Boolean
Dim Aux As String
Dim Cad2 As String
Dim J As Integer
    On Error GoTo EListadoRemesas
    ListadoRemesasBanco = False
    
    SQL = ""
    RC = CampoABD(txtRem(0), "N", "codigo", True)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(1), "N", "codigo", False)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(2), "N", "anyo", True)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(3), "N", "anyo", False)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    'Tipo remesa
    RC = RemesaSeleccionTipoRemesa(chkTipoRemesa(0).Value = 1, chkTipoRemesa(1).Value = 1, chkTipoRemesa(2).Value = 1)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    Set RS = New ADODB.Recordset
    
    'ANTES
    RC = "SELECT remesas.*,nommacta from remesas,cuentas "
    RC = RC & " WHERE remesas.codmacta = cuentas.codmacta"
    If SQL <> "" Then RC = RC & " AND " & SQL
    
    
    'AHORA
    RC = "Select codigo,anyo, fecremesa,tiporemesa.descripcion as d1,descsituacion,remesas.codmacta,nommacta,"
    RC = RC & " Importe , remesas.descripcion, remesas.Tipo,situacion,tiporem"
    RC = RC & " from cuentas,tiposituacionrem,remesas left join tiporemesa on remesas.tipo=tiporemesa.tipo where remesas.codmacta=cuentas.codmacta"
    RC = RC & " and situacio=situacion"
    If SQL <> "" Then RC = RC & " AND " & SQL
    
    
    RS.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        MsgBox "Ninguna remesa para listar", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Function
    End If
    
    SQL = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "Delete from Usuarios.ztmplibrodiario where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    Set miRsAux = New ADODB.Recordset
    
    
    'Para insertar en una                       codigo /año desc                tipo   situacion
    SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2,texto3,texto4, texto5,importe1,  fecha1,observa1) VALUES ("
    
    
    
    TotalRegistros = 0
    
    While Not RS.EOF
        'Insertamos la cabecera de la remesas
        'Para ello el codigo sera: 200500001   es decir remesa 1 del año 2005 ...
        CONT = RS!Anyo * 100000 + RS!Codigo
        
        
        RC = vUsu.Codigo & "," & CONT & ",'" & DevNombreSQL(DBLet(RS!Descripcion, "T")) & "','" & DevNombreSQL(RS!Nommacta) & "',"
        'TIPO   situacion
        
        RC = RC & "'" & DBLet(RS!d1, "t") & "','" & DBLet(RS!descsituacion, "T") & "','"
        
        'Tipo remesa
        If RS!Tiporem = 2 Then
            RC = RC & "PAG"
        ElseIf RS!Tiporem = 3 Then
            RC = RC & "TAL"
        Else
            RC = RC & "EFE"
        End If
        RC = RC & "'," & TransformaComasPuntos(CStr(RS!Importe)) & ",'" & Format(RS!fecremesa, FormatoFecha) & "','"
        
        Cad2 = "Select * from ctabancaria where codmacta = '" & RS!codmacta & "'"
        miRsAux.Open Cad2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad2 = "NO ENCONTRADO"
        If Not miRsAux.EOF Then
            Cad2 = Trim(DBLet(miRsAux!IBAN, "T") & " ") & Format(DBLet(miRsAux!Entidad, "N"), "0000") & " " & Format(DBLet(miRsAux!Oficina, "N"), "0000") & " "
            If IsNull(miRsAux!Control) Then
                Cad2 = Cad2 & "*"
            Else
                Cad2 = Cad2 & miRsAux!Control
            End If
            Cad2 = Cad2 & " " & Format(DBLet(miRsAux!CtaBanco, "N"), "0000000000")
        End If
        miRsAux.Close
        RC = RC & Cad2 & "')"
        'ctabancaria(entidad,oficina,control,ctabanco)
        Cad2 = ""
        
        RC = SQL & RC
        Conn.Execute RC
       
        I = 1
        
            'Voy a comprobar que existen
            RC = "SELECT codmacta,reftalonpag FROM scobro "
            RC = RC & "  WHERE codrem=" & RS!Codigo & " AND anyorem =" & RS!Anyo
            RC = RC & " GROUP BY 1,2 "
            miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad2 = ""
            While Not miRsAux.EOF
                Cad2 = Cad2 & " scarecepdoc.codmacta = '" & miRsAux!codmacta & "' AND numeroref = '" & DevNombreSQL(miRsAux!reftalonpag) & "'|"
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            If Cad2 = "" Then
                MsgBox "Error obteniendo cuenta/referenciatalon", vbExclamation
                RS.Close
                Exit Function
            End If
                
            'Comprobare que existen y de paso los inserto
            While Cad2 <> ""
                J = InStr(1, Cad2, "|")
                Aux = Mid(Cad2, 1, J - 1)
                Cad2 = Mid(Cad2, J + 1)
                
                'RC = "SELECT * FROM scarecepdoc ,slirecepdoc,cuentas WHERE codigo=id AND scarecepdoc.codmacta=cuentas.codmacta AND " & Aux
                RC = "SELECT * FROM scarecepdoc ,cuentas WHERE  scarecepdoc.codmacta=cuentas.codmacta AND " & Aux
               
                miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If miRsAux.EOF Then
                    MsgBox "No se encuentra la referencia; " & Aux, vbExclamation
                    miRsAux.Close
                    RS.Close
                    Exit Function
                End If
                
                While Not miRsAux.EOF
            
                
                
                
                
                    
                    'Para insertar en la otra
                    cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien, fechaent,linliapu,  nommacta,codmacta, numdocum, ampconce, debe,haber) VALUES ("
                
                    'Trampas:  Entre codmacta numdocum llevare el numero de talon. Ya que suman 20 y reftal es len20
                    RC = vUsu.Codigo & "," & CONT & ",'" & Format(miRsAux!fechavto, FormatoFecha) & "',"
                    RC = RC & I & ",'" & DevNombreSQL(miRsAux!Nommacta) & "','"
                    Importe = DBLet(miRsAux!Importe, "N")
                    
                    'Referencia talon
                    Aux = DevNombreSQL(miRsAux!numeroref)
                    If Len(Aux) > 10 Then
                        RC = RC & Mid(Aux, 1, 10) & "','" & Mid(Aux, 11)
                    Else
                        RC = RC & Aux & "','"
                    End If
                    'Banco
                    RC = RC & "','" & DevNombreSQL(miRsAux!Banco) & "',"
                    
                    'Talon / Pagare. Si tiene numero puesto lo pondre
                    RC = RC & TransformaComasPuntos(CStr(Importe))
                    
                    'En el haber pongo el ascii de la serie
                    '--------------------------------------
                    RC = RC & ",0)"
                    RC = cad & RC
                
                    Conn.Execute RC
                
                    'Sig
                    I = I + 1
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
            Wend

        TotalRegistros = TotalRegistros + 1
        RS.MoveNext
    Wend
    
    RS.Close
    
    
    
    
      Set RS = Nothing
    Set miRsAux = Nothing
    
    If Me.chkRem(0).Value = 1 Then
        If TotalRegistros = 0 Then
            MsgBox "No hay vencimientos asociados a las remesas", vbExclamation
            Exit Function
        End If
    End If
    ListadoRemesasBanco = True
    Exit Function
EListadoRemesas:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
    Set miRsAux = Nothing

End Function




'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'
'               CREDITO CAUCION
'
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

Private Function ListadoTransferencias() As Boolean
Dim Importe As Currency

    On Error GoTo EListadoTransferencias
    ListadoTransferencias = False
    
    SQL = ""
    RC = CampoABD(txtNumero(0), "N", "codigo", True)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtNumero(1), "N", "codigo", False)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    
    cad = RC
    
    Set RS = New ADODB.Recordset
    
    RC = "SELECT stransfer.*,nommacta from stransfer"
    If Opcion = 13 Then RC = RC & "cob"
    RC = RC & " as stransfer,cuentas "
    RC = RC & " WHERE stransfer.codmacta = cuentas.codmacta"
    If SQL <> "" Then RC = RC & " AND " & SQL
    
    RS.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        MsgBox "Ninguna valor para listar", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Function
    End If
    
    SQL = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "Delete from Usuarios.ztmplibrodiario where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    If Opcion = 13 Then Conn.Execute "Delete from usuarios.zcuentas where codusu =" & vUsu.Codigo
        
    
    Set miRsAux = New ADODB.Recordset
    
    
    'Para insertar en una                       codigo /año desc
    SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2, importe1,  fecha1) VALUES ("
    'Para insertar en la otra
    cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien, fechaent,linliapu, codmacta, nommacta, numdocum, ampconce, debe) VALUES ("
    
    
    

    
    While Not RS.EOF
        'Insertamos la cabecera de la remesas
        'Para ello el codigo sera: 200500001   es decir remesa 1 del año 2005 ...
        CONT = RS!Codigo
        
        
        RC = vUsu.Codigo & "," & CONT & ",'" & DevNombreSQL(DBLet(RS!Descripcion, "T")) & "','" & DevNombreSQL(RS!Nommacta) & "',"
        RC = RC & TransformaComasPuntos("0") & ",'" & Format(RS!Fecha, FormatoFecha) & "')"
    
        RC = SQL & RC
        Conn.Execute RC
       
        I = 1
     
            
            If Opcion = 13 Then
                RC = "scobro"
            Else
                RC = "spagop"
            End If
            RC = "SELECT " & RC & ".*,nommacta from cuentas," & RC
            RC = RC & " WHERE transfer = " & RS!Codigo
            RC = RC & " AND cuentas.codmacta = "
            If Opcion = 13 Then
                RC = RC & " scobro.codmacta "
                RC = RC & " ORDER BY scobro.codmacta,fecfaccl"
            Else
                RC = RC & " spagop.ctaprove "
                RC = RC & " ORDER BY ctaprove,fecfactu"
            End If
            miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                'cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien,
                'fechaent,linliapu, codmacta, nommacta, numdocum, ampconce, debe "
                If Opcion = 13 Then
                    Fecha = miRsAux!fecfaccl
                Else
                    Fecha = miRsAux!FecFactu
                End If
                RC = vUsu.Codigo & "," & CONT & ",'" & Format(Fecha, FormatoFecha) & "',"
                RC = RC & I & ",'"
                If Opcion = 13 Then
                    RC = RC & miRsAux!codmacta
                Else
                    RC = RC & miRsAux!ctaprove
                End If
                
                RC = RC & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
                
                
                'Cuenta
                If Opcion <> 13 Then
                    RC = RC & DevNombreSQL(miRsAux!NumFactu) & "','"
                    
                    'Noviembre 2013
                    'Añadimos el IBAN
                    
                    RC = RC & Trim(DBLet(miRsAux!IBAN, "T") & " " & Format(DBLet(miRsAux!Entidad, "T"), "0000")) & " " & Format(DBLet(miRsAux!Oficina, "T"), "0000") & " "
                    RC = RC & Mid(DBLet(miRsAux!CC, "T") & "**", 1, 2) & " " & Right(String(10, "0") & DBLet(miRsAux!Cuentaba, "T"), 10)
                    Importe = miRsAux!ImpEfect - (DBLet(miRsAux!imppagad, "N"))
                    RC = RC & "'," & TransformaComasPuntos(CStr(Importe)) & ")"
                Else
                    RC = RC & DevNombreSQL(miRsAux!codfaccl) & "','"
                    
                    CadenaDesdeOtroForm = "NO"
                    If DBLet(miRsAux!codbanco, "N") > 0 Then
                        'Es especifico para ESCALONO, pero no molesta a nadie
                        If DBLet(miRsAux!Cuentaba, "T") = "8888888888" Then
                            'Seguira poniendo  cuenta no domiciliada
                        Else
                            CadenaDesdeOtroForm = ""
                        End If
                    End If
                    If CadenaDesdeOtroForm = "" Then
                        'OK, ponemos la cuenta
                        CadenaDesdeOtroForm = Trim(DBLet(miRsAux!IBAN, "T") & " " & Format(DBLet(miRsAux!codbanco, "N"), "0000")) & " " & Format(DBLet(miRsAux!codsucur, "N"), "0000") & " "
                        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "**  ******" & Right(String(4, "0") & DBLet(miRsAux!Cuentaba, "T"), 4)
                        
                    Else
                        'CUENTANODOMICILIADA
                        CadenaDesdeOtroForm = "NODOM"  'en el rpt haremos un replace
                    End If
                    RC = RC & CadenaDesdeOtroForm
                    Importe = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N") - DBLet(miRsAux!impcobro, "N")
                    RC = RC & "'," & TransformaComasPuntos(CStr(Importe)) & ")"
                End If
                RC = cad & RC
            
                Conn.Execute RC
            
                'Sig
                I = I + 1
                miRsAux.MoveNext
            Wend
            miRsAux.Close
       
'        Else
'            'Tenemos k insertar una unica linea a blancos
'            RC = vUsu.Codigo & "," & CONT & ",''," & I & ",'','','','',0)"
'            RC = Cad & RC
'
'            Conn.Execute RC
'        End If
        RS.MoveNext
    Wend
    RS.Close
    CadenaDesdeOtroForm = ""
    
    Set RS = Nothing
    Set miRsAux = Nothing
    
    If Opcion = 13 Then
        'Puede ser carta
        If chkCartaAbonos.Value Then
            'En nommacta pongo la provincia  (desprovi)
            cad = "INSERT INTO usuarios.zcuentas(codusu,codmacta,nommacta,razosoci,dirdatos,codposta,despobla,nifdatos)"
            cad = cad & " Select " & vUsu.Codigo & ",codmacta,desprovi,razosoci,dirdatos,codposta,despobla,nifdatos FROM cuentas WHERE "
            cad = cad & " codmacta IN (select distinct(codmacta) from usuarios.ztmplibrodiario where codusu =" & vUsu.Codigo & ")"
            Ejecuta cad
        
        
            cad = "apoderado"
            RC = DevuelveDesdeBD("contacto", "empresa2", "1", "1", "N", cad)
            If RC = "" Then RC = cad
            If RC <> "" Then
                cad = "UPDATE usuarios.ztesoreriacomun SET observa1='" & DevNombreSQL(RC) & "'"
                cad = cad & " WHERE codusu = " & vUsu.Codigo
                Conn.Execute cad
            End If
        End If
    End If
    
    If Me.chkRem(0).Value = 1 Then
        If I = 1 Then
            MsgBox "No hay vencimientos asociados a las transferencias", vbExclamation
            Exit Function
        End If
    End If
    ListadoTransferencias = True
    Exit Function
EListadoTransferencias:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
    Set miRsAux = Nothing
End Function





Private Function ListAseguBasico() As Boolean
    On Error GoTo EListAseguBasico
    ListAseguBasico = False
    
    cad = "DELETE FROM Usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    cad = "Select * from cuentas where numpoliz<>"""""
    SQL = ""
    RC = CampoABD(Text3(21), "F", "fecsolic", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(22), "F", "fecconce", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    If SQL <> "" Then cad = cad & SQL
        
    
    'ORDENACION
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    cad = cad & " ORDER BY " & RC
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,fecha1,fecha2,importe1,"
    cad = cad & "importe2,observa1,observa2) VALUES (" & vUsu.Codigo & ","
        
    While Not miRsAux.EOF
        CONT = CONT + 1
        SQL = CONT & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','" & DBLet(miRsAux!nifdatos, "T") & "','"
        SQL = SQL & DevNombreSQL(miRsAux!numpoliz) & "',"
        'Fecha sol y concesion
        SQL = SQL & CampoBD_A_SQL(miRsAux!fecsolic, "F", True) & "," & CampoBD_A_SQL(miRsAux!fecconce, "F", True) & ","
        'Importes sol y concesion
        SQL = SQL & CampoBD_A_SQL(miRsAux!credisol, "N", True) & "," & CampoBD_A_SQL(miRsAux!credicon, "N", True) & ","
        'Observaciones
        RC = Memo_Leer(miRsAux!observa)
        If Len(RC) = 0 Then
            'Los dos campos NULL
            SQL = SQL & "NULL,NULL"
        Else
            If Len(RC) < 255 Then
                SQL = SQL & "'" & DevNombreSQL(RC) & "',NULL"
            Else
                SQL = SQL & "'" & DevNombreSQL(Mid(RC, 1, 255))
                RC = Mid(RC, 256)
                SQL = SQL & "','" & DevNombreSQL(Mid(RC, 1, 255)) & "'"
            End If
        End If
        
        SQL = SQL & ")"
        Conn.Execute cad & SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If CONT > 0 Then
        ListAseguBasico = True
    Else
        MsgBox "Ningun datos con esos valores", vbExclamation
    End If
    Exit Function
EListAseguBasico:
    MuestraError Err.Number, "ListAseguBasico"
End Function





Private Function ListAsegFacturacion() As Boolean
Dim FP As Integer 'Forma de pago
Dim Cadpago As String
    On Error GoTo EListAsegFacturacion
    ListAsegFacturacion = False
    
    cad = "DELETE FROM Usuarios.zpendientes  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    
    If Me.optFecgaASig(0).Value Then
        cad = "fecfaccl"
    Else
        cad = "fecvenci"
    End If
        
    SQL = ""
    RC = CampoABD(Text3(21), "F", cad, True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(22), "F", cad, False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    
    
    
    cad = "Select scobro.*,nommacta,numpoliz,nomforpa,forpa from scobro,cuentas,sforpa where scobro.codmacta= cuentas.codmacta AND numpoliz<>"""""
    cad = cad & " and scobro.codforpa=sforpa.codforpa AND sforpa.tipforpa <> " & vbEfectivo 'EL EFECTIVO NO se comprueba
    If SQL <> "" Then cad = cad & SQL
        
    
    'ORDENACION
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    cad = cad & " ORDER BY " & RC
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0

    cad = "INSERT INTO Usuarios.zpendientes (codusu, serie_cta, factura, fecha, numorden,"
    cad = cad & "codforpa, nomforpa, codmacta, nombre, fecVto, importe,"
    cad = cad & "Situacion,pag_cob, vencido,  gastos) VALUES (" & vUsu.Codigo & ","
    Cadpago = ","
    While Not miRsAux.EOF
        CONT = CONT + 1
        SQL = "'" & miRsAux!NUmSerie & "','" & Format(miRsAux!codfaccl, "000000000") & "','" & Format(miRsAux!fecfaccl, FormatoFecha) & "',"
        FP = miRsAux!codforpa
        If optFP(1).Value Then
            If DBLet(miRsAux!Forpa, "N") > 0 Then
                FP = miRsAux!Forpa
                If InStr(1, Cadpago, "," & FP & ",") = 0 Then Cadpago = Cadpago & FP & ","
            End If
        End If
        SQL = SQL & miRsAux!numorden & "," & FP & ",'" & DevNombreSQL(miRsAux!nomforpa) & "','" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta)
        SQL = SQL & "','" & Format(miRsAux!FecVenci, FormatoFecha) & "',"
        'IMporte
        Importe = miRsAux!ImpVenci
        If Not IsNull(miRsAux!Gastos) Then Importe = Importe + miRsAux!Gastos
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        SQL = SQL & TransformaComasPuntos(CStr(Importe))
        'Situacion tengo numpoliza
        SQL = SQL & ",'" & DevNombreSQL(miRsAux!numpoliz) & "',"
        'Gastos e imvenci van a la columna pag_cob   Julio 2009
        Importe = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N")
        SQL = SQL & TransformaComasPuntos(CStr(Importe))
        'El resto
        SQL = SQL & ",0,NULL)"
        
        Conn.Execute cad & SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    If CONT = 0 Then
        MsgBox "Ningun datos con esos valores", vbExclamation
        Exit Function
    End If
    
    
    'Si ha cambiado valores en la forma de pago
    If Len(Cadpago) > 1 Then
        Cadpago = Mid(Cadpago, 2)
        Cadpago = Mid(Cadpago, 1, Len(Cadpago) - 1)
        cad = "select codforpa,nomforpa from sforpa where codforpa in (" & Cadpago & ") GROUP by  codforpa"
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        cad = " WHERE codusu = " & vUsu.Codigo & " AND codforpa = "
        While Not miRsAux.EOF
            SQL = "UPDATE Usuarios.zpendientes SET nomforpa = '" & DevNombreSQL(miRsAux!nomforpa) & "'" & cad & miRsAux!codforpa
            If Not Ejecuta(SQL) Then MsgBox "Error actualizando tmp.  Forpa: " & miRsAux!codforpa & " " & miRsAux!nomforpa, vbExclamation
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    ListAsegFacturacion = True
    
    
    Exit Function
EListAsegFacturacion:
    MuestraError Err.Number, "ListAseguBasico"
End Function



Private Function ListAsegImpagos() As Boolean
    On Error GoTo EListAsegImpagos
    ListAsegImpagos = False
    
    cad = "DELETE FROM Usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    cad = "Select numserie,codfaccl,numorden,fecvenci,impvenci,impcobro,gastos,scobro.codmacta,nommacta,despobla,desprovi,numpoliz,nomforpa from "
    cad = cad & "scobro,cuentas,sforpa where scobro.codmacta= cuentas.codmacta AND numpoliz<>"""""
    cad = cad & " and scobro.codforpa=sforpa.codforpa AND sforpa.tipforpa <> " & vbEfectivo 'EL EFECTIVO NO se comprueba
    'Impagados
    cad = cad & " AND devuelto = 1"
    SQL = ""
    RC = CampoABD(Text3(21), "F", "fecvenci", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(22), "F", "fecvenci", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    If SQL <> "" Then cad = cad & SQL
        
    
    'ORDENACION
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    cad = cad & " ORDER BY " & RC
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,texto5,texto6,fecha1,importe1) VALUES (" & vUsu.Codigo & ","
        
    While Not miRsAux.EOF
        CONT = CONT + 1
        SQL = CONT & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','" & DevNombreSQL(DBLet(miRsAux!desPobla, "T")) & "','"
        SQL = SQL & DevNombreSQL(DBLet(miRsAux!desProvi, "T")) & "','" & DevNombreSQL(miRsAux!numpoliz) & "','"
        SQL = SQL & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000000") & "',"
        'Fecha vto
        SQL = SQL & CampoBD_A_SQL(miRsAux!FecVenci, "F", True) & ","
        'Importes sol y concesion
        Importe = miRsAux!ImpVenci
        If Not IsNull(miRsAux!Gastos) Then Importe = Importe + miRsAux!Gastos
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        SQL = SQL & TransformaComasPuntos(CStr(Importe))
        
    
        SQL = SQL & ")"
        Conn.Execute cad & SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If CONT > 0 Then
        ListAsegImpagos = True
    Else
        MsgBox "Ningun datos con esos valores", vbExclamation
    End If
    Exit Function
EListAsegImpagos:
    MuestraError Err.Number, "ListAsegImpagos"
End Function


Private Function ListAsegEfectos() As Boolean
Dim TotalCred As Currency

    On Error GoTo EListAsegEfectos
    ListAsegEfectos = False
    
    cad = "DELETE FROM Usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    cad = "Select numserie,codfaccl,numorden,fecvenci,impvenci,impcobro,gastos,fecfaccl,devuelto,scobro.codmacta,nommacta,credicon from "
    cad = cad & "scobro,cuentas,sforpa where scobro.codmacta= cuentas.codmacta AND numpoliz<>"""""
    cad = cad & " and scobro.codforpa=sforpa.codforpa AND sforpa.tipforpa <> " & vbEfectivo 'EL EFECTIVO NO se comprueba

    SQL = ""
    RC = CampoABD(Text3(21), "F", "fecvenci", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(22), "F", "fecvenci", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    If SQL <> "" Then cad = cad & SQL
        
    
    'ORDENACION
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    cad = cad & " ORDER BY codmacta,fecvenci"
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    'Seran:                                                     codmac,nomma,credicon,numfac,fecfac,fvto,impvto,disponible,vencida
    cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,fecha1,fecha2,importe1,importe2,opcion) VALUES (" & vUsu.Codigo & ","
    RC = ""
    
    While Not miRsAux.EOF
        If RC <> miRsAux!codmacta Then
            RC = miRsAux!codmacta
            TotalCred = DBLet(miRsAux!credicon, "N")
            CadenaDesdeOtroForm = ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
            If IsNull(miRsAux!credicon) Then
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & "0,00','"
            Else
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Format(miRsAux!credicon, FormatoImporte) & "','"
            End If
        End If
        CONT = CONT + 1
        SQL = CONT & CadenaDesdeOtroForm
        SQL = SQL & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000000") & "',"
        'Fecha fac
        SQL = SQL & CampoBD_A_SQL(miRsAux!fecfaccl, "F", True) & ","
        'Fecha vto
        SQL = SQL & CampoBD_A_SQL(miRsAux!FecVenci, "F", True) & ","
        'Importes sol y concesion
        Importe = miRsAux!ImpVenci
        If Not IsNull(miRsAux!Gastos) Then Importe = Importe + miRsAux!Gastos
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        SQL = SQL & TransformaComasPuntos(CStr(Importe))
        TotalCred = TotalCred - Importe
        SQL = SQL & "," & TransformaComasPuntos(CStr(TotalCred))
       
        'Devuelto
        SQL = SQL & "," & DBLet(miRsAux!Devuelto, "N") & ")"
    
        Conn.Execute cad & SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If CONT > 0 Then
        ListAsegEfectos = True
    Else
        MsgBox "Ningun datos con esos valores", vbExclamation
    End If
    Exit Function
EListAsegEfectos:
    MuestraError Err.Number, "ListAsegEfec"
End Function



Private Sub GeneraComboCuentas()
'
'    If Opcion = 1 Then
'        'COBROS PENDIENTES
'    Else: Pagos
'
        cmbCuentas(Opcion - 1).Clear
        cmbCuentas(Opcion - 1).AddItem "Sin especificar"
        
        cmbCuentas(Opcion - 1).AddItem "Crear selección"
              
        'En el tag tendremos las cuentas seleccionadas
        If Me.cmbCuentas(Opcion - 1).Tag <> "" Then cmbCuentas(Opcion - 1).AddItem "Cuentas seleccionadas"


    'Cargo aqui tb los checks
    CargaTextosTipoPagos False
End Sub



Private Sub CargaTextosTipoPagos(Reclamaciones As Boolean)
    
    Set miRsAux = New ADODB.Recordset
    cad = "Select tipoformapago, descformapago,siglas from stipoformapago order by tipoformapago "
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If Reclamaciones Then
            chkTipPagoRec(miRsAux!tipoformapago).Caption = miRsAux!siglas
            chkTipPagoRec(miRsAux!tipoformapago).Visible = True
            chkTipPagoRec(miRsAux!tipoformapago).Value = 1
        
        Else
            chkTipPago(miRsAux!tipoformapago).Caption = miRsAux!siglas
            chkTipPago(miRsAux!tipoformapago).Visible = True
            chkTipPago(miRsAux!tipoformapago).Value = 1
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



'Para conceptos y diarios
'Opcion: 0- Diario
'        1- Conceptos
'        2- Centros de coste
'        3- Gastos fijos
'        4. Hco compensaciones
Private Sub LanzaBuscaGrid(Indice As Integer, OpcionGrid As Byte)


'--monica
'
'    Select Case OpcionGrid
'    Case 0
'    'Diario
'        DevfrmCCtas = "0"
'        cad = "Número|numdiari|N|30·"
'        cad = cad & "Descripción|desdiari|T|60·"
'
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = "Tiposdiario"
'        frmB.vSQL = ""
'
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = "Diario"
'        frmB.vSelElem = 0
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        If DevfrmCCtas <> "" Then
'           Me.txtDiario(Indice) = RecuperaValor(DevfrmCCtas, 1)
'           Me.txtDescDiario(Indice) = RecuperaValor(DevfrmCCtas, 2)
'        End If
' Case 1
'        'Conceptos
'        DevfrmCCtas = "0"
'        cad = "Codigo|codconce|N|30·"
'        cad = cad & "Descripción|nomconce|T|60·"
'
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = "Conceptos"
'        frmB.vSQL = ""
'
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = "CONCEPTOS"
'        frmB.vSelElem = 0
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        If DevfrmCCtas <> "" Then
'           Me.txtConcpto(Indice) = RecuperaValor(DevfrmCCtas, 1)
'           Me.txtDescConcepto(Indice) = RecuperaValor(DevfrmCCtas, 2)
'        End If
'
'    Case 2
'        'Centros de coste
'        DevfrmCCtas = "0"
'        cad = "Codigo|codccost|T|30·"
'        cad = cad & "Descripción|nomccost|T|60·"
'
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = "cabccost"
'        frmB.vSQL = ""
'
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = "Centros de coste"
'        frmB.vSelElem = 0
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        If DevfrmCCtas <> "" Then
'
'           txtCCost(Indice) = RecuperaValor(DevfrmCCtas, 1)
'           txtDescCCoste(Indice) = RecuperaValor(DevfrmCCtas, 2)
'        End If
'
'    Case 3
'        'Gasto fijos  sgastfij codigo Descripcion
'        DevfrmCCtas = "0"
'        cad = "Código|codigo|T|30·"
'        cad = cad & "Descripción|Descripcion|T|60·"
'
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = "sgastfij"
'        frmB.vSQL = ""
'
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vTitulo = "Gastos fijos"
'        frmB.vSelElem = 0
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        If DevfrmCCtas <> "" Then
'
'           txtGastoFijo(Indice) = RecuperaValor(DevfrmCCtas, 1)
'           txtDescGastoFijo(Indice) = RecuperaValor(DevfrmCCtas, 2)
'        End If
'
'    Case 4
'        'Gasto fijos  sgastfij codigo Descripcion
'        '-------------------------------------------
'        DevfrmCCtas = "0"
'        cad = "Código|codigo|T|10·"
'        cad = cad & "Fecha|fecha|T|26·"
'        cad = cad & "Cuenta|codmacta|T|14·"
'        cad = cad & "Nombre|nommacta|T|50·"
'
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = "scompenclicab"
'        frmB.vSQL = ""
'
'        '###A mano
'        frmB.vDevuelve = "0|"
'        frmB.vTitulo = "Hco. compensaciones cliente"
'        frmB.vSelElem = 0
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        If DevfrmCCtas <> "" Then
'            DevfrmCCtas = RecuperaValor(DevfrmCCtas, 1)
'            If DevfrmCCtas = "" Then DevfrmCCtas = "0"
'            If Val(DevfrmCCtas) Then
'                CONT = Val(DevfrmCCtas)
'                ImprimiCompensacion CONT
'            End If
'
'        End If
'    End Select
End Sub

                                       '                Para saber el index del listview
Public Sub InsertaItemComboCompensaVto(TEXTO As String, Indice As Integer)
    Me.cboCompensaVto.AddItem TEXTO
    Me.cboCompensaVto.ItemData(Me.cboCompensaVto.NewIndex) = Indice
End Sub


Private Function GeneraDatosTalPag() As Boolean
Dim B As Boolean

    'Borramos los tmp
    SQL = "DELETE FROM usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo
    Conn.Execute SQL

    If chkLstTalPag(3).Value = 1 Then
        B = GeneraDatosTalPagDesglosado
    Else
        'Sin desglosar
        B = GeneraDatosTalPagSinDesglose
    End If
    GeneraDatosTalPag = B
End Function

Private Function GeneraDatosTalPagDesglosado() As Boolean
    On Error GoTo EGeneraDatosTalPag
    GeneraDatosTalPagDesglosado = False
    
    

       
       
    SQL = "select slirecepdoc.*,scarecepdoc.*,nommacta,nifdatos from slirecepdoc,scarecepdoc,cuentas "
    SQL = SQL & " where slirecepdoc.id =scarecepdoc.codigo and scarecepdoc.codmacta=cuentas.codmacta"
    If Text3(24).Text <> "" Then SQL = SQL & " AND fecharec >= '" & Format(Text3(24).Text, FormatoFecha) & "'"
    If Text3(25).Text <> "" Then SQL = SQL & " AND fecharec <= '" & Format(Text3(25).Text, FormatoFecha) & "'"
    'Contabilizado
    'SQL = SQL & " AND Contabilizada =  1"
    'Si esta llevada a banco o no
    If cboListPagare.ListIndex >= 1 Then SQL = SQL & " AND LlevadoBanco = " & Abs(cboListPagare.ListIndex = 1)
    '------------------------------------------------------------------------
    I = -1
    If (chkLstTalPag(0).Value = 1) Xor (chkLstTalPag(1).Value = 1) Then
        'Solo uno seleccionado
        I = 0
        If (chkLstTalPag(0).Value = 1) Then I = 1
            
    Else
        If (chkLstTalPag(0).Value = 0) Then
            MsgBox "Seleccione talon o pagare", vbExclamation
            Exit Function
        End If
    End If
    If I >= 0 Then SQL = SQL & " AND talon = " & I

    'Si ID
    If txtnumfac(2).Text <> "" Then SQL = SQL & " AND codigo >= " & txtnumfac(2).Text
    If txtnumfac(3).Text <> "" Then SQL = SQL & " AND codigo <= " & txtnumfac(3).Text

    Set RS = New ADODB.Recordset
    
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    RC = ""
    While Not RS.EOF
        I = I + 1
        'ztesoreriacomun (`codusu`,`codigo
        SQL = ", (" & vUsu.Codigo & "," & I & ","
        
        'texto1`,`texto2`,`texto3`,y el 4
        SQL = SQL & "'" & DevNombreSQL(RS!numeroref) & "','" & DevNombreSQL(RS!Banco) & "','"
        SQL = SQL & DevNombreSQL(RS!codmacta) & "','" & DevNombreSQL(RS!Nommacta) & "','"
        
        
        '5 Serie y numero factura
        SQL = SQL & RS!NUmSerie & Format(RS!numfaccl, "000000") & "',"
        '`importe1`
        'SQL = SQL & TransformaComasPuntos(CStr(Rs!Importe)) & ","
        SQL = SQL & TransformaComasPuntos(CStr(RS.Fields(5))) & ",'"   'La columna 5 es sli.importe
        
        'texto6=nifdatos
        SQL = SQL & DevNombreSQL(DBLet(RS!nifdatos, "N"))
        
        '`fecha1`,`fecha2`,`fecha3`
        SQL = SQL & "','" & Format(RS!fecharec, FormatoFecha) & "',"
        SQL = SQL & "'" & Format(RS!fechavto, FormatoFecha) & "',"
        SQL = SQL & "'" & Format(RS!fecfaccl, FormatoFecha) & "')"
    
        RC = RC & SQL
        
        RS.MoveNext
    Wend
    RS.Close
    
    If I > 0 Then
        RC = Mid(RC, 3) 'QUITO LA PRIMERA COMA
        'OK hay datos. Insertamos
        SQL = "INSERT INTO usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,`texto2`,`texto3`,"
        SQL = SQL & "`texto4`,`texto5`,`importe1`,texto6,`fecha1`,`fecha2`,`fecha3`) VALUES "
        SQL = SQL & RC
        Conn.Execute SQL
    
        'Si estamos emitiendo el justicante de recepcion, guardare en z340 los campos
        'fiscales del cliente para su impresion
        If Me.chkLstTalPag(2).Value = 1 Then
            SQL = "DELETE FROM usuarios.z347 WHERE codusu = " & vUsu.Codigo
            Conn.Execute SQL
            
            SQL = "DELETE FROm Usuarios.z347carta WHERE codusu = " & vUsu.Codigo
            Conn.Execute SQL
            
            espera 0.3
            
            
            'En texto3 esta la codmacta
            SQL = "Select texto3 from usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo & " GROUP BY texto3"
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            RC = ""
            While Not RS.EOF
                RC = RC & ", '" & RS!texto3 & "'"
                RS.MoveNext
            Wend
            RS.Close
            
            
            
            
            
            'No puede ser EOF
            RC = Trim(Mid(RC, 2))
            'Monto un superselect
            'pongo el IGNORE por si acaso hay cuentas con el mismo NIF
            SQL = "insert ignore into usuarios.z347 (`codusu`,`cliprov`,`nif`,`razosoci`,`dirdatos`,`codposta`,`despobla`,`Provincia`)"
            SQL = SQL & " SELECT " & vUsu.Codigo & ",0,nifdatos,razosoci,dirdatos,codposta,despobla,desprovi FROM cuentas where codmacta in (" & RC & ")"
            Conn.Execute SQL
    
    
    
            'Ahora meto los datos de la empresa
            cad = "INSERT INTO Usuarios.z347carta (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir,"
            cad = cad & "contacto) VALUES ("
            cad = cad & vUsu.Codigo
                
                
            'Monta Datos Empresa
            RS.Open "empresa2", Conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
            If RS.EOF Then
                MsgBox "Error en los datos de la empresa " & vEmpresa.nomempre
                RC = ",'','','','','',''"  '6 campos
            Else
                RC = DBLet(RS!siglasvia) & " " & DBLet(RS!Direccion) & "  " & DBLet(RS!numero) & ", " & DBLet(RS!puerta)
                RC = ",'" & DBLet(RS!nifempre) & "','" & vEmpresa.nomempre & "','" & RC & "','"
                RC = RC & DBLet(RS!codpos) & "','" & DBLet(RS!Poblacion) & "','" & DBLet(RS!provincia) & "','" & DBLet(RS!contacto) & "')"
            End If
            RS.Close
            cad = cad & RC
            Conn.Execute cad
            
            
            
    
        End If
        GeneraDatosTalPagDesglosado = True
    Else
        'I>0
        MsgBox "No hay datos", vbExclamation
    End If

EGeneraDatosTalPag:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Function



Private Function GeneraDatosTalPagSinDesglose() As Boolean
    On Error GoTo EGeneraDatosTalPag
    GeneraDatosTalPagSinDesglose = False
    
    

       
       
    SQL = "select scarecepdoc.*,nommacta from scarecepdoc,cuentas "
    SQL = SQL & " where  scarecepdoc.codmacta=cuentas.codmacta"
    If Text3(24).Text <> "" Then SQL = SQL & " AND fecharec >= '" & Format(Text3(24).Text, FormatoFecha) & "'"
    If Text3(25).Text <> "" Then SQL = SQL & " AND fecharec <= '" & Format(Text3(25).Text, FormatoFecha) & "'"
    'Contabilizado
    'SQL = SQL & " AND Contabilizada =  1"
    'Si esta llevada a banco o no
    'SQL = SQL & " AND LlevadoBanco = " & Abs(chkLstTalPag(2).Value)
    If cboListPagare.ListIndex >= 1 Then SQL = SQL & " AND LlevadoBanco = " & Abs(cboListPagare.ListIndex = 1)
    
    I = -1
    If (chkLstTalPag(0).Value = 1) Xor (chkLstTalPag(1).Value = 1) Then
        'Solo uno seleccionado
        I = 0
        If (chkLstTalPag(0).Value = 1) Then I = 1
            
    Else
        If (chkLstTalPag(0).Value = 0) Then
            MsgBox "Seleccione talon o pagare", vbExclamation
            Exit Function
        End If
    End If
    If I >= 0 Then SQL = SQL & " AND talon = " & I
    'Si ID
    If txtnumfac(2).Text <> "" Then SQL = SQL & " AND codigo >= " & txtnumfac(2).Text
    If txtnumfac(3).Text <> "" Then SQL = SQL & " AND codigo <= " & txtnumfac(3).Text



    Set RS = New ADODB.Recordset
    
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    RC = ""
    While Not RS.EOF
        I = I + 1
        'ztesoreriacomun (`codusu`,`codigo
        SQL = ", (" & vUsu.Codigo & "," & I & ","
        
        'texto1`,`texto2`,`texto3`,y el 4
        SQL = SQL & "'" & DevNombreSQL(RS!numeroref) & "','" & DevNombreSQL(RS!Banco) & "','"
        SQL = SQL & DevNombreSQL(RS!codmacta) & "','" & DevNombreSQL(RS!Nommacta) & "','"
        
        
        '5 Serie y numero factura
        SQL = SQL & "',"
        '`importe1`
        'SQL = SQL & TransformaComasPuntos(CStr(Rs.Fields(5))) & ","   'La columna 5 es sli.importe
        SQL = SQL & TransformaComasPuntos(CStr(RS!Importe)) & ","
        
        '
        '`fecha1`,`fecha2`,`fecha3`
        SQL = SQL & "'" & Format(RS!fecharec, FormatoFecha) & "',"
        SQL = SQL & "'" & Format(RS!fechavto, FormatoFecha) & "',"
        SQL = SQL & "'" & Format(Now, FormatoFecha) & "')"
    
        RC = RC & SQL
        
        RS.MoveNext
    Wend
    RS.Close
    
    If I > 0 Then
        RC = Mid(RC, 3) 'QUITO LA PRIMERA COMA
        'OK hay datos. Insertamos
        SQL = "INSERT INTO usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,`texto2`,`texto3`,"
        SQL = SQL & "`texto4`,`texto5`,`importe1`,`fecha1`,`fecha2`,`fecha3`) VALUES "
        SQL = SQL & RC
        Conn.Execute SQL
        GeneraDatosTalPagSinDesglose = True
    Else
        MsgBox "No hay datos", vbExclamation
    End If
    
    

EGeneraDatosTalPag:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Function





Private Function ListadoOrdenPago() As Boolean
    On Error GoTo EListadoOrdenPago
    ListadoOrdenPago = False

    'Borramos
    cad = "DELETE from usuarios.zlistadopagos WHERE codusu = " & vUsu.Codigo
    Conn.Execute cad
    Set miRsAux = New ADODB.Recordset
    'Inserttamos
    RC = ""
    If txtCtaBanc(3).Text <> "" Or txtCtaBanc(4).Text <> "" Then
        If txtCtaBanc(3).Text <> "" Then RC = " codmacta >= '" & txtCtaBanc(3).Text & "'"
        
        If txtCtaBanc(4).Text <> "" Then
            If RC <> "" Then RC = RC & " AND "
            RC = RC & " codmacta <= '" & txtCtaBanc(4).Text & "'"
        End If
        cad = "Select codmacta from ctabancaria where " & RC
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        RC = ""
        While Not miRsAux.EOF
            RC = RC & ", '" & miRsAux!codmacta & "'"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If RC = "" Then
            MsgBox "Ningún banco con esos valores", vbExclamation
            Exit Function
        End If
           
        RC = Mid(RC, 2)
    End If
    
    
    SQL = ""
    If Text3(26).Text <> "" Then SQL = SQL & " AND fecefect >= '" & Format(Text3(26).Text, FormatoFecha) & "'"
    If Text3(27).Text <> "" Then SQL = SQL & " AND fecefect <= '" & Format(Text3(27).Text, FormatoFecha) & "'"
    If RC <> "" Then SQL = SQL & " AND ctabanc1 in (" & RC & ")"
    
    
    'JULIO2013
    'La variable contdocu grabaremos emitdocum, y en el listado sabremos si es de talon/pagere
    'para poder separalos
    'Antes: contdocu, ahora emitdocum
    
    'Agosto 2014
    'Tipo de pago
    cad = "select " & vUsu.Codigo & ",`ctaprove`,`numfactu`,`fecfactu`,`numorden`,`fecefect`,`impefect`-coalesce(imppagad,0),"
    cad = cad & " `ctabanc1`,`ctabanc2`,`emitdocum`,spagop.entidad,spagop.oficina,spagop.CC,spagop.cuentaba,"
    cad = cad & " nommacta,'error','error',descformapago "
    
    cad = cad & " from spagop,cuentas,sforpa,stipoformapago "
    cad = cad & " WHERE spagop.ctaprove = cuentas.codmacta AND spagop.codforpa=sforpa.codforpa and tipoformapago=tipforpa"
    'Ponemos un check si salen negativos o no
    RC = " AND impefect >=0"
    If Me.chkPagBanco(0).Value = 1 And Me.chkPagBanco(1).Value = 1 Then RC = "" 'Salen todos
    cad = cad & RC 'todos o solo positivos
    cad = cad & SQL
    
    SQL = "INSERT INTO usuarios.zlistadopagos (`codusu`,`ctaprove`,`numfactu`,`fecfactu`,`numorden`,`fecefect`,"
    SQL = SQL & " `impefect`,`ctabanc1`,`ctabanc2`,`contdocu`,`entidad`,`oficina`,`CC`,`cuentaba`,"
    SQL = SQL & " `nomprove`,`nombanco`,`cuentabanco`,TipoForpa) " & cad
    Conn.Execute SQL
    
    cad = DevuelveDesdeBD("count(*)", "usuarios.zlistadopagos", "codusu", vUsu.Codigo)
    If Val(cad) = 0 Then
        MsgBox "Ningun vencimiento con esos valores", vbExclamation
        Exit Function
    End If
    
    'Actualizo los datos de los bancos `nombanco`,`cuentabanco`
    cad = "Select ctabanc1 from usuarios.zlistadopagos WHERE codusu = " & vUsu.Codigo
    cad = cad & " AND ctabanc1 <>'' GROUP BY ctabanc1"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    While Not miRsAux.EOF
        cad = cad & miRsAux!ctabanc1 & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    While cad <> ""
        I = InStr(1, cad, "|")
        If I = 0 Then
            cad = ""
        Else
            RC = Mid(cad, 1, I - 1)
            cad = Mid(cad, I + 1)
            
            SQL = "Select ctabancaria.codmacta,ctabancaria.entidad, ctabancaria.oficina, ctabancaria.control, ctabancaria.ctabanco,"
            SQL = SQL & " ctabancaria.descripcion,nommacta from  ctabancaria,cuentas where ctabancaria.codmacta=cuentas.codmacta "
            SQL = SQL & " AND ctabancaria.codmacta ='" & RC & "'"
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                SQL = "Cuenta banco erronea: " & vbCrLf & "Hay vencimientos asociados a la cuenta " & RC & " que no esta en bancos"
                MsgBox SQL, vbExclamation
            Else
                SQL = DBLet(miRsAux!Descripcion, "T")
                If SQL = "" Then SQL = miRsAux!Nommacta
                SQL = DevNombreSQL(SQL) & "|"
                
                'Enti8dad...
                I = DBLet(miRsAux!Entidad, "0")
                SQL = SQL & Format(I, "0000")
                                'Oficina...
                I = DBLet(miRsAux!Oficina, "0")
                SQL = SQL & Format(I, "0000")
                                'CC...
                RC = DBLet(miRsAux!Control, "T")
                If RC = "" Then RC = "**"
                SQL = SQL & RC
                'cuenta
                RC = DBLet(miRsAux!CtaBanco, "T")
                If RC = "" Then RC = "    **"
                SQL = SQL & RC & "|"
                
                
                RC = "UPDATE usuarios.zlistadopagos set `nombanco`='" & RecuperaValor(SQL, 1)
                RC = RC & "',`cuentabanco`='" & RecuperaValor(SQL, 2) & "'"
                RC = RC & " WHERE ctabanc1 = '" & miRsAux!codmacta & "' AND codusu = " & vUsu.Codigo
                Conn.Execute RC
                
            End If
            miRsAux.Close
        End If
    Wend
    
    ListadoOrdenPago = True
    Set miRsAux = Nothing
    Exit Function
EListadoOrdenPago:
    MuestraError Err.Number, "ListadoOrdenPago"
End Function



Private Function ListadoReclamas() As Boolean

On Error GoTo EListadoReclamas

    ListadoReclamas = False
        

    SQL = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = ""
    cad = ""
    
    If Text3(28).Text <> "" Or Text3(29).Text <> "" Then
        RC = DesdeHasta("F", 28, 29, "F.Reclama")
        If RC <> "" Then cad = cad & " " & RC
            
        RC = CampoABD(Text3(28), "F", "fecreclama", True)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
        RC = CampoABD(Text3(29), "F", "fecreclama", False)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
    End If
    
    
    If txtCta(15).Text <> "" Or txtCta(16).Text <> "" Then
        RC = DesdeHasta("C", 15, 16, "Cta")
        If RC <> "" Then cad = cad & " " & RC
            
        RC = CampoABD(txtCta(15), "T", "codmacta", True)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
        RC = CampoABD(txtCta(16), "T", "codmacta", False)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
    End If
    If SQL <> "" Then SQL = " WHERE " & SQL
    SQL = "Select * from shcocob" & SQL
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    RC = "insert into usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,`texto2`,`"
    RC = RC & "texto3`,`texto4`,`texto5`,`texto6`,`importe1`,`importe2`,`fecha1`,`fecha2`,"
    RC = RC & "`fecha3`,`texto`,`observa2`,`opcion`) VALUES "
    SQL = ""
    While Not RS.EOF
        NumRegElim = NumRegElim + 1
        SQL = SQL & ", (" & vUsu.Codigo & "," & NumRegElim & ",'" & RS!codmacta & "','"
        'text 2 y 3
        SQL = SQL & DevNombreSQL(RS!Nommacta) & "','" & RS!NUmSerie & Format(RS!codfaccl, "000000") & "','"
        '4 y 5
        SQL = SQL & RS!numorden & "','"
        If Val(RS!carta) = 1 Then
            SQL = SQL & "Email"
        ElseIf Val(RS!carta) = 2 Then
            SQL = SQL & "Teléfono"
        Else
            SQL = SQL & "Carta"
        End If
        'Text6, importe 1 y 2
        SQL = SQL & "',''," & TransformaComasPuntos(CStr(RS!ImpVenci)) & ",NULL,"
        'Fec1 reclama fec2 factra   fec3
        SQL = SQL & "'" & Format(RS!Fecreclama, FormatoFecha) & "','" & Format(RS!fecfaccl, FormatoFecha) & "',NULL,"
        DevfrmCCtas = Memo_Leer(RS!observaciones)
        If DevfrmCCtas = "" Then
            DevfrmCCtas = "NULL"
        Else
            DevfrmCCtas = "'" & DevNombreSQL(DevfrmCCtas) & "'"
        End If
        SQL = SQL & DevfrmCCtas & ",NULL,0)"


        'Siguiente
        RS.MoveNext
        
        
        If Len(SQL) > 100000 Then
            SQL = Mid(SQL, 2) 'QUITO LA COMA
            SQL = RC & SQL
            Conn.Execute SQL
            SQL = ""
        End If
            
        
    Wend
    RS.Close
        If SQL <> "" Then
            SQL = Mid(SQL, 2) 'QUITO LA COMA
            SQL = RC & SQL
            Conn.Execute SQL
        End If
        
        
    If NumRegElim > 0 Then
        ListadoReclamas = True
    Else
        MsgBox "Ningun dato devuelto", vbExclamation
    End If
    Exit Function
EListadoReclamas:
    MuestraError Err.Number, Err.Description
End Function





'******************************************************************************************
'
'   Listado gastos fijos

Private Function ListadoGastosFijos() As Boolean

On Error GoTo EListadoGF

    ListadoGastosFijos = False
        

    SQL = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = ""
    cad = ""
    
    
    DevfrmCCtas = "" ' ON del left join , NO al WHERE
    If Text3(30).Text <> "" Or Text3(31).Text <> "" Then
        RC = DesdeHasta("F", 30, 31, "Fecha")
        If RC <> "" Then cad = cad & " " & Trim(RC)
            
        RC = CampoABD(Text3(30), "F", "fecha", True)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
        RC = CampoABD(Text3(31), "F", "fecha", False)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
    End If
    DevfrmCCtas = SQL
    SQL = ""
    
    'Este si que va al where
    If txtGastoFijo(0).Text <> "" Or txtGastoFijo(1).Text <> "" Then
        RC = DesdeHasta("GF", 0, 1, "Codigo")
        If RC <> "" Then
            If cad <> "" Then
                'Ya esta la fecha
                If Len(cad & RC) > 100 Then cad = cad & """ + chr(13) + """
            End If
            cad = cad & " " & Trim(RC)
        End If
            
        RC = CampoABD(txtGastoFijo(0), "N", "sgastfij.codigo", True)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
        RC = CampoABD(txtGastoFijo(1), "N", "sgastfij.codigo", False)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
    End If
    
   
   
    RC = " FROM sgastfij left join sgastfijd ON sgastfij.Codigo = sgastfijd.Codigo"
    If DevfrmCCtas <> "" Then RC = RC & " AND " & DevfrmCCtas
    If SQL <> "" Then RC = RC & " WHERE " & SQL
    SQL = "SELECT sgastfij.codigo,descripcion,ctaprevista,fecha,importe" & RC
    
    

    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    RC = "insert into usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,`texto2`,`"
    RC = RC & "texto3`,`importe1`,`fecha1`) VALUES "
    SQL = ""
    While Not RS.EOF
        NumRegElim = NumRegElim + 1
        SQL = SQL & ", (" & vUsu.Codigo & "," & NumRegElim & ",'" & Format(RS!Codigo, "00000") & "','"
        'text 2 y 3
        SQL = SQL & DevNombreSQL(RS!Descripcion) & "','" & RS!Ctaprevista & "',"
       
  
        'Detalla
        If IsNull(RS!Fecha) Then
            SQL = SQL & "0,'" & Format(Now, FormatoFecha) & "'"
        Else
            SQL = SQL & TransformaComasPuntos(DBLet(RS!Importe, "N")) & ",'" & Format(RS!Fecha, FormatoFecha) & "'"
        End If
        SQL = SQL & ")"
        
        'Siguiente
        RS.MoveNext
            
        
    Wend
    RS.Close
    If SQL <> "" Then
        SQL = Mid(SQL, 2) 'QUITO LA COMA
        SQL = RC & SQL
        Conn.Execute SQL
    End If
        
        
    If NumRegElim = 0 Then
        MsgBox "Ningun dato devuelto", vbExclamation
        Exit Function
    End If
    
    
    'Updateo la cuenta bancaria
    RC = "Select texto3 from usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo & " GROUP BY 1"
    RS.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not RS.EOF
        SQL = SQL & RS!texto3 & "|"
        RS.MoveNext
    Wend
    RS.Close
    
    While SQL <> ""
        NumRegElim = InStr(1, SQL, "|")
        If NumRegElim = 0 Then
            SQL = ""
        Else
            RC = Mid(SQL, 1, NumRegElim - 1)
            SQL = Mid(SQL, NumRegElim + 1)
            
            RC = "Select codmacta,nommacta from cuentas where codmacta='" & RC & "'"
            RS.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                RC = "UPDATE usuarios.ztesoreriacomun SET texto4='" & DevNombreSQL(RS!Nommacta) & "' WHERE codusu =" & vUsu.Codigo & " AND texto3='" & RS!codmacta & "'"
                Conn.Execute RC
            End If
            RS.Close
        End If
    Wend
    ListadoGastosFijos = True
    Exit Function
EListadoGF:
    MuestraError Err.Number, Err.Description
End Function






'Listadoas aseguradoas
Private Function AvisosAseguradora() As Boolean



    On Error GoTo EListAsegEfectos
    AvisosAseguradora = False
    
    cad = "DELETE FROM Usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    'feccomunica,fecprorroga,fecsiniestro
    SQL = ""
    If Me.optAsegAvisos(0).Value Then
        cad = "feccomunica"
    ElseIf Me.optAsegAvisos(1).Value Then
        cad = "fecprorroga"
    Else
        cad = "fecsiniestro"
    End If
    RC = CampoABD(Text3(21), "F", cad, True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(22), "F", cad, False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    'Significa que no ha puesto fechas
    If InStr(1, SQL, cad) = 0 Then SQL = SQL & " AND " & cad & ">='1900-01-01'"
    
    'ORDENACION
    If Me.optAsegAvisos(0).Value Then
        RC = "feccomunica"
    ElseIf Me.optAsegAvisos(1).Value Then
        RC = "fecprorroga"
    Else
        RC = "fecsiniestro"
    End If
    
    cad = "Select numserie,codfaccl,numorden,fecvenci,impvenci,impcobro,gastos,fecfaccl,devuelto,scobro.codmacta,nommacta,numpoliz"
    cad = cad & ",credicon," & RC & " LaFecha" 'alias
    cad = cad & "  FROM scobro,cuentas,sforpa where scobro.codmacta= cuentas.codmacta AND numpoliz<>"""""
    cad = cad & " and scobro.codforpa=sforpa.codforpa "
    If SQL <> "" Then cad = cad & SQL
    
    
    
    

    cad = cad & " ORDER BY " & RC & ","
    'ORDENACION 2º nivel
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    cad = cad & RC
    
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    'Seran:                                                     codmac,nomma,credicon,numfac,fecfac,faviso,fvto,impvto,disponible,vencida
    cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,fecha1,fecha2,fecha3,importe1,importe2,opcion) VALUES "
    RC = ""
    
    While Not miRsAux.EOF
        If Len(RC) > 500 Then
            RC = Mid(RC, 2)
            Conn.Execute cad & RC
            RC = ""
        End If
        CONT = CONT + 1
        SQL = ", (" & vUsu.Codigo & "," & CONT & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
        SQL = SQL & DevNombreSQL(miRsAux!numpoliz) & "'"
        SQL = SQL & ",'" & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000000") & "',"  'texto4
        'Fecha fac
        SQL = SQL & CampoBD_A_SQL(miRsAux!fecfaccl, "F", True) & ","
        'Fecha aviso
        SQL = SQL & CampoBD_A_SQL(miRsAux!lafecha, "F", True) & ","
        'Fecha vto
        SQL = SQL & CampoBD_A_SQL(miRsAux!FecVenci, "F", True)
        
        SQL = SQL & "," & TransformaComasPuntos(CStr(miRsAux!ImpVenci))
        SQL = SQL & "," & TransformaComasPuntos(CStr(DBLet(miRsAux!Gastos, "N")))
        'Devuelto
        SQL = SQL & "," & DBLet(miRsAux!Devuelto, "N") & ")"
    
        RC = RC & SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If RC <> "" Then
        RC = Mid(RC, 2)
        Conn.Execute cad & RC
    End If
    
    
    If CONT > 0 Then
        AvisosAseguradora = True
    Else
        MsgBox "Ningun datos con esos valores", vbExclamation
    End If
    Exit Function
EListAsegEfectos:
    MuestraError Err.Number, "Avisos aseguradoras"
End Function



'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'
'       Compensaciones Cliente. Abonos vs Cobros
'
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

Private Sub PonerVtosCompensacionCliente()
Dim IT


    lwCompenCli.ListItems.Clear
    Me.txtimpNoEdit(0).Tag = 0
    Me.txtimpNoEdit(1).Tag = 0
    Me.txtimpNoEdit(0).Text = ""
    Me.txtimpNoEdit(1).Text = ""
    If Me.txtCta(17).Text = "" Then Exit Sub
    Set Me.lwCompenCli.SmallIcons = frmPpal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    cad = "Select scobro.*,nomforpa from scobro,sforpa where scobro.codforpa=sforpa.codforpa "
    cad = cad & " AND codmacta = '" & Me.txtCta(17).Text & "'"
    cad = cad & " AND (transfer =0 or transfer is null) and codrem is null"
    cad = cad & " and estacaja=0 and recedocu=0"
    cad = cad & " ORDER BY fecvenci"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwCompenCli.ListItems.Add()
        IT.Text = miRsAux!NUmSerie
        IT.SubItems(1) = Format(miRsAux!codfaccl, "000000")
        IT.SubItems(2) = Format(miRsAux!fecfaccl, "dd/mm/yyyy")
        IT.SubItems(3) = miRsAux!numorden
        IT.SubItems(4) = miRsAux!FecVenci
        IT.SubItems(5) = miRsAux!nomforpa
    
        Importe = DBLet(miRsAux!Gastos, "N")
        Importe = Importe + miRsAux!ImpVenci
        
        'Si ya he cobrado algo
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        
        If Importe > 0 Then
            IT.SubItems(6) = Format(Importe, FormatoImporte)
            IT.SubItems(7) = " "
        Else
            IT.SubItems(6) = " "
            IT.SubItems(7) = Format(-Importe, FormatoImporte)
        End If
        IT.Tag = Abs(Importe)  'siempre valor absoluto
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



Private Sub RealizarCompensacionAbonosClientes()
Dim Borras As Boolean
    
    If BloqueoManual(True, "COMPEABONO", "1") Then

        cad = DevuelveDesdeBD("max(codigo)", "scompenclicab", "1", "1")
        If cad = "" Then cad = "0"
        CONT = Val(cad) + 1 'ID de la operacion
        
        cad = "INSERT INTO scompenclicab(codigo,fecha,login,PC,codmacta,nommacta) VALUES (" & CONT
        cad = cad & ",now(),'" & DevNombreSQL(vUsu.Login) & "','" & DevNombreSQL(vUsu.PC)
        cad = cad & "','" & txtCta(17).Text & "','" & DevNombreSQL(DtxtCta(17).Text) & "')"
        
        Set miRsAux = New ADODB.Recordset
        Borras = True
        If Ejecuta(cad) Then
            
            Borras = Not RealizarProcesoCompensacionAbonos
        
        End If


        Set miRsAux = Nothing
        If Borras Then
            Conn.Execute "DELETE FROM scompenclilin WHERE codigo = " & CONT
            Conn.Execute "DELETE FROM scompenclicab WHERE codigo = " & CONT
            
        End If

        'Desbloquamos proceso
        BloqueoManual False, "COMPEABONO", ""
        DevfrmCCtas = ""
        
        PonerVtosCompensacionCliente   'Volvemos a cargar los vencimientos
        
        'El nombre del report
        CadenaDesdeOtroForm = Me.Tag
        Me.Tag = ""
        If Not Borras Then
            ImprimiCompensacion CONT
            
        
        End If
        
        Set miRsAux = Nothing
    Else
        MsgBox "Proceso bloqueado", vbExclamation
    End If

End Sub



Private Sub ImprimiCompensacion(CodigoCompensacion As Long)

    On Error GoTo EImprimiCompensacion
        
        'CadenaDesdeOtroForm:  lleva el nombre del report
        
        
        'Ha ido bien. Imprimiremos la hoja por si quiere crear PDF
        Conn.Execute "DELETE FROM Usuarios.ztmpfaclin WHERE codusu =" & vUsu.Codigo
        Conn.Execute "Delete from Usuarios.zTesoreriaComun where codusu = " & vUsu.Codigo
        Conn.Execute "Delete from Usuarios.z347carta where codusu = " & vUsu.Codigo
        
        'insert into `ztmpfaclin` (`codusu`,`codigo`,`Numfac`,`Fecha`,`cta`,`Cliente`,`NIF`,
        '`Imponible`,`IVA`,`ImpIVA`,`Total`,`retencion`,`TipoIva`)
        SQL = "INSERT INTO usuarios.ztmpfaclin(`codusu`,`codigo`,`Numfac`,`Fecha`,`cta`,`Cliente`,`NIF`,`Imponible`,`ImpIVA`,`retencion`,`Total`,`IVA`,TipoIva)"
        SQL = SQL & "select " & vUsu.Codigo & ",@rownum:=@rownum+1 AS rownum,"
        SQL = SQL & "concat(numserie,right(concat(""000000"",codfaccl),8)) fecha,date_format(fecfaccl,'%d/%m/%Y') ffaccl,"
        SQL = SQL & "scompenclilin.codmacta,if (nommacta is null,nomclien,nommacta) nomcli,"
        SQL = SQL & "date_format(fecvenci,'%d/%m/%Y') venci,impvenci,gastos,impcobro,"
        SQL = SQL & "impvenci + coalesce(gastos,0) + coalesce(impcobro,0)  tot_al"
        SQL = SQL & ",if(fecultco is null,null,date_format(fecultco,'%d/&m')) fecco ,destino"
        SQL = SQL & " From (scompenclilin left join cuentas on scompenclilin.codmacta=cuentas.codmacta)"
        SQL = SQL & ",(SELECT @rownum:=0) r WHERE codigo=" & CONT & " order by destino desc,numserie,codfaccl"
        Conn.Execute SQL
            
        
            
        
   
    
    
        
    
    
    
    
        'Datos carta
        'Datos basicos de la empresa para la carta
        cad = "INSERT INTO Usuarios.z347carta (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir, "
        cad = cad & "parrafo1, parrafo2, contacto, despedida,saludos,parrafo3, parrafo4, parrafo5, Asunto, Referencia)"
        cad = cad & " VALUES (" & vUsu.Codigo & ", "
        
        'Estos datos ya veremos com, y cuadno los relleno
        Set miRsAux = New ADODB.Recordset
        SQL = "select nifempre,siglasvia,direccion,numero,escalera,piso,puerta,codpos,poblacion,provincia,contacto from empresa2"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'Paarafo1 Parrafo2 contacto
        SQL = "'','',''"
        'sql= "'1234567890A','Ariadna Software ','Franco Tormo 3, Bajo Izda','46007','Valencia'"
        SQL = "'##########','" & vEmpresa.nomempre & "','#############','######','##########','##########'," & SQL
        If Not miRsAux.EOF Then
            SQL = ""
            For I = 1 To 6
                SQL = SQL & DBLet(miRsAux.Fields(I), "T") & " "
            Next I
            SQL = Trim(SQL)
            SQL = "'" & DBLet(miRsAux!nifempre, "T") & "','" & DevNombreSQL(vEmpresa.nomempre) & "','" & DevNombreSQL(SQL) & "'"
            SQL = SQL & ",'" & DBLet(miRsAux!codpos, "T") & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "'"

            'Contaccto
            SQL = SQL & ",NULL,NULL,'" & DevNombreSQL(DBLet(miRsAux!contacto)) & "' "
        End If
        miRsAux.Close
      
        cad = cad & SQL
        cad = cad & ",NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
        
        Conn.Execute cad
        
        
        'Datos CLIENTE
         ', texto3, texto4, texto5,texto6
        cad = DevuelveDesdeBD("codmacta", "scompenclicab", "codigo", CStr(CONT))
        cad = "Select nommacta,razosoci,dirdatos,codposta,despobla,desprovi from cuentas where codmacta ='" & cad & "'"
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        'NO PUEDE SER EOF
        cad = miRsAux!Nommacta
        If Not IsNull(miRsAux!razosoci) Then cad = miRsAux!razosoci
        cad = "'" & DevNombreSQL(cad) & "'"
        'Direccion
        cad = cad & ",'" & DevNombreSQL(CStr(DBLet(miRsAux!dirdatos))) & "'"
        'Poblacion
        SQL = DBLet(miRsAux!codposta)
        If SQL <> "" Then SQL = SQL & " - "
        SQL = SQL & DevNombreSQL(CStr(DBLet(miRsAux!desPobla)))
        cad = cad & ",'" & SQL & "'"
        'Provincia
        cad = cad & ",'" & DevNombreSQL(CStr(DBLet(miRsAux!desProvi))) & "'"
        miRsAux.Close
        

        
        SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2, texto3, texto4,texto5,texto6, observa1, "
        SQL = SQL & "importe1, importe2, fecha1, fecha2, fecha3, observa2, opcion)"
        SQL = SQL & " VALUES (" & vUsu.Codigo & ",1,'',''," & cad
        
        'select Numfac,fecha from usuarios.ztmpfaclin where tipoiva=1 and codusu=2200
        Importe = 0
        RC = "NIF"   'RC = "fecha"   La fecha de VTo esta en el campo: NIF
        cad = DevuelveDesdeBD("numfac", "Usuarios.ztmpfaclin", "codusu =" & vUsu.Codigo & " AND tipoiva", "1", "N", RC)
        If cad = "" Then
            'Significa que la compesacion ha sido total, no quedaba resultante
            
        Else
            cad = "Quedando el resultado en el vencimiento: " & cad & " de " & Format(RC, "dd/mm/yyyy")
            Importe = 1
        End If
        
        If Importe > 0 Then
            RC = "select sum(impvenci + coalesce(gastos,0) + coalesce(impcobro,0)) from  scompenclilin  WHERE codigo=" & CONT
            miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            RC = "0"
            If Not miRsAux.EOF Then Importe = DBLet(miRsAux.Fields(0), "N")
            miRsAux.Close
        Else
            RC = "0"
        End If
        
        'observa 1 texto 6 e importe1
        SQL = SQL & ",'" & cad & "'," & TransformaComasPuntos(CStr(Importe))
        
        
        'importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion
        For I = 1 To 6
            SQL = SQL & ",NULL"
        Next
        SQL = SQL & ")"
        Conn.Execute SQL
        
        With frmImprimir
                .OtrosParametros = ""
                .NumeroParametros = 0
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                
                .Opcion = 91
                .Show vbModal
            End With



Exit Sub
EImprimiCompensacion:
    MuestraError Err.Number, Err.Description
End Sub

Private Function RealizarProcesoCompensacionAbonos() As Boolean
Dim Destino As Byte
Dim J As Integer

    'NO USAR CONT

    RealizarProcesoCompensacionAbonos = False











    'Vamos a seleccionar los vtos
    '(numserie,codfaccl,fecfaccl,numorden)
    'EN SQL
    SQLVtosSeleccionadosCompensacion NumRegElim, False    'todos  -> Numregelim tendr el destino
    
    'Metemos los campos en el la tabla de lineas
    ' Esto guarda el valor en CAD
    FijaCadenaSQLCobrosCompen
    
    
    'Texto compensacion
    DevfrmCCtas = ""
    
    RC = "Select " & cad & " FROM scobro where (numserie,codfaccl,fecfaccl,numorden) IN (" & SQL & ")"
    miRsAux.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "Error. EOF vencimientos devueltos ", vbExclamation
        Exit Function
    End If
    
    
    I = 0
    
    While Not miRsAux.EOF
        I = I + 1
        BACKUP_Tabla miRsAux, RC
        'Quito los parentesis
        RC = Mid(RC, 1, Len(RC) - 1)
        RC = Mid(RC, 2)
        
        Destino = 0
        If miRsAux!NUmSerie = Me.lwCompenCli.ListItems(NumRegElim).Text Then
            If miRsAux!codfaccl = Val(Me.lwCompenCli.ListItems(NumRegElim).SubItems(1)) Then
                If Format(miRsAux!fecfaccl, "dd/mm/yyyy") = Me.lwCompenCli.ListItems(NumRegElim).SubItems(2) Then
                    If miRsAux!numorden = Val(Me.lwCompenCli.ListItems(NumRegElim).SubItems(3)) Then Destino = 1
                End If
            End If
        End If
        
        RC = "INSERT INTO scompenclilin (codigo,linea,destino," & cad & ") VALUES (" & CONT & "," & I & "," & Destino & "," & RC & ")"
        Conn.Execute RC
        
        'Para las observaciones de despues
        Importe = DBLet(miRsAux!Gastos, "N")
        Importe = Importe + miRsAux!ImpVenci
        'Si ya he cobrado algo
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        
        If Destino = 0 Then 'El destino
            DevfrmCCtas = DevfrmCCtas & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000") & " " & Format(miRsAux!fecfaccl, "dd/mm/yy")
            DevfrmCCtas = DevfrmCCtas & " Vto:" & Format(miRsAux!FecVenci, "dd/mm/yy") & " " & Importe
            DevfrmCCtas = DevfrmCCtas & "|"
        Else
            'El DESTINO siempre ira en la primera observacion del texto
            RC = "Importe anterior vto: " & Importe
            DevfrmCCtas = RC & "|" & DevfrmCCtas
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Acutalizaremos el VTO destino
    
    Conn.BeginTrans
        'BORRAREMOS LOS VENCIMIENTOS QUE NO SEAN DESTINO a no ser que el importe restante sea 0
        Destino = 1
        If txtimpNoEdit(0).Text = txtimpNoEdit(1).Text Then Destino = 0
        SQLVtosSeleccionadosCompensacion 0, Destino = 1  'sin o con el destino
        RC = "DELETE FROM scobro where (numserie,codfaccl,fecfaccl,numorden) IN (" & SQL & ")"
        
        'Para saber si ha ido bien
        Destino = 0    '0 mal,1 bien
        If Ejecuta(RC) Then
            If txtimpNoEdit(0).Text = txtimpNoEdit(1).Text Then
                Destino = 1
            Else
                'Updatearemos los campos csb del vto restante. A partir del segundo
                'La variable CadenaDesdeOtroForm  tiene los que vamos a actualizar
                
                cad = ""
                J = 0
                SQL = ""
                
                Do
                    I = InStr(1, DevfrmCCtas, "|")
                    If I = 0 Then
                        DevfrmCCtas = ""
                    Else
                        RC = Mid(DevfrmCCtas, 1, I - 1)
                        If Len(RC) > 40 Then RC = Mid(RC, 1, 40)
                        DevfrmCCtas = Mid(DevfrmCCtas, I + 1)
                        J = J + 1
                        'Antes desde aqui cogia el campo
                        'Ahora desde CadenaDesdeOtroForm que tiene los campos libres
                        'Cad = RecuperaValor("text41csb|text42csb|text43csb|text51csb|text52csb|text53csb|text61csb|text62csb|text63csb|text71csb|text72csb|text73csb|text81csb|text82csb|text83csb|", J)
                        cad = RecuperaValor(CadenaDesdeOtroForm, J)
                        SQL = SQL & ", " & cad & " = '" & DevNombreSQL(RC) & "'"
                
                    End If
                Loop Until DevfrmCCtas = ""
                Importe = CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag)  'txtimpNoEdit(1).Tag es negativo
                RC = "gastos=null, impcobro=null,fecultco=null,impvenci=" & TransformaComasPuntos(CStr(Importe))
                SQL = RC & SQL
                SQL = "UPDATE scobro SET " & SQL
                'WHERE
                RC = ""
                For J = 1 To Me.lwCompenCli.ListItems.Count
                    If Me.lwCompenCli.ListItems(J).Bold Then
                        'Este es el destino
                        RC = "NUmSerie = '" & Me.lwCompenCli.ListItems(J).Text
                        RC = RC & "' AND codfaccl = " & Val(Me.lwCompenCli.ListItems(J).SubItems(1))
                        RC = RC & " AND fecfaccl = '" & Format(Me.lwCompenCli.ListItems(J).SubItems(2), FormatoFecha)
                        RC = RC & "' AND numorden = " & Val(Me.lwCompenCli.ListItems(J).SubItems(3))
                        Exit For
                    End If
                Next
                If RC <> "" Then
                    cad = SQL & " WHERE " & RC
                    If Ejecuta(cad) Then Destino = 1
                Else
                    MsgBox "No encontrado destino", vbExclamation
                    
                End If
            End If
        End If
        If Destino = 1 Then
            Conn.CommitTrans
            RealizarProcesoCompensacionAbonos = True
        Else
            Conn.RollbackTrans
        End If
        
End Function

Private Sub SQLVtosSeleccionadosCompensacion(ByRef RegistroDestino As Long, SinDestino As Boolean)
Dim Insertar As Boolean
    SQL = ""
    For I = 1 To Me.lwCompenCli.ListItems.Count
        If Me.lwCompenCli.ListItems(I).Checked Then
        
            Insertar = True
            If Me.lwCompenCli.ListItems(I).Bold Then
                RegistroDestino = I
                If SinDestino Then Insertar = False
            End If
            If Insertar Then
                SQL = SQL & ", ('" & lwCompenCli.ListItems(I).Text & "'," & lwCompenCli.ListItems(I).SubItems(1)
                SQL = SQL & ",'" & Format(lwCompenCli.ListItems(I).SubItems(2), FormatoFecha) & "'," & lwCompenCli.ListItems(I).SubItems(3) & ")"
            End If
            
        End If
    Next
    SQL = Mid(SQL, 2)
            
End Sub


Private Sub FijaCadenaSQLCobrosCompen()

    cad = "NUmSerie , codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, ctabanc1,"
    cad = cad & "codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, emitdocum, "
    cad = cad & "recedocu, contdocu, text33csb, text41csb, text42csb, text43csb, text51csb, text52csb,"
    cad = cad & "text53csb, text61csb, text62csb, text63csb, text71csb, text72csb, text73csb, text81csb,"
    cad = cad & "text82csb, text83csb, ultimareclamacion, agente, departamento, tiporem, CodRem, AnyoRem,"
    cad = cad & "siturem, Gastos, Devuelto, situacionjuri, noremesar, obs, transfer, estacaja, referencia,"
    cad = cad & "reftalonpag, nomclien, domclien, pobclien, cpclien, proclien, referencia1, referencia2,"
    cad = cad & "feccomunica, fecprorroga, fecsiniestro"
    
End Sub


'******************************************************************************
'******************************************************************************
'
'******************************************************************************
'******************************************************************************



Private Function ComunicaDatosSeguro_() As Boolean
Dim k As Integer

    ComunicaDatosSeguro_ = False
    
   
    NumRegElim = 0
    
    For k = 1 To Me.ListView3.ListItems.Count
        If Me.ListView3.ListItems(k).Checked Then
            DatosSeguroUnaEmpresa CInt(ListView3.ListItems(k).Tag)
      
            SQL = DevuelveDesdeBD("count(*)", "Usuarios.ztesoreriacomun", "codusu", vUsu.Codigo)
            If SQL <> "" Then NumRegElim = Val(SQL)
        End If
    Next
    
    
    
    If NumRegElim > 0 Then
        SQL = "DELETE from Usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo
        SQL = SQL & " AND importe1<=0"
        
        
        
    
    
        '   Conn.Execute SQL
        SQL = DevuelveDesdeBD("count(*)", "Usuarios.ztesoreriacomun", "codusu", vUsu.Codigo)
        If SQL <> "" Then
            NumRegElim = Val(SQL)
        Else
            NumRegElim = 0
        End If
        
        
        ComunicaDatosSeguro_ = NumRegElim > 0
        If NumRegElim > 0 Then
            SQL = "Select texto5 from Usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo & " GROUP BY 1"
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            While Not miRsAux.EOF
                SQL = miRsAux!texto5
                If SQL = "" Then
                    SQL = "ESPAÑA"
                Else
                    If InStr(1, SQL, " ") > 0 Then
                        SQL = Mid(SQL, 3)
                    Else
                        SQL = "" 'no updateamos
                    End If
                End If
                If SQL <> "" Then
                    SQL = "UPDATE Usuarios.ztesoreriacomun set texto5='" & DevNombreSQL(SQL) & "' WHERE codusu ="
                    SQL = SQL & vUsu.Codigo & " AND texto5='" & DevNombreSQL(miRsAux!texto5) & "'"
                    Conn.Execute SQL
                End If
                miRsAux.MoveNext
            Wend
            miRsAux.Close
        End If
        
    End If
End Function

Private Sub DatosSeguroUnaEmpresa(NumConta As Integer)

    
    'select numpoliz,nifdatos,numserie,codfaccl,nommacta,impvenci,gastos,impcobro,credicon from scobro,cuentas where
    'scobro.codmacta = cuentas.codmacta     fecbajcre
    
    'JUlio2013
    'Para fontenas iran por PAIS
    'añadiremos en text05 el pais
    SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2,texto3,texto4,"
    SQL = SQL & " importe1,  importe2,texto5) "
    
'    RC = "select " & vUsu.Codigo & ",@rownum:=@rownum+1, numpoliz,nifdatos,concat(numserie,right(concat('0000000',codfaccl),8)),nommacta, "
'
'    RC = RC & "impvenci + if(gastos is null,0,gastos) - if( impcobro is null,0,impcobro) ,credicon"
'    RC = RC & "  from conta" & NumConta & ".scobro,conta" & NumConta & ".cuentas,(SELECT @rownum:=" & NumRegElim & ") r "
'    RC = RC & " WHERE scobro.codmacta=cuentas.codmacta  and numpoliz<>''  and "
'    RC = RC & " (fecbajcre  is null or fecbajcre>'" & Format(Text3(35).Text, FormatoFecha) & "')"
'
    
    'ENERO 2013.
    'Despues de hablar con BERNIA, en este listado salen
    RC = "select " & vUsu.Codigo & ",@rownum:=@rownum+1, numpoliz,nifdatos,concat(numserie,right(concat('0000000',codfaccl),8)),nommacta, "
    
    RC = RC & " totfaccl ,credicon,if(pais is null,'',pais)"    'JUL13 añadimos PAIS
    RC = RC & " from conta" & NumConta & ".cabfact,conta" & NumConta & ".cuentas,(SELECT @rownum:=" & NumRegElim & ") r "
    RC = RC & " WHERE cabfact.codmacta=cuentas.codmacta  and numpoliz<>''  and "
    
    RC = RC & " (fecbajcre  is null or fecbajcre>'" & Format(Text3(35).Text, FormatoFecha) & "')"
    
    'Contemplamos facturas desde la fecha de concesion
    RC = RC & " and fecfaccl>= fecconce"
    
    'D/H fecha factura
    If Me.Text3(34).Text <> "" Then RC = RC & " AND fecfaccl >='" & Format(Text3(34).Text, FormatoFecha) & "'"
    If Me.Text3(35).Text <> "" Then RC = RC & " AND fecfaccl <='" & Format(Text3(35).Text, FormatoFecha) & "'"
    
    
    
    
    
    
    SQL = SQL & RC
    Conn.Execute SQL
End Sub


Private Function GeneraDatosFrasAsegurados() As Boolean
Dim NumConta As Byte

    NumConta = CByte(vEmpresa.codempre)
    GeneraDatosFrasAsegurados = False

    SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2,texto3,texto4,"
    SQL = SQL & " importe1,  importe2,fecha1,fecha2) "
    
    RC = "select " & vUsu.Codigo & ",@rownum:=@rownum+1, numpoliz,nifdatos,concat(numserie,right(concat('0000000',codfaccl),8)),nommacta, "
    
    RC = RC & "impvenci + if(gastos is null,0,gastos) - if( impcobro is null,0,impcobro) ,if (credicon is null,0,credicon)"
    RC = RC & ",fecfaccl,fecvenci"
    RC = RC & "  from conta" & NumConta & ".scobro,conta" & NumConta & ".cuentas,(SELECT @rownum:=" & NumRegElim & ") r "
    RC = RC & " WHERE scobro.codmacta=cuentas.codmacta  "
    
    If Me.chkVarios(0).Value = 1 Then
        'SOLO asegudaros
        RC = RC & " and numpoliz<>''  and (fecbajcre  is null or fecbajcre>'" & Format(Text3(35).Text, FormatoFecha) & "')"
    End If
    'D/H fecha factura
    If Me.Text3(34).Text <> "" Then RC = RC & " AND fecfaccl >='" & Format(Text3(34).Text, FormatoFecha) & "'"
    If Me.Text3(35).Text <> "" Then RC = RC & " AND fecfaccl <='" & Format(Text3(35).Text, FormatoFecha) & "'"
    
    
    SQL = SQL & RC
    Conn.Execute SQL

    
    
    'Borramos importe cero

    SQL = "DELETE from Usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo
    SQL = SQL & " AND importe1<=0"
    Conn.Execute SQL
    
    SQL = DevuelveDesdeBD("count(*)", "Usuarios.ztesoreriacomun", "codusu", vUsu.Codigo)
    If SQL <> "" Then
        NumRegElim = Val(SQL)
    Else
        NumRegElim = 0
    End If
    GeneraDatosFrasAsegurados = NumRegElim > 0



End Function

'****************************************************************************************
'****************************************************************************************
'
'       NORMA 57
'
'****************************************************************************************
'****************************************************************************************
Private Function procesarficheronorma57() As Boolean
Dim Estado As Byte  '0  esperando cabcerea
                    '1  esperando pie (leyendo lineas)
    
    On Error GoTo eprocesarficheronorma57
    
    
    'insert into tmpconext(codusu,cta,fechaent,Pos)
    Conn.Execute "DELETE FROM tmpconext WHERE codusu = " & vUsu.Codigo
    procesarficheronorma57 = False
    I = FreeFile
    Open cd1.FileName For Input As #I
    SQL = ""
    Estado = 0
    Importe = 0
    TotalRegistros = 0
    While Not EOF(I)
            Line Input #I, SQL
            RC = Mid(SQL, 1, 4)
            Select Case Estado
            Case 0
                'Para saber que el fichero tiene el formato correcto
                If RC = "0270" Then
                        Estado = 1
                        'Voy a buscar si hay un banco
                        
                        RC = "select cuentas.codmacta,nommacta from ctabancaria,cuentas where ctabancaria.codmacta="
                        RC = RC & "cuentas.codmacta AND ctabancaria.entidad = " & Trim(Mid(SQL, 23, 4))
                        Set miRsAux = New ADODB.Recordset
                        miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        TotalRegistros = 0
                        While Not miRsAux.EOF
                            RC = miRsAux!codmacta & "|" & miRsAux!Nommacta & "|"
                            TotalRegistros = TotalRegistros + 1
                            miRsAux.MoveNext
                        Wend
                        miRsAux.Close
                        If TotalRegistros = 1 Then
                            Me.txtCtaBanc(5).Text = RecuperaValor(RC, 1)
                            Me.txtDescBanc(5).Text = RecuperaValor(RC, 2)
                        End If
                        TotalRegistros = 0
                End If
            Case 1
                If RC = "6070" Then
                    'Linea con recibo
                    'Ejemplo:
                    '   6070      46076147000130582263151014000000014067003059                      0000000516142
                    '                                  fecha       impot   socio                      fra      CC codigo de control del codigo de barra
                    'Fecha pago
                    RC = Mid(SQL, 31, 2) & "/" & Mid(SQL, 33, 2) & "/20" & Mid(SQL, 35, 2)
                    Fecha = CDate(RC)
                    'IMporte
                    RC = Mid(SQL, 37, 12)
                    cad = CStr(CCur(Val(RC) / 100))
                    'FRA
                    RC = Mid(SQL, 77, 11)
                    CONT = Val(RC)
                    'Socio
                    RC = Val(Mid(SQL, 50, 6))
                        
                    'Insertamos en tmp
                    TotalRegistros = TotalRegistros + 1
                    SQL = "INSERT INTO tmpconext(codusu,cta,fechaent,Pos,TimporteD,linliapu) VALUES (" & vUsu.Codigo & ",'"
                    SQL = SQL & RC & "','" & Format(Fecha, FormatoFecha) & "'," & CONT & "," & TransformaComasPuntos(cad) & "," & TotalRegistros & ")"
                    Conn.Execute SQL
                    
                    Importe = Importe + CCur(TransformaPuntosComas(cad))
                ElseIf RC = "8070" Then
                    'OK. Final de linea.
                    '
                    'Comprobacion BASICA
                    '8070      46076147000 000010        000000028440
                    '                       vtos-2           importe
                    
                    RC = ""
                    
                    'numero registros
                    cad = Val(Mid(SQL, 24, 5))
                    If Val(cad) = 0 Then
                        RC = RC = RC & vbCrLf & "Linea totales. Nº registros cero. " & SQL
                    Else
                        If Val(cad) - 2 <> TotalRegistros Then RC = "Contador de registros incorrecto"
                    End If
                    'Suma importes
                    cad = CStr(CCur(Mid(SQL, 37, 12) / 100))
                    
                    If CCur(cad) = 0 Then
                        RC = RC = RC & vbCrLf & "Linea totales. Suma importes cero. " & SQL
                    Else
                        If CCur(cad) <> Importe Then RC = RC & vbCrLf & "Suma importes incorrecta"
                    End If
                    
                    
                   
                    
                    
                    If RC <> "" Then
                        Err.Raise 513, , RC
                    Else
                        Estado = 2
                    End If
                End If
            End Select
    Wend
    Close #I
    I = 0 'para que no vuelva a cerrar el fichero
    
    If Estado < 2 Then
        'Errores procesando fichero
        If Estado = 0 Then
            SQL = "No se encuetra la linea de inicio de declarante(6070)"
        Else
            SQL = "No se encuetra la linea de totales(8070)"
        End If

        MsgBox "Error procesando el fichero." & vbCrLf & SQL, vbExclamation
    Else
        espera 0.5
        procesarficheronorma57 = True
    End If
eprocesarficheronorma57:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    If I > 0 Then Close #I
End Function


Private Function BuscarVtosNorma57() As Boolean

    BuscarVtosNorma57 = False
    
    Set miRsAux = New ADODB.Recordset
    
    'Dependiendo del parametro....
    If vParamT.Norma57 = 1 Then
        'ESCALONA.
        'Viene el socio y el numero de factura e importe.
        'Habra que buscar
        BuscarVtosNorma57 = VtosNorma57Escalona

    Else
        MsgBox "En desarrollo", vbExclamation
    End If
    
    Set miRsAux = Nothing
End Function

Private Function VtosNorma57Escalona() As Boolean
Dim RN As ADODB.Recordset
Dim Fin As Boolean
Dim NoEncontrado As Byte
Dim AlgunVtoNoEncontrado As Boolean
On Error GoTo eVtosNorma57Escalona
    
    VtosNorma57Escalona = False
    Set RN = New ADODB.Recordset
    SQL = "select * from tmpconext WHERE codusu =" & vUsu.Codigo & " order by cta,pos "
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    AlgunVtoNoEncontrado = False
    While Not miRsAux.EOF
        'Vto a vto
        'If miRsAux!Linliapu = 9 Then Stop
        RC = RellenaCodigoCuenta("430." & miRsAux!Cta)
        SQL = "Select * from scobro where codmacta = '" & RC & "' AND codfaccl =" & miRsAux!Pos & " and impvenci>0"
        RN.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        CONT = 0
        If RN.EOF Then
            cad = "NO encontrado"
            NoEncontrado = 2
        Else
            'OK encontrado.
            Fin = False
            I = 0
            NoEncontrado = 1
            cad = ""
            
            While Not Fin
            
                I = I + 1
                
                Norma57VencimientoEncontradoEsCorrecto RN, Fin
                
                If Not Fin Then
                    RN.MoveNext
                    If RN.EOF Then Fin = True
                End If
            Wend
        End If
        RN.Close
        SQL = "UPDATE tmpconext SET "
        If CONT = 1 Then
            'OK este es el vto
            'NO hacemos nada. Updateamos los campos de la tmp
            'para buscar despues
            'numdiari numorden       numdocum=fecfaccl     ccost numserie
            SQL = SQL & " nomdocum ='" & Format(Fecha, FormatoFecha)
            SQL = SQL & "', ccost ='" & DevfrmCCtas
            SQL = SQL & "', numdiari = " & I
            SQL = SQL & ", contra = '" & RC & "'"
        Else
            If I > 1 Then cad = "(+1) " & cad
            SQL = SQL & " numasien=  " & NoEncontrado  'para vtos no encontrados o erroneos
            SQL = SQL & ", ampconce ='" & DevNombreSQL(cad) & "'"
            If NoEncontrado = 2 Then AlgunVtoNoEncontrado = True
        End If
        SQL = SQL & " WHERE codusu =" & vUsu.Codigo & " AND linliapu = " & miRsAux!Linliapu
        Conn.Execute SQL
            
 
        
        'Sig
        miRsAux.MoveNext
    Wend
    
    miRsAux.Close
    
    
    
    If AlgunVtoNoEncontrado Then
        'Lo buscamos al reves
        espera 0.5
        SQL = "select * from  tmpconext  WHERE codusu =" & vUsu.Codigo & " AND numasien=2"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            'Miguel angel
            'Puede que en algunos recibos las posciones del fichero vengan cambiadas
            'Donde era la factura es la cta y al reves
            RC = RellenaCodigoCuenta("430." & miRsAux!Pos)
            SQL = "Select * from scobro where codmacta = '" & RC & "' AND codfaccl =" & Val(miRsAux!Cta) & " and impvenci>0"
            RN.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RN.EOF Then
        
                'OK encontrado.
                Fin = False
                CONT = 0
                Norma57VencimientoEncontradoEsCorrecto RN, Fin
                
                
            
            
                'OK este es el vto
                'NO hacemos nada. Updateamos los campos de la tmp
                'para buscar despues
                'numdiari numorden       numdocum=fecfaccl     ccost numserie
                If CONT = 1 Then
                    SQL = SQL & " nomdocum ='" & Format(Fecha, FormatoFecha)
                    SQL = SQL & "', ccost ='" & DevfrmCCtas
                    SQL = SQL & "', numdiari = " & I
                    SQL = SQL & ", contra = '" & RC & "'"
                    SQL = "UPDATE tmpconext SET "
                End If
            End If
            RN.Close
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
    End If
    
    
    
    
    VtosNorma57Escalona = True
eVtosNorma57Escalona:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description, "Buscando vtos Escalona"
    Set RN = Nothing
End Function


Private Sub Norma57VencimientoEncontradoEsCorrecto(ByRef Rss As ADODB.Recordset, ByRef Final As Boolean)
        
        'Ha encontrado el vencimiento. Falta ver si no esta en remesa....
        If Not IsNull(Rss!CodRem) Then
            cad = "En la remesa " & Rss!CodRem
        
        Else
            If Not IsNull(Rss!transfer) Then
                cad = "Transferencia " & Rss!transfer
            Else
                Importe = Rss!ImpVenci + DBLet(Rss!Gastos, "N") - DBLet(Rss!impcobro, "N")
                If Importe <> miRsAux!timported Then
                    'Importe distinto
                    'Veamos si es que esta
                    cad = "Importe distinto"
                Else
                    'OK. Misma factura, socio, importe. SAlimos ya poniendo ""
                    Fecha = Rss!fecfaccl
                    DevfrmCCtas = Rss!NUmSerie
                    I = Rss!numorden
                    cad = ""
                    Final = True
                    CONT = 1
                End If
            End If
        End If
End Sub

Private Sub CargaLWNorma57(Correctos As Boolean)
Dim IT As ListItem

    Set miRsAux = New ADODB.Recordset
    If Correctos Then
        SQL = "select tmpconext.*,nommacta from tmpconext left join cuentas on tmpconext.contra=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
        SQL = SQL & " and numasien=0 order by  ccost,pos  "
    Else
        SQL = "select * from tmpconext WHERE codusu = " & vUsu.Codigo & " and numasien > 0 order by cta,pos "
    End If
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If Correctos Then
            Set IT = Me.lwNorma57Importar(0).ListItems.Add(, "C" & Format(miRsAux!Linliapu, "0000"))
            IT.Text = miRsAux!CCost
            IT.SubItems(1) = miRsAux!Pos
            IT.SubItems(2) = Format(miRsAux!Nomdocum, "dd/mm/yyyy")
            IT.SubItems(3) = miRsAux!Linliapu
            If IsNull(miRsAux!Nommacta) Then
                SQL = "ERRROR GRAVE"
            Else
                SQL = miRsAux!Nommacta
            End If
            IT.SubItems(4) = SQL
            IT.SubItems(5) = Format(miRsAux!timported, FormatoImporte)
            IT.SubItems(6) = Format(miRsAux!FechaEnt, "dd/mm/yyyy")
            IT.Checked = True
        Else
            'ERRORES
            Set IT = Me.lwNorma57Importar(1).ListItems.Add(, "C" & Format(miRsAux!Linliapu, "0000"))
            IT.Text = miRsAux!Cta
            IT.SubItems(1) = miRsAux!Pos
            IT.SubItems(2) = Format(miRsAux!timported, FormatoImporte)
            IT.SubItems(3) = miRsAux!Ampconce
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
End Sub


Private Sub LimpiarDelProceso()
    lwNorma57Importar(0).ListItems.Clear
    lwNorma57Importar(1).ListItems.Clear
    Me.txtCtaBanc(5).Text = ""
    Me.txtDescBanc(5).Text = ""
End Sub



