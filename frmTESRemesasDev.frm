VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESRemesasDev 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "1"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   15690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6660
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameDevlucionRe 
      Height          =   9195
      Left            =   60
      TabIndex        =   11
      Top             =   60
      Width           =   15315
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   13290
         TabIndex        =   36
         Tag             =   "Importe|N|N|||reclama|importes|||"
         Top             =   4710
         Width           =   1815
      End
      Begin VB.Frame FrameConcepto 
         Caption         =   "Datos Contabilizaci�n"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   270
         TabIndex        =   22
         Top             =   1800
         Width           =   14895
         Begin VB.ComboBox Combo2 
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
            ItemData        =   "frmTESRemesasDev.frx":0000
            Left            =   9450
            List            =   "frmTESRemesasDev.frx":0013
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Tag             =   "Ampliacion debe/CLIENTES|N|N|0||stipoformapago|ampdecli|||"
            Top             =   1500
            Width           =   2820
         End
         Begin VB.TextBox txtImporte 
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
            Left            =   4950
            TabIndex        =   4
            Text            =   "Text4"
            Top             =   900
            Width           =   1125
         End
         Begin VB.TextBox Text1 
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
            Index           =   11
            Left            =   9450
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   600
            Width           =   1275
         End
         Begin VB.OptionButton optDevRem 
            Caption         =   "� x Recibo"
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
            Left            =   780
            TabIndex        =   29
            Top             =   810
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.OptionButton optDevRem 
            Caption         =   "% x Recibo"
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
            Left            =   780
            TabIndex        =   28
            Top             =   1170
            Width           =   2205
         End
         Begin VB.TextBox txtImporte 
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
            Left            =   5280
            TabIndex        =   27
            Text            =   "Text4"
            Top             =   1470
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.OptionButton optDevRem 
            Caption         =   "% x  rec. con m�nimo"
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
            Left            =   780
            TabIndex        =   26
            Top             =   1545
            Width           =   2535
         End
         Begin VB.CheckBox chkDevolRemesa2 
            Caption         =   "Importe gastos banco"
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
            Left            =   9450
            TabIndex        =   7
            Top             =   2070
            Width           =   2805
         End
         Begin VB.TextBox txtDConcpeto 
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
            Left            =   10050
            TabIndex        =   25
            Text            =   "Text9"
            Top             =   1050
            Width           =   4725
         End
         Begin VB.TextBox txtConcepto 
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
            Left            =   9450
            TabIndex        =   6
            Text            =   "Text10"
            Top             =   1050
            Width           =   525
         End
         Begin VB.CheckBox chkAgrupadevol2 
            Caption         =   "Agrupa apunte banco"
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
            Left            =   11970
            TabIndex        =   24
            Top             =   600
            Width           =   2475
         End
         Begin VB.TextBox txtImporte 
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
            Left            =   12480
            TabIndex        =   8
            Text            =   "Text4"
            Top             =   2040
            Width           =   1125
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "C�culo Gastos Devoluci�n Cliente"
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
            Height          =   240
            Index           =   0
            Left            =   390
            TabIndex        =   34
            Top             =   390
            Width           =   3630
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Gastos (�)"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   3
            Left            =   3720
            TabIndex        =   33
            Top             =   930
            Width           =   1050
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   11
            Left            =   9180
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Devoluci�n"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   4
            Left            =   7380
            TabIndex        =   32
            Top             =   600
            Width           =   1740
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Imp.Minimo (�)"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   8
            Left            =   3720
            TabIndex        =   31
            Top             =   1515
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.Image imgConcepto 
            Height          =   240
            Index           =   1
            Left            =   9180
            Top             =   1110
            Width           =   240
         End
         Begin VB.Label Label7 
            Caption         =   "Concepto"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   345
            Index           =   9
            Left            =   7380
            TabIndex        =   30
            Top             =   1110
            Width           =   1590
         End
         Begin VB.Label lblAsiento 
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
            Left            =   2550
            TabIndex        =   23
            Top             =   1440
            Width           =   4095
         End
      End
      Begin VB.Frame FrameDevDesdeRemesa 
         Height          =   1185
         Left            =   270
         TabIndex        =   18
         Top             =   540
         Width           =   3585
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
            Index           =   5
            Left            =   990
            TabIndex        =   0
            Text            =   "Text3"
            Top             =   570
            Width           =   915
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
            Index           =   6
            Left            =   2430
            TabIndex        =   1
            Text            =   "Text3"
            Top             =   570
            Width           =   915
         End
         Begin VB.Image imgRem 
            Height          =   240
            Index           =   1
            Left            =   1050
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Remesa"
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
            Height          =   240
            Index           =   5
            Left            =   120
            TabIndex        =   21
            Top             =   210
            Width           =   885
         End
         Begin VB.Label Label6 
            Caption         =   "C�digo"
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
            Left            =   240
            TabIndex        =   20
            Top             =   585
            Width           =   705
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "A�o"
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
            Left            =   1830
            TabIndex        =   19
            Top             =   585
            Width           =   555
         End
      End
      Begin VB.Frame FrameDevDesdeVto 
         Height          =   1215
         Left            =   3990
         TabIndex        =   15
         Top             =   540
         Width           =   5565
         Begin VB.TextBox txtDCtaNormal 
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
            Index           =   11
            Left            =   1560
            TabIndex        =   16
            Text            =   "Text9"
            Top             =   570
            Width           =   3885
         End
         Begin VB.TextBox txtCtaNormal 
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
            Index           =   11
            Left            =   180
            TabIndex        =   2
            Text            =   "Text9"
            Top             =   570
            Width           =   1335
         End
         Begin VB.Image imgCtaNorma 
            Height          =   240
            Index           =   11
            Left            =   1050
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label2 
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
            Index           =   37
            Left            =   180
            TabIndex        =   17
            Top             =   210
            Width           =   825
         End
      End
      Begin VB.Frame FrameDevDesdeFichero 
         Height          =   1215
         Left            =   9630
         TabIndex        =   13
         Top             =   540
         Width           =   5535
         Begin VB.TextBox Text8 
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
            Left            =   150
            TabIndex        =   3
            Text            =   "Text8"
            Top             =   570
            Width           =   5295
         End
         Begin VB.Image Image4 
            Height          =   240
            Left            =   960
            Top             =   210
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fichero"
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
            Height          =   240
            Index           =   10
            Left            =   120
            TabIndex        =   14
            Top             =   210
            UseMnemonic     =   0   'False
            Width           =   780
         End
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   9
         Left            =   13920
         TabIndex        =   10
         Top             =   8730
         Width           =   1215
      End
      Begin VB.CommandButton cmdDevolRem 
         Caption         =   "Devolucion"
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
         Left            =   12510
         TabIndex        =   9
         Top             =   8730
         Width           =   1335
      End
      Begin MSComctlLib.ListView lwCobros 
         Height          =   3915
         Left            =   300
         TabIndex        =   35
         Top             =   4710
         Width           =   12825
         _ExtentX        =   22622
         _ExtentY        =   6906
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   72
         Left            =   13260
         TabIndex        =   37
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   12420
         Picture         =   "frmTESRemesasDev.frx":0087
         ToolTipText     =   "Quitar al Debe"
         Top             =   4440
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   12780
         Picture         =   "frmTESRemesasDev.frx":01D1
         ToolTipText     =   "Puntear al Debe"
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "DEVOLUCION REMESA"
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
         Index           =   3
         Left            =   5100
         TabIndex        =   12
         Top             =   210
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmTESRemesasDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Opcion As Byte
    '
    '9.- Devolucion remesa
        
    '16.- Devolucion remesa desde fichero del banco
    
    '28 .- Devolucion remesa desde un vto
    
    
    
    
Public SubTipo As Byte

    'Para la opcion 22
    '   Remesas cancelacion cliente.
    '       1:  Efectos
    '       2: Talones pagares
    
'Febrero 2010
'Cuando pago proveedores con un talon, y le he indicado el numero
Public NumeroDocumento As String
    
    
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBa As frmBanco
Attribute frmBa.VB_VarHelpID = -1
Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
Private WithEvents frmP As frmFormaPago
Attribute frmP.VB_VarHelpID = -1
Private WithEvents frmRe As frmTESRemesas
Attribute frmRe.VB_VarHelpID = -1
Private WithEvents frmB As frmBasico
Attribute frmB.VB_VarHelpID = -1


Dim RS As ADODB.Recordset
Dim SQL As String
Dim i As Integer
Dim IT As ListItem  'Comun
Dim PrimeraVez As Boolean
Dim Cancelado As Boolean
Dim CuentasCC As String
Dim ImporteQueda As Currency

Dim vRemesa As String
Dim ValoresDevolucionRemesa As String
Dim ImporteRemesa As Currency
Dim vSQL As String


Private Sub chkDevolRemesa2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 21 Or Index = 25 Or Index = 31 Then CadenaDesdeOtroForm = "" 'ME garantizo =""
    If Index = 31 Then
        If MsgBox("�Cancelar el proceso?", vbQuestion + vbYesNo) = vbYes Then SubTipo = 0
    End If
    Unload Me
End Sub



Private Sub cmdDevolRem_Click()
Dim Importe As Currency
Dim GastoDevolGral As Currency
Dim CadenaVencimiento As String
Dim MultiRemesaDevuelta As String
Dim TipoFicheroDevolucion As Byte

    MultiRemesaDevuelta = ""
    CadenaVencimiento = ""
    
    If Text8.Text <> "" Then Opcion = 16
    If Text3(5).Text <> "" Then Opcion = 9
    If txtCtaNormal(11).Text <> "" Then Opcion = 28
    
    
    
    If Opcion = 16 Then
        'DESDE FICHERO
        Text8.Text = Trim(Text8.Text)
        If Text8.Text = "" Then Exit Sub
        If Dir(Text8.Text, vbArchive) = "" Then
            MsgBox "El fichero: " & Text8.Text & "    NO existe", vbExclamation
            Exit Sub
        End If
        Text3(5).Text = ""
        Text3(6).Text = ""
        
        'Si que existe el fichero
        TipoFicheroDevolucion = EsFicheroDevolucionSEPA2(Text8.Text)
        If TipoFicheroDevolucion > 0 Then
            If TipoFicheroDevolucion = 2 Then
                'SEPA xml
                ProcesaFicheroDevolucionSEPA_XML Text8, SQL
            Else
                ProcesaCabeceraFicheroDevolucionSEPA Text8, SQL
            End If
        Else
            'Texto normal
            ProcesaCabeceraFicheroDevolucion Text8.Text, SQL
        End If
        If SQL = "" Then Exit Sub
        
        
    
        
        MultiRemesaDevuelta = SQL
        Text3(5).Text = RecuperaValor(SQL, 1)
        Text3(6).Text = RecuperaValor(SQL, 2)
        
    End If
    If Opcion = 28 Then
        
        'Desde el Vto
        Set RS = New ADODB.Recordset
        
        SQL = ""
        If Me.txtCtaNormal(11).Text <> "" Then SQL = SQL & " AND codmacta='" & Me.txtCtaNormal(11).Text & "'"
        SQL = Mid(SQL, 5)
        
        
        SQL = "Select codrem,anyorem,NUmSerie,codfaccl,numorden from scobro where " & SQL
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RS.EOF Then
            SQL = "Ninguna pertence a ninguna remesa "
            MsgBox SQL, vbExclamation
            RS.Close
            Exit Sub
        End If
        Text3(5).Text = DBLet(RS!CodRem, "T")
        Text3(6).Text = DBLet(RS!AnyoRem, "T")
        CadenaVencimiento = RS!NUmSerie & "|" & RS!codfaccl & "|" & RS!numorden & "|"
        RS.Close
        Set RS = Nothing
    End If
    
    
    SQL = ""
    If Text3(5).Text = "" Or Text3(6).Text = "" Then
        If Opcion = 9 Then
            SQL = "Ponga una remesa."
        Else
            SQL = "ERROR leyendo remesa"
        End If
    Else
        If Not IsNumeric(Text3(5).Text) Or Not IsNumeric(Text3(6).Text) Then SQL = "La remesa debe ser num�rica"
    End If
    
    If Text1(11).Text = "" Then SQL = "Ponga la fecha de abono"
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    'Fecha pertenece a ejercicios contbles
    If FechaCorrecta2(CDate(Text1(11).Text), True) > 1 Then Exit Sub
    
    
    If txtImporte(1).Text = "" Then
        MsgBox "Indique el gasto por recibo", vbExclamation
        Exit Sub
    End If
    '
    If Me.optDevRem(2).Value Then
        If (txtImporte(2).Text = "") Then
            MsgBox "Debe poner valores del  minimo", vbExclamation
            Exit Sub
        End If
        
    End If
    
    If txtImporte(1).Text <> "" Then
        'Hay gravamen por gastos
        'Bloqueariamos la opcion de modificar esa remesa
        Importe = TextoAimporte(txtImporte(1).Text)
        If Me.optDevRem(1).Value Or Me.optDevRem(2).Value Then
            'Porcentual. No puede ser superior al 100%
            If Importe > 100 Then
                MsgBox "Importe no puede ser superior al 100%", vbExclamation
                Exit Sub
            End If
        End If
        
    Else
        Importe = 0
    End If
    
    'Comprobamos los conceptos y ampliaciones
    SQL = ""
    If txtConcepto(1).Text <> "" Then
        If txtDConcpeto(1).Text = "" Then SQL = "Concepto"
    End If
    
    
    If SQL = "" Then
        If Combo2(0).ListIndex = -1 Then SQL = "Ampliacion concepto incorrecta"
    End If
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    'Nuevo Noviembre 2009
    GastoDevolGral = 0
    If Me.chkDevolRemesa2.Value = 1 Then
        'Ha puesto gasto devolucion pero NO indica el gasto
        GastoDevolGral = TextoAimporte(txtImporte(5).Text)
        If GastoDevolGral = 0 Then
            MsgBox "Ha marcado contabilizar gasto y no lo ha indicado", vbExclamation
            Exit Sub
        End If
    
    Else
        If txtImporte(5).Text <> "" Then
            MsgBox "Ha indicado el gasto pero no ha marcado contabilizarlo", vbExclamation
            Exit Sub
        End If
    End If
    'Ahora miramos la remesa. En que sitaucion , y de que tipo es
    SQL = "Select * from remesas where codigo =" & Text3(5).Text
    SQL = SQL & " AND anyo =" & Text3(6).Text
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        SQL = "Ninguna remesa con esos valores."
        If Opcion = 16 Then SQL = SQL & "  Remesa: " & Text3(5).Text & " / " & Text3(6).Text
        MsgBox SQL, vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Sub
    End If
    
    
    'Tiene valor
    If RS!Situacion = "A" Then
        MsgBox "Remesa abierta. Sin llevar al banco.", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Sub
    End If
    
    
    
    If Asc(RS!Situacion) < Asc("Q") Then
        MsgBox "Remesa sin contabilizar.", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Sub
    End If
    
    
    
    
    SQL = RS!Codigo & "|" & RS!Anyo & "|" & RS!codmacta & "|" & Text1(11).Text & "|"
    
    
    Importe = TextoAimporte(txtImporte(1).Text)   ''Levara el gasto por recibo
    If Me.optDevRem(1).Value Or Me.optDevRem(2).Value Then SQL = SQL & "%"
    SQL = SQL & "|"
    If Me.optDevRem(2).Value Then SQL = SQL & TextoAimporte(txtImporte(2).Text)
    SQL = SQL & "|"
    
    
    'SQL llevara hasta ahora
    '        remes    cta ban  fec contb tipo gasto el 1: si tiene valor es el minimo por recibo
    ' Ej:    1|2009|572000005|20/11/2009|%|1|
    
    
    'Si contabilizamos el gasto, o pro contra vendra como factura bancaria desde otro lugar(norma34 p.e.)
    If GastoDevolGral = 0 Then
        'NO HAY GASTO
        SQL = SQL & "0|"
    Else
        SQL = SQL & CStr(GastoDevolGral) & "|"
        If Me.chkDevolRemesa2.Value = 1 Then
            'Voy a contabi�izar los gastos.
            'Vere si tiene CC
            If vParam.autocoste Then
                If DevuelveDesdeBD("codccost", "ctabancaria", "codmacta", RS!codmacta, "T") = "" Then
                    MsgBox "Va a contabilizar los gastos pero no esta configurado el Centro de coste para el banco: " & RS!codmacta, vbExclamation
                    RS.Close
                    Set RS = Nothing
                    Exit Sub
                End If
            End If
        End If
    End If
    
    'Depues del gasto
    'A�adire el fichero, si es autmatico
    If Opcion = 16 Then SQL = SQL & Text8.Text
    SQL = SQL & "|"
    'Nov 2012. En las devoluciones puede ser que el fichero traiga mas de una devolucion
    If Opcion = 16 Then
        If Text8.Text <> "" Then
            'Tengo que subsituir | por #
            MultiRemesaDevuelta = Replace(MultiRemesaDevuelta, "|", "#")
            SQL = SQL & MultiRemesaDevuelta
        End If
    End If
    SQL = SQL & "|"
    

    
    'Cierro aqui
    RS.Close
    
    'Bloqueamos la devolucion
    BloqueoManual True, "Devolrem", vUsu.Codigo
    'Hacemos la devolucion
'--
'    frmRemesas.Opcion = 2
'    frmRemesas.vRemesa = SQL
'    frmRemesas.ImporteRemesa = Importe 'Utilizamos esta variable para indicar el gasto a cargar por recibo
    '28Marzo2007
    'Para la contabilizacion de la devolucion
    'Client
'++
    vRemesa = SQL
    ImporteRemesa = Importe
    
    
    SQL = txtConcepto(1).Text & "|" & Combo2(0).ListIndex & "|"
    'y el banco
    'Agrupa el apunte del banco
    SQL = SQL & Abs(chkAgrupadevol2.Value) & "|"
    
    
'--
'    frmRemesas.ValoresDevolucionRemesa = SQL
'    'Si es desde el vto, para que lo busque
'    frmRemesas.vSQL = CadenaVencimiento
'
'    frmRemesas.Show vbModal
'++
    vSQL = CadenaVencimiento
    
    DevolverRemesa


    'Desbloqueamos
    BloqueoManual False, "Devolrem", vUsu.Codigo

End Sub

Private Sub DevolverRemesa()
Dim Cad As String
Dim jj As Integer
Dim Aux As String

    Cad = ""
    For jj = 1 To Me.lwCobros.ListItems.Count
        If lwCobros.ListItems(jj).Checked Then
            Cad = Cad & "1"
        End If
    Next jj
    If Cad = "" Then
        MsgBox "Seleccione los efectos devueltos", vbExclamation
        Exit Sub
    End If
    Cad = Len(Cad) & " efecto(s)."
    
    'Llegado aqui hago la pregunta
    Cad = "Va a realizar la devoluci�n de " & Cad & vbCrLf
   '--
   'If InStr(1, Label6.Caption, ":") > 0 Then
   '    Cad = Cad & vbCrLf & "Importe total de la devoluci�n: "
   '    Cad = Cad & Mid(Label6.Caption, InStr(1, Label6.Caption, ":")) & "�" & vbCrLf & vbCrLf
   'End If
   '++ sustituido por esto
    If Text1(4).Text <> "" Then
        Cad = Cad & vbCrLf & "Importe total de la devoluci�n: "
        Cad = Cad & Text1(4).Text & "�" & vbCrLf & vbCrLf
    End If
    
    Aux = RecuperaValor(vRemesa, 5)
    If Aux = "%" Then
        Aux = "Porcentaje por recibo: " & ImporteRemesa & "%" & vbCrLf
        If RecuperaValor(vRemesa, 6) <> "" Then
            Aux = Aux & "Gasto m�nimo: " & RecuperaValor(vRemesa, 6) & " �" & vbCrLf
        End If
    Else
        Aux = "Gasto por recibo: " & ImporteRemesa & " �" & vbCrLf
    End If
    
    Cad = Cad & Aux & vbCrLf
    
    'Gasto tramitacion devolucion
    Aux = RecuperaValor(vRemesa, 7)
    If Aux <> "" Then
        Aux = "Gasto bancario : " & Aux & "�" & vbCrLf
        Cad = Cad & vbCrLf & Aux
    End If
    
    Cad = Cad & vbCrLf & "�Desea continuar?"
    If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    If Not RealizarDevolucion Then Exit Sub

    Unload Me

End Sub

Private Function RealizarDevolucion() As Boolean
Dim IncPorcentaje As Boolean
Dim Gasto As Currency
Dim Minimo As Currency
Dim Cad As String
Dim jj As Long

    RealizarDevolucion = False
    'Tipo de aumento del gasto de devolucion
    Cad = RecuperaValor(vRemesa, 5)
    If Cad = "%" Then
        'Porcentual
        IncPorcentaje = True
        Minimo = 0
        Cad = RecuperaValor(vRemesa, 6)
        If Cad <> "" Then Minimo = Cad
    Else
        IncPorcentaje = False
    End If
        
    
    vSQL = "DELETE FROM tmpfaclin WHERE codusu =" & vUsu.Codigo
    Conn.Execute vSQL
    '                                               numero        serie     vto
    vSQL = "INSERT INTO tmpfaclin (codusu, codigo, Numfac, Fecha, numserie, NIF,  "
    vSQL = vSQL & "Imponible,  ImpIVA,total,cta,cliente) VALUES (" & vUsu.Codigo & ","
    For jj = 1 To lwCobros.ListItems.Count
        If Me.lwCobros.ListItems(jj).Checked Then
                                        'cdofaccl
            Cad = jj & "," & Val(lwCobros.ListItems(jj).SubItems(1)) & ",'"
                                    'fecfaccl                                                   SERIE
            Cad = Cad & Format(lwCobros.ListItems(jj).Tag, FormatoFecha) & "','" & lwCobros.ListItems(jj).Text
                                    'numvencimiento numorden
            Cad = Cad & "'," & Val(lwCobros.ListItems(jj).SubItems(2)) & ","
            ImporteQueda = ImporteFormateado(lwCobros.ListItems(jj).SubItems(5))
            Cad = Cad & TransformaComasPuntos(CStr(ImporteQueda)) & ","
            
            'Calculo el gasto
            If IncPorcentaje Then
                'Importe = importe  * (importe * % )/100
                Gasto = Round((ImporteQueda * ImporteRemesa) / 100, 2)
                
                If Minimo > 0 Then If Gasto < Minimo Then Gasto = Minimo
                
                ImporteQueda = ImporteQueda + Gasto
            Else
                'Importe =importe + incremento
                Gasto = ImporteRemesa
                ImporteQueda = ImporteQueda + ImporteRemesa
            End If
            Cad = Cad & TransformaComasPuntos(CStr(Gasto)) & ","
            Cad = Cad & TransformaComasPuntos(CStr(ImporteQueda)) & ",'"
            'Cuenta cliente, y banco
            Cad = Cad & lwCobros.ListItems(jj).SubItems(3) & "','"
            Cad = Cad & RecuperaValor(vRemesa, 3) & "')"
            Cad = vSQL & Cad
            If Not Ejecuta(Cad) Then Exit Function
        End If
    Next jj
    
    
    'OK. Ya tengo grabada la temporal con los recibos que devuelvo. Ahora
    'hare:
    '       - generar un asiento con los datos k devuelvo
    '       - marcar los cobros como devueltos, a�adirle el gasto y insertar en la
    '           tabla de hco de devueltos
    
    jj = Val(RecuperaValor(vRemesa, 7))
    
    If jj = 0 Then
        'Como no se contabilizan los beneficios no hace falta que calcule nada
        Cad = ""
     Else
        'Vya obteneer la cuenta de gastos bancarios
        Cad = RecuperaValor(vRemesa, 3)  'cta contable
        Cad = DevuelveDesdeBD("ctagastos", "ctabancaria", "codmacta", Cad, "T")
        If Cad = "" Then
            'NO esta configurada
            'Veo si esta en parametros
            'ctabenbanc
            Cad = DevuelveDesdeBD("ctabenbanc", "paramtesor", "codigo", "1", "N")
        End If
        If Cad = "" Then
            MsgBox "No esta configurada la gastos  bancarios", vbExclamation
            Exit Function
        End If
    End If
    
    ValoresDevolucionRemesa = txtConcepto(1).Text & "|" & Combo2(0).ListIndex & "|"
    
    If RealizarDevolucionRemesa(CDate(RecuperaValor(vRemesa, 4)), jj > 0, Cad, vRemesa, ValoresDevolucionRemesa) Then
        RealizarDevolucion = True
        Screen.MousePointer = vbHourglass
'        frmActualizar2.OpcionActualizar = 20
'        frmActualizar2.Show vbModal
        Screen.MousePointer = vbDefault
    End If
End Function




Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
    Limpiar Me
    PrimeraVez = True
    Me.Icon = frmPpal.Icon
    
    
    'Cago los iconos
    CargaImagenesAyudas Me.imgCtaNorma, 1, "Seleccionar cuenta"
    CargaImagenesAyudas Me.Image1, 2
    CargaImagenesAyudas imgRem, 1, "Seleccionar remesa"
    CargaImagenesAyudas imgConcepto, 1, "Concepto"
    
    Me.Image4.Picture = frmPpal.imgIcoForms.ListImages(1).Picture


'    FrameDevlucionRe.Visible = False
'    FrameDevDesdeVto.Visible = False
    
    Select Case Opcion
    Case 9, 16, 28
        If SubTipo = 1 Then
            Caption = "EFECTOS"
        Else
            Caption = "TALONES / PAGARES"
        End If
        FrameDevlucionRe.Visible = True
'        FrameDevDesdeFichero.Visible = Opcion = 16
'        Me.FrameDevDesdeVto.Visible = Opcion = 28
        Caption = "Devolucion remesa (" & UCase(Caption) & ")"
        W = FrameDevlucionRe.Width
        H = FrameDevlucionRe.Height
        Text1(11).Text = Format(Now, "dd/mm/yyyy")
        txtImporte(1).Text = 0
    
        'Nuevo 28Marzo2007
'        PonerValoresPorDefectoDevilucionRemesa
        
    End Select
    
    CargaList
    
    If NumeroDocumento <> "" Then
        Text3(5).Text = RecuperaValor(NumeroDocumento, 1)
        Text3(6).Text = RecuperaValor(NumeroDocumento, 2)
        Text3_LostFocus (5)
        PonFoco Text3(5)
    End If
    
    
    Me.Height = H + 360
    Me.Width = W + 90
'--
'    h = Opcion
'    If Opcion = 7 Then h = 6
'    If Opcion = 14 Then h = 13
'    If Opcion = 16 Or Opcion = 28 Then h = 9
'    If Opcion = 22 Or Opcion = 23 Then h = 8
'    Me.cmdCancelar(h).Cancel = True
    
End Sub

Private Sub CargaList()
    lwCobros.ColumnHeaders.Clear

    lwCobros.ColumnHeaders.Add , , "Serie", 720
    lwCobros.ColumnHeaders.Add , , "Factura", 1440
    lwCobros.ColumnHeaders.Add , , "Vto", 850
    lwCobros.ColumnHeaders.Add , , "Cuenta", 1500
    lwCobros.ColumnHeaders.Add , , "Cliente", 4500
    lwCobros.ColumnHeaders.Add , , "Importe", 1950, 1
    lwCobros.ColumnHeaders.Add , , "FechaOrden", 0
    lwCobros.ColumnHeaders.Add , , "ImporteOrden", 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set RS = Nothing
    Set miRsAux = Nothing
    
    NumeroDocumento = "" 'Para reestrablecerlo siempre
End Sub



Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(Image1(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQL = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub Image1_Click(Index As Integer)
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text1(Index).Text <> "" Then frmC.Fecha = CDate(Text1(Index).Text)
    Image1(0).Tag = Index
    frmC.Show vbModal
    Set frmC = Nothing
    If Text1(Index).Text <> "" Then PonerFoco Text1(Index)
End Sub


Private Sub PonerFoco(ByRef o As Object)
    On Error Resume Next
    o.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub ObtenerFoco(ByRef T As TextBox)
    T.SelStart = 0
    T.SelLength = Len(T.Text)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub
Private Sub Image4_Click()
    SQL = ""
    cd1.ShowOpen
    If cd1.FileName <> "" Then SQL = cd1.FileName
    If SQL <> "" Then
        If Dir(SQL, vbArchive) = "" Then
            MsgBox "Fichero NO existe", vbExclamation
            SQL = ""
        End If
    End If
    If SQL <> "" Then Text8.Text = SQL
End Sub


Private Sub imgCheck_Click(Index As Integer)

    If Index < 2 Then
        'Selecciona forma de pago
        For i = 1 To Me.lwCobros.ListItems.Count
            Me.lwCobros.ListItems(i).Checked = Index = 1
        Next
    End If
End Sub

Private Sub imgConcepto_Click(Index As Integer)
  
    LanzaBuscaGrid 1
    If SQL <> "" Then
        txtConcepto(Index).Text = SQL
        txtConcepto_LostFocus Index
    End If
End Sub

Private Sub imgCtaNorma_Click(Index As Integer)

        If Index <> 6 Then

               Set frmCCtas = New frmColCtas
               SQL = ""
               frmCCtas.DatosADevolverBusqueda = "0"
               frmCCtas.Show vbModal
               
               Set frmCCtas = Nothing
               If SQL <> "" Then
                   txtCtaNormal(Index).Text = SQL
                   txtCtaNormal_LostFocus Index
               End If
            
        End If
            
            
End Sub



Private Sub imgRem_Click(Index As Integer)
    i = Index
    Set frmRe = New frmTESRemesas
    frmRe.Tipo = SubTipo  'Para abrir efectos o talonesypagares
    frmRe.DatosADevolverBusqueda = "1|"
    frmRe.Show vbModal
    Set frmRe = Nothing
    'Por si ha puesto los datos
    CamposRemesaAbono
    
End Sub

Private Sub lwCobros_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i As Currency
    Set lwCobros.SelectedItem = Item
    i = ImporteFormateado(Item.SubItems(5))
    If Not Item.Checked Then i = -1 * i
    ImporteQueda = ImporteQueda + i
    Text1(4).Text = Format(ImporteQueda, FormatoImporte)
End Sub

Private Sub optDevRem_Click(Index As Integer)
        txtImporte(2).Visible = Index = 2
        Label4(8).Visible = Index = 2
        If Index <> 0 Then
            Label4(3).Caption = "% aplicado"
        Else
            Label4(3).Caption = "Gastos (�)"
        End If
End Sub

Private Sub optDevRem_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).Text = "" Then Exit Sub
    
    If Not EsFechaOK(Text1(Index)) Then
        MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
        Text1(Index).Text = ""
        PonerFoco Text1(Index)
    End If
    
End Sub



Private Sub Text3_GotFocus(Index As Integer)
    With Text3(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Text3_LostFocus(Index As Integer)
    With Text3(Index)
        .Text = Trim(.Text)
        If .Text = "" Then Exit Sub
        
        If Not IsNumeric(.Text) Then
            MsgBox "Campo debe ser num�rico: " & .Text, vbExclamation
            .Text = ""
            PonerFoco Text3(Index)
        Else
            Opcion = 9
            LimpiarLin Me, "FrameDevDesdeFichero"
            LimpiarLin Me, "FrameDevDesdeVto"
            
            If Text3(5).Text <> "" And Text3(6).Text <> "" Then
                SQL = " from cobros, formapago where cobros.codforpa = formapago.codforpa and codrem= " & DBSet(Text3(5).Text, "N") & " and anyorem = " & DBSet(Text3(6).Text, "N")
                CargaDevolucion 'PonerVtosRemesa SQL, False
            End If
        End If
        
        'Para que vaya a la tabla y traiga datos remesa
'        If Index = 3 Or Index = 4 Then CamposRemesaAbono
    End With
End Sub


Private Sub CargarValores()
Dim Importe As Currency
Dim GastoDevolGral As Currency
Dim CadenaVencimiento As String
Dim MultiRemesaDevuelta As String
Dim TipoFicheroDevolucion As Byte
    
    MultiRemesaDevuelta = ""
'    CadenaVencimiento = ""
    Select Case Opcion
        Case 9
            SQL = "Select * from remesas where codigo =" & Text3(5).Text
            SQL = SQL & " AND anyo =" & Text3(6).Text
            SQL = SQL & " AND situacion = 'Q'"
        
        Case 16
            'DESDE FICHERO
            Text8.Text = Trim(Text8.Text)
            If Text8.Text = "" Then Exit Sub
            If Dir(Text8.Text, vbArchive) = "" Then
                MsgBox "El fichero: " & Text8.Text & "    NO existe", vbExclamation
                Exit Sub
            End If
            Text3(5).Text = ""
            Text3(6).Text = ""
            
            'Si que existe el fichero
            TipoFicheroDevolucion = EsFicheroDevolucionSEPA2(Text8.Text)
            If TipoFicheroDevolucion > 0 Then
                If TipoFicheroDevolucion = 2 Then
                    'SEPA xml
                    ProcesaFicheroDevolucionSEPA_XML Text8, SQL
                Else
                    ProcesaCabeceraFicheroDevolucionSEPA Text8, SQL
                End If
            Else
                'Texto normal
                ProcesaCabeceraFicheroDevolucion Text8.Text, SQL
            End If
            If SQL = "" Then Exit Sub
            
            MultiRemesaDevuelta = SQL
'            Text3(5).Text = RecuperaValor(SQL, 1)
'            Text3(6).Text = RecuperaValor(SQL, 2)
            
        Case 28
            'Desde la cuenta
            Set RS = New ADODB.Recordset
            
            SQL = "situacion = 'Q' "
            If Me.txtCtaNormal(11).Text <> "" Then SQL = SQL & " AND codmacta='" & Me.txtCtaNormal(11).Text & "'"
            SQL = Mid(SQL, 5)
            
            SQL = "Select codrem,anyorem,NUmSerie,numfactu,numorden from cobros where " & SQL
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If RS.EOF Then
                SQL = "Ninguna pertenece a ninguna remesa "
                MsgBox SQL, vbExclamation
                RS.Close
                Exit Sub
            End If
'            Text3(5).Text = DBLet(RS!CodRem, "T")
'            Text3(6).Text = DBLet(RS!AnyoRem, "T")
            CadenaVencimiento = RS!NUmSerie & "|" & RS!codfaccl & "|" & RS!numorden & "|"
            RS.Close
            Set RS = Nothing
            
    End Select
    
    
    
    
'****************
    
    Select Case Opcion
        Case 9
            SQL = "Select * from remesas where codigo =" & Text3(5).Text
            SQL = SQL & " AND anyo =" & Text3(6).Text
            SQL = SQL & " AND situacion = 'Q'"
        Case 28
            SQL = "Select distinct remesas.* from remesas where situacion = 'Q' "
            SQL = SQL & " and (codigo, anyo) in (select codrem, anyorem from cobros where codmacta = " & DBSet(Me.txtCtaNormal(11).Text, "T") & ")"
    End Select
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then

        SQL = RS!Codigo & "|" & RS!Anyo & "|" & RS!codmacta & "|" & Text1(11).Text & "|"
        
        
'        Importe = TextoAimporte(txtImporte(1).Text)   ''Levara el gasto por recibo
        If Me.optDevRem(1).Value Or Me.optDevRem(2).Value Then SQL = SQL & "%"
        SQL = SQL & "|"
        If Me.optDevRem(2).Value Then SQL = SQL & TextoAimporte(txtImporte(2).Text)
        SQL = SQL & "|"
        
        
        'SQL llevara hasta ahora
        '        remes    cta ban  fec contb tipo gasto el 1: si tiene valor es el minimo por recibo
        ' Ej:    1|2009|572000005|20/11/2009|%|1|
        
        
        'Si contabilizamos el gasto, o pro contra vendra como factura bancaria desde otro lugar(norma34 p.e.)
        If GastoDevolGral = 0 Then
            'NO HAY GASTO
            SQL = SQL & "0|"
        Else
            SQL = SQL & CStr(GastoDevolGral) & "|"
            If Me.chkDevolRemesa2.Value = 1 Then
                'Voy a contabi�izar los gastos.
                'Vere si tiene CC
                If vParam.autocoste Then
                    If DevuelveDesdeBD("codccost", "ctabancaria", "codmacta", RS!codmacta, "T") = "" Then
                        MsgBox "Va a contabilizar los gastos pero no esta configurado el Centro de coste para el banco: " & RS!codmacta, vbExclamation
                        RS.Close
                        Set RS = Nothing
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        'Depues del gasto
        'A�adire el fichero, si es autmatico
        If Opcion = 16 Then SQL = SQL & Text8.Text
        SQL = SQL & "|"
        'Nov 2012. En las devoluciones puede ser que el fichero traiga mas de una devolucion
        If Opcion = 16 Then
            If Text8.Text <> "" Then
                'Tengo que subsituir | por #
                MultiRemesaDevuelta = Replace(MultiRemesaDevuelta, "|", "#")
                SQL = SQL & MultiRemesaDevuelta
            End If
        End If
        SQL = SQL & "|"
        
        vRemesa = SQL
    End If
    
    


End Sub




Private Sub CargaDevolucion()
Dim Itm As ListItem
Dim Col As Collection
Dim EfectoSerie As String
Dim EfectoFra As String
Dim EfectoVto As String
Dim EltoItm  As ListItem
Dim EsSepa As Boolean
Dim Cad As String
Dim jj As Long



'***
    CargaList
    
    CargarValores
'***
    
    
    'Si viene de la opcion de devolucion por efecto, este campo tiene valor
    EfectoSerie = ""
    Set EltoItm = Nothing
    If vSQL <> "" Then
        EfectoSerie = RecuperaValor(vSQL, 1)
        EfectoFra = Format(Val(RecuperaValor(vSQL, 2)), "00000000")
        EfectoVto = RecuperaValor(vSQL, 3)
    End If


    'Veremos si viene de un fichero de devolicion, y si trae mas de una remesa
    vSQL = ""
    Cad = RecuperaValor(vRemesa, 8)
    If Cad <> "" Then
        'Fichero de dovocucion
        Cad = RecuperaValor(vRemesa, 9)
        'Vuelvo a susitiuri los # por |
        Cad = Replace(Cad, "#", "|")
        vSQL = ""
        For jj = 1 To Len(Cad)
            If Mid(Cad, jj, 1) = "�" Then vSQL = vSQL & "X"
        Next
        
        If Len(vSQL) > 1 Then
            'Tienen mas de una remesa
            vSQL = ""
            While Cad <> ""
                jj = InStr(1, Cad, "�")
                If jj = 0 Then
                    Cad = ""
                Else
                    vSQL = vSQL & ", (" & RecuperaValor(Mid(Cad, 1, jj - 1), 1) & " , " & RecuperaValor(Mid(Cad, 1, jj - 1), 2) & ")"
                    Cad = Mid(Cad, jj + 1)
                End If
            
            Wend
            vSQL = Mid(vSQL, 2) 'quitammos la preimar coma
        Else
            vSQL = ""
        End If
        
    End If


    If vSQL = "" Then
        'Normal
        vSQL = " AND codrem =" & RecuperaValor(vRemesa, 1)
        vSQL = vSQL & " AND anyorem =" & RecuperaValor(vRemesa, 2)
    
    Else
        'Multi remesa
        vSQL = " AND (codrem,anyorem) IN ( " & vSQL & ")"
        
    End If
    vSQL = "Select cobros.* from cobros where (1=1)" & vSQL
    
    vSQL = vSQL & " ORDER BY numserie,numfactu"
    Set miRsAux = New ADODB.Recordset
    lwCobros.ListItems.Clear
    miRsAux.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    jj = 1
    While Not miRsAux.EOF
        Set Itm = lwCobros.ListItems.Add(, "C" & jj)
        Itm.Text = miRsAux!NUmSerie
        
        Itm.SubItems(1) = Format(DBLet(miRsAux!NumFactu, "N"), "00000000")
        Itm.SubItems(2) = miRsAux!numorden
        Itm.SubItems(3) = miRsAux!codmacta
        Itm.SubItems(4) = miRsAux!nomclien
        ImporteQueda = DBLet(miRsAux!Gastos, "N")
        'No lo pongo con el importe de gastos pq pudiera ser k habiendo sido devuelto, no quiera
        ' cobrarle gastos
        If miRsAux!Devuelto = 1 Then
            Itm.Bold = True
            Itm.ForeColor = vbRed
        End If
        ImporteQueda = ImporteQueda + miRsAux!ImpVenci
        Itm.SubItems(5) = Format(ImporteQueda, FormatoImporte)
        
        'Para la ordenacion
        'Por si ordena por fecha
        'ItmX.SubItems(6) = Format(RS!fecfaccl, "yyyymmdd")
        'Por si ordena por importe
        Itm.SubItems(7) = Format(miRsAux!ImpVenci * 100, "0000000000")
        
        
        
        'En el tag meto la fecha factura
        Itm.Tag = Format(miRsAux!FecFactu, "dd/mm/yyyy")
        
        

        If EfectoSerie <> "" Then
            If EfectoSerie = Itm.Text Then
                If EfectoFra = Itm.SubItems(1) Then
                    If EfectoVto = Itm.SubItems(2) Then
                        Set EltoItm = Itm
                        'Este es. Para que ya no busque mas
                        EfectoSerie = ""
                    End If
                End If
            End If
        End If
        
        
        jj = jj + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    ImporteQueda = 0
    Cad = RecuperaValor(vRemesa, 8)
    If Cad <> "" Then
        'DEVOLUCION CON FICHERO
        
        Me.Tag = Label2(4).Caption
        Label2(4).Caption = "Leyendo fichero de datos......"
        Me.Refresh
        Screen.MousePointer = vbHourglass
        Set Col = New Collection
        
        Dim TipoFicheroSepa As Byte
        
        TipoFicheroSepa = EsFicheroDevolucionSEPA2(Cad)
        
        If TipoFicheroSepa = 2 Then
            'SEPA XML
            ProcesaLineasFicheroDevolucionXML Cad, Col
            EsSepa = True
        Else
            ProcesaLineasFicheroDevolucion Cad, Col, EsSepa
        End If
        
        Me.Refresh
        If Not (Col Is Nothing) Then
            'Si Col no es nothing
            If Col.Count > 0 Then
                '-Aqui iremos recorriendo el COl hasta encontrar slos recibos que
                'Son a devolver.
                RecorreBuscandoRecibo Col, False, EsSepa
                If Col.Count > 0 Then RecorreBuscandoRecibo Col, True, EsSepa
            End If
            Label2(4).Caption = Me.Tag
        End If
        Me.Tag = ""
        
        
        
        'Borraremos los que no esten en el fichero
        ImporteQueda = 0
        For jj = Me.lwCobros.ListItems.Count To 1 Step -1
            If Not Me.lwCobros.ListItems(jj).Checked Then
                Me.lwCobros.ListItems.Remove jj
            Else
                ImporteQueda = ImporteQueda + ImporteFormateado(lwCobros.ListItems(jj).SubItems(5))
            End If
        Next
    Else
        If Not (EltoItm Is Nothing) Then
            'Ha encontrado un vto
            EltoItm.Checked = True
            lwCobros_ItemCheck EltoItm
            
            Set EltoItm = Nothing
        End If
    End If
    
    Me.Refresh
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub RecorreBuscandoRecibo(ByRef Recibos As Collection, EsMensajeNoEncontrados As Boolean, EsSepa As Boolean)
    If EsSepa Then
        RecorreBuscandoReciboSEPA Recibos, EsMensajeNoEncontrados
    Else
        RecorreBuscandoRecibo2 Recibos, EsMensajeNoEncontrados
    End If
End Sub



Private Sub PonerVtosRemesa(vSQL As String, Modificar As Boolean)
Dim IT
Dim ImporteTot As Currency
Dim Cad As String
Dim Importe As Currency


    lwCobros.ListItems.Clear
    If Not Modificar Then Text1(4).Text = ""
    
    ImporteTot = 0
    
    Set Me.lwCobros.SmallIcons = frmPpal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    
    Cad = "Select cobros.*,nomforpa " & vSQL
    Cad = Cad & " ORDER BY fecvenci"
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwCobros.ListItems.Add()
        IT.Text = miRsAux!NUmSerie
        IT.SubItems(1) = Format(miRsAux!NumFactu, "0000000")
        IT.SubItems(2) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
        IT.SubItems(3) = miRsAux!numorden
        IT.SubItems(4) = miRsAux!FecVenci
        IT.SubItems(5) = miRsAux!nomforpa
    
        IT.Checked = False
    
        Importe = DBLet(miRsAux!Gastos, "N")
        Importe = Importe + miRsAux!ImpVenci
        
        'Si ya he cobrado algo
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        
        IT.SubItems(6) = Format(Importe, FormatoImporte)
        
        ImporteTot = ImporteTot + Importe

        IT.Tag = Abs(Importe)  'siempre valor absoluto
            
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    Text1(4).Text = Format(ImporteTot, "###,###,##0.00")
    

End Sub

Private Sub Text8_LostFocus()
    If Text8.Text <> "" Then
        Opcion = 16
        LimpiarLin Me, "FrameDevDesdeVto"
        LimpiarLin Me, "FrameDevDesdeRemesa"
        CargaDevolucion
    End If
End Sub

Private Sub txtConcepto_GotFocus(Index As Integer)
    ObtenerFoco txtConcepto(Index)
End Sub

Private Sub txtConcepto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtConcepto_LostFocus(Index As Integer)
    'Lost focus
    txtConcepto(Index).Text = Trim(txtConcepto(Index).Text)
    SQL = ""
    i = 0
    If txtConcepto(Index).Text <> "" Then
        If Not IsNumeric(txtConcepto(Index).Text) Then
            MsgBox "Campo num�rico", vbExclamation
            i = 1
        Else
            
            SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtConcepto(Index).Text, "N")
            If SQL = "" Then
                MsgBox "Concepto no existe", vbExclamation
                i = 1
            End If
        End If
    End If
    Me.txtDConcpeto(Index).Text = SQL
    If i = 1 Then
        txtConcepto(Index).Text = ""
        PonerFoco txtConcepto(Index)
    Else
        SQL = "select ampdecli from tipofpago where tipoformapago = 4"
        i = DevuelveValor(SQL)
        PosicionarCombo Me.Combo2(0), i
    End If
End Sub

Private Sub txtCtaNormal_GotFocus(Index As Integer)
    ObtenerFoco txtCtaNormal(Index)
End Sub
    
Private Sub txtCtaNormal_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCtaNormal_LostFocus(Index As Integer)
Dim DevfrmCCtas As String
       
        DevfrmCCtas = Trim(txtCtaNormal(Index).Text)
        i = 0
        If DevfrmCCtas <> "" Then
            If CuentaCorrectaUltimoNivel(DevfrmCCtas, SQL) Then
                
            Else
                MsgBox SQL, vbExclamation
                If Index < 3 Or Index = 9 Or Index = 10 Or Index = 11 Then
                    DevfrmCCtas = ""
                    SQL = ""
                End If
            End If
            i = 1
        Else
            SQL = ""
        End If
        
        
        txtCtaNormal(Index).Text = DevfrmCCtas
        txtDCtaNormal(Index).Text = SQL
        If DevfrmCCtas = "" And i = 1 Then
            PonerFoco txtCtaNormal(Index)
        End If
        VisibleCC
        'limpiamos los otros frames
        If txtCtaNormal(11).Text <> "" Then
            Opcion = 28
            LimpiarLin Me, "FrameDevDesdeFichero"
            LimpiarLin Me, "FrameDevDesdeRemesa"
            CargaDevolucion
        End If
        
End Sub



Private Sub txtImporte_GotFocus(Index As Integer)
    With txtImporte(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
 Dim Valor
        txtImporte(Index).Text = Trim(txtImporte(Index))
        If txtImporte(Index).Text = "" Then Exit Sub
        

        If Not EsNumerico(txtImporte(Index).Text) Then
            txtImporte(Index).Text = ""
            Exit Sub
        End If
    
        
        If Index = 6 Or Index = 7 Then
           
            If InStr(1, txtImporte(Index).Text, ",") > 0 Then
                Valor = ImporteFormateado(txtImporte(Index).Text)
            Else
                Valor = CCur(TransformaPuntosComas(txtImporte(Index).Text))
            End If
            txtImporte(Index).Text = Format(Valor, FormatoImporte)
        End If
        
End Sub


Private Sub VisibleCC()
Dim B As Boolean

    B = False
    If vParam.autocoste Then
        If txtCtaNormal(11).Text <> "" Then
                SQL = "|" & Mid(txtCtaNormal(11).Text, 1, 1) & "|"
                If InStr(1, CuentasCC, SQL) > 0 Then B = True
        End If
    End If
End Sub



Private Sub LanzaBuscaGrid(Opcion As Integer)

'No tocar variable SQL
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String


'
'    SQL = ""
'    Screen.MousePointer = vbHourglass
'    Set frmB = New frmBasico
'    frmB.vSQL = ""
'
'    '###A mano
'    frmB.vDevuelve = "0|"   'Siempre el 0
'
'    frmB.vSelElem = 0
'
'    'Ejemplo
'        'Cod Diag.|idDiag|N|10�
'        Select Case Opcion
'        Case 1
'            'CONCEPTO
'            Cad = "Codigo|codconce|N|15�"
'            Cad = Cad & "Descripcion|nomconce|T|60�"
'            frmB.vTabla = "Conceptos"
'            frmB.vTitulo = "Conceptos"
'
'            frmB.vSQL = " codconce <900"
'
'        Case 2
'            'CC
'            Cad = "Codigo|codccost|N|15�"
'            Cad = Cad & "Descripcion|nomccost|T|60�"
'            frmB.vTabla = "cabccost"
'            frmB.vTitulo = "Centros de coste"
'
'        Case 3
'            'Cuentas agrupadas bajo el concepto: grupotesoreria
'            Cad = "Grupo tesoreria|grupotesoreria|T|60�"
'            frmB.vTabla = "cuentas"
'            frmB.vSQL = " grupotesoreria <> '' GROUP BY 1"
'            frmB.vTitulo = "Cuentas grupos tesoreria"
'        End Select
'
'
'        frmB.vCampos = Cad
'
'
'
'
'
''        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'
'
'    Screen.MousePointer = vbDefault
End Sub




Private Sub PonerValoresPorDefectoDevilucionRemesa()
Dim FP As Ctipoformapago

    On Error GoTo EPonerValoresPorDefectoDevilucionRemesa
    
    
    Set FP = New Ctipoformapago
    FP.Leer vbTipoPagoRemesa
    Me.txtConcepto(1).Text = FP.condecli
    'Ampliaciones
    Combo2(0).ListIndex = FP.ampdecli
    
    'Que carge el concepto
    txtConcepto_LostFocus 1
    Set FP = Nothing
    Exit Sub
EPonerValoresPorDefectoDevilucionRemesa:
    MuestraError Err.Number, "PonerValoresPorDefectoDevilucionRemesa"
    Set FP = Nothing
End Sub




Private Sub CamposRemesaAbono()
       
   
   
   If Text3(3) <> "" And Text3(4).Text <> "" Then
        
        Set RS = New ADODB.Recordset
        SQL = "select importe,nommacta from remesas,cuentas where remesas.codmacta=cuentas.codmacta "
        SQL = SQL & " and anyo=" & Text3(4).Text & " and codigo=" & Text3(3).Text
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
'            Me.txtTexto(0).Text = RS!Nommacta
'            Me.txtTexto(1).Text = Format(RS!Importe, FormatoImporte)
        End If
        RS.Close
        Set RS = Nothing
    End If
    
End Sub



Private Sub EliminarEnRecepcionDocumentos()
Dim CtaPte As Boolean
Dim J As Integer
Dim CualesEliminar As String
On Error GoTo EEliminarEnRecepcionDocumentos

    'Comprobaremos si hay datos
    
        'Si no lleva cuenta puente, no hace falta que este contabilizada
        'Es decir. Solo mirare contabilizados si llevo ctapuente
        CuentasCC = ""
        CualesEliminar = ""
        J = 0
        For i = 0 To 1
            ' contatalonpte
            SQL = "pagarecta"
            If i = 1 Then SQL = "contatalonpte"
            CtaPte = (DevuelveDesdeBD(SQL, "paramtesor", "codigo", "1") = "1")
            
            'Repetiremos el proceso dos veces
            SQL = "Select * from scarecepdoc where fechavto<='" & Format(Text1(17).Text, FormatoFecha) & "'"
            SQL = SQL & " AND   talon = " & i
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                    'Si lleva cta puente habra que ver si esta contbilizada
                    J = 0
                    If CtaPte Then
                        If Val(RS!Contabilizada) = 0 Then
                            'Veo si tiene lineas. S
                            SQL = DevuelveDesdeBD("count(*)", "slirecepdoc", "id", CStr(RS!Codigo))
                            If SQL = "" Then SQL = "0"
                            If Val(SQL) > 0 Then
                                CuentasCC = CuentasCC & RS!Codigo & " - No contabilizada" & vbCrLf
                                J = 1
                            End If
                        End If
                    End If
                    If J = 0 Then
                        'Si va benee
                        If Val(DBLet(RS!llevadobanco, "N")) = 0 Then
                            SQL = DevuelveDesdeBD("count(*)", "slirecepdoc", "id", CStr(RS!Codigo))
                            If SQL = "" Then SQL = "0"
                            If Val(SQL) > 0 Then
                                CuentasCC = CuentasCC & RS!Codigo & " - Sin llevar a banco" & vbCrLf
                                J = 1
                            End If
                    
                        End If
                    End If
                    'Esta la borraremos
                    If J = 0 Then CualesEliminar = CualesEliminar & ", " & RS!Codigo
                    
                    RS.MoveNext
            Wend
            RS.Close
            
            
            
        Next i
        
        

        
        If CualesEliminar = "" Then
            'No borraremos ninguna
            If CuentasCC <> "" Then
                CuentasCC = "No se puede eliminar de la recepcion de documentos los siguientes registros: " & vbCrLf & vbCrLf & CuentasCC
                MsgBox CuentasCC, vbExclamation
                
            End If
            Exit Sub
        End If
            
        
        
        'Si k hay para borrar
        CualesEliminar = Mid(CualesEliminar, 2)
        J = 1
        SQL = "X"
        Do
            i = InStr(J, CualesEliminar, ",")
            If i > 0 Then
                J = i + 1
                SQL = SQL & "X"
            End If
        Loop Until i = 0
        
        SQL = "Va a eliminar " & Len(SQL) & " registros de la recepcion de documentos." & vbCrLf & vbCrLf & vbCrLf
        If CuentasCC <> "" Then CuentasCC = "No se puede eliminar de la recepcion de documentos los siguientes registros: " & vbCrLf & vbCrLf & CuentasCC
        SQL = SQL & vbCrLf & CuentasCC
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
            SQL = "DELETE from slirecepdoc where id in (" & CualesEliminar & ")"
            Conn.Execute SQL
            
            SQL = "DELETE from scarecepdoc where codigo in (" & CualesEliminar & ")"
            Conn.Execute SQL
    
        End If

    Exit Sub
EEliminarEnRecepcionDocumentos:
    MuestraError Err.Number, Err.Description
End Sub

'Esta recibo SEPA
Private Sub RecorreBuscandoReciboSEPA(ByRef Recibos As Collection, EsMensajeNoEncontrados As Boolean)
Dim B As Boolean
Dim Cad As String
Dim jj As Integer


    If EsMensajeNoEncontrados Then
            Cad = ""
            For jj = Recibos.Count To 1 Step -1
                'M  0330047820131201001   430000061
                'SER FACTU    FEC   VTO
                
                'ImporteQueda = CCur(Val(Mid(Recibos(jj), 1, 10)) / 100)
                Cad = Cad & jj & ".-Fecha: "
                Cad = Cad & Mid(Recibos(jj), 18, 2) & "/" & Mid(Recibos(jj), 16, 2) & "/" & Mid(Recibos(jj), 12, 4)
                Cad = Cad & " Vto: " & Mid(Recibos(jj), 1, 3) & "/" & Mid(Recibos(jj), 4, 8) & "-" & Mid(Recibos(jj), 20, 3) & vbCrLf
            Next jj
            Cad = "Recibos no encontrados que vienen del fichero." & vbCrLf & vbCrLf & Cad
            MsgBox Cad, vbExclamation
            ImporteQueda = 0
    Else
        
        For jj = Recibos.Count To 1 Step -1
            'M  0330047820131201001   430000061
            'SER FACTU    FEC   VTO
            Cad = Mid(Recibos(jj), 18, 2) & "/" & Mid(Recibos(jj), 16, 2) & "/" & Mid(Recibos(jj), 12, 4)
            
            
            B = EstaElReciboSEPA(Trim(Mid(Recibos(jj), 1, 3)), Mid(Recibos(jj), 4, 8), Cad, Mid(Recibos(jj), 20, 3))

            If B Then Recibos.Remove jj
        Next jj
                
    End If
    
End Sub



Private Sub RecorreBuscandoRecibo2(ByRef Recibos As Collection, EsMensajeNoEncontrados As Boolean)
Dim B As Boolean

Dim EsFormatoAntiguoDevolucion As Boolean
Dim Cad As String
Dim jj As Long

    'Formato antiguo:A020500021
    'En el nuevo : X 00045771 >> serie(2)=X  factura(7)=4577    vto(1)=1
    EsFormatoAntiguoDevolucion = Dir(App.Path & "\DevRecAnt.dat") <> ""
    

    If EsMensajeNoEncontrados Then
            Cad = ""
            For jj = Recibos.Count To 1 Step -1
                'Ejemplo 0047080000004708
                '        251205A020500021
                '        $$$$$$ fecha                       6
                '              $ Serie                      1
                '               $$$$$$$$  Facutra           8
                '                       $  Vencimiento      1
                'La fecha
                ImporteQueda = CCur(Val(Mid(Recibos(jj), 1, 10)) / 100)
                Cad = Cad & jj & ".-Fecha: "
                Cad = Cad & Mid(Recibos(jj), 11, 2) & "/" & Mid(Recibos(jj), 13, 2) & "/20" & Mid(Recibos(jj), 15, 2)
                Cad = Cad & " Vto: " & Mid(Recibos(jj), 17, 1) & "/" & Mid(Recibos(jj), 18, 8) & "-" & Mid(Recibos(jj), 26, 1) & "   Importe: " & Format(ImporteQueda, FormatoImporte) & vbCrLf
            Next jj
            Cad = "Recibos no encontrados que vienen del fichero." & vbCrLf & vbCrLf & Cad
            MsgBox Cad, vbExclamation
            ImporteQueda = 0
    Else
        
        For jj = Recibos.Count To 1 Step -1
            'Ejemplo          0047080000004708
            '       0000001234251205A020500021
            '          ...$$$$    Importe                        10
            '                 $$$$$$ fecha                       6
            '                       $ Serie                      1
            '                        $$$$$$$$  Facutra           8
            '                                $  Vencimiento      1
            'La fecha
            Cad = Mid(Recibos(jj), 11, 2) & "/" & Mid(Recibos(jj), 13, 2) & "/20" & Mid(Recibos(jj), 15, 2)
            'Octubre 2011
            'If Not IsNumeric(Mid(Recibos(jj), 27, 1)) Then
               
            'SEPT 2013
            If Not EsFormatoAntiguoDevolucion Then
                'Alzira. Estaba mal formateado el numfac.
               B = EstaElRecibo(Mid(Recibos(jj), 17, 2), Mid(Recibos(jj), 19, 7), Cad, Mid(Recibos(jj), 26, 1))
            Else
               B = EstaElRecibo(Mid(Recibos(jj), 17, 2), Mid(Recibos(jj), 20, 7), Cad, Mid(Recibos(jj), 27, 1))
            End If
            If B Then Recibos.Remove jj
        Next jj
                
    End If
    
End Sub


Private Function EstaElReciboSEPA(Serie As String, Fac As String, Fec As String, Venci As String) As Boolean
Dim J As Integer

        'Itm.Text = miRsAux!NUmSerie
        'Itm.SubItems(1) = Format(miRsAux!codfaccl, "0000000000")
        'Itm.SubItems(2) = miRsAux!numorden
        'Itm.Tag = miRsAux!fecfaccl
        

    EstaElReciboSEPA = False
    With lwCobros
        For J = 1 To .ListItems.Count
            If Trim(.ListItems(J).Text) = Trim(Serie) Then
                'Misma serie
                If Val(.ListItems(J).SubItems(1)) = Val(Fac) And Val(.ListItems(J).SubItems(2)) = Venci And .ListItems(J).Tag = Fec Then
                        'Este es el recibo
                        .ListItems(J).Checked = True
                        ImporteQueda = ImporteQueda + ImporteFormateado(.ListItems(J).SubItems(5))
                        EstaElReciboSEPA = True
                        Exit For
                End If
            End If
        Next J
    
    End With
End Function


Private Function EstaElRecibo(Serie As String, Fac As String, Fec As String, Venci As String) As Boolean
Dim J As Integer

        'Itm.Text = miRsAux!NUmSerie
        'Itm.SubItems(1) = Format(miRsAux!codfaccl, "0000000000")
        'Itm.SubItems(2) = miRsAux!numorden
        'Itm.Tag = miRsAux!fecfaccl
        

    EstaElRecibo = False
    With lwCobros
        For J = 1 To .ListItems.Count
            If Mid(.ListItems(J).Text, 1, 2) = Trim(Serie) Then
                'Misma serie
                If Val(.ListItems(J).SubItems(1)) = Val(Fac) And .ListItems(J).SubItems(2) = Venci And .ListItems(J).Tag = Fec Then
                        'Este es el recibo
                        .ListItems(J).Checked = True
                        ImporteQueda = ImporteQueda + ImporteFormateado(.ListItems(J).SubItems(5))
                        EstaElRecibo = True
                        Exit For
                End If
            End If
        Next J
    
    
        'Nov 2012
        If Not EstaElRecibo Then
            'Pruebo solo con el numero de vto y que la primera letra d serie sea como la del vto (pueden ser dos)
            'Ademas meto el numero de vto dentro del fac
            For J = 1 To .ListItems.Count
                If Mid(.ListItems(J).Text, 1, 1) = Mid(Trim(Serie), 1, 1) Then
                        'Misma serie
                        If Val(.ListItems(J).SubItems(1)) = Val(Fac & Venci) And .ListItems(J).Tag = Fec Then
                                'Este es el recibo
                                .ListItems(J).Checked = True
                                ImporteQueda = ImporteQueda + ImporteFormateado(.ListItems(J).SubItems(5))
                                EstaElRecibo = True
                                Exit For
                        End If
                End If
            Next
        End If
    End With
End Function

