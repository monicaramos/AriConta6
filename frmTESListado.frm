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
      Left            =   30
      TabIndex        =   4
      Top             =   60
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
   Begin VB.Frame FrameDividVto 
      Height          =   2415
      Left            =   150
      TabIndex        =   30
      Top             =   90
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
         Left            =   150
         TabIndex        =   32
         Top             =   210
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
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameProgreso 
      Height          =   1935
      Left            =   210
      TabIndex        =   0
      Top             =   240
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

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 20 Or Index = 23 Or Index >= 26 Then
        CadenaDesdeOtroForm = "" 'Por si acaso. Tiene que devolve "" para que no haga nada
    End If
    Unload Me
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
    CargaImagenesAyudas Me.ImageAyudaImpcta, 3
    For Each Img In Me.imgConcepto
        Img.ToolTipText = "Concepto"
    Next
    For Each Img In Me.imgDiario
        Img.ToolTipText = "Diario"
    Next
    
    
    
    'Limpiamos el tag
    PrimeraVez = True
    FrameCompensaciones.Visible = False
    FrameDividVto.Visible = False
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

Private Sub frmBa_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CInt(RC)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtFPago(RC).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescFPago(RC).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image2_Click(Index As Integer)
    
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text3(Index).Text <> "" Then frmC.Fecha = CDate(Text3(Index).Text)
    RC = Index
    frmC.Show vbModal
    Set frmC = Nothing
End Sub


Private Sub ImageAyudaImpcta_Click(Index As Integer)
Dim C As String
    Select Case Index
    Case 0
            C = "Compensaciones" & vbCrLf & String(60, "-") & vbCrLf
            C = C & "Cuando compense sobre un vencimiento al marcar la opción " & vbCrLf
            C = C & Space(10) & Me.chkCompensa.Caption & vbCrLf
            C = C & "se modificará el importe vencimiento poniendo el total a compensar  y en importe cobrado un cero"
            
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



Private Sub imgFP_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    'Set frmCta = New frmColCtas
    Set frmP = New frmFormaPago
    RC = Index
    frmP.DatosADevolverBusqueda = "0|1"
    frmP.Show vbModal
    Set frmP = Nothing
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

Private Sub PonerFrameProgressVisible(Optional TEXTO As String)
        If TEXTO = "" Then TEXTO = "Generando datos"
        Me.lblPPAL.Caption = TEXTO
        Me.lbl2.Caption = ""
        Me.ProgressBar1.Value = 0
        Me.FrameProgreso.Visible = True
        Me.Refresh
End Sub


'Para conceptos y diarios
'Opcion: 0- Diario
'        1- Conceptos
'        2- Centros de coste
'        3- Gastos fijos
'        4. Hco compensaciones
Private Sub LanzaBuscaGrid(Indice As Integer, OpcionGrid As Byte)


End Sub

                                       '                Para saber el index del listview
Public Sub InsertaItemComboCompensaVto(TEXTO As String, Indice As Integer)
    Me.cboCompensaVto.AddItem TEXTO
    Me.cboCompensaVto.ItemData(Me.cboCompensaVto.NewIndex) = Indice
End Sub




