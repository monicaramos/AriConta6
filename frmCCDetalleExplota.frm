VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCCDetalleExplota 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   4605
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   6915
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
         Index           =   1
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "imgConcepto"
         Top             =   3900
         Width           =   1305
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
         Index           =   1
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   3900
         Width           =   4185
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
         Index           =   0
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3480
         Width           =   4185
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1050
         Width           =   4185
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1470
         Width           =   4185
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
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   1050
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
         Index           =   1
         Left            =   1230
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1470
         Width           =   1275
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
         Index           =   0
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "imgConcepto"
         Top             =   3480
         Width           =   1305
      End
      Begin VB.TextBox txtFecha 
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
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "imgConcepto"
         Top             =   2640
         Width           =   1305
      End
      Begin VB.TextBox txtFecha 
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
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "imgConcepto"
         Top             =   2220
         Width           =   1305
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   6360
         TabIndex        =   33
         Top             =   240
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
         Index           =   1
         Left            =   270
         TabIndex        =   36
         Top             =   3930
         Width           =   600
      End
      Begin VB.Image imgCCoste 
         Height          =   255
         Index           =   1
         Left            =   900
         Top             =   3930
         Width           =   255
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
         Index           =   0
         Left            =   270
         TabIndex        =   34
         Top             =   3510
         Width           =   600
      End
      Begin VB.Image imgCCoste 
         Height          =   255
         Index           =   0
         Left            =   900
         Top             =   3510
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   0
         Left            =   930
         Top             =   1050
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   1
         Left            =   930
         Top             =   1500
         Width           =   255
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
         Height          =   225
         Index           =   6
         Left            =   240
         TabIndex        =   29
         Top             =   3120
         Width           =   2130
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   930
         Picture         =   "frmCCDetalleExplota.frx":0000
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   930
         Picture         =   "frmCCDetalleExplota.frx":008B
         Top             =   2250
         Width           =   240
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
         Index           =   0
         Left            =   2550
         TabIndex        =   28
         Top             =   990
         Width           =   4095
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
         TabIndex        =   27
         Top             =   1440
         Width           =   4095
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
         TabIndex        =   26
         Top             =   1440
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
         TabIndex        =   25
         Top             =   1080
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
         Left            =   240
         TabIndex        =   24
         Top             =   2640
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
         Left            =   240
         TabIndex        =   23
         Top             =   2280
         Width           =   690
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
         Index           =   7
         Left            =   240
         TabIndex        =   22
         Top             =   690
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
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
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   960
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
      Left            =   5820
      TabIndex        =   8
      Top             =   7530
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
      Left            =   4260
      TabIndex        =   6
      Top             =   7530
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
      TabIndex        =   7
      Top             =   7470
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
      TabIndex        =   9
      Top             =   4710
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
         TabIndex        =   20
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   19
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   18
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCCDetalleExplota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 1005



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

Public Cuenta As String
Public Descripcion As String
Public FecDesde As String
Public FecHasta As String


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCCo As frmBasico
Attribute frmCCo.VB_VarHelpID = -1

Private SQL As String
Dim Cad As String
Dim RC As String
Dim i As Integer
Dim IndCodigo As Integer
Dim PrimeraVez As String


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



Private Sub cmdAccion_Click(Index As Integer)
    
    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme("hlinapu", cadselect) Then Exit Sub
    
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
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
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
    Me.Icon = frmPpal.Icon
        
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
        
        
    'Otras opciones
    Me.Caption = "Detalle de Explotación"

    For i = 0 To 1
        Me.imgCuentas(i).Picture = frmPpal.ImageList3.ListImages(1).Picture
    Next i
    For i = 0 To 0
        Me.imgCCoste(i).Picture = frmPpal.ImageList3.ListImages(1).Picture
    Next i
    
    PrimeraVez = True
     
    
    txtFecha(0).Text = vParam.fechaini
    txtFecha(1).Text = vParam.fechafin
    If Not vParam.FecEjerAct Then
        txtFecha(1).Text = Format(DateAdd("yyyy", 1, vParam.fechafin), "dd/mm/yyyy")
    End If
     
   
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtCuentas(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNCuentas(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
    txtCCoste(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNCCoste(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub ImgCCoste_Click(Index As Integer)

    IndCodigo = Index
    
    Set frmCCo = New frmBasico
    frmCCo.DatosADevolverBusqueda = "0|1|"
    frmCCo.Show vbModal
    Set frmCCo = Nothing
    
    PonFoco txtCCoste(Index)

End Sub


Private Sub imgCuentas_Click(Index As Integer)

    IndCodigo = Index
    
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1|"
    frmC.ConfigurarBalances = 3  'NUEVO
    frmC.Show vbModal
    Set frmC = Nothing

    PonFoco txtCuentas(Index)

End Sub


Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0, 1, 2
        IndCodigo = Index
    
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If txtFecha(Index).Text <> "" Then frmF.Fecha = CDate(txtFecha(Index).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco txtFecha(Index)
        
    End Select
    
    Screen.MousePointer = vbDefault

End Sub


Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
End Sub

Private Sub optVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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

Private Sub txtCuentas_GotFocus(Index As Integer)
    ConseguirFoco txtCuentas(Index), 3
End Sub


Private Sub txtCuentas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0

        LanzaFormAyuda "imgCuentas", Index
    End If
End Sub


Private Sub LanzaFormAyuda(Nombre As String, indice As Integer)
    Select Case Nombre
    Case "imgCCuentas"
        imgCuentas_Click indice
    Case "imgFecha"
        imgFec_Click indice
    Case "imgCCoste"
        ImgCCoste_Click indice
    End Select
    
End Sub


''++
'Private Sub txtCuentas_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = teclaBuscar Then
'        LanzaFormAyuda "imgCuentas", Index
'    Else
'        KEYpress KeyAscii
'    End If
'End Sub


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

Private Sub txtCCoste_GotFocus(Index As Integer)
    ConseguirFoco txtCCoste(Index), 3
End Sub


Private Sub txtCCoste_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0

        LanzaFormAyuda "imgCCoste", Index
    End If
End Sub


Private Sub txtCCoste_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim RC As String
Dim Hasta As Integer

    txtCCoste(Index).Text = Trim(txtCCoste(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    txtCCoste(Index).Text = Trim(txtCCoste(Index).Text)
    If txtCCoste(Index).Text = "" Then
        txtNCCoste(Index).Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtCCoste(Index).Text) Then
        If InStr(1, txtCCoste(Index).Text, "+") = 0 Then MsgBox "El Centro de Coste debe ser numérico: " & txtCCoste(Index).Text, vbExclamation
        txtCCoste(Index).Text = ""
        txtNCCoste(Index).Text = ""
        Exit Sub
    End If



    Select Case Index
        Case 0, 1 'Centros de Coste
            txtNCCoste(Index).Text = DevuelveDesdeBD("nomccost", "ccoste", "codccost", txtCCoste(Index), "N")
            
    End Select

End Sub


Private Sub AccionesCSV()
Dim SQL2 As String

    'Monto el SQL
    SQL = "Select hlinapu.codccost CCoste, ccoste.nomccost Descripción, hlinapu.codmacta as Cuenta, cuentas.nommacta Título, hlinapu.fechaent Fecha,  "
    SQL = SQL & " numdocum Documento, ctacontr Contrapartida, cuentas1.nommacta Descripción , hlinapu.ampconce Concepto, coalesce(timported,0) Debe, coalesce(timporteh,0) Haber, "
    SQL = SQL & " if (@Cta <> hlinapu.codmacta,@Saldo:= coalesce(timported,0) - coalesce(timporteh,0), @Saldo:= @Saldo + coalesce(timported,0) - coalesce(timporteh,0) ) Saldo, "
    SQL = SQL & " @Cta:= hlinapu.codmacta Cta"

    SQL = SQL & " FROM ((hlinapu INNER JOIN cuentas ON hlinapu.codmacta = cuentas.codmacta)  "
    SQL = SQL & " INNER JOIN ccoste ON hlinapu.codccost = ccoste.codccost) "
    SQL = SQL & " INNER JOIN cuentas cuentas1 ON hlinapu.ctacontr = cuentas1.codmacta , (select @Saldo:= 0) aaa, (select @Cta:= '') bbb   "
    SQL = SQL & " where " & cadselect
    SQL = SQL & " ORDER BY 1,2,3,4,5 "
    
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    indRPT = "1006-00" '"CC_X_CTA.rpt"
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu

    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 2
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
End Sub


Private Function MontaSQL() As Boolean
Dim SQL As String
Dim SQL2 As String
Dim RC As String
Dim RC2 As String

    MontaSQL = False
    
    If Not PonerDesdeHasta("hlinapu.fechaent", "F", Me.txtFecha(0), Me.txtFecha(0), Me.txtFecha(1), Me.txtFecha(1), "pDHFecha=""") Then Exit Function
    If Not PonerDesdeHasta("hlinapu.codmacta", "CTA", Me.txtCuentas(0), Me.txtNCuentas(0), Me.txtCuentas(1), Me.txtNCuentas(1), "pDHCuentas=""") Then Exit Function
    If Not PonerDesdeHasta("hlinapu.codccost", "CCO", Me.txtCCoste(0), Me.txtNCCoste(0), Me.txtCCoste(1), Me.txtNCCoste(1), "pDHCCoste=""") Then Exit Function
    
    MontaSQL = CargarTemporal
           
End Function

Private Function CargarTemporal() As Boolean
Dim SQL As String


    On Error GoTo eCargarTemporal
    
    CargarTemporal = False
    
    SQL = "delete from tmplinccexplo where codusu = " & DBSet(vUsu.Codigo, "N")
    Conn.Execute SQL
    
    SQL = "insert into tmplinccexplo (codusu,codccost,codmacta,linapu,docum,fechaent,ampconce,ctactra,desctra,perD,perH) "
    SQL = SQL & " Select " & vUsu.Codigo & ", hlinapu.codccost CCoste,  hlinapu.codmacta as Cuenta, hlinapu.linliapu, hlinapu.numdocum, hlinapu.fechaent,  "
    SQL = SQL & " hlinapu.ampconce, ctacontr Contrapartida, cuentas1.nommacta Descripción , coalesce(timported,0) Debe, coalesce(timporteh,0) Haber "
    SQL = SQL & " FROM (hlinapu INNER JOIN cuentas cuentas1 ON hlinapu.ctacontr = cuentas1.codmacta) "
    SQL = SQL & " INNER JOIN ccoste ON hlinapu.codccost = ccoste.codccost "
    SQL = SQL & " where " & cadselect
    SQL = SQL & " ORDER BY 1,2,3,4,5 "
    
    Conn.Execute SQL
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Tabla Temporal", Err.Description
End Function


Private Sub txtfecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    PonerFormatoFecha txtFecha(Index)
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda "imgFecha", Index
    End If
End Sub

Private Function DatosOK() As Boolean
    
    DatosOK = False
    

    DatosOK = True

End Function

Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub
