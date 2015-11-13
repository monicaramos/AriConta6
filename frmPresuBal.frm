VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPresuBal 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   11670
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
      Height          =   4995
      Left            =   7080
      TabIndex        =   14
      Top             =   0
      Width           =   4455
      Begin VB.CheckBox chkPreAct 
         Caption         =   "Ejercicio siguiente"
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
         Left            =   540
         TabIndex        =   40
         Top             =   1260
         Width           =   2505
      End
      Begin VB.CheckBox chkPreMensual 
         Caption         =   "Mensual"
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
         Left            =   540
         TabIndex        =   39
         Top             =   780
         Width           =   1425
      End
      Begin VB.TextBox txtMes 
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
         Left            =   2640
         TabIndex        =   38
         Top             =   780
         Width           =   1485
      End
      Begin VB.CheckBox chkQuitarApertura 
         Caption         =   "Quitar apertura"
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
         Left            =   540
         TabIndex        =   37
         Top             =   1770
         Width           =   2205
      End
      Begin VB.Frame Frame2 
         Height          =   1815
         Left            =   150
         TabIndex        =   22
         Top             =   2370
         Width           =   4185
         Begin VB.CheckBox Check1 
            Caption         =   "9º nivel"
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
            Index           =   9
            Left            =   120
            TabIndex        =   32
            Top             =   1290
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "8º nivel"
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
            Index           =   8
            Left            =   2850
            TabIndex        =   31
            Top             =   960
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "7º nivel"
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
            Index           =   7
            Left            =   1470
            TabIndex        =   30
            Top             =   960
            Width           =   1305
         End
         Begin VB.CheckBox Check1 
            Caption         =   "6º nivel"
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
            Index           =   6
            Left            =   120
            TabIndex        =   29
            Top             =   930
            Width           =   1305
         End
         Begin VB.CheckBox Check1 
            Caption         =   "5º nivel"
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
            Index           =   5
            Left            =   2850
            TabIndex        =   28
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "4º nivel"
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
            Index           =   4
            Left            =   1470
            TabIndex        =   27
            Top             =   600
            Width           =   1305
         End
         Begin VB.CheckBox Check1 
            Caption         =   "3º nivel"
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
            Index           =   3
            Left            =   120
            TabIndex        =   26
            Top             =   570
            Width           =   1245
         End
         Begin VB.CheckBox Check1 
            Caption         =   "2º nivel"
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
            Index           =   2
            Left            =   2850
            TabIndex        =   25
            Top             =   240
            Width           =   1185
         End
         Begin VB.CheckBox Check1 
            Caption         =   "1er nivel"
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
            Left            =   1470
            TabIndex        =   24
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Último:  "
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
            Index           =   10
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   1  'Checked
            Width           =   1155
         End
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   3750
         TabIndex        =   21
         Top             =   210
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
         Caption         =   "Mes"
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
         Left            =   2040
         TabIndex        =   36
         Top             =   810
         Width           =   690
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
      Height          =   2295
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtNCta 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1050
         Width           =   4185
      End
      Begin VB.TextBox txtNCta 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1470
         Width           =   4185
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
         Index           =   6
         Left            =   1230
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   1050
         Width           =   1275
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
         Index           =   7
         Left            =   1230
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   6
         Left            =   990
         Top             =   1050
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   7
         Left            =   990
         Top             =   1500
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   1080
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
         TabIndex        =   18
         Top             =   690
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
      Left            =   10320
      TabIndex        =   4
      Top             =   5160
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
      TabIndex        =   2
      Top             =   5160
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
      TabIndex        =   3
      Top             =   5130
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
      TabIndex        =   5
      Top             =   2340
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
         TabIndex        =   17
         Top             =   720
         Width           =   1515
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   16
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   15
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
      TabIndex        =   33
      Top             =   5160
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pb4 
      Height          =   285
      Left            =   1560
      TabIndex        =   41
      Top             =   5160
      Visible         =   0   'False
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmPresuBal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 1103

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


Public Cuenta As String
Public Descripcion As String
Public FecDesde As String
Public FecHasta As String


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmDia As frmTiposDiario
Attribute frmDia.VB_VarHelpID = -1
Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCon  As frmConceptos
Attribute frmCon.VB_VarHelpID = -1
Private frmCtas As frmCtasAgrupadas

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


Dim HanPulsadoSalir As Boolean

Dim vFecIni As Date
Dim vFecFin As Date

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
    
    
    'Solo podemos quitar el asiento de apertura para ejercicio actual
    i = 0
    If chkQuitarApertura.Value = 1 Then
        i = 1
        'ejer siguiente
        If chkPreAct.Value = 1 Then
            i = 0
        Else
            'Si es mensual y el mes NO es uno tampoco lo quita
            If chkPreMensual.Value = 1 Then
                If Val(txtMes(2).Text) > 1 Then i = 0
            End If
        End If
    End If
    chkQuitarApertura.Value = i
    
    SQL = ""
    RC = ""
    If txtCta(6).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        RC = "Desde " & txtCta(6).Text & " - " & txtNCta(6).Text
        SQL = SQL & "presupuestos.codmacta >= '" & txtCta(6).Text & "'"
    End If
    
    
    If txtCta(7).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        If RC <> "" Then
            RC = RC & "       h"
        Else
            RC = "H"
        End If
        RC = RC & "asta " & txtCta(7).Text & " - " & txtNCta(7).Text
        SQL = SQL & "presupuestos.codmacta <= '" & txtCta(7).Text & "'"
    End If
'--
'    If Sql <> "" Then Sql = Sql & " AND"
'    i = Year(vParam.fechaini)
'    If chkPreAct.Value Then i = i + 1
'    Sql = Sql & " anopresu =" & i
'++ sustituido por
    If SQL <> "" Then SQL = SQL & " AND"
    If chkPreAct.Value Then
        vFecIni = DateAdd("yyyy", 1, vParam.fechaini)
        vFecFin = DateAdd("yyyy", 1, vParam.fechafin)
        SQL = SQL & " date(concat(right(concat('0000',anopresu),4), right(concat('00',mespresu),2),'01')) between " & DBSet(vFecIni, "F") & " and " & DBSet(vFecFin, "F")
    Else
        vFecIni = vParam.fechaini
        vFecFin = vParam.fechafin
        SQL = SQL & " date(concat(right(concat('0000',anopresu),4), right(concat('00',mespresu),2),'01')) between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    End If
    
    If RC <> "" Then RC = """ + chr(13) +""" & RC
    If chkPreMensual.Value = 1 Then
        If txtMes(2).Text <> "" Then RC = "** " & Format("01/" & txtMes(2).Text & "/1999", "mmmm") & " ** " & RC
        RC = "  MENSUAL " & RC
    End If
    
    
    
    RC = "Ejercicio: " & vFecIni & " " & vFecFin & RC
    CadenaDesdeOtroForm = ""
    
    For Cont = 1 To 10
        If Check1(Cont).Value = 1 Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "- " & Cont
    Next

    RC = RC & " Digitos: " & Mid(CadenaDesdeOtroForm, 2)
    
    If chkQuitarApertura.Value = 1 Then RC = RC & "     Sin apertura"
    CadenaDesdeOtroForm = "CampoSeleccion= """ & RC & """|"

    cadParam = cadParam & CadenaDesdeOtroForm
    numParam = numParam + 1


    RC = ""
    For Cont = 1 To 9
        If Check1(Cont).Value = 1 Then
            If RC = "" Then RC = Cont
        End If
    Next
    If RC = "" Then RC = "11"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & "Remarcar= " & RC & "|"
    

    Me.cmdCancelarAccion.Visible = False
    Me.cmdCancelarAccion.Enabled = False
    
    Me.cmdCancelar.Visible = True
    Me.cmdCancelar.Enabled = True

    
    If Not MontaSQL Then Exit Sub
    
    If Not HayRegParaInforme("tmppresu2", "codusu=" & vUsu.Codigo) Then Exit Sub
    
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
    
    Me.pb4.Visible = False
    
    
End Sub


Public Function GeneraBalancePresupuestario() As Boolean
Dim AUx As String
Dim Importe As Currency
Dim AUX2 As String
Dim vMes  As Integer
Dim Cta As String

On Error GoTo EGeneraBalancePresupuestario

    GeneraBalancePresupuestario = False
    
    If Me.chkPreMensual.Value = 0 Then
        AUx = "select codmacta,sum(imppresu)  from presupuestos "
        If SQL <> "" Then AUx = AUx & " where " & SQL
        AUx = AUx & " group by codmacta"
        
'--      'Para el otro
'        Cad = "Select SUM(impmesde),SUM(impmesha) from hsaldos where anopsald=" & i
'        Cad = Cad & " and codmacta = '"
'++
        Cad = "select sum(coalesce(timported,0)),sum(coalesce(timporteh,0)) from hlinapu where fechaent between " & DBSet(vFecIni, "F") & " and " & DBSet(vFecFin, "F")
        Cad = Cad & " and codmacta = '"

    Else
        AUx = "select codmacta,imppresu,mespresu, anopresu from presupuestos where " & SQL
        If txtMes(2).Text <> "" Then AUx = AUx & " and mespresu = " & txtMes(2).Text
        AUx = AUx & " ORDER BY codmacta,mespresu"
'--      'para luego
'        Cad = "Select impmesde,impmesha from hsaldos where anopsald=" & i
'        If txtMes(2).Text <> "" Then Cad = Cad & " and mespsald = " & txtMes(2).Text
'        Cad = Cad & " and codmacta = '"
'++
        Cad = "select sum(coalesce(timported,0)),sum(coalesce(timporteh,0)) from hlinapu where fechaent between " & DBSet(vFecIni, "F") & " and " & DBSet(vFecFin, "F")
        If txtMes(2).Text <> "" Then Cad = Cad & " and month(fechaent)= " & txtMes(2).Text
        Cad = Cad & " and codmacta = '"
       
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open AUx, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        MsgBox "Ningún registro a mostrar.", vbExclamation
        Rs.Close
        Exit Function
    End If
    
    'Borramos tmp de presu 2
    AUx = "DELETE FROM tmppresu2 where codusu =" & vUsu.Codigo
    Conn.Execute AUx
    
    SQL = "INSERT INTO tmppresu2 (codusu, codigo, cta, titulo,  mes, Presupuesto, realizado, anyo) VALUES ("
    SQL = SQL & vUsu.Codigo & ","
    
    Cont = 0
    Do
        Cont = Cont + 1
        Rs.MoveNext
    Loop Until Rs.EOF
    Rs.MoveFirst
    
    'Ponemos el PB4
    pb4.Max = Cont + 1
    pb4.Value = 0
    If Cont > 3 Then pb4.Visible = True
    Cta = ""
    Cont = 1   'Contador
    While Not Rs.EOF
        If Me.chkPreMensual.Value = 1 Then
            If Cta <> Rs!codmacta Then
                vMes = 1
                Cta = Rs!codmacta
            End If
            
            If Rs!mespresu > vMes Then
                For i = vMes To Rs!mespresu - 1
                
                    AUx = Rs!codmacta  'Aqui pondremos el nombre
                    AUx = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", AUx, "T")
                    AUx = Cont & ",'" & Rs!codmacta & "','" & DevNombreSQL(AUx) & "',"
                    AUx = AUx & i
             
                    AUx = AUx & ",0,"
                    
                    AUX2 = Cad & Rs!codmacta & "'"
                    AUX2 = AUX2 & " AND month(fechaent) =" & i
                    
                    Importe = ImporteBalancePresupuestario(AUX2)
                    
                    AUx = AUx & TransformaComasPuntos(CStr(Importe)) & ","
                    AUx = AUx & DBSet(Rs!anopresu, "N") & ")"
                    
                    If Importe <> 0 Then
                        Conn.Execute SQL & AUx
                        Cont = Cont + 1
                    End If
                Next i
            End If
            
        End If
                
        
    
    
        AUx = Rs!codmacta  'Aqui pondremos el nombre
        AUx = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", AUx, "T")
        AUx = Cont & ",'" & Rs!codmacta & "','" & DevNombreSQL(AUx) & "',"
        If Me.chkPreMensual.Value = 0 Then
            AUx = AUx & "0"
        Else
            AUx = AUx & Rs!mespresu
        End If
        AUx = AUx & "," & TransformaComasPuntos(CStr(Rs.Fields(1))) & ","
        
        'SQL
        AUX2 = Cad & Rs!codmacta & "'"
        If Me.chkPreMensual.Value = 1 Then
            AUX2 = AUX2 & " AND month(fechaent) =" & Rs!mespresu
            'AUmento el mes
            vMes = Rs!mespresu + 1
        End If
        
        
        Importe = ImporteBalancePresupuestario(AUX2)
        'Debug.Print Importe
        AUx = AUx & TransformaComasPuntos(CStr(Importe)) & ","
        If Me.chkPreMensual.Value = 0 Then
            AUx = AUx & "0)"
        Else
            AUx = AUx & DBSet(Rs!anopresu, "N") & ")"
        End If
        Conn.Execute SQL & AUx
        
        'Sig
        pb4.Value = pb4.Value + 1
        Cont = Cont + 1
        Rs.MoveNext
    Wend
    Rs.Close
    
    
        '2013  Junio
    ' QUitaremos si asi lo pide, el saldo de la apertura
    ' Curiosamente, las 6 y 7  NO tienen apertura(perdi y ganacias)
    RC = "" 'Por si quitamos el apunte de apertura. Guardare las cuentas para buscarlas despues en la apertura
    If chkQuitarApertura.Value = 1 Then
        AUx = "SELECT cta from tmppresu2 WHERE codusu = " & vUsu.Codigo & " GROUP BY 1"
        Rs.Open AUx, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            RC = RC & ", '" & Rs!Cta & "'"
            Rs.MoveNext
        Wend
        Rs.Close
        
        
        
        'Subo qui lo de quitar apertura
        If RC <> "" Then
            RC = Mid(RC, 2)
            AUx = " AND codmacta IN (" & RC & ")"
            
            Cad = "SELECT codmacta cta,sum(coalesce(timported,0))-sum(coalesce(timporteh,0)) as importe"
            Cad = Cad & " from hlinapu where codconce=970 and fechaent='" & Format(vParam.fechaini, FormatoFecha) & "'"
            Cad = Cad & AUx
            Cad = Cad & " GROUP BY 1"
            Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                Cad = "UPDATE tmppresu2 SET realizado=realizado-" & TransformaComasPuntos(CStr(Rs!Importe))
                
                Cad = Cad & " WHERE codusu = " & vUsu.Codigo & " AND cta = '" & Rs!Cta & "' AND mes = "
                If Me.chkPreMensual.Value = 1 Then
                    Cad = Cad & " 1"
                Else
                    Cad = Cad & " 0"
                End If
                Conn.Execute Cad
                Rs.MoveNext
            Wend
            Rs.Close
                
            
            
        End If
        
        
        
    End If
    
    
    'Si pide a 3 DIGITOS este es el momemto
    'Sera facil.
    'Hacemos un insert into con substring
 
        'SUBNIVEL
        AUx = ""
        For i = 1 To 9
            If Check1(i).Value = 1 Then
                
                AUx = DevuelveDesdeBD("count(*)", "tmppresu2", "codusu", CStr(vUsu.Codigo))
                Cont = Val(AUx)
                
                '@rownum:=@rownum+1 AS rownum      (SELECT @rownum:=0) r
                AUx = "Select " & vUsu.Codigo & " us,@rownum:=@rownum+1 AS rownum,substring(cta,1," & i & ") as cta2,mes,sum(presupuesto),sum(realizado)"
                AUx = AUx & " FROM tmppresu2,(SELECT @rownum:=" & Cont & ") r WHERE codusu = " & vUsu.Codigo
                
                AUx = AUx & " AND length(cta)=" & vEmpresa.DigitosUltimoNivel
                
                AUx = AUx & " group by cta2,us,mes"
                AUx = "insert into tmppresu2 (codusu, codigo, cta,   mes, Presupuesto, realizado) " & AUx
                'Insertamos
                Conn.Execute AUx
                
                'Quito los de ultimo nivel

                
                AUx = "SELECT cta from tmppresu2 WHERE codusu = " & vUsu.Codigo & " GROUP BY 1"
                Rs.Open AUx, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not Rs.EOF
                    'Actualizo el nommacta
                    AUx = Rs!Cta  'Aqui pondremos el nombre
                    AUx = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", AUx, "T")
                    AUx = "UPDATE tmppresu2  SET titulo = '" & DevNombreSQL(AUx) & "' WHERE codusu = " & vUsu.Codigo & " AND Cta = '" & Rs!Cta & "'"
                    Conn.Execute AUx
                    Rs.MoveNext
                Wend
                Rs.Close
                
                
                
            End If
        Next
        
        
        If Check1(10).Value = 0 Then
            AUx = "DELETE FROM tmppresu2 WHERE codusu = " & vUsu.Codigo & " AND cta like '" & Mid("__________", 1, vEmpresa.DigitosUltimoNivel) & "'"
            Conn.Execute AUx
        End If
        
    
    
    
  
            
        
  
    
    
    Set Rs = Nothing
    GeneraBalancePresupuestario = True
    Exit Function
EGeneraBalancePresupuestario:
    MuestraError Err.Number, "Generar balance presupuestario"
    Set Rs = Nothing
End Function




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
    Me.Caption = "Balance Presupuestario"

    For i = 6 To 7
        Me.imgCuentas(i).Picture = frmPpal.ImageList3.ListImages(1).Picture
    Next i
    
    PrimeraVez = True
     
    CargarComboFecha
     

    
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
    cmdCancelarAccion.Enabled = False
    cmdCancelarAccion.Visible = False
    
    
    PonerNiveles
    
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNCta(IndCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub





Private Sub imgCuentas_Click(Index As Integer)

    IndCodigo = Index
    
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1|"
    frmC.ConfigurarBalances = 3  'NUEVO
    frmC.Show vbModal
    Set frmC = Nothing

    PonFoco txtCta(Index)

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


Private Sub txtCta_GotFocus(Index As Integer)
    ConseguirFoco txtCta(Index), 3
End Sub


Private Sub txtCta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0

        LanzaFormAyuda "imgCuentas", Index
    End If
End Sub


Private Sub LanzaFormAyuda(Nombre As String, indice As Integer)
    Select Case Nombre
    Case "imgCuentas"
        imgCuentas_Click indice
    End Select
    
End Sub


''++
'Private Sub txtcta_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = teclaBuscar Then
'        LanzaFormAyuda "imgCuentas", Index
'    Else
'        KEYpress KeyAscii
'    End If
'End Sub


Private Sub txtCta_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim RC As String
Dim Hasta As Integer

    txtCta(Index).Text = Trim(txtCta(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    txtCta(Index).Text = Trim(txtCta(Index).Text)
    If txtCta(Index).Text = "" Then
        txtNCta(Index).Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(txtCta(Index).Text) Then
        If InStr(1, txtCta(Index).Text, "+") = 0 Then MsgBox "La cuenta debe ser numérica: " & txtCta(Index).Text, vbExclamation
        txtCta(Index).Text = ""
        txtNCta(Index).Text = ""
        Exit Sub
    End If



    Select Case Index
        Case 6, 7 'Cuentas
            'lblCuentas(Index).Caption = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", txtcta(Index), "T")
            
            RC = txtCta(Index).Text
            If CuentaCorrectaUltimoNivelSIN(RC, SQL) Then
                txtCta(Index) = RC
                txtNCta(Index).Text = SQL
            Else
                MsgBox SQL, vbExclamation
                txtCta(Index).Text = ""
                txtNCta(Index).Text = ""
                PonFoco txtCta(Index)
            End If
            
            If Index = 0 Then Hasta = 1
            If Hasta >= 1 Then
                txtCta(Hasta).Text = txtCta(Index).Text
                txtNCta(Hasta).Text = txtNCta(Index).Text
            End If
    End Select

'    PierdeFocoTiposDiario Me.txtTiposDiario(Index), Me.lblTiposDiario(Index)
End Sub



Private Sub AccionesCSV()
Dim SQL2 As String
Dim Tipo As Byte
        
    SQL = "SELECT `tmppresu2`.`cta` Cuenta, `tmppresu2`.`titulo` Nombre, `tmppresu2`.`anyo` Anyo, `tmppresu2`.`mes` Mes, `tmppresu2`.`Presupuesto` Presupuesto, `tmppresu2`.`realizado` Reslizado "
    SQL = SQL & " FROM  `tmppresu2` `tmppresu2`"
    SQL = SQL & " where codusu = " & DBSet(vUsu.Codigo, "N")
    SQL = SQL & " ORDER BY `tmppresu2`.`cta`"
        
        
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim Tipo As Byte
Dim UltimoNivel As Integer
Dim indRPT As String
Dim nomDocu As String


    '------------------------------
    'Numero de niveles
    'Para cada nivel marcado veremos si tiene cuentas en la tmp
    Cont = 0
    UltimoNivel = 0
    For i = 1 To 10
        If Check1(i).Visible Then
'                If Check2(I).Value = 1 Then Cont = Cont + 1
            If Check1(i).Value = 1 Then
                If i = 10 Then
                    Cad = vEmpresa.DigitosUltimoNivel
                Else
                    Cad = CStr(DigitosNivel(i))
                End If
            End If
        End If
    Next i
    Cad = "numeroniveles= " & Cont & "|"
    SQL = SQL & Cad
    'Otro parametro mas
    Cad = "vUltimoNivel= " & UltimoNivel & "|"
    
    cadParam = cadParam & Cad
    numParam = numParam + 2

    
    vMostrarTree = False
    conSubRPT = False
        
    If chkPreMensual = 1 Then
        indRPT = "1103-00"
    Else
        indRPT = "1103-01"
    End If
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu '"SumasySaldos.rpt"

    cadFormula = "{tmppresu2.codusu}=" & vUsu.Codigo

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
    
'    If Not PonerDesdeHasta("hlinapu.fechaent", "F", Me.txtFecha(0), Me.lblFecha(0), Me.txtFecha(1), Me.lblFecha(1), "pDHFecha=""") Then Exit Function

'    If Not PonerDesdeHasta("hlinapu.codmacta", "CTA", Me.txtCta(0), Me.lblCuentas(0), Me.txtCta(1), Me.lblCuentas(1), "pDHCuentas=""") Then Exit Function
    
    MontaSQL = GeneraBalancePresupuestario
           
End Function



Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    If Not ComprobarCuentas(6, 7) Then Exit Function
    
    SQL = ""
    For i = 1 To Me.Check1.Count
        If Me.Check1(i).Value Then SQL = SQL & "&"
    Next i
    If Len(SQL) <> 1 Then
        If chkPreMensual.Value = 1 Then
            MsgBox "Seleccione uno, y solo uno, de los niveles contables.", vbExclamation
            Exit Function
        End If
    End If
    
    
    If txtMes(2).Text <> "" And Me.chkPreMensual.Value = 0 Then
        
        MsgBox "Si indica el mes debe marcar la opcion ""mensual""", vbExclamation
        Exit Function
    End If
    
    If txtMes(2).Text <> "" Then
        If Val(txtMes(2).Text) < 1 Or Val(txtMes(2).Text) > 12 Then
            MsgBox "Mes incorrecto: " & txtMes(2).Text, vbExclamation
            Exit Function
        End If
    End If
    

    DatosOK = True

End Function

Private Sub CargarComboFecha()
Dim J As Integer

    
    For i = 1 To vEmpresa.numnivel - 1
        J = DigitosNivel(i)
        Check1(i).Visible = True
        Check1(i).Caption = "Digitos: " & J
    Next i

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



Private Function TieneCuentasEnTmpBalance(DigitosNivel As String) As Boolean
Dim Rs As ADODB.Recordset
Dim C As String

    Set Rs = New ADODB.Recordset
    TieneCuentasEnTmpBalance = False
    C = Mid("__________", 1, CInt(DigitosNivel))
    C = "Select count(*) from tmpbalancesumas  where cta like '" & C & "'"
    C = C & " AND codusu = " & vUsu.Codigo
    Rs.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            If Rs.Fields(0) > 0 Then TieneCuentasEnTmpBalance = True
        End If
    End If
    Rs.Close
End Function

Private Sub PonerNiveles()
Dim i As Integer
Dim J As Integer


    Frame2.Visible = True
    Check1(10).Visible = True
    For i = 1 To vEmpresa.numnivel - 1
        J = DigitosNivel(i)
        Cad = "Digitos: " & J
        Check1(i).Visible = True
        Me.Check1(i).Caption = Cad
    Next i
    
    For i = vEmpresa.numnivel To 9
        Check1(i).Visible = False
    Next i
    
    
End Sub


Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

