VERSION 5.00
Begin VB.Form frmTESGastosFijos2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTESGastosFijos2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAltaGastoFijo 
      Caption         =   "Alta Gasto Fijo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4545
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10395
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   0
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   38
         Tag             =   "imgConcepto"
         Top             =   450
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   5
         Left            =   2670
         TabIndex        =   19
         Top             =   3300
         Width           =   1365
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   5
         Left            =   2670
         TabIndex        =   18
         Text            =   "99/99/9999"
         Top             =   2130
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   1
         Left            =   4020
         MaxLength       =   50
         TabIndex        =   17
         Tag             =   "Descripción|T|N|||remesas|descripción|||"
         Top             =   450
         Width           =   6045
      End
      Begin VB.TextBox txtCuentas 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   0
         Left            =   2670
         TabIndex        =   16
         Top             =   1020
         Width           =   1275
      End
      Begin VB.TextBox txtCuentas 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   1
         Left            =   2700
         TabIndex        =   15
         Top             =   1560
         Width           =   1275
      End
      Begin VB.TextBox txtNCuentas 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   0
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1020
         Width           =   6015
      End
      Begin VB.TextBox txtNCuentas 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   1
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1560
         Width           =   5985
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   8820
         TabIndex        =   12
         Top             =   3810
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptarAltaCab 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   7530
         TabIndex        =   11
         Top             =   3840
         Width           =   1155
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         Index           =   0
         ItemData        =   "frmTESGastosFijos2.frx":000C
         Left            =   2670
         List            =   "frmTESGastosFijos2.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2730
         Width           =   2100
      End
      Begin VB.Label Label3 
         Caption         =   "Código"
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
         Height          =   315
         Index           =   4
         Left            =   210
         TabIndex        =   39
         Top             =   480
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Importe"
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
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   25
         Top             =   3360
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Periodicidad"
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
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   5
         Left            =   2310
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de gasto"
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
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   23
         Top             =   2190
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta Prevista"
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
         Height          =   285
         Index           =   6
         Left            =   210
         TabIndex        =   22
         Top             =   1020
         Width           =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción"
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
         Height          =   315
         Index           =   8
         Left            =   2730
         TabIndex        =   21
         Top             =   450
         Width           =   1380
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   0
         Left            =   2310
         Top             =   1050
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   1
         Left            =   2310
         Top             =   1590
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Contrapartida"
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
         Height          =   285
         Index           =   11
         Left            =   210
         TabIndex        =   20
         Top             =   1620
         Width           =   1890
      End
   End
   Begin VB.Frame FrameAltaModLineaGasto 
      Caption         =   "Alta/modificacion Línea Gasto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6675
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   0
         Left            =   2550
         TabIndex        =   4
         Text            =   "99/99/9999"
         Top             =   1140
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   360
         Index           =   4
         Left            =   2580
         TabIndex        =   3
         Text            =   "99/99/9999"
         Top             =   1710
         Width           =   1245
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   5280
         TabIndex        =   2
         Top             =   2910
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   3840
         TabIndex        =   1
         Top             =   2910
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Linea a modificar"
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
         Height          =   315
         Index           =   9
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   5460
      End
      Begin VB.Label lblFecha1 
         Height          =   255
         Index           =   1
         Left            =   2580
         TabIndex        =   7
         Top             =   3990
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de gasto"
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
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2190
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Importe"
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
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   1770
         Width           =   1935
      End
   End
   Begin VB.Frame FrameModGastoFijo 
      Caption         =   "Modificacion Gasto Fijo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   10395
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   7530
         TabIndex        =   33
         Top             =   3060
         Width           =   1155
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   8820
         TabIndex        =   32
         Top             =   3060
         Width           =   1095
      End
      Begin VB.TextBox txtNCuentas 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   3
         Left            =   3810
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2190
         Width           =   6015
      End
      Begin VB.TextBox txtNCuentas 
         BackColor       =   &H80000018&
         Height          =   360
         Index           =   2
         Left            =   3810
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1650
         Width           =   6015
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   10
         Left            =   2490
         TabIndex        =   29
         Top             =   2190
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   9
         Left            =   2460
         TabIndex        =   28
         Top             =   1650
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   8
         Left            =   2460
         MaxLength       =   50
         TabIndex        =   27
         Tag             =   "Descripción|T|N|||remesas|descripción|||"
         Top             =   1020
         Width           =   7365
      End
      Begin VB.Label Label3 
         Caption         =   "Código"
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
         Height          =   315
         Index           =   3
         Left            =   300
         TabIndex        =   37
         Top             =   480
         Width           =   2280
      End
      Begin VB.Label Label3 
         Caption         =   "Contrapartida"
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
         Height          =   285
         Index           =   2
         Left            =   300
         TabIndex        =   36
         Top             =   2250
         Width           =   1770
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   3
         Left            =   2100
         Top             =   2220
         Width           =   255
      End
      Begin VB.Image imgCuentas 
         Height          =   255
         Index           =   2
         Left            =   2100
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción"
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
         Height          =   315
         Index           =   1
         Left            =   270
         TabIndex        =   35
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta Prevista"
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
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   34
         Top             =   1650
         Width           =   1860
      End
   End
End
Attribute VB_Name = "frmTESGastosFijos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '1.- Alta cabecera gasto fijo
    '2.- Modificacion cabecera gasto fijo
    '3.- Alta linea gasto fijo
    '4.- Modificacion linea gasto fijo
    '5.- Contabilizacion del gasto
    
    
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
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private PrimeraVez As Boolean

Dim I As Integer
Dim SQL As String
Dim RS As Recordset
Dim ItmX As ListItem
Dim Errores As String
Dim NE As Integer
Dim Ok As Integer

Dim CampoOrden As String
Dim Orden As Boolean
Dim Indice As Integer


Private Function DatosOK() As Boolean
Dim B As Boolean
Dim Cta As String

    DatosOK = False
    
    Select Case Opcion
        Case 1 ' insertar cabecera
            If Text1(1).Text = "" Then
                MsgBox "Debe introducir el concepto.", vbExclamation
                PonFoco Text1(1)
                Exit Function
            End If
            If txtCuentas(0).Text = "" Then
                MsgBox "Debe introducir una cuenta prevista. Reintroduzca.", vbExclamation
                PonFoco txtCuentas(0)
                Exit Function
            Else
                Cta = (txtCuentas(0).Text)
                                    '********
                B = CuentaCorrectaUltimoNivelSIN(Cta, SQL)
                If B = 0 Then
                    MsgBox "NO existe la cuenta: " & txtCuentas(0).Text, vbExclamation
                    PonFoco txtCuentas(0)
                    Exit Function
                End If
            End If
            
            
            If txtCuentas(1).Text <> "" Then
                Cta = (txtCuentas(1).Text)
                                    '********
                B = CuentaCorrectaUltimoNivelSIN(Cta, SQL)
                If B = 0 Then
                    MsgBox "NO existe la cuenta: " & txtCuentas(1).Text, vbExclamation
                    PonFoco txtCuentas(1)
                    Exit Function
                End If
            End If
            
            
            If txtFecha(5).Text = "" Then
                MsgBox "Debe introducir una fecha de gasto.", vbExclamation
                PonFoco txtFecha(5)
                Exit Function
            End If
            If Combo1(0).ListIndex = -1 Then
                MsgBox "Debe introducir una periodicidad", vbExclamation
                PonerFocoCmb Combo1(0)
                Exit Function
            End If
    End Select
    
    DatosOK = True
    
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub





Private Sub cmdAceptarAltaCab_Click()


    If Not DatosOK Then Exit Sub
    
    If GenerarGasto Then
        MsgBox "Proceso realizado correctamente", vbExclamation
        Unload Me
    End If


End Sub


Private Function GenerarGasto() As Boolean
Dim Perio As Integer
Dim nVeces As Integer
Dim SqlValues As String
Dim SqlInsert As String
Dim Fecha As Date
Dim NumGasto As Long

    On Error GoTo eGenerarGasto

    GenerarGasto = False
    
    Conn.BeginTrans
    
    NumGasto = SugerirCodigoSiguiente
    
    SQL = "insert into gastosfijos (codigo, descripcion ,ctaprevista,contrapar) values ( " & DBSet(NumGasto, "N") & ","
    SQL = SQL & DBSet(Text1(1), "T") & "," & DBSet(txtCuentas(0).Text, "T") & "," & DBSet(txtCuentas(1).Text, "T") & ")"
    
    Conn.Execute SQL
    
    Select Case Combo1(0).ListIndex
        Case 1 ' mensual
            Perio = 1
            nVeces = 12
        Case 2 ' bimensual
            Perio = 2
            nVeces = 6
        Case 3 ' trimestral
            Perio = 3
            nVeces = 4
        Case 4 ' semestral
            Perio = 6
            nVeces = 2
        Case 5 ' anual
            Perio = 12
            nVeces = 1
    End Select
    
    SqlInsert = "insert into gastosfijos_recibos(codigo, fecha, importe, contabilizado) values "
    SqlValues = ""
    
    Fecha = CDate(txtFecha(5).Text)
    
    For I = 1 To nVeces
        SqlValues = "(" & DBSet(NumGasto, "N") & "," & DBSet(Fecha, "F") & "," & DBSet(Text1(5).Text, "N") & "0),"
        
        Fecha = DateAdd("m", Perio, Fecha)
    Next I
    
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        Conn.Execute SqlInsert & SqlValues
    End If
    
    Conn.CommitTrans
    GenerarGasto = True
    Exit Function

eGenerarGasto:
    Conn.RollbackTrans
    MuestraError Err.Number, "Generar Gasto", Err.Description
End Function

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
            Case 1 ' alta cabecera
                Text1(0).Text = SugerirCodigoSiguiente
            Case 2  ' modificacion cabecera
            
            Case 3 ' alta linea
            
            Case 4 ' modificacion linea
            
            Case 5 ' contabilizacion del gasto
            
        
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Function SugerirCodigoSiguiente() As String
    Dim SQL As String
    Dim RS As ADODB.Recordset
    
    SQL = "Select Max(codigo) from gastosfijos"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, , , adCmdText
    SQL = "1"
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            SQL = CStr(RS.Fields(0) + 1)
        End If
    End If
    RS.Close
    SugerirCodigoSiguiente = SQL
End Function


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
    PrimeraVez = True
    
    Me.imgCuentas(2).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Me.imgCuentas(3).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    
    Me.imgFec(0).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Me.imgFec(1).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    
    
    
    Me.FrameAltaGastoFijo.Visible = False
    Me.FrameModGastoFijo.Visible = False
    Me.FrameAltaModLineaGasto.Visible = False
    
    
    Select Case Opcion
    Case 1
        Me.Caption = "Nuevo Gasto Fijo"
        W = Me.FrameAltaGastoFijo.Width
        H = Me.FrameAltaGastoFijo.Height
        Me.FrameAltaGastoFijo.Visible = True
        
        CargarCombo

    Case 2
        Me.Caption = "Modificación Gasto Fijo"
        W = Me.FrameModGastoFijo.Width
        H = Me.FrameModGastoFijo.Height + 150
        Me.FrameModGastoFijo.Visible = True
    
    Case 3, 4
        Me.Caption = "Nueva Linea de Gasto"
        W = Me.FrameAltaModLineaGasto.Width
        H = Me.FrameAltaModLineaGasto.Height + 200
        Me.FrameAltaModLineaGasto.Visible = True
    End Select
    
    Me.Width = W + 120
    Me.Height = H + 120
End Sub









Private Sub imgFec_Click(Index As Integer)
    'FECHA FACTURA
    Indice = Index
    
    Set frmF = New frmCal
    frmF.Fecha = Now
    If txtFecha(Indice).Text <> "" Then frmF.Fecha = CDate(txtFecha(Indice).Text)
    frmF.Show vbModal
    Set frmF = Nothing
    PonFoco txtFecha(Indice)

End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim Cta As String
Dim B As Byte
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    If Text1(Index).Text = "" Then
        Exit Sub
    End If
    
    Select Case Index
        Case 5  ' importe
            PonerFormatoDecimal Text1(Index), 1
            
            
    End Select

End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub



Private Sub txtCuentas_GotFocus(Index As Integer)
    ConseguirFoco txtCuentas(Index), 3
End Sub

Private Sub txtCuentas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
'        LanzaFormAyuda txtCuentas(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub txtCuentas_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCuentas_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim Cta As String
Dim B As Boolean
Dim SQL As String
Dim Hasta As Integer   'Cuando en cuenta pongo un desde, para poner el hasta

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
    
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'cuentas
            Cta = (txtCuentas(Index).Text)
                                    '********
            B = CuentaCorrectaUltimoNivelSIN(Cta, SQL)
            If B = 0 Then
                MsgBox "NO existe la cuenta: " & txtCuentas(Index).Text, vbExclamation
                txtCuentas(Index).Text = ""
                txtNCuentas(Index).Text = ""
            Else
                txtCuentas(Index).Text = Cta
                txtNCuentas(Index).Text = SQL
                If B = 1 Then
                    txtNCuentas(Index).Tag = ""
                Else
                    txtNCuentas(Index).Tag = SQL
                End If
                Hasta = -1
                If Index = 6 Then
                    Hasta = 7
                Else
                    If Index = 0 Then
                        Hasta = 1
                    Else
                        If Index = 5 Then
                            Hasta = 4
                        Else
                            If Index = 23 Then Hasta = 24
                        End If
                    End If
                    
                End If
                    
                If Hasta >= 0 Then
                    txtCuentas(Hasta).Text = txtCuentas(Index).Text
                    txtNCuentas(Hasta).Text = txtNCuentas(Index).Text
                End If
            End If
    
    End Select
    
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



Private Sub EjecutarSQL()
    On Error Resume Next
    
    Conn.Execute SQL
    If Err.Number <> 0 Then
        If Conn.Errors(0).Number = 1062 Then
            Err.Clear
        Else
            'MuestraError Err.Number, Err.Description
        End If
        Err.Clear
    End If
End Sub


Private Sub frmF_Selec(vFecha As Date)
    txtFecha(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub CargarCombo()
Dim RS As ADODB.Recordset
Dim SQL As String
Dim J As Long
    
    Combo1(0).Clear

    Combo1(0).AddItem "Mensual "
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Bimensual "
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    Combo1(0).AddItem "Trimestral "
    Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    Combo1(0).AddItem "Semestral "
    Combo1(0).ItemData(Combo1(0).NewIndex) = 4
    Combo1(0).AddItem "Anual "
    Combo1(0).ItemData(Combo1(0).NewIndex) = 5



End Sub


