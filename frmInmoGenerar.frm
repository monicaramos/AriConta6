VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInmoGenerar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   Icon            =   "frmInmoGenerar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton CmdCancel 
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
         Left            =   3630
         TabIndex        =   5
         Top             =   2940
         Width           =   1275
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   3720
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdCalcula 
         Caption         =   "Aceptar"
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
         Left            =   2130
         TabIndex        =   3
         Top             =   2940
         Width           =   1275
      End
      Begin VB.TextBox txtFecAmo 
         Alignment       =   2  'Center
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
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text4"
         Top             =   1590
         Width           =   1695
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   4680
         TabIndex        =   6
         Top             =   330
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
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   240
         Index           =   2
         Left            =   2400
         Picture         =   "frmInmoGenerar.frx":000C
         Top             =   1620
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha amortizacion"
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
         Left            =   360
         TabIndex        =   2
         Top             =   1620
         Width           =   2025
      End
   End
End
Attribute VB_Name = "frmInmoGenerar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 509

Public opcion As Byte
    '0.- Parametros
    '1.- Simular
    '2.- Cálculo amort.
    '3.- Venta/Baja inmovilizado
    '---------------------------
    'los siguiente utilizan el mismo frame, con opciones
    '4.- Listado estadisticas
    '5.- Ficha elementos
    '6.- Entre fechas


    '10.- Deshacer ultima amortizacion

Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmD As frmTiposDiario
Attribute frmD.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim Rs As Recordset
Dim Cad As String
Dim i As Byte
Dim B As Boolean
Dim Importe As Currency
'
'Desde parametros
Dim Contabiliza As Boolean
Dim UltAmor As Date
Dim DivMes As Integer
Dim ParametrosContabiliza As String
Dim Mc As Contadores

'Tipo de IVA
Dim TipoIva As String
Dim AUX2 As String


'Contador para las lineas de apuntes
Dim Cont As Integer

Private Function ActivadoParametro()
Dim SQL As String

    SQL = "select intcont from paramamort "
    ActivadoParametro = (DevuelveValor(SQL) = 1)

End Function


Private Sub cmdCalcula_Click()

    If Not ActivadoParametro Then
        If MsgBox("No tiene activada la Contabilización Automática de la Amortización." & vbCrLf & vbCrLf & " ¿ Desea continuar ? " & vbCrLf, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
    End If

    If MsgBox("Seguro que desea realizar la amortización a fecha: " & txtFecAmo.Text & " ?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    If txtFecAmo.Text = "" Then
        MsgBox "Fecha incorrecta", vbExclamation
        Exit Sub
    End If
    If Me.Tag <> "" Then
        If CDate(Me.txtFecAmo.Text) < CDate(Me.Tag) Then
            MsgBox "Fecha no puede ser menor que la ultima fecha de amortizacion: " & Me.Tag, vbExclamation
            Exit Sub
        End If
    End If
    i = FechaCorrecta2(CDate(txtFecAmo.Text))
    If i > 1 Then
        If i = 2 Then
            MsgBox varTxtFec, vbExclamation
        Else
            If i = 2 Then
                MsgBox "Fecha de amortización pertence a un ejercicio cerrado.", vbExclamation
            Else
                MsgBox "Fecha amortización pertenece a un ejercicio todavía no abierto", vbExclamation
            End If
        End If
        Exit Sub
    End If
    'Leemos los parametros
    If ObtenerparametrosAmortizacion(DivMes, UltAmor, ParametrosContabiliza) = False Then Exit Sub
    Contabiliza = RecuperaValor(ParametrosContabiliza, 1) = "1"
    'Si contabilizamos hay k conseguir el numero de asiento
    Set Mc = New Contadores
    If Contabiliza Then
        B = (Mc.ConseguirContador("0", (i = 0), True) = 0)
    Else
        B = True
    End If
    
    If B Then
        Screen.MousePointer = vbHourglass
        
        'Grabamos el LOG
        Cad = "Fecha amortización: " & txtFecAmo.Text
        If Mc.Contador > 0 Then Cad = Cad & " Asiento asignado: " & Mc.Contador
        vLog.Insertar 13, vUsu, Cad

        
        
        PreparaBloquear
            Conn.BeginTrans
            Cad = "Select * from inmovele where   inmovele.fecventa is null and inmovele.valoradq > inmovele.amortacu and situacio=1"
            'Fecha adq
            Cad = Cad & " and fechaadq <='" & Format(CDate(txtFecAmo.Text), FormatoFecha) & "'"
            Cad = Cad & " for update "
            B = GeneraCalculoInmovilizado(Cad, 2)
            If B Then
                Conn.CommitTrans
            Else
                Conn.RollbackTrans
            End If
        TerminaBloquear
        pb1.Visible = False
        Screen.MousePointer = vbDefault
        If B Then
            'ha ido bien
            MsgBox "El cálculo se ha realizado con éxito. En introducción de Asientos está el asiento generado.", vbExclamation
            Set Mc = Nothing
            Unload Me
            Exit Sub
        Else
            If Contabiliza Then Mc.DevolverContador "0", (i = 0), Mc.Contador
        End If
    End If
    Set Mc = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++


Private Sub Form_Activate()
If PrimeraVez Then
    PrimeraVez = False

End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    Me.Icon = frmPpal.Icon

    Set miTag = New CTag
    Limpiar Me
    pb1.Visible = False
    PrimeraVez = True
    
    
    Frame2.Visible = False
    Select Case opcion
    Case 2
        txtFecAmo.Text = SugerirFechaNuevo
        txtFecAmo.Enabled = vUsu.Nivel < 2
        Frame2.Visible = True
        Me.Width = Frame2.Width + 150
        Me.Height = Frame2.Height + 150
        Caption = "Cálculo y contabilización amortización"
    End Select
        
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
        
End Sub

Private Function SugerirFechaNuevo() As String
Dim RC As String
    RC = "tipoamor"
    Cad = DevuelveDesdeBD("ultfecha", "paramamort", "codigo", "1", "N", RC)

    If Cad <> "" Then
        Me.Tag = Cad   'Ultima actualizacion
        Select Case Val(RC)
        Case 2
            'Semestral
            i = 6
            'Siempre es la ultima fecha de mes
        Case 3
            'Trimestral
            i = 3
        Case 4
            'Mensual
            i = 1
        Case Else
            'Anual
            i = 12
        End Select
        RC = PonFecha
    Else
        Cad = "01/01/1991"
        RC = Format(Now, "dd/mm/yyyy")
    End If
    SugerirFechaNuevo = Format(RC, "dd/mm/yyyy")
    
End Function



Private Function PonFecha() As Date
Dim d As Date
'Dada la fecha en Cad y los meses k tengo k sumar
'Pongo la fecha
d = DateAdd("m", i, CDate(Cad))
Select Case Month(d)
Case 2
    If ((Year(d) - 2000) Mod 4) = 0 Then
        i = 29
    Else
        i = 28
    End If
Case 1, 3, 5, 7, 8, 10, 12
    '31
        i = 31
Case Else
    '30
        i = 30
End Select
Cad = i & "/" & Month(d) & "/" & Year(d)
PonFecha = CDate(Cad)
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set miTag = Nothing
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Cad = Format(vFecha, "dd/mm/yyyy")
    Select Case i
    Case 2
        txtFecAmo.Text = Cad
    End Select
End Sub

Private Sub Image1_Click(Index As Integer)
    Set frmF = New frmCal
    frmF.Fecha = Now
    i = Index
    Select Case Index
    Case 2
        If txtFecAmo.Text <> "" Then frmF.Fecha = CDate(txtFecAmo.Text)
    End Select
    frmF.Show vbModal
    Set frmF = Nothing
End Sub


Private Function ParaBD(ByRef T As TextBox) As String
If T.Text = "" Then
    ParaBD = "NULL"
Else
    ParaBD = T.Text
End If
End Function


Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub


Private Sub txtFecAmo_GotFocus()
With txtFecAmo
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtFecAmo_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtFecAmo_KeyPress(KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        KEYFecAmo KeyAscii
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub




Private Sub txtFecAmo_LostFocus()
With txtFecAmo
    .Text = Trim(.Text)
    If .Text = "" Then Exit Sub
    If Not EsFechaOK(txtFecAmo) Then
        MsgBox "Fecha incorrecta: " & .Text, vbExclamation
        .Text = ""
        .SetFocus
    End If
End With
End Sub


Private Sub KEYFecAmo(KeyAscii As Integer)
    KeyAscii = 0
    Image1_Click (2)
End Sub


'++


'TIPO:
'       0.- Venta
'       1.- Baja
'       2.- Calculo de amortizacion
Private Function GeneraCalculoInmovilizado(ByRef SeleccionInmovilizado As String, Tipo As Byte) As Boolean
Dim Codinmov As Long
Dim B As Boolean
On Error GoTo EGen

    GeneraCalculoInmovilizado = False
    If Tipo = 2 Then
        'Para el calculo del amortizado
        Set Rs = New ADODB.Recordset
        Rs.Open SeleccionInmovilizado, Conn, adOpenKeyset, adLockPessimistic, adCmdText
        If Rs.EOF Then
            MsgBox "Ningun registro", vbExclamation
            Rs.Close
            Exit Function
        End If
    End If
    'Vemos cuantos hay
    Cont = 0
    While Not Rs.EOF
        Cont = Cont + 1
        Rs.MoveNext
    Wend
    Rs.MoveFirst
    If Cont > 3 Then pb1.Visible = True
    pb1.Max = Cont + 1
    pb1.Value = 0
    
    
    
    'Vemos si contabilizamos
    'Insertamos cabecera del asiento
    If Contabiliza Then GeneracabeceraApunte (Tipo)
    Cont = 1
    While Not Rs.EOF
        Codinmov = Rs!Codinmov
       
        'La fecha depende si estamos calculando normal o estamos vendiendo
        If opcion = 3 Then
'            Cad = Text4(0).Text
        Else
            Cad = Me.txtFecAmo.Text
        End If
      
        B = CalculaAmortizacion(Codinmov, CDate(Cad), DivMes, UltAmor, ParametrosContabiliza, Mc.Contador, Cont, Tipo < 2)
        If Not B Then
            Rs.Close
            Exit Function
        End If
        
        'Siguiente
        pb1.Value = pb1.Value + 1
        Cont = Cont + 1
        Rs.MoveNext
    Wend
    'Actualizamos la fecha de ultima amortizacion en paraemtros
    If opcion <> 3 Then
        Cad = "UPDATE paramamort SET ultfecha= '" & Format(Cad, FormatoFecha)
        Cad = Cad & "' WHERE codigo=1"
        Conn.Execute Cad
        Rs.Close
    Else
        'Estamos dando de baja o vendiendo un inmovilizado. Solo hay uno y hay k situarlo
        'en el primero
        Rs.Requery
        Rs.MoveFirst
    End If
    GeneraCalculoInmovilizado = True
    Exit Function
EGen:
    MuestraError Err.Number
End Function


Private Function GeneracabeceraApunte(vTipo As Byte) As Boolean
Dim Fecha As Date
On Error GoTo EGeneracabeceraApunte
        GeneracabeceraApunte = False
        Cad = "INSERT INTO hcabapu (numdiari, fechaent, numasien,  obsdiari) VALUES ("
        Cad = Cad & RecuperaValor(ParametrosContabiliza, 4) & ",'"
        If opcion = 3 Then
'            Fecha = CDate(Text4(0).Text)
        Else
            Fecha = CDate(txtFecAmo.Text)
        End If
        Cad = Cad & Format(Fecha, FormatoFecha)
        Cad = Cad & "'," & Mc.Contador
        Cad = Cad & ",'"
        'Segun sea VENTA, BAJA, o calculo de inmovilizado pondremos una cosa u otra
        Select Case vTipo
        Case 0, 1
            'VENTA
            If vTipo = 0 Then
                Cad = Cad & "Venta de "
            Else
                Cad = Cad & "Baja de "
            End If
            Cad = Cad & DevNombreSQL(Rs!nominmov)
        Case Else
            Cad = Cad & "Amortización: " & Fecha
        End Select
        Cad = Cad & "')"
        Conn.Execute Cad
        GeneracabeceraApunte = True
        Exit Function
EGeneracabeceraApunte:
     MuestraError Err.Number, "Genera cabecera Apunte"
     Set Rs = Nothing
End Function

