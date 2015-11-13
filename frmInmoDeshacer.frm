VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInmoDeshacer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Icon            =   "frmInmoDeshacer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrDeshacer 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton cmdDeshaz 
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
         Index           =   1
         Left            =   3000
         TabIndex        =   2
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdDeshaz 
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
         Index           =   0
         Left            =   4320
         TabIndex        =   1
         Top             =   2040
         Width           =   1095
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   5280
         TabIndex        =   5
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
      Begin VB.Label Label13 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   6
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   5175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Deshacer última amortización"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   23
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5205
      End
   End
End
Attribute VB_Name = "frmInmoDeshacer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 510


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


Private Sub cmdDeshaz_Click(Index As Integer)
    If Index = 1 Then
        'Hacemos deshacer
        Cad = "¿Seguro que desea deshacer la última amortizacion con fecha: " & Format(UltAmor, "dd/mm/yyyy")
        If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Screen.MousePointer = vbHourglass
        Set Rs = New ADODB.Recordset
        
        Me.Tag = Label13(6).Caption
        DeshacerUltimaAmortizacion
        'Ha habido error
        If Me.cmdDeshaz(1).Enabled Then
'            Label13(6).Caption = Me.Tag
        Else
'            Me.cmdDeshaz(0).Caption = "Salir"
            Unload Me
        End If
        Set Rs = Nothing
        Screen.MousePointer = vbDefault
    Else
        Unload Me
    End If
    
End Sub





'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++


Private Sub Form_Activate()
If PrimeraVez Then
    PrimeraVez = False

    'Deshacer ultima amortizacion
    CargarDatosAmortizacion
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    Me.Icon = frmPpal.Icon

    Set miTag = New CTag
    Limpiar Me
    PrimeraVez = True
    
    FrDeshacer.Visible = True
    Me.Width = FrDeshacer.Width + 150
    Me.Height = FrDeshacer.Height + 500
    Caption = "Deshacer"
    
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
    'If Simulacion Then
    '     txtFecha.Text = Format(RC, "dd/mm/yyyy")
    'Else
    '     txtFecAmo.Text = Format(RC, "dd/mm/yyyy")
    '     'Dejamos cambiar la fecha, si , y solo si, es administrador
    '     txtFecAmo.Enabled = vUsu.Nivel < 2
        
    'End If
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


Private Function CargarDatos() As Boolean
On Error GoTo ECargarDatos
    CargarDatos = False
    Set Rs = New ADODB.Recordset
    Cad = "Select * from paramamort where codigo=1"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        CargarDatos = True
        '------------------  Ponemos los datos
    End If
    Rs.Close
ECargarDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando parametros"
    Set Rs = Nothing
End Function


Private Sub Form_Unload(Cancel As Integer)
    Set miTag = Nothing
End Sub



'-----------------------------------------------------------------------------
'
'
'       Deshacer ultima amortizacion
'
Private Sub CargarDatosAmortizacion()

    'Obtengo la ultima fecha a partir de la amortizacion y ultima fecha amortizada
    UltAmor = "01/01/1901"
    Cad = DevuelveDesdeBD("ultfecha", "paramamort", "codigo", "1", "N")
    If Cad <> "" Then UltAmor = CDate(Cad)
    
    If ObtenerparametrosAmortizacion(DivMes, UltAmor, ParametrosContabiliza) Then
        AUX2 = Format(UltAmor, "dd/mm/yyyy")
        B = True
    Else
        B = False
        AUX2 = "### ERROR obten. fecha ###"
    End If
    Cad = "Fecha última amortización:"
    Cad = Cad & Space(10) & AUX2
    Label13(6).Caption = Cad
              
    'Habilitamos o no el boton de deshacer
    cmdDeshaz(1).Enabled = B
End Sub



Private Sub DeshacerUltimaAmortizacion()

    'Constara de varios pasos
    '-------------------------------------------------------------------------------
    'Algunas comprobaciones. Ejercicios contables, que nos se ha vendido ni dado de baja....
    If Not Datosok_Deshacer Then Exit Sub
    


    'Deshacemos en inmovele_his y en inmovele. En los inmovilizados propiamente dicho
    'Transaccionamos esta accion
    PreparaBloquear
    Conn.BeginTrans
    
    If EliminarAmortizacion Then
        B = True
        Conn.CommitTrans
        Me.cmdDeshaz(1).Enabled = False
        'Grabamos el LOG
        Cad = "Fecha ult amortizacion: " & UltAmor
        vLog.Insertar 14, vUsu, Cad

    Else
        B = False
        Conn.RollbackTrans
    End If
    TerminaBloquear
    'Si da error nos piramos
    If Not B Then MsgBox "Se han producido errores.", vbExclamation
        

End Sub



Private Function EliminarAmortizacion() As Boolean
Dim Valor As Currency
Dim F As Date
Dim SQL As String

    On Error GoTo EEliminarAmortizacion

    EliminarAmortizacion = False
    
    Label13(6).Caption = "Comprobar datos"
    Me.Refresh
    DoEvents
    
    'Compreubo cuantos hay. Para que no haya errores
    Cad = "Select count(*) from inmovele_his where fechainm = '" & Format(UltAmor, FormatoFecha) & "'"
    Cont = 0
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then Cont = DBLet(Rs.Fields(0), "N")
    Rs.Close
    AUX2 = CStr(Cont)
    
    If Cont = 0 Then
        MsgBox "Error: NUmero de registos de hcoinmovilizado con fecha " & UltAmor & " es cero", vbExclamation
        Exit Function
    End If
    
    'Abro el rs para actualizar
    Cad = "select l.codinmov,imporinm,amortacu,valoradq,nominmov from inmovele_his l,inmovele where l.codinmov=inmovele.codinmov "
    Cad = Cad & " and fechainm = '" & Format(UltAmor, FormatoFecha) & "'"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cont = 0
    While Not Rs.EOF
        Label13(6).Caption = Rs!Codinmov & " " & Rs!nominmov
        Label13(6).Refresh
        
        'Para cada elemento le sumo lo que ha amortizado
        Importe = DBLet(Rs!amortacu, "N")
        Importe = Importe - Rs!imporinm
        
        'Control auxiliar
        If Importe < 0 Then Importe = 0
        
        'Creo SQL update
        Cad = "UPDATE inmovele set amortacu=" & TransformaComasPuntos(CStr(Importe))
        Cad = Cad & ", situacio= 1"
        Cad = Cad & " WHERE codinmov=" & Rs!Codinmov
        
        'Muevo al siguiente
        Rs.MoveNext
        'Updateo
        Conn.Execute Cad
        'cont++
        Cont = Cont + 1
        
        
    Wend
    Rs.Close
    
    
    If Cont <> Val(AUX2) Then
        'ERROR. Iban a ser val(aux2)  registros y solo se han preocesado cont
        Cad = "Registros del count(*)= " & AUX2 & vbCrLf & "Registros procesados= " & Cont
        Cad = "Error. " & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Exit Function
    End If
    
    Label13(6).Caption = "Restaurando datos situacion anterior"
    Label13(6).Refresh
    
    
    'Borramos todos los datos de inmovele_his con esta fecha
    Cad = "DELETE from inmovele_his where fechainm = '" & Format(UltAmor, FormatoFecha) & "'"
    Conn.Execute Cad
    
    'ACtualizamos la fecha de ultamor
    '--------------------------------
    AUX2 = "tipoamor"
    Cad = DevuelveDesdeBD("intcont", "paramamort", "codigo", "1", "N", AUX2)
    Contabiliza = (Cad = 1)
    DivMes = Val(AUX2)
    Select Case DivMes
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
    F = DateAdd("m", -i, UltAmor)
    i = DiasMes(CByte(Month(UltAmor)), Year(UltAmor))
    If i = Day(UltAmor) Then
        'Es ultimo dia mes
        'Leugo la fecha sera el ultimo dia de mes
        i = DiasMes(CByte(Month(F)), Year(F))
        F = CDate(i & "/" & Month(F) & "/" & Year(F))
        
    End If
    Cad = Format(F, FormatoFecha)
    Cad = "UPDATE paramamort set ultfecha='" & Cad & "'"
    Conn.Execute Cad
    
    

    If Not Contabiliza Then
        'Proceso finalizado con exito. No busco el asiento
        AUX2 = "Proceso finalizado correctamente"
    Else
        'Si contabiliza tratamos de indicarle cual fue el asiento generado.
        'Busco el cabapu que cuadra con fechaent='uktamor' y en observaciones lleva amortizacion
        Cad = "hlinapu where fechaent = '" & Format(UltAmor, FormatoFecha) & "' AND idcontab = 'CONTAI'"
        Cad = Cad & " and (numdiari, fechaent, numasien) in (select numdiari, fechaent, numasien from hcabapu where fechaent = " & DBSet(UltAmor, "F") & " and obsdiari like '%mortiza%')"
        Cont = 0
        'En introduccion
        AUX2 = "Select * from " & Cad
        Rs.Open AUX2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            'LO HE ENCONTRADO
            Cont = 1
            Cad = "Asiento: " & Rs!NumAsien & "      Diario: " & Rs!NumDiari & "      Fecha: " & Rs!FechaEnt & vbCrLf '& "Observaciones: " & DBMemo(Rs!obsdiari)
        End If
        
        'Si cont>0 entonces SI que lo ha encontrado
        
        If Cont > 0 Then
            If MsgBox("Se ha encontrado el asiento de contabilización de la amortización." & vbCrLf & vbCrLf & "¿ Desea eliminarlo ?" & vbCrLf, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                'lineas
                SQL = "delete from hlinapu where numdiari = " & DBSet(Rs!NumDiari, "N") & " and numasien = " & DBSet(Rs!NumAsien, "N") & " and fechaent = " & DBSet(Rs!FechaEnt, "F")
                Conn.Execute SQL
                
                'cabecera
                SQL = "delete from hcabapu where numdiari = " & DBSet(Rs!NumDiari, "N") & " and numasien = " & DBSet(Rs!NumAsien, "N") & " and fechaent = " & DBSet(Rs!FechaEnt, "F")
                Conn.Execute SQL
                
                Cad = "Eliminado el " & Cad
            Else
                Cad = "No se ha eliminado el asiento de contabilización"
            End If
        Else
            Cad = "El asiento NO ha sido encontrado"
        End If
        Rs.Close
        
        AUX2 = "Proceso finalizado correctamente." & vbCrLf & vbCrLf & vbCrLf & Cad
    End If
    
    MsgBox AUX2, vbInformation


    Label13(6).Caption = "" 'AUX2

    EliminarAmortizacion = True
    Exit Function
EEliminarAmortizacion:
    MuestraError Err.Number, "Eliminar Amortización", Err.Description
End Function


Private Function Datosok_Deshacer() As Boolean
    On Error GoTo Edatosok_deshacer
    Datosok_Deshacer = False


    varFecOk = FechaCorrecta2(UltAmor)
    Cad = ""
    If varFecOk > 1 Then
        If varFecOk = 2 Then
            Cad = Mid(varTxtFec, 6)
        Else
            Cad = " fuera de ejercicios. "
        End If
    End If
    If Cad <> "" Then
        Cad = "Fecha última amortizacion " & LCase(Cad)
        MsgBox Cad, vbExclamation
        Exit Function
    End If

    Cad = "select distinct(inmovele_his.codinmov) from inmovele_his, inmovele where inmovele_his.codinmov=inmovele.codinmov and"
    Cad = Cad & " fechainm>='" & Format(UltAmor, FormatoFecha) & "'  and fecventa >='" & Format(UltAmor, FormatoFecha) & "'"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cont = 0
    While Not Rs.EOF
        Cont = Cont + 1
        Rs.MoveNext
    Wend
    Rs.Close
    
    If Cont > 0 Then
        Cad = "Hay " & Cont & " elemento(s) de inmovilizado que están en el hco inmovilizado  y han sido vendidos o dados de baja"
        MsgBox Cad, vbExclamation
        Exit Function
    End If
    
    
    Cad = "select distinct(inmovele_his.codinmov) from inmovele_his where  fechainm > '" & Format(UltAmor, FormatoFecha) & "'"
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cont = 0
    While Not Rs.EOF
        Cont = Cont + 1
        Rs.MoveNext
    Wend
    Rs.Close
    
    If Cont > 0 Then
        Cad = "Hay " & Cont & " elemento(s) de inmovilizado que están en el hco inmovilizado."
        MsgBox Cad, vbExclamation
        Exit Function
    End If

    Datosok_Deshacer = True
    Exit Function
Edatosok_deshacer:
    MuestraError Err.Number, Err.Description
End Function

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub
