VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'ACCIONES DISPONIBLES
'
'
'

'
'       1.- Crear asiento manual
'       2.- Modificar asiento
'       3.- Eliminar asiento

'       4,5,6   FRACLI
'       7,8,9   FRAPRO

'       10.- Actualizar asientos
'       11.- Contabilizar FRACLI
'       12.- Contabilizar FRAPRO

'       13.- Realizar amortizacion
'       14.- Deshacer  " "
'       15.- 340 a fichero
'       16.- Acciones del CC. En la descripcioin IRA lo que he heco
'       17.- Punteo de extractos
'       18.-  proceso cierre [RENUMERAR] [CIERRE]  [DESHACER]

'       19.-  VARIOS.  Para quedar constancia pero no crear un "exproceso"

Private Const AccionesRegistradas = 19
Private mvarDatosDescripcion As String

Dim Sql As String



Public Property Let DatosDescripcion(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nivel = 5
    mvarDatosDescripcion = vData
End Property


Public Property Get DatosDescripcion() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nivel
    DatosDescripcion = mvarDatosDescripcion
End Property



Public Function Insertar(Accion As Byte, ByRef ElUsuario As Usuario, Descripcion As String) As Boolean


    On Error GoTo EI
    Insertar = False
    
    Sql = "insert into `slog` (`fecha`,`accion`,`usuario`,`pc`,`descripcion`) values ( "
    Sql = Sql & " now()," & Accion & " ,'" & DevNombreSQL(ElUsuario.Login) & "','"
    Sql = Sql & DevNombreSQL(ElUsuario.PC) & "',"
    If Descripcion = "" Then
        Sql = Sql & "NULL)"
    Else
        If Len(Descripcion) > 255 Then Descripcion = Mid(Descripcion, 1, 250) & "[MAS]"
        Sql = Sql & "'" & DevNombreSQL(Descripcion) & "')"
    End If
    Conn.Execute Sql
    
    
    Insertar = True
    Exit Function
EI:
    MuestraError Err.Number, "Registro LOG de acciones"

End Function

Private Function DimeAccion(N As Integer) As String
Dim H As Integer
Dim C As String

    If N <= 9 Then
        'las 9 primeras son mantenimientos basicos
        H = N Mod 3
        If H = 1 Then
            C = "Nuevo"
        ElseIf H = 2 Then
            C = "Modificar"
        Else
            C = "Eliminar"
        End If
        H = (N - 1) \ 3
        If H = 0 Then
            C = C & " Asiento"
        ElseIf H = 1 Then
            C = C & " Factura Cliente"
        Else
            C = C & " Factura Proveedor"
        End If
            
    Else
        'El resto de acciones con un select case
        'ACTUALIZACION (CONTABILIZACION)
        If N <= 12 Then
            C = "CONTAB "
            If N = 10 Then
                C = C & " Asiento"
            ElseIf N = 11 Then
                C = C & " Factura Cliente"
            Else
                C = C & " Factura Proveedor"
            End If
            
            
            
            
        Else
            'Aqui iran yendo resto acciones
            If N <= 16 Then
                Select Case N
                Case 13
                    C = "Amortización"
                Case 14
                    C = "Deshacer Amortización"
                Case 15
                    C = "Mod340"
'                Case 16
'                    C = "ControlCC"
                End Select
        
            Else
                'n=17
                If N = 17 Then
                    C = "Punteo Extracto"
                ElseIf N = 18 Then
                    'Llevara la primera linea como que se ha hecho(
                    '[CIERRE-RENUMERAR-DESHACER}
                    C = "Proceso Cierre"
                    
                Else
                    C = "VARIOS"
                End If
            End If
        End If
        
        
        
    End If
    DimeAccion = C
End Function


Public Function DevuelveAcciones(ByRef ListaAcciones As Collection) As Boolean
Dim i As Integer

 On Error GoTo EDevuelveAcciones
    For i = 1 To AccionesRegistradas
        Sql = i & "|" & DimeAccion(i) & "|"
        ListaAcciones.Add Sql
    Next i
    
    DevuelveAcciones = True
    
    Exit Function
EDevuelveAcciones:
    MuestraError Err.Number
End Function



Public Sub VolcarAFichero2()
Dim F As Date
Dim Rs As ADODB.Recordset




    On Error GoTo EVolcar

    

    Set Rs = New ADODB.Recordset
    F = Now   'Por si acaso esta el reloj mal
    Rs.Open "select curdate()", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then F = DBLet(Rs.Fields(0), "F")
    Rs.Close
    
    
    
    
    'Ya tengo la fecha.
    F = DateAdd("yyyy", -1, F)  'Le quito 1 año
    F = CDate("01/" & Month(F) & "/" & Year(F))   'Pongo el primer dia del mes anterior
    'Le quito un dia
    F = DateAdd("d", -1, F)
    
    
    Sql = "Select * from slog where fecha < '" & Format(F, "yyyy-mm-dd") & "' "
'    SQL = "INSERT INTO slog_old " & SQL
'    Conn.Execute SQL
    Conn.Execute "DELETE from slog where fecha <'" & Format(F, "yyyy-mm-dd") & "'"
        
    
    
    Set Rs = Nothing
    
    
    Exit Sub
EVolcar:
    Sql = "Error volcado datos acciones: " & vbCrLf & Err.Description
    Sql = Sql & vbCrLf & vbCrLf & "El programa continuará." & vbCrLf & vbCrLf & "Avise a soporte técnico."
    MsgBox Sql, vbExclamation
    Err.Clear
    Set Rs = Nothing
End Sub





'Dos metodos sobre la variable mvarDatosDescripcion
' Inicializar y añadir texto
Public Sub InicializarDatosDesc()
    mvarDatosDescripcion = ""
End Sub


Public Sub AnyadeTextoDatosDes(T As String)
    
    If mvarDatosDescripcion <> "" Then mvarDatosDescripcion = mvarDatosDescripcion & ":"
    mvarDatosDescripcion = mvarDatosDescripcion & T
End Sub
