Attribute VB_Name = "ModInformes"
Option Explicit


Public AbiertoOtroFormEnListado As Boolean  'Para saber si ha abieto un from desde el forms de listados



'Los reports
Public cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Public cadParam As String 'Cadena con los parametros para Crystal Report
Public numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Public cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Public cadNomRPT As String 'Nombre del informe a Imprimir
Public conSubRPT As Boolean 'Si el informe tiene subreports

Public cadPDFrpt As String 'Nombre del informe a enviar por email
Public vMostrarTree As Boolean
Public ExportarPDF As Boolean
Public SoloImprimir As Boolean


Dim Rs As Recordset
Dim cad As String
Dim SQL As String
Dim i As Integer


'Esto sera para el pb general
Dim TotalReg As Long
Dim Actual As Long


'Esta funcion lo que hace es genera el valor del campo
'El campo lo coge del recordset, luego sera field(i), y el tipo es para añadirle
'las coimllas, o quitarlas comas
'  Si es numero viene un 1 si no nada
Private Function ParaBD(ByRef Campo As ADODB.Field, Optional EsNumerico As Byte) As String
    
    If IsNull(Campo) Then
        ParaBD = "NULL"
    Else
        Select Case EsNumerico
        Case 1
            ParaBD = TransformaComasPuntos(CStr(Campo))
        Case 2
            'Fechas
            ParaBD = "'" & Format(CStr(Campo), "dd/mm/yyyy") & "'"
        Case Else
            ParaBD = "'" & Campo & "'"

            
        End Select
    End If
    ParaBD = "," & ParaBD
End Function


'/----------------------------------------------------------
'/----------------------------------------------------------
'/----------------------------------------------------------
'/----------------------------------------------------------

'   En este modulo se crearan los datos para los informes
'   Con lo cual cada Function Generara los datos en la tabla



'/----------------------------------------------------------
'/----------------------------------------------------------
'/----------------------------------------------------------
'/----------------------------------------------------------











Public Function InformeConceptos(ByRef vSQL As String) As Boolean

On Error GoTo EGI_Conceptos
    InformeConceptos = False
    'Borramos los anteriores
    Conn.Execute "Delete from Usuarios.zconceptos where codusu = " & vUsu.Codigo
    cad = "INSERT INTO Usuarios.zconceptos (codusu, codconce, nomconce,tipoconce) VALUES ("
    cad = cad & vUsu.Codigo & ",'"
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        SQL = cad & Format(Rs.Fields(0), "000")
        SQL = SQL & "','" & Rs.Fields(1) & "','" & Rs.Fields(3) & "')"
        Conn.Execute SQL
        'Siguiente
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    InformeConceptos = True
Exit Function
EGI_Conceptos:
    MuestraError Err.Number
    Set Rs = Nothing
End Function





Public Function ListadoEstadisticas(ByRef vSQL As String) As Boolean
On Error GoTo EListadoEstadisticas
    ListadoEstadisticas = False
    Conn.Execute "Delete from Usuarios.zestadinmo1 where codusu = " & vUsu.Codigo
    'Sentencia insert
    cad = "INSERT INTO Usuarios.zestadinmo1 (codusu, codigo, codconam, nomconam, codinmov, nominmov,"
    cad = cad & "tipoamor, porcenta, codprove, fechaadq, valoradq, amortacu, fecventa, impventa) VALUES ("
    cad = cad & vUsu.Codigo & ","
    
    'Empezamos
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    i = 1
    While Not Rs.EOF
        
        SQL = i & ParaBD(Rs!codconam, 1) & ParaBD(Rs!nomconam)
        SQL = SQL & ParaBD(Rs!Codinmov) & ",'" & DevNombreSQL(Rs!nominmov) & "'"
        
        SQL = SQL & ParaBD(Rs!tipoamor) & ParaBD(Rs!coeficie) & ParaBD(Rs!codprove)
        SQL = SQL & ParaBD(Rs!fechaadq, 2) & ParaBD(Rs!valoradq, 1) & ParaBD(Rs!amortacu, 1)
        SQL = SQL & ParaBD(Rs!fecventa, 2) 'FECHA
        SQL = SQL & ParaBD(Rs!impventa, 1) & ")"
        Conn.Execute cad & SQL
        
        'Sig
        Rs.MoveNext
        i = i + 1
    Wend
    ListadoEstadisticas = True
EListadoEstadisticas:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set Rs = Nothing
End Function





Public Function ListadoFichaInmo(ByRef vSQL As String) As Boolean
On Error GoTo Err1
    ListadoFichaInmo = False
    Conn.Execute "Delete from Usuarios.zfichainmo where codusu = " & vUsu.Codigo
    'Sentencia insert
    cad = "INSERT INTO Usuarios.zfichainmo (codusu, codigo, codinmov, nominmov, fechaadq, valoradq, Fechaamor,Importe, porcenta) VALUES ("
    cad = cad & vUsu.Codigo & ","
    
    'Empezamos
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    i = 1
    While Not Rs.EOF
        SQL = Rs!nominmov
        NombreSQL SQL
        SQL = i & ParaBD(Rs!Codinmov) & ",'" & SQL & "'"
        SQL = SQL & ParaBD(Rs!fechaadq, 2) & ParaBD(Rs!valoradq, 1) & ParaBD(Rs!fechainm, 2)
        SQL = SQL & ParaBD(Rs!imporinm, 1) & ParaBD(Rs!porcinm, 1)
        SQL = SQL & ")"
        Conn.Execute cad & SQL
        
        'Sig
        Rs.MoveNext
        i = i + 1
    Wend
    Rs.Close
    If i > 1 Then
        ListadoFichaInmo = True
    Else
        MsgBox "Ningún registro con esos valores", vbExclamation
    End If
Err1:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set Rs = Nothing
End Function




Public Function GenerarDatosCuentas(ByRef vSQL As String) As Boolean
On Error GoTo EGen
    GenerarDatosCuentas = False
    cad = "Delete FROM Usuarios.zCuentas where codusu =" & vUsu.Codigo
    Conn.Execute cad
    cad = "INSERT INTO Usuarios.zcuentas (codusu, codmacta, nommacta, razosoci,nifdatos, dirdatos, codposta, despobla, apudirec,model347) "
    cad = cad & " SELECT " & vUsu.Codigo & ",ctas.codmacta, ctas.nommacta, ctas.razosoci, ctas.nifdatos, ctas.dirdatos, ctas.codposta, ctas.despobla,ctas.apudirec,ctas.model347"
    cad = cad & " FROM " & vUsu.CadenaConexion & ".cuentas as ctas "
    If vSQL <> "" Then cad = cad & " WHERE " & vSQL
    Conn.Execute cad
    GenerarDatosCuentas = True
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
 
End Function



Public Function GenerarDiarios() As Boolean
On Error GoTo EGen
    GenerarDiarios = False
    cad = "Delete FROM Usuarios.ztiposdiario where codusu =" & vUsu.Codigo
    Conn.Execute cad
    cad = "INSERT INTO Usuarios.ztiposdiario (codusu, numdiari, desdiari)"
    cad = cad & " SELECT " & vUsu.Codigo & ",d.numdiari,d.desdiari"
    cad = cad & " FROM " & vUsu.CadenaConexion & ".tiposdiario as d;"
    Conn.Execute cad
    GenerarDiarios = True
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number

End Function


Public Function GeneraraExtractos() As Boolean
On Error GoTo EGen
    GeneraraExtractos = False
    cad = "Delete FROM Usuarios.ztmpconextcab where codusu =" & vUsu.Codigo
    Conn.Execute cad
    cad = "Delete FROM Usuarios.ztmpconext where codusu =" & vUsu.Codigo
    Conn.Execute cad
    cad = "INSERT INTO Usuarios.ztmpconextcab "
    cad = cad & "(codusu, cuenta, fechini, fechfin, acumantD, acumantH, acumantT, acumperD, acumperH, acumperT, acumtotD, acumtotH, acumtotT, cta)"
    cad = cad & " SELECT " & vUsu.Codigo & ",t.cuenta, t.fechini, t.fechfin, t.acumantD, t.acumantH, t.acumantT, t.acumperD, t.acumperH, t.acumperT, t.acumtotD, t.acumtotH, t.acumtotT, t.cta"
    cad = cad & " FROM " & vUsu.CadenaConexion & ".tmpconextcab as t where t.codusu =" & vUsu.Codigo & ";"
    Conn.Execute cad
    
    
    'Las lineas
    cad = "INSERT INTO Usuarios.ztmpconext (codusu, cta, numdiari, Pos, fechaent, numasien, linliapu, nomdocum, ampconce, timporteD, timporteH, saldo, Punteada, contra, ccost)"
    cad = cad & " SELECT " & vUsu.Codigo & ",t.cta, t.numdiari, t.Pos, t.fechaent, t.numasien, t.linliapu, t.nomdocum, t.ampconce, t.timporteD, t.timporteH, t.saldo, t.Punteada, t.contra, t.ccost"
    cad = cad & " FROM " & vUsu.CadenaConexion & ".tmpconext as t where t.codusu =" & vUsu.Codigo & ";"
    Conn.Execute cad
    GeneraraExtractos = True
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number

End Function


'Para la impresion de extractos y demas, MAYOR, etc
Public Function GeneraraExtractosListado(Cuenta As String) As Boolean
On Error GoTo EGen
    GeneraraExtractosListado = False
    cad = "INSERT INTO Usuarios.ztmpconextcab "
    cad = cad & "(codusu, cuenta, fechini, fechfin, acumantD, acumantH, acumantT, acumperD, acumperH, acumperT, acumtotD, acumtotH, acumtotT, cta)"
    cad = cad & " SELECT " & vUsu.Codigo & ",t.cuenta, t.fechini, t.fechfin, t.acumantD, t.acumantH, t.acumantT, t.acumperD, t.acumperH, t.acumperT, t.acumtotD, t.acumtotH, t.acumtotT, t.cta"
    cad = cad & " FROM " & vUsu.CadenaConexion & ".tmpconextcab as t where t.codusu =" & vUsu.Codigo & " AND cta ='" & Cuenta & "';"
    Conn.Execute cad
    
    'Las lineas
    cad = "INSERT INTO Usuarios.ztmpconext (codusu, cta, numdiari, Pos, fechaent, numasien, linliapu, nomdocum, ampconce, timporteD, timporteH, saldo, Punteada, contra, ccost)"
    cad = cad & " SELECT " & vUsu.Codigo & ",t.cta, t.numdiari, t.Pos, t.fechaent, t.numasien, t.linliapu, t.nomdocum, t.ampconce, t.timporteD, t.timporteH, t.saldo, t.Punteada, t.contra, t.ccost"
    cad = cad & " FROM " & vUsu.CadenaConexion & ".tmpconext as t where t.codusu =" & vUsu.Codigo & " AND cta ='" & Cuenta & "';"
    Conn.Execute cad
    GeneraraExtractosListado = True
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number

End Function


Public Function IAsientosErrores(ByRef vSQL As String) As Boolean
On Error GoTo EGen
    IAsientosErrores = False
    cad = "Delete FROM Usuarios.zdiapendact  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    'Las lineas
    cad = "INSERT INTO Usuarios.zdiapendact (codusu, numdiari, desdiari, fechaent, numasien, linliapu, codmacta, nommacta, numdocum,"
    cad = cad & " ampconce, timporteD, timporteH, codccost)"
    cad = cad & " SELECT " & vUsu.Codigo
    cad = cad & ",linapue.numdiari, tiposdiario.desdiari, linapue.fechaent, linapue.numasien, linapue.linliapu, linapue.codmacta, cuentas.nommacta, linapue.numdocum, linapue.ampconce, linapue.timporteD, linapue.timporteH, linapue.codccost"
    cad = cad & " FROM (linapue LEFT JOIN tiposdiario ON linapue.numdiari = tiposdiario.numdiari) LEFT JOIN cuentas ON linapue.codmacta = cuentas.codmacta"
    If vSQL <> "" Then cad = cad & " WHERE " & vSQL
    
    
    
    
    Conn.Execute cad
    
    Set Rs = New ADODB.Recordset
    cad = "select count(*) FROM Usuarios.zdiapendact  where codusu =" & vUsu.Codigo
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0), "N") > 0 Then cad = ""
    End If
    Rs.Close
    Set Rs = Nothing
    
    If cad <> "" Then
        MsgBox "Ningun registro por mostrar.", vbExclamation
        Exit Function
    End If
    IAsientosErrores = True
    
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function


Public Function IDiariosPendientes(ByRef vSQL As String) As Boolean
On Error GoTo EGen
    IDiariosPendientes = False
    cad = "Delete FROM Usuarios.zdiapendact  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    'Las lineas
    cad = "INSERT INTO Usuarios.zdiapendact (codusu, numdiari, desdiari, fechaent, numasien, linliapu, codmacta, nommacta, numdocum,"
    cad = cad & " ampconce, timporteD, timporteH, codccost)"
    cad = cad & " SELECT " & vUsu.Codigo
    cad = cad & ",cabapu_0.numdiari, tiposdiario_0.desdiari, cabapu_0.fechaent, cabapu_0.numasien, linapu_0.linliapu,"
    cad = cad & " linapu_0.codmacta, cuentas_0.nommacta, linapu_0.numdocum, linapu_0.ampconce, linapu_0.timporteD, "
    cad = cad & " linapu_0.timporteH, linapu_0.codccost  FROM cabapu cabapu_0, cuentas cuentas_0, linapu linapu_0, tiposdiario "
    cad = cad & " tiposdiario_0  WHERE linapu_0.fechaent = cabapu_0.fechaent AND linapu_0.numasien = cabapu_0.numasien AND "
    cad = cad & " linapu_0.numdiari = cabapu_0.numdiari AND tiposdiario_0.numdiari = cabapu_0.numdiari AND"
    cad = cad & " tiposdiario_0.numdiari = linapu_0.numdiari AND cuentas_0.codmacta = linapu_0.codmacta"
    If vSQL <> "" Then cad = cad & " AND " & vSQL
    
    
    
    
    Conn.Execute cad
    
    Set Rs = New ADODB.Recordset
    cad = "select count(*) FROM Usuarios.zdiapendact  where codusu =" & vUsu.Codigo
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0), "N") > 0 Then cad = ""
    End If
    Rs.Close
    Set Rs = Nothing
    
    If cad <> "" Then
        MsgBox "Ningun registro por mostrar.", vbExclamation
        Exit Function
    End If
    IDiariosPendientes = True
    
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function




Public Function ITotalesCtaConcepto(ByRef vSQL As String, tabla As String) As Boolean
On Error GoTo EGen
    ITotalesCtaConcepto = False
    cad = "Delete FROM Usuarios.ztotalctaconce  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    'Las lineas
    cad = "INSERT INTO Usuarios.ztotalctaconce (codusu, codmacta, nommacta, nifdatos, fechaent, timporteD, timporteH, codconce)"
    cad = cad & " SELECT " & vUsu.Codigo
    cad = cad & " ," & tabla & ".codmacta, nommacta, nifdatos, fechaent,"
    cad = cad & " timporteD,timporteH, codconce"
    cad = cad & " FROM " & vUsu.CadenaConexion & ".cuentas ,"
    cad = cad & vUsu.CadenaConexion & "." & tabla & " WHERE cuentas.codmacta = " & tabla & ".codmacta"
    If vSQL <> "" Then cad = cad & " AND " & vSQL
    
    
    
    
    Conn.Execute cad
    
    
    'Inserto en ztmpdiarios para los que sean TODOS los conceptos
    cad = "Delete FROM Usuarios.ztiposdiario  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    cad = "INSERT INTO Usuarios.ztiposdiario SELECT " & vUsu.Codigo & ",codconce,nomconce FROM conceptos"
    Conn.Execute cad
    
    
    'Contamos para ver cuantos hay
    cad = "Select count(*) from Usuarios.ztotalctaconce WHERE codusu =" & vUsu.Codigo
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            If Rs.Fields(0) > 0 Then i = 1
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    If i > 0 Then
        ITotalesCtaConcepto = True
    Else
        MsgBox "Ningún registro con esos valores.", vbExclamation
    End If
    
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function


Public Function IAsientosPre(ByRef vSQL As String) As Boolean
On Error GoTo EGen
    IAsientosPre = False
    cad = "Delete FROM Usuarios.zasipre  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    cad = "INSERT INTO Usuarios.zasipre (codusu, numaspre, nomaspre, linlapre, codmacta, nommacta"
    cad = cad & ", ampconce, timporteD, timporteH, codccost)"

    cad = cad & " SELECT " & vUsu.Codigo
    cad = cad & ", t1.numaspre, t1.nomaspre, t2.linlapre,t2.codmacta, t3.nommacta,t2.ampconce,"
    cad = cad & "t2.timported,t2.timporteh,t2.codccost FROM "
    cad = cad & vUsu.CadenaConexion & ".asipre as t1,"
    cad = cad & vUsu.CadenaConexion & ".asipre_lineas as t2,"
    cad = cad & vUsu.CadenaConexion & ".cuentas as t3 WHERE "
    cad = cad & " t1.numaspre=t2.numaspre AND t2.codmacta=t3.codmacta"
    If vSQL <> "" Then cad = cad & " AND " & vSQL
    
    
    Conn.Execute cad
    IAsientosPre = True
    
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function





Public Function IHcoApuntes(ByRef vSQL As String, NumeroTabla As String) As Boolean
On Error GoTo EGen
    IHcoApuntes = False
    cad = "Delete FROM Usuarios.zhistoapu  where codusu =" & vUsu.Codigo
    Conn.Execute cad
    
    
    cad = " INSERT INTO Usuarios.zhistoapu (codusu, numdiari, desdiari, fechaent, numasien, linliapu, codmacta, nommacta, numdocum, ampconce,"
    cad = cad & " timporteD, timporteH, codccost) "
    cad = cad & "SELECT " & vUsu.Codigo & ",hcabapu" & NumeroTabla & ".numdiari, tiposdiario.desdiari, hcabapu" & NumeroTabla & ".fechaent, hcabapu" & NumeroTabla & ".numasien, hlinapu" & NumeroTabla & ".linliapu,"
    cad = cad & " hlinapu" & NumeroTabla & ".codmacta, cuentas.nommacta, hlinapu" & NumeroTabla & ".numdocum, hlinapu" & NumeroTabla & ".ampconce, hlinapu" & NumeroTabla & ".timporteD,"
    cad = cad & " hlinapu" & NumeroTabla & ".timporteH, hlinapu" & NumeroTabla & ".codccost  "
    cad = cad & " FROM " & vUsu.CadenaConexion & ".cuentas , " & vUsu.CadenaConexion & ".hcabapu" & NumeroTabla & " , " & vUsu.CadenaConexion & ".hlinapu" & NumeroTabla & ", " & vUsu.CadenaConexion & ".tiposdiario"
    cad = cad & " WHERE hlinapu" & NumeroTabla & ".fechaent = hcabapu" & NumeroTabla & ".fechaent AND hlinapu" & NumeroTabla & ".numasien = hcabapu" & NumeroTabla & ".numasien AND"
    cad = cad & " hlinapu" & NumeroTabla & ".numdiari = hcabapu" & NumeroTabla & ".numdiari AND cuentas.codmacta = hlinapu" & NumeroTabla & ".codmacta AND tiposdiario.numdiari ="
    cad = cad & " hcabapu" & NumeroTabla & ".numdiari AND tiposdiario.numdiari = hlinapu" & NumeroTabla & ".numdiari"
    If vSQL <> "" Then cad = cad & " AND " & vSQL
    
    
    
    
    Conn.Execute cad
    
    cad = DevuelveDesdeBD("count(*)", "Usuarios.zhistoapu", "codusu", vUsu.Codigo, "N")
    If Val(cad) = 0 Then
        MsgBox "Ningun registro seleccionado", vbExclamation
    Else
        IHcoApuntes = True
    End If
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function

'                           formateada
'Vienen empipados numasien|  fechaent    |numdiari|
Public Function IHcoApuntesAlActualizarModificar(Cadena1 As String) As Boolean
On Error GoTo EGen
    IHcoApuntesAlActualizarModificar = False
    
    'No borramos, borraremos antes de llamar a esta funcion
    'Cad = "Delete FROM Usuarios.zhistoapu  where codusu =" & vUsu.Codigo
    'Conn.Execute Cad
    
    
    cad = " INSERT INTO Usuarios.zhistoapu (codusu, numdiari, desdiari, fechaent, numasien, linliapu, codmacta, nommacta, numdocum, ampconce,"
    cad = cad & " timporteD, timporteH, codccost) "
    cad = cad & "SELECT " & vUsu.Codigo & ",hcabapu.numdiari, tiposdiario.desdiari, hcabapu.fechaent, hcabapu.numasien, hlinapu.linliapu,"
    cad = cad & " hlinapu.codmacta, cuentas.nommacta, hlinapu.numdocum, hlinapu.ampconce, hlinapu.timporteD,"
    cad = cad & " hlinapu.timporteH, hlinapu.codccost  "
    cad = cad & " FROM cuentas , hcabapu,hlinapu,tiposdiario"
    cad = cad & " WHERE hlinapu.fechaent = hcabapu.fechaent AND hlinapu.numasien = hcabapu.numasien AND"
    cad = cad & " hlinapu.numdiari = hcabapu.numdiari AND cuentas.codmacta = hlinapu.codmacta AND tiposdiario.numdiari ="
    cad = cad & " hcabapu.numdiari AND tiposdiario.numdiari = hlinapu.numdiari"
    cad = cad & " AND hcabapu.numasien  =" & RecuperaValor(Cadena1, 1)
    cad = cad & " AND hcabapu.fechaent  ='" & RecuperaValor(Cadena1, 2)
    cad = cad & "' AND hcabapu.numdiari =" & RecuperaValor(Cadena1, 3)
    
    Conn.Execute cad
    IHcoApuntesAlActualizarModificar = True
    
EGen:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function




Public Function GeneraDatosHcoInmov(ByRef vSQL As String) As Boolean

On Error GoTo EGeneraDatosHcoInmov
    GeneraDatosHcoInmov = False
        
    'Borramos tmp
    cad = "Delete from Usuarios.zfichainmo where codusu = " & vUsu.Codigo
    Conn.Execute cad
    'Abrimos datos
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        MsgBox "Ningún dato a mostrar", vbExclamation
    Else
        cad = "INSERT INTO Usuarios.zfichainmo (codusu, codigo, codinmov, nominmov, fechaadq, valoradq, fechaamor, Importe, porcenta) VALUES (" & vUsu.Codigo & ","
        TotalReg = 0
        While Not Rs.EOF
           
            TotalReg = TotalReg + 1
            'Metemos los nuevos datos
            SQL = TotalReg & ParaBD(Rs!Codinmov, 1) & ",'" & DevNombreSQL(CStr(Rs!nominmov)) & "'" & ParaBD(Rs!fechainm, 2)
            SQL = SQL & ",NULL,NULL" & ParaBD(Rs!imporinm, 1) & ParaBD(Rs!porcinm, 1) & ")"
            SQL = cad & SQL
            Conn.Execute SQL
            Rs.MoveNext
        Wend
        GeneraDatosHcoInmov = True
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EGeneraDatosHcoInmov:
    MuestraError Err.Number, Err.Description
    Set Rs = Nothing
End Function





Public Function GeneraDatosConceptosInmov() As Boolean

On Error GoTo EGeneraDatosConceptosInmov
    GeneraDatosConceptosInmov = False
        
    'Borramos tmp
    cad = "Delete from Usuarios.ztmppresu1 where codusu = " & vUsu.Codigo
    Conn.Execute cad
    'Abrimos datos
    SQL = "Select * from inmovcon"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        MsgBox "Ningún dato a mostrar", vbExclamation
    Else
        cad = "INSERT INTO Usuarios.ztmppresu1 (codusu, codigo, cta, titulo, ano, mes, Importe) VALUES (" & vUsu.Codigo
        While Not Rs.EOF
            'Metemos los nuevos datos
            SQL = ParaBD(Rs!codconam, 1) & ",'" & Format(Rs!codconam, "0000") & "'" & ParaBD(Rs!nomconam)
            SQL = SQL & ",0" & ParaBD(Rs!perimaxi, 1) & ParaBD(Rs!coefimaxi, 1) & ")"
            SQL = cad & SQL
            Conn.Execute SQL
            Rs.MoveNext
        Wend
        GeneraDatosConceptosInmov = True
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
EGeneraDatosConceptosInmov:
    MuestraError Err.Number, Err.Description
    Set Rs = Nothing
End Function


'#################################################
'###########    A Ñ A D I D O     ################  DE  NUEVA CONTA DE DAVID
'#################################################
Public Sub PonerDatosPorDefectoImpresion(ByRef formu As Form, SoloImpresora As Boolean, Optional NombreArchivoEx As String)
On Error Resume Next
'        AbiertoOtroFormEnListado = False
        
        formu.txtTipoSalida(0).Text = Printer.DeviceName
        If Err.Number <> 0 Then
            formu.txtTipoSalida(0).Text = "No hay impresora instalada"
            Err.Clear
        End If
        If SoloImpresora Then Exit Sub
        
        formu.txtTipoSalida(1).Text = App.Path & "\Exportar\" & NombreArchivoEx & ".csv"
        formu.txtTipoSalida(2).Text = App.Path & "\Exportar\" & NombreArchivoEx & ".pdf"
        
        If Err.Number <> 0 Then Err.Clear
    
End Sub


'PDF=true   CSV=false
Public Function EliminarDocum(PDF As Boolean) As Boolean
    On Error Resume Next
    If PDF Then
        If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then Kill App.Path & "\docum.pdf"
    Else
        If Dir(App.Path & "\docum.csv", vbArchive) <> "" Then Kill App.Path & "\docum.csv"
    End If
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
        Err.Clear
        EliminarDocum = False
    Else
        EliminarDocum = True
    End If
End Function


Public Sub ponerLabelBotonImpresion(ByRef BotonAcept As CommandButton, ByRef BotonImpr As CommandButton, SelectorImpresion As Integer)
    If SelectorImpresion = 0 Then
        BotonAcept.Caption = "&Vista previa"
    Else
        BotonAcept.Caption = "&Aceptar"
    End If
    BotonImpr.Visible = SelectorImpresion = 0
End Sub

Public Function PonerDesdeHasta(Campo As String, Tipo As String, ByRef Desde As TextBox, ByRef DesD As TextBox, ByRef Hasta As TextBox, ByRef DesH As TextBox, param As String) As Boolean
Dim Devuelve As String
Dim cad As String
Dim Subtipo As String 'F: fecha   N: numero   T: texto  H: HORA



    PonerDesdeHasta = False
    
    Select Case Tipo
    Case "F", "FEC"
        'Campos fecha
        Subtipo = "F"
    
    Case "CONC", "TDIA", "BIC", "AGE", "COI", "INM", "FRA"
        'concepto
        Subtipo = "N"
        
    Case "CTA", "BAN", "CCO", "SER", "CRY"
        Subtipo = "T"
        
    Case "ASIP", "ASI"
        Subtipo = "N"
        
    Case "TIVA", "DIA"
        Subtipo = "N"
        
    Case "TPAG", "FPAG"
        Subtipo = "N"
        
    End Select
    
    Devuelve = CadenaDesdeHasta(Desde, Hasta, Campo, Subtipo)
    If Devuelve = "Error" Then
        PonFoco Desde
        Exit Function
    End If
    If Not AnyadirAFormula(cadFormula, Devuelve) Then Exit Function
    
    If Devuelve = "" Then
        PonerDesdeHasta = True
        Exit Function
    End If
    
    'QUITO LAS LLAVES
    Devuelve = Replace(Devuelve, "{", "")
    Devuelve = Replace(Devuelve, "}", "")
    
    If Subtipo <> "F" And Subtipo <> "FH" Then
        'Fecha para Crystal Report

        If Not AnyadirAFormula(cadselect, Devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        cad = CadenaDesdeHastaBD(Desde.Text, Hasta.Text, Campo, Subtipo)
        cad = Replace(cad, "{", "")
        cad = Replace(cad, "}", "")
        If Not AnyadirAFormula(cadselect, cad) Then Exit Function
    End If
    
    If Devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, Desde, Hasta, DesD, DesH) & """|"
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Public Function AnyadirAFormula(ByRef cadFormula As String, arg As String) As Boolean
'Concatena los criterios del WHERE para pasarlos al Crystal como FormulaSelection
    If arg = "Error" Then
        AnyadirAFormula = False
        Exit Function
    ElseIf arg <> "" Then
        If cadFormula <> "" Then
            cadFormula = cadFormula & " AND (" & arg & ")"
        Else
            cadFormula = arg
        End If
    End If
    AnyadirAFormula = True
End Function



Private Function AnyadirParametroDH(cad As String, ByRef TextoDESDE As TextBox, TextoHasta As TextBox, ByRef TD As TextBox, ByRef TH As TextBox) As String
On Error Resume Next
    
    
    If Not TextoDESDE Is Nothing Then
         If TextoDESDE.Text <> "" Then
            cad = cad & "desde " & TextoDESDE.Text
'            If TD.Caption <> "" Then Cad = Cad & " - " & TD.Caption
        End If
    End If
    If Not TextoHasta Is Nothing Then
        If TextoHasta.Text <> "" Then
            cad = cad & "  hasta " & TextoHasta.Text
'            If TH.Caption <> "" Then Cad = Cad & " - " & TH.Caption
        End If
    End If
    
    AnyadirParametroDH = cad
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function GeneraFicheroCSV(cadSQL As String, Salida As String) As Boolean
Dim NF As Integer
Dim i  As Integer

    On Error GoTo eGeneraFicheroCSV
    GeneraFicheroCSV = False
    
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cadSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "Ningun dato generado", vbExclamation
        cadSQL = ""
    Else
        NF = FreeFile
        Open App.Path & "\docum.csv" For Output As #NF
        'Cabecera
        cadSQL = ""
        For i = 0 To miRsAux.Fields.Count - 1
            cadSQL = cadSQL & ";""" & miRsAux.Fields(i).Name & """"
        Next i
        Print #NF, Mid(cadSQL, 2)
    
    
        'Lineas
        While Not miRsAux.EOF
            cadSQL = ""
            For i = 0 To miRsAux.Fields.Count - 1
                cadSQL = cadSQL & ";""" & DBLet(miRsAux.Fields(i).Value, "T") & """"
            Next i
            Print #NF, Mid(cadSQL, 2)
            
            
            
            miRsAux.MoveNext
        Wend
        cadSQL = "OK"
    End If
    miRsAux.Close
    Close #NF

    If cadSQL = "OK" Then
        If CopiarFicheroASalida(True, Salida) Then GeneraFicheroCSV = True
    End If
    
    Exit Function
eGeneraFicheroCSV:
    MuestraError Err.Number, Err.Description
End Function


Public Function CopiarFicheroASalida(csv As Boolean, Salida As String, Optional SinMensaje As Boolean) As Boolean
    CopiarFicheroASalida = False
    If csv Then
        FileCopy App.Path & "\docum.csv", Salida
    Else
        FileCopy App.Path & "\docum.pdf", Salida
    End If
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Copiando " & Salida
    Else
        If Not SinMensaje Then
            MsgBox "Fichero:  " & Salida & vbCrLf & "Generado con éxito.", vbInformation
        End If
        CopiarFicheroASalida = True
    End If
End Function

Public Function ImprimeGeneral() As Boolean
    
    Dim HaPulsadoImprimir As Boolean

    Screen.MousePointer = vbHourglass


    'Prueba DAVID
    frmPpal.SkinFramework.AutoApplyNewWindows = False
    frmPpal.SkinFramework.AutoApplyNewThreads = False


    HaPulsadoImprimir = False
    cadPDFrpt = cadNomRPT
    With frmVisReport
        .Informe = App.Path & "\Informes\"
        If ExportarPDF Then
            'PDF
            .Informe = .Informe & cadPDFrpt
        Else
            'IMPRIMIR
            .Informe = .Informe & cadNomRPT
        End If
        .FormulaSeleccion = cadFormula
        .SoloImprimir = False
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .ConSubInforme = conSubRPT

        .NumCopias2 = 1

        .SoloImprimir = SoloImprimir
        .ExportarPDF = ExportarPDF
        .MostrarTree = vMostrarTree
        .Show vbModal
        HaPulsadoImprimir = .EstaImpreso
        
      End With
    
    
     'DAVID
    frmPpal.SkinFramework.AutoApplyNewWindows = True
    frmPpal.SkinFramework.AutoApplyNewThreads = True
    
End Function



Public Sub QuitarPulsacionMas(ByRef T As TextBox)
Dim i As Integer

    Do
        i = InStr(1, T.Text, "+")
        If i > 0 Then T.Text = Mid(T.Text, 1, i - 1) & Mid(T.Text, i + 1)
    Loop Until i = 0
        
End Sub



Public Sub LanzaProgramaAbrirOutlook(outTipoDocumento As Integer)
Dim NombrePDF As String
Dim Aux As String
Dim Lanza As String

    On Error GoTo ELanzaProgramaAbrirOutlook

    If Not PrepararCarpetasEnvioMail(True) Then Exit Sub
    
    If Not ExisteARIMAILGES Then Exit Sub

    'Primer tema. Copiar el docum.pdf con otro nombre mas significatiov
    Select Case outTipoDocumento
    Case 1
        'Conceptos
        Aux = "Conceptos.pdf"
    Case 2
        'Cuentas contables
        Aux = "Cuentas.pdf"
    Case 3
        'Asientos Predefinidos
        Aux = "Asientos Predefinidos.pdf"
    Case 4
        Aux = "Tipos Diario.pdf"
    Case 5
        Aux = "Asientos.pdf"
    Case 6
        Aux = "Tipos de Iva.pdf"
    Case 7
        Aux = "Tipos de Pago.pdf"
    Case 8
        Aux = "Formas de Pago.pdf"
    Case 9
        Aux = "Bancos.pdf"
    Case 10
        Aux = "Bic/Swift.pdf"
    Case 11
        Aux = "Agentes.pdf"
    Case 12
        Aux = "Informes.pdf"
    Case 13
        Aux = "Bancos.pdf"
    Case 14
        Aux = "AsientosHco.pdf"
    Case 15
        Aux = "Listado de Facturas de Cliente.pdf"
    Case 16
        Aux = "Relación de Clientes por Cta Ventas.pdf"
    Case 17
        Aux = "Listado de Facturas de Proveedores.pdf"
    Case 18
        Aux = "Relación de Proveedores por Cta Gastos.pdf"
    Case 19
        Aux = "Modelo 303.pdf"
    Case 20
        Aux = "Modelo 340.pdf"
    Case 21
        Aux = "Modelo 347.pdf"
    Case 51
        
    Case 100
        
    End Select
    NombrePDF = App.Path & "\temp\" & Aux
    If Dir(NombrePDF, vbArchive) <> "" Then Kill NombrePDF
    FileCopy App.Path & "\docum.pdf", NombrePDF
    
    Aux = FijaDireccionEmail(outTipoDocumento)
    Lanza = Aux & "|"
    Aux = ""
    Select Case outTipoDocumento
    Case 1
        Aux = "Conceptos"
    Case 2
        Aux = "Cuentas"
    Case 3
        'Asientos Predefinidos
        Aux = "Asientos Predefinidos"
    Case 4
        Aux = "Tipos Diario"
    Case 5
        Aux = "Asientos"
    Case 6
        Aux = "Tipos de Iva"
    Case 7
        Aux = "Tipos de Pago"
    Case 8
        Aux = "Formas de Pago"
    Case 9
        Aux = "Bancos"
    Case 10
        Aux = "Bic/Swift"
    Case 11
        Aux = "Agentes"
    Case 12
        Aux = "Informes"
    Case 13
        Aux = "Bancos"
    Case 14
        Aux = "AsientosHco"
    Case 15
        Aux = "Listado de Facturas de Cliente"
    Case 16
        Aux = "Relación de Clientes por Cta Ventas"
    Case 17
        Aux = "Listado de Facturas de Proveedores"
    Case 18
        Aux = "Relación de Proveedores por Cta Gastos"
    Case 19
        Aux = "Modelo 303"
    Case 20
        Aux = "Modelo 340"
    Case 21
        Aux = "Modelo 347"
        
        
    '--------------------------------------------------
    Case 51
        Aux = "Pedido proveedor nº: " ' & outClaveNombreArchiv
        
    Case 100
        Aux = "Factura nº" '& outClaveNombreArchiv
        
    End Select
    
    Lanza = Lanza & Aux & "|"
    
    'Aqui pondremos lo del texto del BODY
    Aux = ""
    Lanza = Lanza & Aux & "|"
    
    
    'Envio o mostrar
    Lanza = Lanza & "0"   '0. Display   1.  send
    
    'Campos reservados para el futuro
    Lanza = Lanza & "||||"
    
    'El/los adjuntos
    Lanza = Lanza & NombrePDF & "|"
    
    Aux = App.Path & "\ARIMAILGES.EXE" & " " & Lanza  '& vParamAplic.ExeEnvioMail & " " & Lanza
    Shell Aux, vbNormalFocus
    
    Exit Sub
ELanzaProgramaAbrirOutlook:
    MuestraError Err.Number, Err.Description
End Sub

Private Function FijaDireccionEmail(outTipoDocumento As Integer) As String
Dim campoemail As String
Dim otromail As String


    FijaDireccionEmail = ""
    
    
    If outTipoDocumento < 50 Then
        campoemail = ""
        
    Else
        If outTipoDocumento < 100 Then
            'Para provedores
            If outTipoDocumento = 51 Or outTipoDocumento = 52 Or outTipoDocumento = 53 Then
                campoemail = "maiprov1"
                otromail = "maiprov2"
            Else
                campoemail = "maiprov2"
                otromail = "maiprov1"
            End If
'            campoemail = DevuelveDesdeBDNew(cpconta, "proveedor", campoemail, "codprove", Me.outCodigoCliProv, "N", otromail)
            If campoemail = "" Then campoemail = otromail
        Else
            'Para Socios
            If outTipoDocumento >= 100 Then
                campoemail = "maisocio"
                otromail = "maisocio"
            Else
                campoemail = "maisocio"
                otromail = "maisocio"
            End If
'            campoemail = DevuelveDesdeBDNew(cAgro, "rsocios", campoemail, "codsocio", Me.outCodigoCliProv, "N") ' , otromail)
            If campoemail = "" Then campoemail = otromail
        End If
    End If
    FijaDireccionEmail = campoemail
End Function




'--------------------------------------------------------------------
'-------------------------------------------------------------------
'Para el envio de los mails
Public Function PrepararCarpetasEnvioMail(Optional NoBorrar As Boolean) As Boolean
    On Error GoTo EPrepararCarpetasEnvioMail
    PrepararCarpetasEnvioMail = False

    If Dir(App.Path & "\temp", vbDirectory) = "" Then
        MkDir App.Path & "\temp"
    Else
        If Not NoBorrar Then
            If Dir(App.Path & "\temp\*.*", vbArchive) <> "" Then Kill App.Path & "\temp\*.*"
        End If
    End If


    PrepararCarpetasEnvioMail = True
    Exit Function
EPrepararCarpetasEnvioMail:
    MuestraError Err.Number, "", "Preparar Carpetas Envio Mail "
End Function


Public Function LanzaMailGnral(dirMail As String) As Boolean
'LLama al Programa de Correo (Outlook,...)
Dim Aux As String
Dim Lanza As String

On Error GoTo ELanzaHome

    LanzaMailGnral = False

    If Not ExisteARIMAILGES Then Exit Function


    If dirMail = "" Then
        MsgBox "No hay dirección e-mail a la que enviar.", vbExclamation
        Exit Function
    End If

    Aux = dirMail
    Lanza = Lanza & Aux & "||"

    'Aqui pondremos lo del texto del BODY
    Aux = ""
    Lanza = Lanza & Aux & "|"

    'Envio o mostrar
    Lanza = Lanza & "0"   '0. Display   1.  send

    'Campos reservados para el futuro
    Lanza = Lanza & "||||"

    'El/los adjuntos
    Lanza = Lanza & "|"

    Aux = App.Path & "\ARIMAILGES.EXE" & " " & Lanza  '& vParamAplic.ExeEnvioMail & " " & Lanza
    Shell Aux, vbNormalFocus

    LanzaMailGnral = True

ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, vbCrLf & Err.Description
'    CadenaDesdeOtroForm = ""
End Function


Public Function ExisteARIMAILGES()
Dim SQL As String

    If Dir(App.Path & "\ArimailGes.exe") = "" Then
        MsgBox "No existe el programa ArimailGes.exe. Llame a Ariadna.", vbExclamation
        ExisteARIMAILGES = False
    Else
        ExisteARIMAILGES = True
    End If
End Function



Public Function HayRegParaInforme(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
    
    SQL = "Select count(*) FROM " & cTabla
    If cWhere <> "" Then
        SQL = SQL & " WHERE " & cWhere
    End If
    
    If TotalRegistros(SQL) = 0 Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegParaInforme = False
    Else
        HayRegParaInforme = True
    End If
End Function

Public Function PonerParamRPT(indice As String, nomDocu As String) As Boolean
Dim cad As String
Dim Encontrado As Boolean

        nomDocu = ""
        Encontrado = False
        PonerParamRPT = False
        
        cad = "select informe from scryst where codigo = " & DBSet(indice, "T")
        
        Set Rs = New ADODB.Recordset
        Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs.EOF Then
            nomDocu = DBLet(Rs!Informe, "T")
            Encontrado = True
        End If
        
        If Encontrado = False Or nomDocu = "" Then
            cad = "No se han podido cargar los Parámetros de Tipos de Documentos." & vbCrLf
            MsgBox cad & "Debe configurar la aplicación.", vbExclamation
            PonerParamRPT = False
            Exit Function
        End If
        
        PonerParamRPT = True
    

End Function

