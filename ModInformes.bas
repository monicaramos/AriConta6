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


Dim RS As Recordset
Dim cad As String
Dim SQL As String
Dim I As Integer


'Esto sera para el pb general
Dim TotalReg As Long
Dim Actual As Long


'Esta funcion lo que hace es genera el valor del campo
'El campo lo coge del recordset, luego sera field(i), y el tipo es para a�adirle
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
    Set RS = New ADODB.Recordset
    RS.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        SQL = cad & Format(RS.Fields(0), "000")
        SQL = SQL & "','" & RS.Fields(1) & "','" & RS.Fields(3) & "')"
        Conn.Execute SQL
        'Siguiente
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    InformeConceptos = True
Exit Function
EGI_Conceptos:
    MuestraError Err.Number
    Set RS = Nothing
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
    Set RS = New ADODB.Recordset
    RS.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    I = 1
    While Not RS.EOF
        
        SQL = I & ParaBD(RS!codconam, 1) & ParaBD(RS!nomconam)
        SQL = SQL & ParaBD(RS!Codinmov) & ",'" & DevNombreSQL(RS!nominmov) & "'"
        
        SQL = SQL & ParaBD(RS!tipoamor) & ParaBD(RS!coeficie) & ParaBD(RS!codprove)
        SQL = SQL & ParaBD(RS!fechaadq, 2) & ParaBD(RS!valoradq, 1) & ParaBD(RS!amortacu, 1)
        SQL = SQL & ParaBD(RS!fecventa, 2) 'FECHA
        SQL = SQL & ParaBD(RS!impventa, 1) & ")"
        Conn.Execute cad & SQL
        
        'Sig
        RS.MoveNext
        I = I + 1
    Wend
    ListadoEstadisticas = True
EListadoEstadisticas:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set RS = Nothing
End Function





Public Function ListadoFichaInmo(ByRef vSQL As String) As Boolean
On Error GoTo Err1
    ListadoFichaInmo = False
    Conn.Execute "Delete from Usuarios.zfichainmo where codusu = " & vUsu.Codigo
    'Sentencia insert
    cad = "INSERT INTO Usuarios.zfichainmo (codusu, codigo, codinmov, nominmov, fechaadq, valoradq, Fechaamor,Importe, porcenta) VALUES ("
    cad = cad & vUsu.Codigo & ","
    
    'Empezamos
    Set RS = New ADODB.Recordset
    RS.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    I = 1
    While Not RS.EOF
        SQL = RS!nominmov
        NombreSQL SQL
        SQL = I & ParaBD(RS!Codinmov) & ",'" & SQL & "'"
        SQL = SQL & ParaBD(RS!fechaadq, 2) & ParaBD(RS!valoradq, 1) & ParaBD(RS!fechainm, 2)
        SQL = SQL & ParaBD(RS!imporinm, 1) & ParaBD(RS!porcinm, 1)
        SQL = SQL & ")"
        Conn.Execute cad & SQL
        
        'Sig
        RS.MoveNext
        I = I + 1
    Wend
    RS.Close
    If I > 1 Then
        ListadoFichaInmo = True
    Else
        MsgBox "Ning�n registro con esos valores", vbExclamation
    End If
Err1:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set RS = Nothing
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
    
    Set RS = New ADODB.Recordset
    cad = "select count(*) FROM Usuarios.zdiapendact  where codusu =" & vUsu.Codigo
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If DBLet(RS.Fields(0), "N") > 0 Then cad = ""
    End If
    RS.Close
    Set RS = Nothing
    
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
    
    Set RS = New ADODB.Recordset
    cad = "select count(*) FROM Usuarios.zdiapendact  where codusu =" & vUsu.Codigo
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If DBLet(RS.Fields(0), "N") > 0 Then cad = ""
    End If
    RS.Close
    Set RS = Nothing
    
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
    Set RS = New ADODB.Recordset
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If RS.Fields(0) > 0 Then I = 1
        End If
    End If
    RS.Close
    Set RS = Nothing
    If I > 0 Then
        ITotalesCtaConcepto = True
    Else
        MsgBox "Ning�n registro con esos valores.", vbExclamation
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
Public Function IHcoApuntesAlActualizarModificar(cadena1 As String) As Boolean
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
    cad = cad & " AND hcabapu.numasien  =" & RecuperaValor(cadena1, 1)
    cad = cad & " AND hcabapu.fechaent  ='" & RecuperaValor(cadena1, 2)
    cad = cad & "' AND hcabapu.numdiari =" & RecuperaValor(cadena1, 3)
    
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
    Set RS = New ADODB.Recordset
    RS.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        MsgBox "Ning�n dato a mostrar", vbExclamation
    Else
        cad = "INSERT INTO Usuarios.zfichainmo (codusu, codigo, codinmov, nominmov, fechaadq, valoradq, fechaamor, Importe, porcenta) VALUES (" & vUsu.Codigo & ","
        TotalReg = 0
        While Not RS.EOF
           
            TotalReg = TotalReg + 1
            'Metemos los nuevos datos
            SQL = TotalReg & ParaBD(RS!Codinmov, 1) & ",'" & DevNombreSQL(CStr(RS!nominmov)) & "'" & ParaBD(RS!fechainm, 2)
            SQL = SQL & ",NULL,NULL" & ParaBD(RS!imporinm, 1) & ParaBD(RS!porcinm, 1) & ")"
            SQL = cad & SQL
            Conn.Execute SQL
            RS.MoveNext
        Wend
        GeneraDatosHcoInmov = True
    End If
    RS.Close
    Set RS = Nothing
    Exit Function
EGeneraDatosHcoInmov:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Function





Public Function GeneraDatosConceptosInmov() As Boolean

On Error GoTo EGeneraDatosConceptosInmov
    GeneraDatosConceptosInmov = False
        
    'Borramos tmp
    cad = "Delete from Usuarios.ztmppresu1 where codusu = " & vUsu.Codigo
    Conn.Execute cad
    'Abrimos datos
    SQL = "Select * from inmovcon"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        MsgBox "Ning�n dato a mostrar", vbExclamation
    Else
        cad = "INSERT INTO Usuarios.ztmppresu1 (codusu, codigo, cta, titulo, ano, mes, Importe) VALUES (" & vUsu.Codigo
        While Not RS.EOF
            'Metemos los nuevos datos
            SQL = ParaBD(RS!codconam, 1) & ",'" & Format(RS!codconam, "0000") & "'" & ParaBD(RS!nomconam)
            SQL = SQL & ",0" & ParaBD(RS!perimaxi, 1) & ParaBD(RS!coefimaxi, 1) & ")"
            SQL = cad & SQL
            Conn.Execute SQL
            RS.MoveNext
        Wend
        GeneraDatosConceptosInmov = True
    End If
    RS.Close
    Set RS = Nothing
    Exit Function
EGeneraDatosConceptosInmov:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Function


'#################################################
'###########    A � A D I D O     ################  DE  NUEVA CONTA DE DAVID
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
Dim SubTipo As String 'F: fecha   N: numero   T: texto  H: HORA



    PonerDesdeHasta = False
    
    Select Case Tipo
    Case "F", "FEC"
        'Campos fecha
        SubTipo = "F"
    
    Case "CONC", "TDIA", "BIC", "AGE", "COI", "INM", "FRA"
        'concepto
        SubTipo = "N"
        
    Case "CTA", "BAN", "CCO", "SER", "CRY"
        SubTipo = "T"
        
    Case "ASIP", "ASI", "AGE", "DPTO"
        SubTipo = "N"
        
    Case "TIVA", "DIA"
        SubTipo = "N"
        
    Case "TPAG", "FPAG", "REM", "ANYO"
        SubTipo = "N"
        
    End Select
    
    Devuelve = CadenaDesdeHasta(Desde, Hasta, Campo, SubTipo)
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
    
    If SubTipo <> "F" And SubTipo <> "FH" Then
        'Fecha para Crystal Report

        If Not AnyadirAFormula(cadselect, Devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        cad = CadenaDesdeHastaBD(Desde.Text, Hasta.Text, Campo, SubTipo)
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
Dim I  As Integer

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
        For I = 0 To miRsAux.Fields.Count - 1
            cadSQL = cadSQL & ";""" & miRsAux.Fields(I).Name & """"
        Next I
        Print #NF, Mid(cadSQL, 2)
    
    
        'Lineas
        While Not miRsAux.EOF
            cadSQL = ""
            For I = 0 To miRsAux.Fields.Count - 1
                cadSQL = cadSQL & ";""" & DBLet(miRsAux.Fields(I).Value, "T") & """"
            Next I
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
            MsgBox "Fichero:  " & Salida & vbCrLf & "Generado con �xito.", vbInformation
        End If
        CopiarFicheroASalida = True
    End If
End Function

Public Function ImprimeGeneral() As Boolean
    
    Dim HaPulsadoImprimir As Boolean

    Screen.MousePointer = vbHourglass


    'Prueba DAVID
'    frmPpal.SkinFramework.AutoApplyNewWindows = False
   ' frmPpal.SkinFramework.AutoApplyNewThreads = False


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
'    frmPpal.SkinFramework.AutoApplyNewWindows = True
'    frmPpal.SkinFramework.AutoApplyNewThreads = True
    
End Function



Public Sub QuitarPulsacionMas(ByRef T As TextBox)
Dim I As Integer

    Do
        I = InStr(1, T.Text, "+")
        If I > 0 Then T.Text = Mid(T.Text, 1, I - 1) & Mid(T.Text, I + 1)
    Loop Until I = 0
        
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
        Aux = "Relaci�n de Clientes por Cta Ventas.pdf"
    Case 17
        Aux = "Listado de Facturas de Proveedores.pdf"
    Case 18
        Aux = "Relaci�n de Proveedores por Cta Gastos.pdf"
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
        Aux = "Relaci�n de Clientes por Cta Ventas"
    Case 17
        Aux = "Listado de Facturas de Proveedores"
    Case 18
        Aux = "Relaci�n de Proveedores por Cta Gastos"
    Case 19
        Aux = "Modelo 303"
    Case 20
        Aux = "Modelo 340"
    Case 21
        Aux = "Modelo 347"
        
        
    '--------------------------------------------------
    Case 51
        Aux = "Pedido proveedor n�: " ' & outClaveNombreArchiv
        
    Case 100
        Aux = "Factura n�" '& outClaveNombreArchiv
        
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
        MsgBox "No hay direcci�n e-mail a la que enviar.", vbExclamation
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

Public Function PonerParamRPT(Indice As String, nomDocu As String) As Boolean
Dim cad As String
Dim Encontrado As Boolean

        nomDocu = ""
        Encontrado = False
        PonerParamRPT = False
        
        cad = "select informe from scryst where codigo = " & DBSet(Indice, "T")
        
        Set RS = New ADODB.Recordset
        RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RS.EOF Then
            nomDocu = DBLet(RS!Informe, "T")
            Encontrado = True
        End If
        
        If Encontrado = False Or nomDocu = "" Then
            cad = "No se han podido cargar los Par�metros de Tipos de Documentos." & vbCrLf
            MsgBox cad & "Debe configurar la aplicaci�n.", vbExclamation
            PonerParamRPT = False
            Exit Function
        End If
        
        PonerParamRPT = True
    

End Function

Public Sub LanzaProgramaAbrirOutlookMasivo(outTipoDocumento As Integer, Cuerpo As String)
Dim NombrePDF As String
Dim Aux As String
Dim Lanza As String

    On Error GoTo ELanzaProgramaAbrirOutlook

    
    If Not ExisteARIMAILGES Then Exit Sub

    'Primer tema. Copiar el docum.pdf con otro nombre mas significatiov
    Select Case outTipoDocumento
    Case 1
        'Reclamacion
        Aux = "Reclamacion.pdf"
    End Select
    
    SQL = "select tmp347.*, cuentas.razosoci, cuentas.maidatos from tmp347, cuentas "
    SQL = SQL & " where codusu = " & vUsu.Codigo & " and importe <> 0 and tmp347.cta = cuentas.codmacta"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
    
        NombrePDF = App.Path & "\temp\" & RS!NIF
        
        'direccion email
        Aux = DBLet(RS!maidatos)
        Lanza = Aux & "|"
        
        'asunto
        Aux = ""
        Select Case outTipoDocumento
        Case 1 ' reclamaciones
            Aux = RecuperaValor(Cuerpo, 1)
        End Select
        
        Lanza = Lanza & Aux & "|"
        
        
        
    If LCase(Mid(cadNomRPT, 1, 3)) = "esc" Then
    ' para el caso de escalona
    
        cad = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">"
        cad = cad & "<HTML><HEAD><TITLE>Mensaje</TITLE></HEAD>"
        cad = cad & "<TR><TD VALIGN=""TOP""><P><FONT FACE=""Tahoma""><FONT SIZE=3>"
        cad = cad & RecuperaValor(Cuerpo, 2)
        'FijarTextoMensaje
        
        cad = cad & "</FONT></FONT></P></TD></TR><TR><TD VALIGN=""TOP"">"
        cad = cad & "<p class=""MsoNormal""><b><i>"
        cad = cad & "<span style=""font-size: 7.5pt; font-family: Arial,sans-serif; color: #9999FF"">C."
        cad = cad & "R. Reial S�quia Escalona</span></i></b></p>"
        cad = cad & "<p class=""MsoNormal""><em><b>"
        cad = cad & "<span style=""font-size: 7.5pt; font-family: Arial,sans-serif; color: #9999FF"">"
        cad = cad & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; La Junta</span></b></em><span style=""font-size: 10.0pt; font-family: Arial,sans-serif; color: black"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><span style=""font-size: 7.5pt; font-family: Arial,sans-serif; color: #9999FF"">&nbsp;</span></p>"
        cad = cad & "<p class=""MsoNormal"">"
        cad = cad & "<span style=""font-size: 13.5pt; font-family: Arial,sans-serif; color: #9999FF"">"
        cad = cad & "********************</span></p>"
        cad = cad & "<p class=MsoNormal><b>"
         cad = cad & "<span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'>Confidencialidad"
         cad = cad & "</span></b><span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'><br>"
         cad = cad & "Este mensaje y sus archivos adjuntos van dirigidos exclusivamente a su destinatario, "
         cad = cad & "pudiendo contener informaci�n confidencial sometida a secreto profesional. No est� permitida su reproducci�n o "
         cad = cad & "distribuci�n sin la autorizaci�n expresa de Real Acequia Escalona. Si usted no es el destinatario final por favor "
         cad = cad & "elim�nelo e inf�rmenos por esta v�a.<o:p></o:p></span></p><p class=MsoNormal style='mso-margin-top-alt:6.0pt;"
         cad = cad & "margin-right:0cm;margin-bottom:6.0pt;margin-left:0cm;text-align:justify'><span style='font-size:8.0pt;"
         cad = cad & "font-family:""Comic Sans MS"";color:black'>De acuerdo con la Ley 34/2002 (LSSI) y la Ley 15/1999 (LOPD), "
         cad = cad & "usted tiene derecho al acceso, rectificaci�n y cancelaci�n de sus datos personales informados en el fichero del que es "
         cad = cad & "titular Real Acequia Escalona. Si desea modificar sus datos o darse de baja en el sistema de comunicaci�n electr�nica "
         cad = cad & "env�e un correo a</span> <span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'>"
         cad = cad & "<a href=""mailto:escalona@acequiaescalona.org"">escalona@acequiaescalona.org</a> </span><span style='font-size:8.0pt;"
         cad = cad & "font-family:""Comic Sans MS""'>, <span style='color:black'>indicando en la l�nea de <b>&#8220;Asunto&#8221;</b> el derecho "
         cad = cad & "que desea ejercitar. <o:p></o:p></span></span></p><p class=MsoNormal><o:p>&nbsp;</o> "
         
         'ahora en valenciano
         cad = cad & ""
         cad = cad & "<p class=MsoNormal><b>"
         cad = cad & "<span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'>Confidencialitat"
         cad = cad & "</span></b><span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'><br>"
         cad = cad & "Aquest missatge i els seus arxius adjunts van dirigits exclusivamente al seu destinatari, "
         cad = cad & "podent contindre informaci� confidencial sotmesa a secret professional. No est� permesa la seua reproducci� o "
         cad = cad & "distribuci� sense la autoritzaci� expressa de Reial S�quia Escalona. Si vost� no �s el destinatari final per favor "
         cad = cad & "elim�neu-lo e informe-nos per aquesta via.<o:p></o:p></span></p><p class=MsoNormal style='mso-margin-top-alt:6.0pt;"
         cad = cad & "margin-right:0cm;margin-bottom:6.0pt;margin-left:0cm;text-align:justify'><span style='font-size:8.0pt;"
         cad = cad & "font-family:""Comic Sans MS"";color:black'>D'acord amb la Llei 34/2002 (LSSI) i la Llei 15/1999 (LOPD), "
         cad = cad & "vost� t� dret a l'acc�s, rectificaci� i cancelaci� de les seues dades personals informats en el ficher del qu� �s "
         cad = cad & "titolar Reial S�quia Escalona. Si vol modificar les seues dades o donar-se de baixa en el sistema de comunicaci� electr�nica "
         cad = cad & "env�e un correu a</span> <span style='font-size:8.0pt;font-family:""Comic Sans MS"";color:black'>"
         cad = cad & "<a href=""mailto:escalona@acequiaescalona.org"">escalona@acequiaescalona.org</a> </span><span style='font-size:8.0pt;"
         cad = cad & "font-family:""Comic Sans MS""'>, <span style='color:black'>indicant en la l�nea de <b>&#8220;Asumpte&#8221;</b> el dret "
         cad = cad & "que desitja exercitar. <o:p></o:p></span></span></p><p class=MsoNormal><o:p>&nbsp;</o:p></p> "
        
        
        cad = cad & "</TR></BODY></HTML>"
        
        
    Else
    
        cad = RecuperaValor(cad, 2)
        
    End If
        
    ' end
        
        
        'Aqui pondremos lo del texto del BODY
        
        Aux = ""
        Select Case outTipoDocumento
        Case 1 ' reclamaciones
            Aux = cad
        End Select
        Lanza = Lanza & Aux & "|"
        
        'Envio o mostrar
        Lanza = Lanza & "1"   '0. Display   1.  send
        
        'Campos reservados para el futuro
        Lanza = Lanza & "||||"
        
        'El/los adjuntos
        Lanza = Lanza & NombrePDF & "|"
        
        Aux = App.Path & "\ARIMAILGES.EXE" & " " & Lanza  '& vParamAplic.ExeEnvioMail & " " & Lanza
        Shell Aux, vbNormalFocus
        
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    
    
    Exit Sub
ELanzaProgramaAbrirOutlook:
    MuestraError Err.Number, Err.Description
End Sub

