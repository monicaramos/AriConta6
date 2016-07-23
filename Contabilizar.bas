Attribute VB_Name = "Contabilizar"
Option Explicit


        'Se ha añadido un concepto mas a la ampliacion  26 Abril 2007
        '------------------------------------------------------------
        ' De momento lo resolveremos con un simple
        '    devuelvedesdebd.   Si se realentiza mucho deberiamos obtener un recodset
        '    con las titlos de las contrapartidas si el tipo de ampliacion es el 4



Public Sub InsertaTmpActualizar(NumAsien, NumDiari, FechaEnt)
Dim SQL As String
        SQL = "INSERT INTO tmpactualizar (numdiari, fechaent, numasien, codusu) VALUES ("
        SQL = SQL & NumDiari & ",'" & Format(FechaEnt, FormatoFecha) & "'," & NumAsien
        SQL = SQL & "," & vUsu.Codigo & ")"
        Conn.Execute SQL
End Sub


'TipoRemesa:
'           0. Efecto
'           1. Pagare
'           2. Talon
'
' El abono(CONTABILIZACION) de la remesa sera la 572 contra 5208(del banco)
'Y punto, como mucho los gastos si quiero contabilizarlis
Public Function ContabilizarRecordsetRemesa(TipoRemesa As Byte, Norma19 As Boolean, Codigo As Integer, Anyo As Integer, CtaBanco As String, FechaAbono As Date, GastosBancarios As Currency) As Boolean
'Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
Dim Gastos As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim RS As ADODB.Recordset
Dim AmpRemesa As String
Dim CtaParametros As String
Dim Cuenta As String
Dim CuentaPuente As Boolean

'Dim ImporteTalonPagare As Currency    'beneficiosPerdidasTalon: por si hay diferencias entre vtos y total talon
Dim ImpoAux As Currency
Dim VaAlHaber As Boolean
Dim Aux As String
Dim GastosGeneralesRemesasDescontadosDelImporte As Boolean
Dim LCta As Integer
'Noviembre 2009.
'Paramero nuevo
'Contabiliza contra cuenta efectos comerciales decontados
'Son DOS apuntes en el abono
Dim LlevaCtaEfectosComDescontados As Boolean
Dim CtaEfectosComDescontados As String

Dim Obs As String

    On Error GoTo ECon
    ContabilizarRecordsetRemesa = False

    
    GastosGeneralesRemesasDescontadosDelImporte = False
    Cuenta = "GastRemDescontad" 'gastos tramtiaacion remesa descontados importe
    CtaParametros = DevuelveDesdeBD("ctaefectosdesc", "bancos", "codmacta", RecuperaValor(CtaBanco, 1), "T", Cuenta)
    GastosGeneralesRemesasDescontadosDelImporte = Cuenta = "1"
    If GastosGeneralesRemesasDescontadosDelImporte Then
        'Si no tiene gastos generales pongo esto a false tb
        If GastosBancarios = 0 Then GastosGeneralesRemesasDescontadosDelImporte = False
    End If
    Cuenta = ""
    LlevaCtaEfectosComDescontados = False   'Solo sera para efectos bancarios. Tipo FONTENAS
    
    'La forma de pago
    Set vCP = New Ctipoformapago
    If TipoRemesa = 1 Then
        Linea = vbTipoPagoRemesa
        Cuenta = "Efectos"
        
    ElseIf TipoRemesa = 2 Then
        Linea = vbPagare
        Cuenta = "Pagarés"
        'CtaParametros = "pagarecta"
        CuentaPuente = vParamT.PagaresCtaPuente
        
    Else
        Linea = vbTalon
        Cuenta = "Talones"
        'CtaParametros = "taloncta"
        CuentaPuente = vParamT.TalonesCtaPuente
    End If
    
    
    
    If CuentaPuente Then
        If CtaParametros = "" Then
            MsgBox "Mal configurado el banco. Falta configurar cuenta efectos descontados del banco: " & Cuenta, vbExclamation
            Exit Function
        End If
    End If
            
            
    
    
    
            
    'Si llevamos las dos cuentas de efectos descontados, la de cancelacion YA las combrpobo en el proceso de cancelacion
    'ahora tenemos que comprobar la de efectos descontados pendientes de cobro
    If LlevaCtaEfectosComDescontados Then
        Set RS = New ADODB.Recordset
        LCta = Len(CtaEfectosComDescontados)
        If LCta < vEmpresa.DigitosUltimoNivel Then
        
            Conn.Execute "DELETE from tmpcierre1 where codusu = " & vUsu.Codigo
                
            Ampliacion = ",CONCAT('" & CtaEfectosComDescontados & "',SUBSTRING(codmacta," & LCta + 1 & ")" & ")"
            ''SQL = "Select " & vUsu.Codigo & Ampliacion & " from scarecepdoc where codigo=" & IdRecepcion
                
            SQL = "Select " & vUsu.Codigo & Ampliacion & " from cobros WHERE codrem=" & Codigo & " AND anyorem = " & Anyo
            SQL = SQL & " GROUP BY codmacta"
            'INSERT
            SQL = "INSERT INTO tmpcierre1(codusu,cta) " & SQL
            Conn.Execute SQL
            
            'Ahora monto el select para ver que cuentas 430 no tienen la 4310
            SQL = "Select cta,codmacta from tmpcierre1 left join cuentas on tmpcierre1.cta=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
            SQL = SQL & " HAVING codmacta is null"
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            Linea = 0
            While Not RS.EOF
                Linea = Linea + 1
                SQL = SQL & RS!Cta & "     "
                If Linea = 5 Then
                    SQL = SQL & vbCrLf
                    Linea = 0
                End If
                RS.MoveNext
            Wend
            RS.Close
            
            If SQL <> "" Then
                
                AmpRemesa = "Abono remesa"
                
                SQL = "Cuentas " & AmpRemesa & ".  No existen las cuentas: " & vbCrLf & String(90, "-") & vbCrLf & SQL
                SQL = SQL & vbCrLf & "¿Desea crearlas?"
                Linea = 1
                If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                    'Ha dicho que si desea crearlas
                    
                    Ampliacion = "CONCAT('" & CtaEfectosComDescontados & "',SUBSTRING(codmacta," & LCta + 1 & ")) "
                
                    'SQL = "Select codmacta," & Ampliacion & " from scarecepdoc where codigo=" & IdRecepcion
                    SQL = "Select codmacta," & Ampliacion & " from cobros WHERE codrem=" & Codigo & " AND anyorem = " & Anyo
                    SQL = SQL & " and " & Ampliacion & " in "
                    SQL = SQL & "(Select cta from tmpcierre1 left join cuentas on tmpcierre1.cta=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
                    SQL = SQL & " AND codmacta is null)"
                    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not RS.EOF
                    
                         SQL = "INSERT IGNORE INTO  cuentas(codmacta,nommacta ,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos) SELECT '"
                                    ' CUenta puente
                         SQL = SQL & RS.Fields(1) & "',nommacta ,'S',0,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos from cuentas where codmacta = '"
                                    'Cuenta en la scbro (codmacta)
                         SQL = SQL & RS.Fields(0) & "'"
                         Conn.Execute SQL
                         RS.MoveNext
                         
                    Wend
                    RS.Close
                    Linea = 0
                End If
                If Linea = 1 Then GoTo ECon
            End If
            
        Else
            'Cancela contra UNA unica cuenta todos los vencimientos
            SQL = DevuelveDesdeBD("codmacta", "cuentas", "codmacta", CtaEfectosComDescontados, "T")
            If SQL = "" Then
                MsgBox "No existe la cuenta efectos comerciales descontados : " & CtaEfectosComDescontados, vbExclamation
                GoTo ECon
            End If
        End If
        Set RS = Nothing
    End If  'de comprobar cta efectos comerciales descontados
            
            
    If vCP.Leer(Linea) = 1 Then GoTo ECon
    
    
    Set Mc = New Contadores
    
    
    If Mc.ConseguirContador("0", FechaAbono <= vParam.fechafin, True) = 1 Then Exit Function
    
    
    
    'Insertamos la cabera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion, usucreacion, desdeaplicacion) VALUES ("
    SQL = SQL & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ", '"
    SQL = SQL & "Abono remesa: " & Codigo & " / " & Anyo & "   " & Cuenta & vbCrLf
    SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy") & " por " & vUsu.Nombre & "',"
    
    Obs = Codigo & " / " & Anyo & "   " & Cuenta
    
    SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Abono remesa: " & Obs & "');"
    If Not Ejecuta(SQL) Then Exit Function
    
    
    Linea = 1
    Importe = 0
    Gastos = 0
    Set RS = New ADODB.Recordset
    
    
    
    
    'La ampliacion para el banco
    AmpRemesa = ""
    SQL = "Select * from remesas WHERE codigo=" & Codigo & " AND anyo = " & Anyo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    'NO puede ser EOF
    
    
    Importe = RS!Importe

    
    If Not IsNull(RS!Descripcion) Then AmpRemesa = RS!Descripcion
    
    
    If AmpRemesa = "" Then
        AmpRemesa = " Remesa: " & Codigo & "/" & Anyo
    Else
        AmpRemesa = " " & AmpRemesa
    End If
    
    RS.Close
    
    'AHORA Febrero 2009
    '572 contra  5208  Efectos descontados
    '-------------------------------------
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada, "
    SQL = SQL & " numserie,numfaccl,fecfactu,numorden,tipforpa, tiporem,codrem,anyorem) "
    SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"


    Gastos = 0
    If CuentaPuente Then
        
        'DOS LINEAS POR APUNTE, banco contra efectos descontados
        'A no ser que sea TAL/PAG y pueden haber beneficios o perdidas por diferencias de importes
        SQL = SQL & CtaParametros & "','RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.conhacli
    
        Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
        Ampliacion = Ampliacion & " RE. " & Codigo & "/" & Anyo
        SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
    
    
        SQL = SQL & "NULL," & TransformaComasPuntos(CStr(Importe)) & ",NULL,"
    
        If vCP.ctrhacli = 1 Then
            If CuentaPuente And Not LlevaCtaEfectosComDescontados Then
                SQL = SQL & "'" & RecuperaValor(CtaBanco, 1) & "',"
            Else
                'NO lleva cuenta puente
                'Directamente contra el cliente
                If Not LlevaCtaEfectosComDescontados Then
                    SQL = SQL & "'" & RS!codmacta & "',"
                Else
                    SQL = SQL & "NULL,"
                End If
            End If
        Else
            SQL = SQL & "NULL,"
        End If
        SQL = SQL & "'COBROS',0,"
        
        '###FALTA1
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
        

        If Not Ejecuta(SQL) Then Exit Function
  
        Linea = Linea + 1
    
    
    
       'Lleva cta efectos comerciales descontados
        If LlevaCtaEfectosComDescontados Then
            'AQUI
            'Para cada efecto cancela la 5208 contra las CtaEfectosComDescontados(4311x)
 
            
            Aux = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
            
            
            SQL = "Select * from cobros WHERE codrem=" & Codigo & " AND anyorem = " & Anyo
            RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            While Not RS.EOF
        
                'Banco contra cliente
                'La linea del banco
                SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
                SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
                SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada,numserie,numfaccl,fecfactu,numorden,tipforpa, tiporem,codrem,anyorem) "
                SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"
            
                'Cuenta
                SQL = SQL & CtaEfectosComDescontados
                If LCta <> vEmpresa.DigitosUltimoNivel Then SQL = SQL & Mid(RS!codmacta, LCta + 1)
                
                SQL = SQL & "','" & Format(RS!NumFactu, "000000000") & "'," & vCP.conhacli
            
            
                
                Ampliacion = Aux & " "
            
                                'Neuvo dato para la ampliacion en la contabilizacion
                Select Case vCP.amphacli
                Case 2
                   Ampliacion = Ampliacion & Format(RS!FecVenci, "dd/mm/yyyy")
                Case 4
                    'Contrapartida BANCO
                    Cuenta = RecuperaValor(CtaBanco, 1)
                    Cuenta = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cuenta, "T")
                    Ampliacion = Ampliacion & AmpRemesa
                Case Else
                   If vCP.amphacli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
                   Ampliacion = Ampliacion & RS!NUmSerie & "/" & RS!NumFactu
                End Select
                SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
                
                
                ' debe timporteH, codccost, ctacontr, idcontab, punteada
                'Importe
                SQL = SQL & TransformaComasPuntos(RS!ImpVenci) & ",NULL,NULL,"
            
                If vCP.ctrdecli = 1 Then
                    SQL = SQL & "'" & CtaParametros & "',"
                Else
                    SQL = SQL & "NULL,"
                End If
                SQL = SQL & "'COBROS',0,"
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & "," & ValorNulo & ValorNulo & "," & ValorNulo & ")"
                '###FALTA1
                
                
                If Not Ejecuta(SQL) Then Exit Function
                
                Linea = Linea + 1
                RS.MoveNext
            Wend
            RS.Close
            
        End If   'de lleva cta de efectos comerciales descontados
        
        
    Else
        
        
        
        Importe = 0
        SQL = "Select * from cobros WHERE codrem=" & Codigo & " AND anyorem = " & Anyo
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RS.EOF
        
            'Banco contra cliente
            'La linea del banco
            SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
            SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
            SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada,numserie,numfaccl,fecfactu,numorden,tipforpa, tiporem,codrem,anyorem) "
            SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"
        
            'Cuenta
            SQL = SQL & RS!codmacta & "','" & RS!NUmSerie & Format(RS!NumFactu, "0000000") & "'," & vCP.conhacli
    
            
            
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
            Ampliacion = Ampliacion & " "
                   
            'Neuvo dato para la ampliacion en la contabilizacion
            Select Case vCP.amphacli
            Case 2
               Ampliacion = Ampliacion & Format(RS!FecVenci, "dd/mm/yyyy")
            Case 4
                'Contrapartida BANCO
                Cuenta = RecuperaValor(CtaBanco, 1)
                Cuenta = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cuenta, "T")
                Ampliacion = Ampliacion & AmpRemesa
            Case Else
               If vCP.amphacli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
               Ampliacion = Ampliacion & RS!NUmSerie & Format(RS!NumFactu, "0000000")
            End Select
            SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
            
            Importe = Importe + RS!ImpVenci
                
            Gastos = Gastos + DBLet(RS!Gastos, "N")
            
            ' timporteH, codccost, ctacontr, idcontab, punteada
            'Importe
            SQL = SQL & "NULL," & TransformaComasPuntos(RS!ImpVenci) & ",NULL,"
        
            If vCP.ctrdecli = 1 Then
                SQL = SQL & "'" & RecuperaValor(CtaBanco, 1) & "',"
            Else
                SQL = SQL & "NULL,"
            End If
            SQL = SQL & "'COBROS',0,"
            
            'los datos de la factura (solo en el apunte del cliente)
            Dim TipForpa As Byte
            TipForpa = DevuelveDesdeBD("tipforpa", "formapago", "codforpa", RS!codforpa, "N")
            
            SQL = SQL & DBSet(RS!NUmSerie, "T") & "," & DBSet(RS!NumFactu, "N") & "," & DBSet(RS!FecFactu, "F") & "," & DBSet(RS!numorden, "N") & "," & DBSet(TipForpa, "N") & ","
            SQL = SQL & TipoRemesa & "," & Codigo & "," & Anyo & ")"
            
            If Not Ejecuta(SQL) Then Exit Function
            
            Linea = Linea + 1
            RS.MoveNext
        
        Wend
        RS.Close
            
    End If
    
    'La linea del banco
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
    SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & ","

    
    'Gastos de los recibos.
    'Si tiene alguno de los efectos remesados gastos
    If Gastos > 0 Then
        Linea = Linea + 1
        Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
        Ampliacion = "RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.conhacli & ",'" & Ampliacion & " " & Codigo & "/" & Anyo & "'"



        Ampliacion = Linea & ",'" & RecuperaValor(CtaBanco, 4) & "','" & Ampliacion & ",NULL,"
        Ampliacion = Ampliacion & TransformaComasPuntos(CStr(Gastos)) & ","

        If RecuperaValor(CtaBanco, 3) = "" Then
            Ampliacion = Ampliacion & "NULL"
        Else
            Ampliacion = Ampliacion & "'" & RecuperaValor(CtaBanco, 3) & "'"
        End If
        
        Ampliacion = Ampliacion & ",NULL,'COBROS',0)"

        Ampliacion = SQL & Ampliacion
        If Not Ejecuta(Ampliacion) Then Exit Function
        Linea = Linea + 1
    End If
    
  
    'AGOSTO 2009
    'Importe final banco
    'Y desglose en TAL/PAG de los beneficios o perdidas si es que tuviera
    
    ImpoAux = Importe + Gastos
    
    'NOV 2009
    'Gastos tramitacion descontados del importe
    If GastosGeneralesRemesasDescontadosDelImporte And GastosBancarios > 0 Then
        ImpoAux = ImpoAux - GastosBancarios
        'Para que la linea salga al final del asiento, juego con numlinea
        Linea = Linea + 1
    End If
    
    Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
    Ampliacion = Ampliacion & AmpRemesa
    Ampliacion = Linea & ",'" & RecuperaValor(CtaBanco, 1) & "','RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.condecli & ",'" & Ampliacion & "',"
    Ampliacion = Ampliacion & TransformaComasPuntos(CStr(ImpoAux)) & ",NULL,NULL,"
    
    If vCP.ctrdecli = 0 Then
        Ampliacion = Ampliacion & "NULL"
    Else
        If CuentaPuente Then
            If Not LlevaCtaEfectosComDescontados Then
                Ampliacion = Ampliacion & "'" & CtaParametros & "'"
            Else
                Ampliacion = Ampliacion & "NULL"
            End If
        Else
            Ampliacion = Ampliacion & "NULL"
        End If
       
    End If
    Ampliacion = Ampliacion & ",'COBROS',0)"


    Ampliacion = SQL & Ampliacion
    If Not Ejecuta(Ampliacion) Then Exit Function
    
    'Juego con la linea
    
    'Gastos bancarios derivados de la tramitacion de la remesa.
    'Metemos dos lineas mas. Podriamos meter una si en el importe anterior le restamos los gastos bancarios
    If GastosBancarios > 0 Then
        SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
        SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
        SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada,tiporem,codrem,anyorem) "
        SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & ","
        
        
        
        'imporeted timporteH, codccost, ctacontr, idcontab, punteada) "
        If GastosGeneralesRemesasDescontadosDelImporte Then
            'He jugado con el orden para k la linea anterior salga la ultima
            Linea = Linea - 1
        Else
            Linea = Linea + 1
        End If
        Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
        Ampliacion = Ampliacion & " Gastos Remesa:" & Codigo & " / " & Anyo
        Ampliacion = DevNombreSQL(Ampliacion)
    
        ' numdocum, codconce, ampconce
        Ampliacion = "'RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.condecli & ",'" & Ampliacion & "',"
        Ampliacion = Linea & ",'" & RecuperaValor(CtaBanco, 2) & "'," & Ampliacion
        
        Ampliacion = Ampliacion & TransformaComasPuntos(CStr(GastosBancarios)) & ",NULL,"
        'CENTRO DE COSTE
        If vParam.autocoste Then
            Ampliacion = Ampliacion & "'" & RecuperaValor(CtaBanco, 3) & "'"
        Else
            Ampliacion = Ampliacion & "NULL"
        End If
        Ampliacion = Ampliacion & ",'" & RecuperaValor(CtaBanco, 1) & "','COBROS',0," & TipoRemesa & "," & DBSet(Codigo, "N") & "," & DBSet(Anyo, "N") & ")"
        Ampliacion = SQL & Ampliacion
        
        If Not Ejecuta(Ampliacion) Then Exit Function
        
        If Not GastosGeneralesRemesasDescontadosDelImporte Then
            SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
            SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
            SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
            SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & ","
            
            
            
            Linea = Linea + 1
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
            Ampliacion = Ampliacion & " Gastos Remesa: " & Codigo & " / " & Anyo
            Ampliacion = Linea & ",'" & RecuperaValor(CtaBanco, 1) & "','RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.conhacli & ",'" & Ampliacion & "',"
            Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(GastosBancarios)) & ",NULL,'"
            Ampliacion = Ampliacion & RecuperaValor(CtaBanco, 2) & "','COBROS',0)"
            Ampliacion = SQL & Ampliacion
            If Not Ejecuta(Ampliacion) Then Exit Function
        End If
            
        If GastosGeneralesRemesasDescontadosDelImporte Then Linea = Linea + 2
    End If
    
    
    'Noviembre 2009
    '-------------------------------------------
    'Efectos. Si lleva cta puente, y lleva la segunda cuenta puente
    If LlevaCtaEfectosComDescontados Then
    
        'Crearemos n x 2 lineas de apunte de los efectos remesados
        'siendo
        '       CtaEfectosComDescontados        contra   CtaParametros (431x)
        '            y el aseinto de contrapartida
    
        Aux = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
        CtaEfectosComDescontados = DevuelveDesdeBD("RemesaCancelacion", "paramtesor", "codigo", "1")
        LCta = Len(CtaEfectosComDescontados)
        If LCta = 0 Then
            MsgBox "Deberia tener valor el paremtro de cta puente", vbCritical
            LCta = Val(RS!davidadavi) 'QUE GENERE UN ERROR
        End If
        
        CtaParametros = RecuperaValor(CtaBanco, 1) 'Cuenta del banco para la contrpartida
        Linea = Linea + 1
        SQL = "Select * from cobros WHERE codrem=" & Codigo & " AND anyorem = " & Anyo
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RS.EOF
        
            SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
            SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
            SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
            SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"
        
            'Cuenta
            SQL = SQL & CtaEfectosComDescontados
            If LCta <> vEmpresa.DigitosUltimoNivel Then SQL = SQL & Mid(RS!codmacta, LCta + 1)
            
            SQL = SQL & "','" & RS!NUmSerie & Format(RS!NumFactu, "0000000") & "'," & vCP.conhacli
        
        
            
            Ampliacion = Aux & " "
        
                            'Neuvo dato para la ampliacion en la contabilizacion
            Select Case vCP.amphacli
            Case 2
               Ampliacion = Ampliacion & Format(RS!FecVenci, "dd/mm/yyyy")
            Case 4
                'Contrapartida BANCO
                Cuenta = RecuperaValor(CtaBanco, 1)
                Cuenta = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cuenta, "T")
                Ampliacion = Ampliacion & AmpRemesa
            Case Else
               If vCP.amphacli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
               Ampliacion = Ampliacion & RS!NUmSerie & Format(RS!NumFactu, "0000000")
            End Select
            SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
            
            
            ' timporteH, codccost, ctacontr, idcontab, punteada
            'Importe
            SQL = SQL & "NULL," & TransformaComasPuntos(RS!ImpVenci) & ",NULL,"
        
            If vCP.ctrdecli = 1 Then
                SQL = SQL & "'" & CtaParametros & "',"
            Else
                SQL = SQL & "NULL,"
            End If
            SQL = SQL & "'COBROS',0)"
            
            If Not Ejecuta(SQL) Then Exit Function
            Linea = Linea + 1
            
            RS.MoveNext
        Wend
        RS.Close
    
    End If
    

    'AHora actualizamos los efectos.
    SQL = "UPDATE cobros SET"
    SQL = SQL & " siturem= 'Q'"
    SQL = SQL & ", situacion = 1 "
'    SQL = SQL & ", ctabanc2= '" & RecuperaValor(CtaBanco, 1) & "'"
'    SQL = SQL & ", contdocu= 1"   'contdocu indica k se ha contabilizado
    SQL = SQL & " WHERE codrem=" & Codigo
    SQL = SQL & " and anyorem=" & Anyo
'++ la he añadido yo, antes no estaba
    SQL = SQL & " and tiporem = " & TipoRemesa
    
    Conn.Execute SQL

    Dim MaxLin As Integer

    'Insertamos para pasar a hco
    InsertaTmpActualizar Mc.Contador, vCP.diaricli, FechaAbono
    
    'Todo OK
    ContabilizarRecordsetRemesa = True
    
ECon:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    
    End If
    Set RS = Nothing
    Set Mc = Nothing
    Set vCP = Nothing
End Function


'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'   DEVOLUCION DE REGISTROS
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------


    'OK. Ya tengo grabada la temporal con los recibos que devuelvo. Ahora
    'hare:
    '       - generar un asiento con los datos k devuelvo
    '       - marcar los cobros como devueltos, añadirle el gasto y insertar en la
    '           tabla de hco de devueltos
    
    'La variable remesa traera todos los valores
    
    '21 Octubre 2011
    'Desdoblaremos el procedimiento de deolucion
    'de talones-pagares frente a efectos
Public Function RealizarDevolucionRemesa(FechaDevolucion As Date, ContabilizoGastoBanco As Boolean, CtaBenBancarios As String, Remesa As String, DatosContabilizacionDevolucion As String) As Boolean
Dim C As String
    
    C = RecuperaValor(Remesa, 10)
    
    CtaBenBancarios = DevuelveDesdeBD("ctagastos", "bancos", "codmacta", RecuperaValor(Remesa, 3), "T")
    If CtaBenBancarios = "" Then
        CtaBenBancarios = DevuelveDesdeBD("ctabenbanc", "paramtesor", "codigo", "1", "N")
    End If
    
    
    If C = "1" Then
        RealizarDevolucionRemesa = RealizarDevolucionRemesaEfectos(FechaDevolucion, ContabilizoGastoBanco, CtaBenBancarios, Remesa, DatosContabilizacionDevolucion)
    Else
        RealizarDevolucionRemesa = RealizarDevolucionRemesaTalPag(FechaDevolucion, ContabilizoGastoBanco, CtaBenBancarios, Remesa, DatosContabilizacionDevolucion)
    End If
    
End Function


Public Function RealizarDevolucionRemesaEfectos(FechaDevolucion As Date, ContabilizoGastoBanco As Boolean, CtaBenBancarios As String, Remesa As String, DatosContabilizacionDevolucion As String) As Boolean

'Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim RS As ADODB.Recordset
Dim Amp11 As String
Dim DescRemesa As String
Dim CuentaPuente As Boolean
Dim TipoRemesa As Byte
Dim SubCtaPte As String
'Dim AgrupaApunteBanco As Boolean
Dim GastoDevolucion As Currency
Dim DescuentaImporteDevolucion As Boolean
Dim GastoVto As Currency
Dim Gastos As Currency  'de cada recibo/vto
Dim Aux As String
Dim Importeauxiliar As Currency
Dim CtaBancoGastos As String
Dim CCBanco As String
Dim Agrupa431x As Boolean
Dim Agrupa4311x As Boolean   'Segunad cuenta de cancelacion TIPO fontenas
Dim CtaEfectosComDescontados As String   '   tipo FONTENAS
Dim LINAPU As String

    On Error GoTo ECon
    RealizarDevolucionRemesaEfectos = False
    
    
    'La forma de pago
    Set vCP = New Ctipoformapago
    Set RS = New ADODB.Recordset
    
    
    'Leo la descipcion de la remesa si alguna de las ampliaciones me la solicita
    DescRemesa = ""
    Aux = RecuperaValor(Remesa, 8)
    If Aux <> "" Then
        'OK viene de fichero
        Aux = RecuperaValor(Remesa, 9)
        'Vuelvo a susitiuri los # por |
        Aux = Replace(Aux, "#", "|")
        SQL = ""
        For Linea = 1 To Len(Aux)
            If Mid(Aux, Linea, 1) = "·" Then SQL = SQL & "X"
        Next
        
        If Len(SQL) > 1 Then
            'Tienen mas de una remesa
            SQL = ""
            While Aux <> ""
                Linea = InStr(1, Aux, "·")
                If Linea = 0 Then
                    Aux = ""
                Else
                    SQL = SQL & ",    " & Format(RecuperaValor(Mid(Aux, 1, Linea - 1), 1), "000") & "/" & RecuperaValor(Mid(Aux, 1, Linea - 1), 2) & ""
                    Aux = Mid(Aux, Linea + 1)
                End If
            
            Wend
            Aux = RecuperaValor(Remesa, 8)
            SQL = "Devolución remesas: " & Trim(Mid(SQL, 2))
            DescRemesa = SQL & vbCrLf & "Fichero: " & Aux
        End If
        
    End If

    
    
    DescRemesa = RecuperaValor(Remesa, 9)
    TipoRemesa = RecuperaValor(Remesa, 10)
    
    
    If TipoRemesa = 1 Then
        Linea = vbTipoPagoRemesa

        'Noviembre 2009. Tipo FONTENAS
        SubCtaPte = "RemesaCancelacion"
        SQL = "ctaefectcomerciales"
    Else
        If TipoRemesa = 2 Then
            Linea = vbPagare
            SQL = "pagarecta"
            
        Else
            Linea = vbTalon
            SQL = "taloncta"
        End If

    End If
    
    If vCP.Leer(Linea) = 1 Then GoTo ECon


    'Los parametros de contbilizacion se le pasan en el frame de pedida de datos
    'Ahora se los asignaremos a la formma de pago
    vCP.condecli = RecuperaValor(DatosContabilizacionDevolucion, 1)
    vCP.ampdecli = RecuperaValor(DatosContabilizacionDevolucion, 2)
    vCP.conhacli = RecuperaValor(DatosContabilizacionDevolucion, 1) '3)
    vCP.amphacli = RecuperaValor(DatosContabilizacionDevolucion, 2) '4)
    SQL = RecuperaValor(DatosContabilizacionDevolucion, 5)  'agupa o no
    Agrupa431x = SQL = "1"
    
    
    
    SQL = RecuperaValor(Remesa, 7)
    GastoDevolucion = TextoAimporte(SQL)
    DescuentaImporteDevolucion = False
    If GastoDevolucion > 0 Then
        CtaBancoGastos = "CtaIngresos"
        SQL = RecuperaValor(Remesa, 3)
        SQL = DevuelveDesdeBD("GastRemDescontad", "bancos", "codmacta", SQL, "T")
        If SQL = "1" Then DescuentaImporteDevolucion = True
    End If
    
    'Datos del banco
    SQL = RecuperaValor(Remesa, 3)
    SQL = "Select * from bancos where codmacta ='" & SQL & "'"
    CCBanco = ""
    CtaBancoGastos = ""
    CtaEfectosComDescontados = ""
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        SQL = "No se ha encontrado banco: " & vbCrLf & SQL
        Err.Raise 516, SQL
    End If
    CCBanco = DBLet(RS!codccost, "T")
    CtaBancoGastos = DBLet(RS!ctagastos, "T")
    If Not vParam.autocoste Then CCBanco = ""  'NO lleva analitica
    RS.Close
    
    If TipoRemesa = 1 Then
        CtaEfectosComDescontados = DevuelveDesdeBD("ctaefectcomerciales", "paramtesor", "codigo", "1")
    Else
        CtaEfectosComDescontados = ""
    End If
    Agrupa4311x = False 'La de fontenas
    If Agrupa431x Then
        'QUIERE AGRUPAR. Veremos is por la longitud de las puentes, puede agrupar
        Agrupa4311x = True
        If Len(SubCtaPte) <> vEmpresa.DigitosUltimoNivel Then Agrupa431x = False 'NO puede agrupar
        If Len(CtaEfectosComDescontados) <> vEmpresa.DigitosUltimoNivel Then Agrupa4311x = False 'NO puede agrupar
        
    End If
    
    'EMPEZAMOS
    'Borramos tmpactualizar
    SQL = "DELETE FROM tmpactualizar where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    
    'Cargaremos los registros a devolver que estaran en la tabla temporal
    'para codusu
    SQL = "Select * from tmpfaclin where codusu=" & vUsu.Codigo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        MsgBox "EOF.  NO se han cargado datos devolucion", vbExclamation
        RS.Close
        GoTo ECon
    End If

    Set Mc = New Contadores


    If Mc.ConseguirContador("0", FechaDevolucion <= vParam.fechafin, True) = 1 Then GoTo ECon


    'Insertamos la cabera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
    SQL = SQL & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ",'"
    
    'Ahora esta en desc remesa
    DescRemesa = DescRemesa & vbCrLf & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy hh:nn") & " por " & vUsu.Nombre
    SQL = SQL & DevNombreSQL(DescRemesa) & "',"
    'SQL = SQL & "'Devolucion remesa: " & Format(RecuperaValor(Remesa, 1), "0000") & " / " & RecuperaValor(Remesa, 2)
    'SQL = SQL & vbCrLf & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy") & " por " & vUsu.Nombre & "')"
    SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Devolución efectos')"

    
    If Not Ejecuta(SQL) Then GoTo ECon




    Linea = 1
    Importe = 0

    If vCP.ampdecli = 3 Then
        Amp11 = DescRemesa
    Else
        Amp11 = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
    End If
    
    'Lo meto en una VAR
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada, "
    SQL = SQL & " numserie,numfaccl,fecfactu,numorden,tipforpa,fecdevol,coddevol,gastodev,tiporem,codrem,anyorem,esdevolucion) "
    SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","
    LINAPU = SQL
    
    While Not RS.EOF

        'Lineas de apuntes .
         SQL = LINAPU & Linea & ",'"
         SQL = SQL & RS!Cta
         SQL = SQL & "','" & RS!NUmSerie & Format(RS!NumFac, "0000000") & "'," & vCP.condecli

        Ampliacion = Amp11 & " "
    
        If vCP.ampdecli = 3 Then
            'NUEVA forma de ampliacion
            'No hacemos nada pq amp11 ya lleva lo solicitado
            
        Else
            If vCP.ampdecli = 4 Then
                'COntrapartida
                Ampliacion = Ampliacion & DevuelveDesdeBD("nommacta", "cuentas", "codmacta", RS!Cta, "T")
                
            Else
                If vCP.ampdecli = 2 Then
                   Ampliacion = Ampliacion & Format(RS!Fecha, "dd/mm/yyyy")
                Else
                   If vCP.ampdecli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
                   'Ampliacion = Ampliacion & RS!NUmSerie & "/" & RS!codfaccl
                   Ampliacion = Ampliacion & RS!NUmSerie & Format(RS!NumFac, "0000000") ' & "/" & RS!NumFac
                   
                End If
            End If
        End If
        SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"

        Importe = Importe + RS!imponible


        GastoVto = 0
        Aux = " numserie='" & RS!NUmSerie & "' AND numfactu=" & RS!NumFac
        Aux = Aux & " AND fecfactu='" & Format(RS!Fecha, FormatoFecha) & "' AND numorden"
        Aux = DevuelveDesdeBD("gastos", "cobros", Aux, CStr(RS!NIF), "N")
        
        If Aux <> "" Then GastoVto = CCur(Aux)
        Gastos = Gastos + GastoVto

        ' timporteH, codccost, ctacontr, idcontab, punteada
        Importeauxiliar = RS!imponible - GastoVto
        SQL = SQL & TransformaComasPuntos(CCur(Importeauxiliar)) & ",NULL,NULL,"

        If vCP.ctrdecli = 1 Then
            SQL = SQL & "'" & RS!Cliente & "',"
        Else
            SQL = SQL & "NULL,"
        End If
        SQL = SQL & "'COBROS',0,"
        
        '%%%%% aqui van todos los datos de la devolucion en la linea de cuenta
        SQL = SQL & DBSet(RS!NUmSerie, "T") & "," & DBSet(RS!NumFac, "N") & "," & DBSet(RS!Fecha, "F") & "," & DBSet(RS!NIF, "N") & ","
            
         '-------------------------------------------------------------------------------------
         'Ahora
         '-------------------------------------------------------------------------------------
         'Lo pongo en uno
             'Actualizamos el registro. Ponemos la marca de devuelto. Y aumentamos el importe de gastos
         'Si es que hay
         Dim SqlCobro As String
         Dim RsCobro As ADODB.Recordset
         Dim ImporteNue As Currency
         
         SqlCobro = "select tipforpa, tiporem, codrem, anyorem, gastos from cobros inner join formapago on cobros.codforpa = formapago.codforpa "
         SqlCobro = SqlCobro & " WHERE numserie='" & RS!NUmSerie & "' AND numfactu=" & RS!NumFac
         SqlCobro = SqlCobro & " AND fecfactu='" & Format(RS!Fecha, FormatoFecha) & "' AND numorden=" & RS!NIF
         
         Set RsCobro = New ADODB.Recordset
         RsCobro.Open SqlCobro, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
         If Not RsCobro.EOF Then
         
'    SQL = SQL & " numserie,numfaccl,fecfactu,numorden,tipforpa,fecdevol,coddevol,gastodev,tiporem,codrem,anyorem) "
            SQL = SQL & DBSet(RsCobro!TipForpa, "N") & "," & DBSet(FechaDevolucion, "F") & "," & DBSet(RS!CtaBase, "T", "S") & ","
            SQL = SQL & DBSet(RS!ImpIva, "N") & "," & DBSet(RsCobro!Tiporem, "N") & "," & DBSet(RsCobro!CodRem, "N") & "," & DBSet(RsCobro!AnyoRem, "N") & ",1)"
              
         
            Ampliacion = "UPDATE cobros SET "
            Ampliacion = Ampliacion & " Devuelto = 1, situacion = 0   "
            ImporteNue = RS!Total - RS!imponible '- Rs!impiva
            
            ImporteNue = DBLet(RsCobro!Gastos, "N")
            If DBLet(RS!ImpIva, "N") > 0 Then
            
                If ImporteNue = 0 Then
                    Ampliacion = Ampliacion & " , Gastos = " & TransformaComasPuntos(CStr(RS!ImpIva))
                Else
                    Ampliacion = Ampliacion & " , Gastos = Gastos + " & TransformaComasPuntos(CStr(RS!ImpIva))
                End If
            End If
            Ampliacion = Ampliacion & " ,impcobro=NULL,codrem= NULL, anyorem = NULL, siturem = NULL,tiporem=NULL,fecultco=NULL,recedocu=0"
            Ampliacion = Ampliacion & " WHERE numserie='" & RS!NUmSerie & "' AND numfactu=" & RS!NumFac
            Ampliacion = Ampliacion & " AND fecfactu='" & Format(RS!Fecha, FormatoFecha) & "' AND numorden=" & RS!NIF
            
            Ejecuta Ampliacion
             
         End If
         Set RsCobro = Nothing

        '%%%%% hasta aqui
        

        If Not Ejecuta(SQL) Then GoTo ECon

        Linea = Linea + 1
        
        
        'Gasto.
        ' Si tiene y no agrupa
        '-------------------------------------------------------
        If GastoVto > 0 And Not Agrupa4311x And Not Agrupa431x Then
        
           'Lineas de apuntes .
            SQL = LINAPU & Linea & ",'"
    
    
            SQL = SQL & CtaBancoGastos & "','" & RS!NUmSerie & Format(RS!NumFac, "0000000") & "'," & vCP.condecli
            SQL = SQL & ",'Gastos vto.'"
            
            
            'Importe al debe
            SQL = SQL & "," & TransformaComasPuntos(CStr(GastoVto)) & ",NULL,"
            
            If CCBanco <> "" Then
                SQL = SQL & "'" & DevNombreSQL(CCBanco) & "'"
            Else
                SQL = SQL & "NULL"
            End If
                
            'Contra partida
            'Si no lleva cuenta puente contabiliza los gastos
            Aux = "NULL"
           
            SQL = SQL & "," & Aux & ",'COBROS',0,"
            SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",1)"
            If Not Ejecuta(SQL) Then Exit Function
            
            Linea = Linea + 1
        
        End If
        
        RS.MoveNext
    Wend
    
    
    RS.MoveFirst



    'Linea de los gastos de cada RECIBO
    'Gastos de los recibos.
    'Si tiene alguno de los efectos remesados gastos
    If Gastos > 0 And (Agrupa4311x Or Agrupa431x) Then
        
        If CtaBancoGastos = "" Then CtaBancoGastos = DevuelveDesdeBD("ctabenbanc", "paramtesor", "codigo", "1")
        
        Aux = "RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2)
        
        SQL = LINAPU & Linea & ",'"
        SQL = SQL & CtaBancoGastos & "','" & Aux & "'," & vCP.condecli
        SQL = SQL & ",'Gastos vtos. " & Format(RecuperaValor(Remesa, 1), "0000") & " / " & RecuperaValor(Remesa, 2) '"
        
        
        'Importe al debe
        SQL = SQL & "'," & TransformaComasPuntos(CStr(Gastos)) & ",NULL,"
        
        If CCBanco <> "" Then
            SQL = SQL & "'" & DevNombreSQL(CCBanco) & "'"
        Else
            SQL = SQL & "NULL"
        End If
            
        'Contra partida
        SQL = SQL & ",NULL,'COBROS',0,"
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",1)"
        
        
        If Not Ejecuta(SQL) Then Exit Function
        
        Linea = Linea + 1
    
    End If

    'La linea del banco
    '*********************************************************************
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
    SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","

    'Nuevo tipo ampliacion
    If vCP.amphacli = 3 Then
        Ampliacion = DescRemesa
    Else
        Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
    End If
    
    Ampliacion = Ampliacion & " Dev.rem:" & Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
    
    Amp11 = RS!Cliente  'cta banco

    'Lleva gasto pero lo descontamos de aqui
    If GastoDevolucion > 0 And DescuentaImporteDevolucion And ContabilizoGastoBanco Then
        Importe = Importe + GastoDevolucion
        'Para que la linea salga al fina
        Linea = Linea + 2
    End If
    Ampliacion = Linea & ",'" & Amp11 & "','RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2) & "'," & vCP.conhacli & ",'" & Ampliacion & "',"
    Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(Importe)) & ",NULL,"
    If CuentaPuente Then
        Ampliacion = Ampliacion & "'" & SubCtaPte & "'"
    Else
        'Nulo
        Ampliacion = Ampliacion & "NULL"
    End If
    Ampliacion = Ampliacion & ",'COBROS',0)"
    Ampliacion = SQL & Ampliacion
    If Not Ejecuta(Ampliacion) Then GoTo ECon
    If GastoDevolucion > 0 And DescuentaImporteDevolucion And ContabilizoGastoBanco Then
        Linea = Linea - 2
        'Dejo el importe como estaba
        Importe = Importe - GastoDevolucion
    Else
        Linea = Linea + 1
    End If
    
    
    'SI hay que contabilizar los gastos de devolucion
    If ContabilizoGastoBanco Then
        
         If GastoDevolucion > 0 And DescuentaImporteDevolucion And ContabilizoGastoBanco Then
         Else
            Linea = Linea + 1
         End If
         SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
         SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
         SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
         SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"

         'Cuenta
         SQL = SQL & CtaBenBancarios & "','RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2) & "'," & vCP.condecli
         'SQL = SQL & Rs!Cta & "','REM" & Format(Rs!numfac, "000000000") & "'," & vCP.condecli
        

        If vCP.ampdecli = 3 Then
            Ampliacion = DescRemesa
        Else
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
            Ampliacion = Ampliacion & " Gasto remesa:" & Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
        End If
        SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"


        ' timporteH, codccost, ctacontr, idcontab, punteada
        'Importe.  Va al debe
        SQL = SQL & TransformaComasPuntos(CStr(GastoDevolucion)) & ",NULL,"
        
        'Centro de coste.
        '--------------------------
        Amp11 = "NULL"
        If vParam.autocoste Then
            Amp11 = DevuelveDesdeBD("codccost", "bancos", "codmacta", RS!Cliente, "T")
            Amp11 = "'" & Amp11 & "'"
        End If
        SQL = SQL & Amp11 & ","

        
        SQL = SQL & "'" & RS!Cliente & "',"
        SQL = SQL & "'COBROS',0)"

        If Not Ejecuta(SQL) Then GoTo ECon

        
        
    
        'El total del banco..
        
        'La linea del banco
        'Rs.MoveFirst
        'Si no agrupa dto importe
        If Not DescuentaImporteDevolucion Then
            Linea = Linea + 1
            SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
            SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
            SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
            SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","
        
            
            If vCP.amphacli = 3 Then
                Ampliacion = DescRemesa
            Else
                Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
                Ampliacion = Ampliacion & " Gasto remesa:" & Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
            End If
            
            Ampliacion = Linea & ",'" & RS!Cliente & "','RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2) & "'," & vCP.conhacli & ",'" & Ampliacion & "',"
            'Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(Importe)) & ",NULL,'" & CtaBenBancarios & "','CONTAB',0)"
            Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(GastoDevolucion)) & ",NULL,'" & CtaBenBancarios & "','COBROS',0)"
            Ampliacion = SQL & Ampliacion
            If Not Ejecuta(Ampliacion) Then GoTo ECon
            
        End If
      
    
    End If

    'Ya tenemos generado el apunte de devolucion
    'Insertamos para su actualziacion
    InsertaTmpActualizar Mc.Contador, vCP.diaricli, FechaDevolucion
    
    
    RealizarDevolucionRemesaEfectos = True
ECon:
    If Err.Number <> 0 Then
        
        Amp11 = "Devolución remesa: " & Remesa & vbCrLf
        If Not Mc Is Nothing Then Amp11 = Amp11 & "MC.cont: " & Mc.Contador & vbCrLf
        Amp11 = Amp11 & Err.Description
        MuestraError Err.Number, Amp11
        
    End If
    Set RS = Nothing
    Set Mc = Nothing
    Set vCP = Nothing
End Function


'*************************************************************************************
Public Function RealizarDevolucionRemesaTalPag(FechaDevolucion As Date, ContabilizoGastoBanco As Boolean, CtaBenBancarios As String, Remesa As String, DatosContabilizacionDevolucion As String) As Boolean

'Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim RS As ADODB.Recordset
Dim Amp11 As String
Dim DescRemesa As String
Dim CuentaPuente As Boolean
Dim TipoRemesa2 As Byte
Dim SubCtaPte As String
'Dim AgrupaApunteBanco As Boolean
Dim GastoDevolucion As Currency
Dim DescuentaImporteDevolucion As Boolean
Dim GastoVto As Currency
Dim Gastos As Currency  'de cada recibo/vto
Dim Aux As String
Dim Importeauxiliar As Currency
Dim CtaBancoGastos As String
Dim CCBanco As String
Dim CtaEfectosComDescontados As String   '   tipo FONTENAS
Dim LINAPU As String

Dim Obs As String

    On Error GoTo ECon
    RealizarDevolucionRemesaTalPag = False
    
    
    'La forma de pago
    Set vCP = New Ctipoformapago
    
    
    'Leo la descipcion de la remesa si alguna de las ampliaciones me la solicita
    SQL = "Select descripcion,tiporem from remesas where codigo =" & RecuperaValor(Remesa, 1)
    SQL = SQL & " AND anyo =" & RecuperaValor(Remesa, 2)
    
    DescRemesa = "Remesa: " & RecuperaValor(Remesa, 1) & " / " & RecuperaValor(Remesa, 2)
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TipoRemesa2 = RS!Tiporem
    If Not IsNull(RS.Fields(0)) Then DescRemesa = DevNombreSQL(RS.Fields(0))
    RS.Close
    
    CuentaPuente = False
    
    
    If TipoRemesa2 = 2 Then
        Linea = vbPagare
        SQL = "pagarecta"
        CuentaPuente = vParamT.PagaresCtaPuente
        
    Else
        Linea = vbTalon
        SQL = "taloncta"
        CuentaPuente = vParamT.TalonesCtaPuente
    End If

    If CuentaPuente Then
     
        SubCtaPte = DevuelveDesdeBD(SQL, "paramtesor", "codigo", "1")
             
        If SubCtaPte = "" Then
            MsgBox "Falta por configurar en parametros", vbExclamation
            Exit Function
           
        End If
    End If

    
    If vCP.Leer(Linea) = 1 Then GoTo ECon


    'Los parametros de contbilizacion se le pasan en el frame de pedida de datos
    'Ahora se los asignaremos a la formma de pago
    vCP.condecli = RecuperaValor(DatosContabilizacionDevolucion, 1)
    vCP.ampdecli = RecuperaValor(DatosContabilizacionDevolucion, 2)
    vCP.conhacli = RecuperaValor(DatosContabilizacionDevolucion, 1)
    vCP.amphacli = RecuperaValor(DatosContabilizacionDevolucion, 2)
    
    
    
    
    SQL = RecuperaValor(Remesa, 7)
    GastoDevolucion = TextoAimporte(SQL)
    DescuentaImporteDevolucion = False
    If GastoDevolucion > 0 Then
        CtaBancoGastos = "CtaIngresos"
        SQL = RecuperaValor(Remesa, 3)
        SQL = DevuelveDesdeBD("GastRemDescontad", "bancos", "codmacta", SQL, "T")
        If SQL = "1" Then DescuentaImporteDevolucion = True
    End If
    
    'Datos del banco
    SQL = RecuperaValor(Remesa, 3)
    SQL = "Select * from bancos where codmacta ='" & SQL & "'"
    CCBanco = ""
    CtaBancoGastos = ""
    CtaEfectosComDescontados = ""
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        SQL = "No se ha encontrado banco: " & vbCrLf & SQL
        Err.Raise 516, SQL
    End If
    CCBanco = DBLet(RS!codccost, "T")
    CtaBancoGastos = DBLet(RS!ctagastos, "T")
    If Not vParam.autocoste Then CCBanco = ""  'NO lleva analitica
    RS.Close
    

    CtaEfectosComDescontados = ""


    
    'EMPEZAMOS
    'Borramos tmpactualizar
    SQL = "DELETE FROM tmpactualizar where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    
    'Cargaremos los registros a devolver que estaran en la tabla temporal
    'para codusu
    SQL = "Select * from tmpfaclin where codusu=" & vUsu.Codigo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        MsgBox "EOF.  NO se han cargado datos devolucion", vbExclamation
        RS.Close
        GoTo ECon
    End If

    Set Mc = New Contadores


    If Mc.ConseguirContador("0", FechaDevolucion <= vParam.fechafin, True) = 1 Then GoTo ECon


    'Insertamos la cabera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
    SQL = SQL & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ", '"
    SQL = SQL & "Devolucion remesa(T/P): " & Format(RecuperaValor(Remesa, 1), "0000") & " / " & RecuperaValor(Remesa, 2)
    SQL = SQL & vbCrLf & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy") & " por " & vUsu.Nombre & "',"
    
    
    SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Devolución remesa(T/P)" & Format(RecuperaValor(Remesa, 1), "0000") & " / " & RecuperaValor(Remesa, 2) & "')"
    
    
    If Not Ejecuta(SQL) Then GoTo ECon


    Linea = 1
    Importe = 0

    If vCP.ampdecli = 3 Then
        Amp11 = DescRemesa
    Else
        Amp11 = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
    End If
    
    'Lo meto en una VAR
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada,  numserie,numfaccl,fecfactu,numorden,tipforpa,fecdevol,coddevol,gastodev,tiporem,codrem,anyorem,esdevolucion) "
    SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","
    LINAPU = SQL
    
    While Not RS.EOF

        'Lineas de apuntes .
         SQL = LINAPU & Linea & ",'"
         SQL = SQL & RS!Cta
         SQL = SQL & "','" & Format(RS!NumFac, "0000000") & "'," & vCP.condecli

        Ampliacion = Amp11 & " "
    
        If vCP.ampdecli = 3 Then
            'NUEVA forma de ampliacion
            'No hacemos nada pq amp11 ya lleva lo solicitado
            
        Else
            If vCP.ampdecli = 4 Then
                'COntrapartida
                Ampliacion = Ampliacion & DevuelveDesdeBD("nommacta", "cuentas", "codmacta", RS!Cta, "T")
                
            Else
                If vCP.ampdecli = 2 Then
                   Ampliacion = Ampliacion & Format(RS!Fecha, "dd/mm/yyyy")
                Else
                   If vCP.ampdecli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
                   'Ampliacion = Ampliacion & RS!NUmSerie & "/" & RS!codfaccl
                   Ampliacion = Ampliacion & RS!iva & "/" & RS!NumFac
                   
                End If
            End If
        End If
        SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"

        Importe = Importe + RS!imponible


        GastoVto = 0
        Aux = " numserie='" & RS!iva & "' AND numfactu=" & RS!NumFac
        Aux = Aux & " AND fecfactu='" & Format(RS!Fecha, FormatoFecha) & "' AND numorden"
        Aux = DevuelveDesdeBD("gastos", "cobros", Aux, CStr(RS!NIF), "N")
        
        If Aux <> "" Then GastoVto = CCur(Aux)
        Gastos = Gastos + GastoVto

        ' timporteH, codccost, ctacontr, idcontab, punteada
        Importeauxiliar = RS!imponible - GastoVto
        SQL = SQL & TransformaComasPuntos(CCur(Importeauxiliar)) & ",NULL,NULL,"

        If vCP.ctrdecli = 1 Then
            If CuentaPuente Then
                If Len(SubCtaPte) = vEmpresa.DigitosUltimoNivel Then
                    SQL = SQL & "'" & SubCtaPte & "',"
                Else
                    SQL = SQL & "'" & SubCtaPte & Mid(RS!Cta, Len(SubCtaPte) + 1) & "',"
                End If
            Else
                SQL = SQL & "'" & RS!Cliente & "',"
            End If
        Else
            SQL = SQL & "NULL,"
        End If
        SQL = SQL & "'COBROS',0,"
        
        '%%%%% aqui van todos los datos de la devolucion en la linea de cuenta
        SQL = SQL & DBSet(RS!NUmSerie, "T") & "," & DBSet(RS!NumFac, "N") & "," & DBSet(RS!Fecha, "F") & "," & DBSet(RS!NIF, "N") & ","

         '-------------------------------------------------------------------------------------
         'Ahora
         '-------------------------------------------------------------------------------------
         'Lo pongo en uno
             'Actualizamos el registro. Ponemos la marca de devuelto. Y aumentamos el importe de gastos
         'Si es que hay
         Dim SqlCobro As String
         Dim RsCobro As ADODB.Recordset
         Dim ImporteNue As Currency
         
         SqlCobro = "select tipforpa, tiporem, codrem, anyorem, gastos from cobros inner join formapago on cobros.codforpa = formapago.codforpa "
         SqlCobro = SqlCobro & " WHERE numserie='" & RS!NUmSerie & "' AND numfactu=" & RS!NumFac
         SqlCobro = SqlCobro & " AND fecfactu='" & Format(RS!Fecha, FormatoFecha) & "' AND numorden=" & RS!NIF
         
         Set RsCobro = New ADODB.Recordset
         RsCobro.Open SqlCobro, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
         If Not RsCobro.EOF Then
         
'    SQL = SQL & " numserie,numfaccl,fecfactu,numorden,tipforpa,fecdevol,coddevol,gastodev,tiporem,codrem,anyorem) "
            SQL = SQL & DBSet(RsCobro!TipForpa, "N") & "," & DBSet(FechaDevolucion, "F") & "," & DBSet(RS!CtaBase, "T", "S") & ","
            SQL = SQL & DBSet(RS!ImpIva, "N") & "," & DBSet(RsCobro!Tiporem, "N") & "," & DBSet(RsCobro!CodRem, "N") & "," & DBSet(RsCobro!AnyoRem, "N") & ",1)"
              
         
            Ampliacion = "UPDATE cobros SET "
            Ampliacion = Ampliacion & " Devuelto = 1, situacion = 0   "
            ImporteNue = RS!Total - RS!imponible '- Rs!impiva
            
            ImporteNue = DBLet(RsCobro!Gastos, "N")
            If DBLet(RS!ImpIva, "N") > 0 Then
            
                If ImporteNue = 0 Then
                    Ampliacion = Ampliacion & " , Gastos = " & TransformaComasPuntos(CStr(RS!ImpIva))
                Else
                    Ampliacion = Ampliacion & " , Gastos = Gastos + " & TransformaComasPuntos(CStr(RS!ImpIva))
                End If
            End If
            Ampliacion = Ampliacion & " ,impcobro=NULL,codrem= NULL, anyorem = NULL, siturem = NULL,tiporem=NULL,fecultco=NULL,recedocu=0"
            Ampliacion = Ampliacion & " WHERE numserie='" & RS!NUmSerie & "' AND numfactu=" & RS!NumFac
            Ampliacion = Ampliacion & " AND fecfactu='" & Format(RS!Fecha, FormatoFecha) & "' AND numorden=" & RS!NIF
            
            If Not Ejecuta(Ampliacion) Then GoTo ECon
             
         End If
         Set RsCobro = Nothing

        '%%%%% hasta aqui

        If Not Ejecuta(SQL) Then GoTo ECon

        Linea = Linea + 1
        
 
        'Lineas de apuntes del GASTO del vto en curso
        SQL = LINAPU & Linea & ",'"


        SQL = SQL & CtaBancoGastos & "','" & Format(RS!NumFac, "000000000") & "'," & vCP.condecli
        SQL = SQL & ",'Gastos vto.'"
        
        
        'Importe al debe
        SQL = SQL & "," & TransformaComasPuntos(CStr(GastoVto)) & ",NULL,"
        
        If CCBanco <> "" Then
            SQL = SQL & "'" & DevNombreSQL(CCBanco) & "'"
        Else
            SQL = SQL & "NULL"
        End If
            
        'Contra partida
        'Si no lleva cuenta puente contabiliza los gastos
        Aux = "NULL"
       
        SQL = SQL & "," & Aux & ",'COBROS',0,"
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",1)"
        If GastoVto <> 0 Then
            If Not Ejecuta(SQL) Then Exit Function
        
            Linea = Linea + 1
        End If

        
        'Si tiene cuenta puente cancelo la puente tb
        If CuentaPuente Then
                
            'Si lleva cta efectos comerciales descontados, tipo fontenas, NO HACE este contrapunte
            If CtaEfectosComDescontados = "" Then
                'Lineas de apuntes .
                 SQL = LINAPU & Linea & ",'"
              
                 If Len(SubCtaPte) = vEmpresa.DigitosUltimoNivel Then
                     SQL = SQL & SubCtaPte
                 Else
                     SQL = SQL & SubCtaPte & Mid(RS!Cta, Len(SubCtaPte) + 1)
                 End If
                 SQL = SQL & "','" & Format(RS!NumFac, "0000000") & "'," & vCP.conhacli
    
                
                Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli) & " "
            
                If vCP.amphacli = 3 Then
                    'NUEVA forma de ampliacion
                    'No hacemos nada pq amp11 ya lleva lo solicitado
                    
                Else
                    If vCP.amphacli = 4 Then
                        'COntrapartida
                        Ampliacion = Ampliacion & DevuelveDesdeBD("nommacta", "cuentas", "codmacta", RS!Cta, "T")
                        
                    Else
                        If vCP.amphacli = 2 Then
                           Ampliacion = Ampliacion & Format(RS!Fecha, "dd/mm/yyyy")
                        Else
                           If vCP.amphacli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
                           'Ampliacion = Ampliacion & RS!NUmSerie & "/" & RS!codfaccl
                           Ampliacion = Ampliacion & RS!iva & "/" & RS!NumFac
                           
                        End If
                    End If
                End If
                SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',NULL,"
        
                SQL = SQL & TransformaComasPuntos(RS!imponible) & ",NULL,"
        
                If vCP.ctrhacli = 1 Then
                    SQL = SQL & "'" & RS!Cta & "'"
                Else
                    SQL = SQL & "NULL"
                End If
                SQL = SQL & ",'COBROS',0,"
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",1)"
            
                            
                If Not Ejecuta(SQL) Then GoTo ECon
                Linea = Linea + 1
            End If 'de eefctosdescontados=""
        End If 'de ctapte
            
        RS.MoveNext
    Wend
    
    
    RS.MoveFirst









    If CuentaPuente Then
        SubCtaPte = RS!Cliente
        SubCtaPte = DevuelveDesdeBD("ctaefectosdesc", "bancos", "codmacta", SubCtaPte, "T")
        If SubCtaPte = "" Then
            MsgBox "Cuenta efectos descontados erronea. Revisar apunte " & Mc.Contador, vbExclamation
            SubCtaPte = RS!Cliente
        End If
    End If
    
    'La linea del banco
    '*********************************************************************
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
    SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","

    'Nuevo tipo ampliacion
    If vCP.amphacli = 3 Then
        Ampliacion = DescRemesa
    Else
        Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
    End If
    
    Ampliacion = Ampliacion & " Dev.rem:" & Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
    
    Amp11 = RS!Cliente  'cta banco

    'Lleva gasto pero lo descontamos de aqui
    If GastoDevolucion > 0 And DescuentaImporteDevolucion And ContabilizoGastoBanco Then
        Importe = Importe + GastoDevolucion
        'Para que la linea salga al fina
        Linea = Linea + 2
    End If
    Ampliacion = Linea & ",'" & Amp11 & "','RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2) & "'," & vCP.conhacli & ",'" & Ampliacion & "',"
    Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(Importe)) & ",NULL,"
    If CuentaPuente Then
        Ampliacion = Ampliacion & "'" & SubCtaPte & "'"
    Else
        'Nulo
        Ampliacion = Ampliacion & "NULL"
    End If
    Ampliacion = Ampliacion & ",'COBROS',0)"
    Ampliacion = SQL & Ampliacion
    If Not Ejecuta(Ampliacion) Then GoTo ECon
    If GastoDevolucion > 0 And DescuentaImporteDevolucion And ContabilizoGastoBanco Then
        Linea = Linea - 2
        'Dejo el importe como estaba
        Importe = Importe - GastoDevolucion
    Else
        Linea = Linea + 1
    End If
    If CuentaPuente Then
        'EL ANTERIOR contrapuenteado
        SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
        SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
        SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
        SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","
    
        'Nuevo tipo ampliacion
        If vCP.ampdecli = 3 Then
            Ampliacion = DescRemesa
        Else
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
        End If
        
        Ampliacion = Ampliacion & " Dev.rem:" & Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
        
        
        Amp11 = SubCtaPte  'cta efectos dtos
        
        Ampliacion = Linea & ",'" & Amp11 & "','RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2) & "'," & vCP.condecli & ",'" & Ampliacion & "',"
        Ampliacion = Ampliacion & TransformaComasPuntos(CStr(Importe)) & ",NULL,NULL,"
        'Cta efectos descontados
        Ampliacion = Ampliacion & "'" & RS!Cliente & "'"

        Ampliacion = Ampliacion & ",'COBROS',0)"
        Ampliacion = SQL & Ampliacion
        If Not Ejecuta(Ampliacion) Then GoTo ECon
        Linea = Linea + 1
  
    End If
    
    
    'SI hay que contabilizar los gastos de devolucion
    If ContabilizoGastoBanco Then
        
             If GastoDevolucion > 0 And DescuentaImporteDevolucion And ContabilizoGastoBanco Then
             
             Else
                Linea = Linea + 1
             End If
             SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
             SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
             SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
             SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"
    
             'Cuenta
             SQL = SQL & CtaBenBancarios & "','RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2) & "'," & vCP.condecli
             'SQL = SQL & Rs!Cta & "','REM" & Format(Rs!numfac, "000000000") & "'," & vCP.condecli
            
    
            If vCP.ampdecli = 3 Then
                Ampliacion = DescRemesa
            Else
                Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
                Ampliacion = Ampliacion & " Gasto remesa:" & Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
            End If
            SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
    
    
            ' timporteH, codccost, ctacontr, idcontab, punteada
            'Importe.  Va al debe
            SQL = SQL & TransformaComasPuntos(CStr(GastoDevolucion)) & ",NULL,"
            
            'Centro de coste.
            '--------------------------
            Amp11 = "NULL"
            If vParam.autocoste Then
                Amp11 = DevuelveDesdeBD("codccost", "bancos", "codmacta", RS!Cliente, "T")
                Amp11 = "'" & Amp11 & "'"
            End If
            SQL = SQL & Amp11 & ","
    
            
            SQL = SQL & "'" & RS!Cliente & "',"
            SQL = SQL & "'COBROS',0)"
    
            If Not Ejecuta(SQL) Then GoTo ECon
    
            
            
  
            'Si no agrupa dto importe
            If Not DescuentaImporteDevolucion Then
                Linea = Linea + 1
                SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
                SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
                SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
                SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaDevolucion, FormatoFecha) & "'," & Mc.Contador & ","
            
                
                If vCP.amphacli = 3 Then
                    Ampliacion = DescRemesa
                Else
                    Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
                    Ampliacion = Ampliacion & " Gasto remesa:" & Format(RecuperaValor(Remesa, 1), "0000") & "/" & RecuperaValor(Remesa, 2)
                End If
                
                Ampliacion = Linea & ",'" & RS!Cliente & "','RE" & Format(RecuperaValor(Remesa, 1), "0000") & RecuperaValor(Remesa, 2) & "'," & vCP.conhacli & ",'" & Ampliacion & "',"
                'Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(Importe)) & ",NULL,'" & CtaBenBancarios & "','CONTAB',0)"
                Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(GastoDevolucion)) & ",NULL,'" & CtaBenBancarios & "','COBROS',0)"
                Ampliacion = SQL & Ampliacion
                If Not Ejecuta(Ampliacion) Then GoTo ECon
                
            End If
      
    
    End If

    'Ya tenemos generado el apunte de devolucion
    'Insertamos para su actualziacion
    InsertaTmpActualizar Mc.Contador, vCP.diaricli, FechaDevolucion
    
    
    
    'Cerramos RS
    RS.Close
    Set miRsAux = Nothing
    
    RealizarDevolucionRemesaTalPag = True
ECon:
    If Err.Number <> 0 Then
        
        Amp11 = "Devolución remesa: " & Remesa & vbCrLf
        If Not Mc Is Nothing Then Amp11 = Amp11 & "MC.cont: " & Mc.Contador & vbCrLf
        Amp11 = Amp11 & Err.Description
        MuestraError Err.Number, Amp11
        
    End If
    Set RS = Nothing
    Set Mc = Nothing
    Set vCP = Nothing
End Function








'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'
'           Contabilizar cierre caja        A N T I G U O
'
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
Public Function ContabilizarCierreCaja(FechaCierre As Date, Caja As String, ByRef CtaPendAplicar1 As String) As Boolean
Dim Mc As Contadores
Dim Linea As Integer
Dim Debe As Currency
Dim Haber As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim RS As ADODB.Recordset
Dim Fechas As String    'Meteremos todas las fechas a cerrar
Dim VaAlDebe As Boolean

    On Error GoTo ECon
    ContabilizarCierreCaja = False
    
    'Voy a ver todas las fechas a cerrar
    Set RS = New ADODB.Recordset
    SQL = "Select feccaja from scaja where ctacaja='" & Caja & "' and feccaja <='" & Format(FechaCierre, FormatoFecha) & "' group by feccaja"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        'Esto no deberia haber pasado.
        MsgBox "Error en el cierre. Datos no encontrados", vbCritical
        RS.Close
        GoTo ECon
    End If
    Fechas = ""
    While Not RS.EOF
        Fechas = Fechas & RS.Fields(0) & "|"
        RS.MoveNext
    Wend
    RS.Close



    'Borro de hco lo anterior a dos meses
    FechaCierre = DateAdd("m", -2, FechaCierre)
    SQL = "DELETE from shcaja where feccaja<'" & Format(FechaCierre, FormatoFecha) & "'"
    Conn.Execute SQL

    'Borro tmpactualizar
    SQL = "DELETE from tmpactualizar where codusu = " & vUsu.Codigo
    Conn.Execute SQL

    'La forma de pago
    Set vCP = New Ctipoformapago
    If vCP.Leer(vbEfectivo) = 1 Then GoTo ECon
    
    
    Set Mc = New Contadores
    
    
    'Para cada fecha..
    While Fechas <> ""
    
        Linea = InStr(1, Fechas, "|")
        
        FechaCierre = CDate(Mid(Fechas, 1, Linea - 1))
        Fechas = Mid(Fechas, Linea + 1)   'Para la siguiente
        
        
        
        If Mc.ConseguirContador("0", FechaCierre <= vParam.fechafin, True) = 1 Then Exit Function
        
        
        
        'Insertamos la cabera
        SQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) VALUES ("
        SQL = SQL & vCP.diaricli & ",'" & Format(FechaCierre, FormatoFecha) & "'," & Mc.Contador
        SQL = SQL & ", 1, NULL, '"
        SQL = SQL & "Cierre: " & FechaCierre & " - " & Caja & vbCrLf
        SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy") & " por " & vUsu.Nombre & "');"
        If Not Ejecuta(SQL) Then Exit Function
        
        
        
        
        Linea = 1
        Debe = 0
        Haber = 0
        
        

        SQL = "Select * from scaja where feccaja='" & Format(FechaCierre, FormatoFecha) & "' AND ctacaja = '" & Caja & "'"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            
        
        
            'Lineas de apuntes .
             SQL = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, "
             SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
             SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
             SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaCierre, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"
             
            'Cuenta
            Select Case RS!NUmSerie
            Case "0"
               'Suministro. codmacta la coje de parametros
                'SQL = SQL & Rs!codmacta & "','SUMINIS',"
                SQL = SQL & CtaPendAplicar1 & "','SUMINIS',"
                VaAlDebe = False
                
            Case "1"
                'TRASPASO ENTRE CAJAS
                SQL = SQL & RS!codmacta & "','TRAS_CAJA',"
                If RS!ImpEfect < 0 Then
                    VaAlDebe = True
                Else
                    VaAlDebe = False
                End If
            Case "2"
                'FACTURAS PROVEEDORES
                SQL = SQL & RS!codmacta & "','" & RS!NumFactu & "',"
                If RS!ImpEfect < 0 Then
                    VaAlDebe = True
                Else
                    VaAlDebe = False
                End If
            Case Else
                'FACTURAS CLIENTES
                SQL = SQL & RS!codmacta & "','" & Format(RS!NumFactu, "000000000") & "',"
                If RS!ImpEfect < 0 Then
                    VaAlDebe = True
                Else
                    VaAlDebe = False
                End If
                
            End Select
            SQL = SQL & vCP.condecli
               
                    
                    
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
            Ampliacion = Ampliacion & " "
                   
            If vCP.amphacli = 2 Then
               Ampliacion = Ampliacion & Format(FechaCierre, "dd/mm/yyyy")
            Else
               If vCP.amphacli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
               If RS!NUmSerie = "0" Or RS!NUmSerie = "1" Then
                    Ampliacion = Ampliacion & RS!Ampliacion
               Else
                    If RS!NUmSerie = "2" Then
                        Ampliacion = Ampliacion & RS!NumFactu
                    Else
                        Ampliacion = Ampliacion & RS!NUmSerie & "/" & RS!NumFactu
                    End If
                End If
            End If
            SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
            
            
            If VaAlDebe Then
                Debe = Debe + Abs(RS!ImpEfect)
            Else
                Haber = Haber + Abs(RS!ImpEfect)
            End If
            
            ' timporteH, codccost, ctacontr, idcontab, punteada
            If VaAlDebe Then
                SQL = SQL & TransformaComasPuntos(Abs(RS!ImpEfect)) & ",NULL,NULL,"       'y el CC tambien
            Else
                SQL = SQL & "NULL," & TransformaComasPuntos(Abs(RS!ImpEfect)) & ",NULL,"
            End If
            
            
            If vCP.ctrdecli = 1 Then
                SQL = SQL & "'" & Caja & "',"
            Else
                SQL = SQL & "NULL,"
            End If
            SQL = SQL & "'CONTAB',0)"
            
            If Not Ejecuta(SQL) Then Exit Function
            
            Linea = Linea + 1
            RS.MoveNext
        Wend
        RS.Close
    
       
        
        'La linea del banco
        SQL = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, "
        SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
        SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
        SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaCierre, FormatoFecha) & "'," & Mc.Contador & ","
        
        

        
        'TOTAL
        Linea = Linea + 1
        Debe = -1 * (Debe - Haber)
        If Debe < 0 Then
            VaAlDebe = False
        Else
            VaAlDebe = True
        End If
        Debe = Abs(Debe)
        If Debe <> 0 Then
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
            Ampliacion = Ampliacion & " CIERRE  " & FechaCierre
            Ampliacion = Linea & ",'" & Caja & "',''," & vCP.condecli & ",'" & Ampliacion & "',"
            If VaAlDebe Then
                Ampliacion = Ampliacion & TransformaComasPuntos(CStr(Debe)) & ",NULL,"
            Else
                Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(Debe)) & ","
            End If
            Ampliacion = Ampliacion & "NULL,NULL,'CONTAB',0)"
            Ampliacion = SQL & Ampliacion
            If Not Ejecuta(Ampliacion) Then Exit Function
        End If
    
        'Insertamos para pasar a hco
        InsertaTmpActualizar Mc.Contador, vCP.diaricli, FechaCierre
        
        
        
        
        'Insertamos en hco de caja
        SQL = "insert into shcaja Select * from scaja where feccaja='" & Format(FechaCierre, FormatoFecha) & "' AND ctacaja = '" & Caja & "'"
        Conn.Execute SQL
        
        'Borramos
        SQL = "DELETE from scaja where feccaja = '" & Format(FechaCierre, FormatoFecha) & "' AND ctacaja = '" & Caja & "'"
        Conn.Execute SQL
    Wend
    
    
    'Cojemos y desbloqueamos los apuntes
    SQL = "Select * from tmpactualizar where codusu = " & vUsu.Codigo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        SQL = "UPDATE cabapu set bloqactu=0 where numasien=" & RS!NumAsien
        SQL = SQL & " AND numdiari=" & RS!NumDiari & " AND fechaent ='" & Format(RS!FechaEnt, FormatoFecha) & "'"
        Conn.Execute SQL
    
        RS.MoveNext
    Wend
    RS.Close
    
    'Borro tmpactualizar
    SQL = "DELETE from tmpactualizar where codusu = " & vUsu.Codigo
    Conn.Execute SQL

    
    
    'Todo OK
    ContabilizarCierreCaja = True
    
    
ECon:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    
    End If
    Set RS = Nothing
    Set Mc = Nothing
    Set vCP = Nothing

End Function










'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'
'
'           NUEVO CIERRE CAJA
'
'
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
Public Function ContabilizarCierreCajaNuevo(FechaCierre As Date, Caja As String, ByRef CtaPendAplicar1 As String, Diario As Integer) As Boolean
Dim Mc As Contadores
Dim Linea As Integer
Dim Debe As Currency
Dim Haber As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim RS As ADODB.Recordset
Dim Fechas As String    'Meteremos todas las fechas a cerrar
Dim TextoFactura As String

    On Error GoTo ECon
    ContabilizarCierreCajaNuevo = False
    
    'Voy a ver todas las fechas a cerrar
    Set RS = New ADODB.Recordset
    SQL = "Select feccaja from slicaja where codusu=" & (vUsu.Codigo Mod 100) & " and feccaja <='" & Format(FechaCierre, FormatoFecha) & "' group by feccaja"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        'Esto no deberia haber pasado.
        MsgBox "Error en el cierre. Datos no encontrados", vbCritical
        RS.Close
        GoTo ECon
    End If
    Fechas = ""
    While Not RS.EOF
        Fechas = Fechas & RS.Fields(0) & "|"
        RS.MoveNext
    Wend
    RS.Close



    'Borro de hco lo anterior a dos meses
    FechaCierre = DateAdd("m", -2, FechaCierre)
    SQL = "DELETE from shcaja where feccaja<'" & Format(FechaCierre, FormatoFecha) & "'"
    Conn.Execute SQL

    'Borro tmpactualizar
    SQL = "DELETE from tmpactualizar where codusu = " & vUsu.Codigo
    Conn.Execute SQL

    'La forma de pago
    Set vCP = New Ctipoformapago
    If vCP.Leer(vbEfectivo) = 1 Then GoTo ECon
    
    
    Set Mc = New Contadores
    
    
    'Para cada fecha..
    While Fechas <> ""
    
        Linea = InStr(1, Fechas, "|")
        
        FechaCierre = CDate(Mid(Fechas, 1, Linea - 1))
        Fechas = Mid(Fechas, Linea + 1)   'Para la siguiente
        
        
        
        If Mc.ConseguirContador("0", FechaCierre <= vParam.fechafin, True) = 1 Then Exit Function
        
        
        
        'Insertamos la cabera
        SQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) VALUES ("
        SQL = SQL & Diario & ",'" & Format(FechaCierre, FormatoFecha) & "'," & Mc.Contador
        SQL = SQL & ", 1, NULL, '"
        SQL = SQL & "Cierre: " & FechaCierre & " - Usuario caja:" & Caja & vbCrLf
        SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy") & " por " & vUsu.Nombre & "');"
        If Not Ejecuta(SQL) Then Exit Function
        
        
        
        
        Linea = 1
        Debe = 0
        Haber = 0
        
        

        SQL = "Select * from slicaja where feccaja='" & Format(FechaCierre, FormatoFecha) & "' AND codusu = '" & (vUsu.Codigo Mod 100) & "'"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            
        
        
            'Lineas de apuntes .
             SQL = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, "
             SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
             SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
             SQL = SQL & "VALUES (" & Diario & ",'" & Format(FechaCierre, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"
             
            'Cuenta
            TextoFactura = ""
            Select Case RS!tipomovi
            Case 2
               'Suministro. codmacta la coje de parametros
                'SQL = SQL & Rs!codmacta & "','SUMINIS',"
                SQL = SQL & CtaPendAplicar1 & "','SUMINIS',"
       
                
            Case 3
                'TRASPASO ENTRE CAJAS
                SQL = SQL & RS!codmacta & "','TRAS_CAJA',"

            Case 1
                'FACTURAS PROVEEDORES
                If Not IsNull(RS!NumFacpr) Then
                    TextoFactura = DevNombreSQL(RS!NumFacpr)
                Else
                    TextoFactura = ""
                End If
                
                SQL = SQL & RS!codmacta & "','" & TextoFactura & "',"

            Case Else
                'FACTURAS CLIENTES
                
                If Not IsNull(RS!numfaccl) Then
                    'TextoFactura = DBLet(RS!NUmSerie, "T") & Format(RS!numfaccl, "000000000")
                    TextoFactura = SerieNumeroFactura(10, RS!NUmSerie, RS!numfaccl)
                    
                Else
                    TextoFactura = ""
                End If
                SQL = SQL & RS!codmacta & "','" & TextoFactura & "',"
                
                
                
                
            End Select
            SQL = SQL & vCP.condecli
               
                    
                    
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
            Ampliacion = Ampliacion & " "
                   
            If vCP.amphacli = 2 Then
               Ampliacion = Ampliacion & Format(FechaCierre, "dd/mm/yyyy")
            Else
               If vCP.amphacli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
               If RS!tipomovi > 1 Then
                'ES UN TRASPASO o SUMINISTRO
                    'No hago nada
                    'Ampliacion = Ampliacion & ""
               Else
                    Ampliacion = Ampliacion & TextoFactura
               End If
            End If
            SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
            
            
            If IsNull(RS!ImporteH) Then
                Debe = Debe + RS!ImporteD
            Else
                Haber = Haber + RS!ImporteH
            End If
            
            ' timporteH, codccost, ctacontr, idcontab, punteada
            If IsNull(RS!ImporteH) Then
                SQL = SQL & TransformaComasPuntos(RS!ImporteD) & ",NULL,NULL,"       'y el CC tambien
            Else
                SQL = SQL & "NULL," & TransformaComasPuntos(RS!ImporteH) & ",NULL,"
            End If
            
            
            If vCP.ctrdecli = 1 Then
                SQL = SQL & "'" & Caja & "',"
            Else
                SQL = SQL & "NULL,"
            End If
            SQL = SQL & "'CONTAB',0)"
            
            If Not Ejecuta(SQL) Then Exit Function
            
            Linea = Linea + 1
            RS.MoveNext
        Wend
        RS.Close
    
       
        
        'La linea del banco
        SQL = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, "
        SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
        SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
        SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FechaCierre, FormatoFecha) & "'," & Mc.Contador & ","
        
        

        
        'TOTAL
        Linea = Linea + 1
        Debe = -1 * (Debe - Haber)
        If Debe <> 0 Then
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
            Ampliacion = Ampliacion & " CIERRE  " & FechaCierre
            Ampliacion = Linea & ",'" & Caja & "',''," & vCP.condecli & ",'" & Ampliacion & "',"
            If Debe > 0 Then
                Ampliacion = Ampliacion & TransformaComasPuntos(CStr(Debe)) & ",NULL,"
            Else
                Ampliacion = Ampliacion & "NULL," & TransformaComasPuntos(CStr(Abs(Debe))) & ","
            End If
            Ampliacion = Ampliacion & "NULL,NULL,'CONTAB',0)"
            Ampliacion = SQL & Ampliacion
            If Not Ejecuta(Ampliacion) Then Exit Function
            
            
            'Actualizo el saldo de caja del usuario
            Haber = 0
            Ampliacion = "Select saldo from susucaja where codusu = " & CStr(vUsu.Codigo Mod 100)
            RS.Open Ampliacion, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then Haber = DBLet(RS!Saldo, "N")
            RS.Close
            
            Haber = Haber + Debe 'Saldo arrastrado
            Ampliacion = "UPDATE susucaja SET saldo =" & TransformaComasPuntos(CStr(Haber))
            Ampliacion = Ampliacion & " WHERE  codusu = " & CStr(vUsu.Codigo Mod 100)
            Conn.Execute Ampliacion
        End If
    
    
    
    
    
        'Para ACtualizar.        NO ACTUALZIAMOS
        'InsertaTmpActualizar Mc.Contador, vCP.diaricli, FechaCierre
        
        

        
        'Borramos de la cabecera y las lineas
        'Y actualizamos el saldo
        SQL = "DELETE FROM slicaja where feccaja='" & Format(FechaCierre, FormatoFecha) & "' AND codusu = " & (vUsu.Codigo Mod 100)
        Conn.Execute SQL
        
        SQL = "DELETE FROM scacaja where feccaja='" & Format(FechaCierre, FormatoFecha) & "' AND codusu = " & (vUsu.Codigo Mod 100)
        Conn.Execute SQL
        
        
        
    Wend
    
    
    'Cojemos y desbloqueamos los apuntes
    SQL = "Select * from tmpactualizar where codusu = " & vUsu.Codigo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        SQL = "UPDATE cabapu set bloqactu=0 where numasien=" & RS!NumAsien
        SQL = SQL & " AND numdiari=" & RS!NumDiari & " AND fechaent ='" & Format(RS!FechaEnt, FormatoFecha) & "'"
        Conn.Execute SQL
    
        RS.MoveNext
    Wend
    RS.Close
    
    
    'NO ACTUALIZMOS DEJAMOS EL ASTO EN LA INTRODUCCION
    'Ahora actualizamos los registros que estan en tmpactualziar
    'frmActualizar2.OpcionActualizar = 20
    'frmActualizar2.Show vbModal
    
    
        
    'Borro tmpactualizar
    SQL = "DELETE from tmpactualizar where codusu = " & vUsu.Codigo
    Conn.Execute SQL

    
    
    'Todo OK
    ContabilizarCierreCajaNuevo = True
    
    
ECon:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    
    End If
    Set RS = Nothing
    Set Mc = Nothing
    Set vCP = Nothing

End Function












'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'
'   COMPENSACIONES.
'       Contabilizara las compensaciones. Es decir. Desde el FORM de las compensaciones
'       le mandara el conjunto de cobros, el de pagos
'       cta bancaria
'
'       Y generara un UNICO apunte eliminando todos los cobros y pagos seleccionados
'       excepto si la compensacion se efectua sobre un determinado VTO
'       que sera updateado
'       Si AumentaElImporteDelVto significa que el vto aumenta ;)
Public Function ContabilizarCompensaciones(ByRef ColCobros As Collection, ByRef ColPagos As Collection, ByVal DatosAdicionales As String, AumentaElImporteDelVto As Boolean) As Boolean
Dim SQL As String
Dim Mc As Contadores
Dim CadenaSQL As String
Dim FechaContab As Date
Dim I As Integer
Dim Obs As String

Dim SqlNue As String
Dim RsNue As ADODB.Recordset


    On Error GoTo EEContabilizarCompensaciones

    ContabilizarCompensaciones = False
    
    
    'Fecha contabilizacion
    FechaContab = RecuperaValor(DatosAdicionales, 4)
    
    'Borro tmpactualizar
    SQL = "DELETE from tmpactualizar where codusu = " & vUsu.Codigo
    Conn.Execute SQL


    Conn.BeginTrans    'TRANSACCION
    Set Mc = New Contadores
    If Mc.ConseguirContador("0", FechaContab <= vParam.fechafin, True) = 1 Then GoTo EEContabilizarCompensaciones
        
        
    'Insertamos la cabera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion, usucreacion, desdeaplicacion) VALUES ("
    SQL = SQL & CInt(RecuperaValor(DatosAdicionales, 3)) & ",'" & Format(FechaContab, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ", '"
    SQL = SQL & "Compensa: " & DevNombreSQL(RecuperaValor(DatosAdicionales, 7)) & vbCrLf
    If AumentaElImporteDelVto Then SQL = SQL & "Aumento VTO" & vbCrLf
    SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy hh:nn") & " por " & vUsu.Nombre & "',"
    
    Obs = "ARICONTA 6: Compensa: " & RecuperaValor(DatosAdicionales, 7)
    SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(Obs, "T") & ");"
    Conn.Execute SQL
    
    
    
    'Insertamos para pasar a hco
    InsertaTmpActualizar Mc.Contador, RecuperaValor(DatosAdicionales, 3), FechaContab
    
    

    'Añadimos las facturas de clientes
    'Lineas de apuntes .
    CadenaSQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    CadenaSQL = CadenaSQL & "codconce, numdocum, ampconce , "
    'Toda esta linea viene juntita
    CadenaSQL = CadenaSQL & "codmacta, timporteD,timporteH,"
    'Numdocum viene con otro valor
    CadenaSQL = CadenaSQL & " ctacontr, codccost, idcontab, punteada, "
    CadenaSQL = CadenaSQL & " numserie, numfaccl, numfacpr, fecfactu, numorden, tipforpa) "
    CadenaSQL = CadenaSQL & "VALUES (" & RecuperaValor(DatosAdicionales, 3) & ",'" & Format(FechaContab, FormatoFecha) & "'," & Mc.Contador & ","
    

    NumRegElim = 1
    'Los cobros
    For I = 1 To ColCobros.Count
        
        SQL = NumRegElim & "," & RecuperaValor(ColCobros.Item(I), 1) & "NULL,'COBROS',0,"
        
        'parte donde indicamos en el apunte que se ha cobrado
        SqlNue = "select * from cobros " & RecuperaValor(ColCobros.Item(I), 3)
        Set RsNue = New ADODB.Recordset
        RsNue.Open SqlNue, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RsNue.EOF Then
            SQL = SQL & DBSet(RsNue!NUmSerie, "T") & ","
            SQL = SQL & DBSet(RsNue!NumFactu, "N") & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & DBSet(RsNue!FecFactu, "F") & ","
            SQL = SQL & DBSet(RsNue!numorden, "N") & ","
            SQL = SQL & DevuelveValor("select tipforpa from formapago where codforpa = " & DBSet(RsNue!codforpa, "N")) & ")"
        Else
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ")"
        End If
        Set RsNue = Nothing
        
        Conn.Execute CadenaSQL & SQL
        
        
        NumRegElim = NumRegElim + 1
        'Borro el cobro pago
        SQL = RecuperaValor(ColCobros.Item(I), 2)
        If Mid(SQL, 1, 6) = "UPDATE" Then
            'UPDATEAMOS
            Conn.Execute SQL
        Else
            ' ATENCION !!!!!
            ' ya no borramos hemos de darlo como cobrado
'            Conn.Execute "DELETE FROM cobros " & Sql
'
'            'Borramos de efectos devueltos... por si acaso sefecdev
'            Ejecuta "DELETE FROM sefecdev " & Sql
            SqlNue = "update cobros set fecultco = " & DBSet(FechaContab, "F") & ", impcobro = coalesce(impcobro,0) + impvenci + coalesce(gastos,0), situacion = 1 "
            SqlNue = SqlNue & SQL

            Ejecuta SqlNue
        End If
    Next I


    
    'Los pagos
    For I = 1 To ColPagos.Count
        SQL = NumRegElim & "," & RecuperaValor(ColPagos.Item(I), 1) & "NULL,'PAGOS',0,"
        
        'parte donde indicamos en el apunte que se ha pagado
        SqlNue = "select * from pagos " & RecuperaValor(ColPagos.Item(I), 3)
        Set RsNue = New ADODB.Recordset
        RsNue.Open SqlNue, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RsNue.EOF Then
            SQL = SQL & DBSet(RsNue!NUmSerie, "T") & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & DBSet(RsNue!NumFactu, "T") & ","
            SQL = SQL & DBSet(RsNue!FecFactu, "F") & ","
            SQL = SQL & DBSet(RsNue!numorden, "N") & ","
            SQL = SQL & DevuelveValor("select tipforpa from formapago where codforpa = " & DBSet(RsNue!codforpa, "N")) & ")"
        Else
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ","
            SQL = SQL & ValorNulo & ")"
        End If
        Set RsNue = Nothing
        
        
        Conn.Execute CadenaSQL & SQL
        NumRegElim = NumRegElim + 1
        'Borro el  pago   La linea del banco va aqui dentro, con lo cual
        'Si tengo que comprobar si es la linea del banco o no para borrar
        SQL = RecuperaValor(ColPagos.Item(I), 2)
        If SQL <> "" Then
            If Mid(SQL, 1, 6) = "UPDATE" Then
                'UPDATEAMOS
                Conn.Execute SQL
            Else
                ' ATENCION !!!!!
                ' ya no borramos hemos de darlo como pagado
'                Conn.Execute "DELETE FROM pagos " & Sql
            
                SqlNue = "update pagos set fecultpa = " & DBSet(FechaContab, "F") & ",imppagad = coalesce(imppagad,0) + impefect, situacion = 1 "
                SqlNue = SqlNue & SQL
    
                Ejecuta SqlNue
            End If

        End If
    Next I

    Conn.CommitTrans   'TODO HA IDO BIEN
    

'    'Actualizacion
'    'Ahora actualizamos los registros que estan en tmpactualziar
'    frmTESActualizar.OpcionActualizar = 20
'    frmTESActualizar.Show vbModal
        
    'Borro tmpactualizar
    SQL = "DELETE from tmpactualizar where codusu = " & vUsu.Codigo
    Ejecuta SQL
    
    'Marco para indicar que TODO ha ido de P.M.
    CadenaDesdeOtroForm = ""
    Exit Function
EEContabilizarCompensaciones:
    If Err.Number <> 0 Then MuestraError Err.Number
    Conn.RollbackTrans
    
End Function





'----------------------------------------------------------------------------------------------------
' NORMA 32,58, Febrero 2009: TOoooodas las remesas
' ================================================
'
'
'Mod Nov 2009
Public Function RemesasCancelacionEfectos(Codigo As Integer, Anyo As Integer, CtaBanco As String, FechaAbono As Date, GastosBancarios As Currency, AgrupaCancelacion As Boolean) As Boolean
'Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
Dim Gastos As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim RS As ADODB.Recordset
Dim AmpRemesa As String
Dim CtaCancelacion As String
Dim Cuenta As String
Dim RaizCuentasCancelacionConfirmacion As String
Dim LCta As Integer
Dim ImporteTotal As Currency
Dim ImpDelVto As Currency
Dim CtaBancoIngresos As String

    On Error GoTo ERemesa_CancelacionCliente
    RemesasCancelacionEfectos = False
    
    Cuenta = "RemesaConfirmacion"
    AmpRemesa = DevuelveDesdeBD("RemesaCancelacion", "paramtesor", "codigo", "1", "N", Cuenta)
    RaizCuentasCancelacionConfirmacion = RaizCuentasCancelacionConfirmacion & AmpRemesa & "|" & Cuenta & "|"
    
    CtaBancoIngresos = RecuperaValor(CtaBanco, 4)
    'Comprobacion.  Para todos los efectos de la 43.... se cancelan con la 4310....
    '
    'Tendre que ver que existen estas cuentas
    Set RS = New ADODB.Recordset
    
    SQL = "DELETE FROM tmpcierre1 where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
        
 
    
        

        AmpRemesa = RecuperaValor(RaizCuentasCancelacionConfirmacion, 1)

        LCta = Len(AmpRemesa)
            
            
        If LCta <> vEmpresa.DigitosUltimoNivel Then '//Para cuenta puente raiz
            Ampliacion = ",CONCAT('" & AmpRemesa & "',SUBSTRING(codmacta," & LCta + 1 & ")" & ")"
            
            SQL = "Select " & vUsu.Codigo & Ampliacion & " from scobro where codrem=" & Codigo & " AND anyorem = " & Anyo
            SQL = SQL & " GROUP BY codmacta"
            'INSERT
            SQL = "INSERT INTO tmpcierre1(codusu,cta) " & SQL
            Conn.Execute SQL
            
            'Ahora monto el select para ver que cuentas 430 no tienen la 4310
            SQL = "Select cta,codmacta from tmpcierre1 left join cuentas on tmpcierre1.cta=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
            SQL = SQL & " HAVING codmacta is null"
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            Linea = 0
            While Not RS.EOF
                Linea = Linea + 1
                SQL = SQL & RS!Cta & "     "
                If Linea = 6 Then
                    SQL = SQL & vbCrLf
                    Linea = 0
                End If
                RS.MoveNext
            Wend
            RS.Close
            
            If SQL <> "" Then
               
                AmpRemesa = "CANCELACION remesa"
   
                SQL = "Cuentas " & AmpRemesa & ".  No existen las cuentas: " & vbCrLf & String(90, "-") & vbCrLf & SQL
                SQL = SQL & vbCrLf & "¿Desea crearlas?"
                
                'FEBRERO 09
                Linea = 1
                If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                    'Ha dicho que si desea crearlas
                    
                    
 
                    AmpRemesa = RecuperaValor(RaizCuentasCancelacionConfirmacion, 1)
 
                    LCta = Len(AmpRemesa)
                    Ampliacion = "CONCAT('" & AmpRemesa & "',SUBSTRING(codmacta," & LCta + 1 & ")) "
                
                    SQL = "Select distinct(codmacta)," & Ampliacion & " from scobro where codrem=" & Codigo & " AND anyorem = " & Anyo
                    SQL = SQL & " and " & Ampliacion & " in "
                    SQL = SQL & "(Select cta from tmpcierre1 left join cuentas on tmpcierre1.cta=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
                    SQL = SQL & " AND codmacta is null)"
                    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    While Not RS.EOF
                    
                         SQL = "INSERT  IGNORE INTO cuentas(codmacta,nommacta ,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos) SELECT '"
                                    ' CUenta puente
                         SQL = SQL & RS.Fields(1) & "',nommacta ,'S',0,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos from cuentas where codmacta = '"
                                    'Cuenta en la scbro (codmacta)
                         SQL = SQL & RS.Fields(0) & "'"
                         Conn.Execute SQL
                         RS.MoveNext
                         
                    Wend
                    RS.Close
                    Linea = 0
                End If
                If Linea = 1 Then GoTo ERemesa_CancelacionCliente
            
            End If

        Else  'de cuenta unica
        
        
            'Para la eliminacion de efectos, la cuenta de cancelacion es
            ' ctaefectosdesc , del banco
            SQL = "select ctaefectosdesc from remesas r,ctabancaria b where r.codmacta=b.codmacta and codigo=" & Codigo & " AND anyo = " & Anyo
            CtaCancelacion = ""
            RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            If Not RS.EOF Then
                If Not IsNull(RS.Fields(0)) Then CtaCancelacion = RS.Fields(0)
            End If
            RS.Close
            If CtaCancelacion = "" Then
                MsgBox "Falta configurar la cuenta de efectos descontados para el banco de la remesa", vbExclamation
                GoTo ERemesa_CancelacionCliente
            End If
        End If 'De la comprobacion
    
    'La forma de pago
    Set vCP = New Ctipoformapago
    If vCP.Leer(vbTipoPagoRemesa) = 1 Then GoTo ERemesa_CancelacionCliente
    
    
    Set Mc = New Contadores
    
    
    If Mc.ConseguirContador("0", FechaAbono <= vParam.fechafin, True) = 1 Then Exit Function
    
    
    
    'Insertamos la cabera
    SQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) VALUES ("
    SQL = SQL & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ", 1, NULL, '"

    SQL = SQL & "Cancelacion cliente"

    SQL = SQL & " remesa: " & Codigo & " / " & Anyo & vbCrLf
    SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy") & " por " & vUsu.Nombre & "');"
    If Not Ejecuta(SQL) Then Exit Function
    
    
    
    
    Linea = 1
    Importe = 0
    Gastos = 0
    
    vCP.descformapago = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
    
    SQL = "Select * from scobro where codrem=" & Codigo & " AND anyorem = " & Anyo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Trozo comun
    AmpRemesa = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, "
    AmpRemesa = AmpRemesa & "codmacta, numdocum, codconce, ampconce,timporteD,"
    AmpRemesa = AmpRemesa & " timporteH, codccost, ctacontr, idcontab, punteada) "
    AmpRemesa = AmpRemesa & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & ","

    
    While Not RS.EOF
        
        'Ampliacion
        Ampliacion = vCP.descformapago & " "
               
        'Neuvo dato para la ampliacion en la contabilizacion
        Select Case vCP.amphacli
        Case 2, 4
            'La opcion Contrapartida BANCO NO vale ahora, pq no hay apunte a banco
            Ampliacion = Ampliacion & Format(RS!FecVenci, "dd/mm/yyyy")
            
        Case Else
           If vCP.amphacli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
           Ampliacion = Ampliacion & RS!NUmSerie & "/" & RS!codfaccl
        End Select
        

        Cuenta = RS!codmacta
        
        'Contra que cancela
        CtaCancelacion = RecuperaValor(RaizCuentasCancelacionConfirmacion, 1)
        If LCta <> vEmpresa.DigitosUltimoNivel Then CtaCancelacion = CtaCancelacion & Mid(RS!codmacta, LCta + 1)
        

    
        
         
        'Cuenta
        SQL = Linea & ",'" & Cuenta & "','" & Format(RS!codfaccl, "000000000") & "'," & vCP.conhacli
        SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
        Importe = Importe + RS!ImpVenci
        Gastos = Gastos + DBLet(RS!Gastos, "N")
        
        

        SQL = SQL & "NULL," & TransformaComasPuntos(RS!ImpVenci) & ",NULL,"
    
        'Contra partida
        SQL = SQL & "'" & CtaCancelacion & "','CONTAB',0)"
        SQL = AmpRemesa & SQL
        If Not Ejecuta(SQL) Then Exit Function
        Linea = Linea + 1
        

        'La contrapartida
        If Not AgrupaCancelacion Then
            ImpDelVto = RS!ImpVenci + DBLet(RS!Gastos, "N")
            'Si no agrupa cancelacion, los gastos VAN separados para cada cuenta
            If DBLet(RS!Gastos, "N") > 0 Then
                'Tiene gastos
            
            
                            
                SQL = Linea & ",'" & CtaBancoIngresos & "','" & Format(RS!codfaccl, "000000000") & "'," & vCP.conhacli
                SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
                
                
                'Importe al haber
                SQL = SQL & "NULL," & TransformaComasPuntos(RS!Gastos) & ","
                If RecuperaValor(CtaBanco, 3) <> "" Then
                    SQL = SQL & "'" & DevNombreSQL(RecuperaValor(CtaBanco, 3)) & "'"
                Else
                    SQL = SQL & "NULL"
                End If
                    
                'Contra partida
                SQL = SQL & ",'" & CtaCancelacion & "','CONTAB',0)"
                SQL = AmpRemesa & SQL
                If Not Ejecuta(SQL) Then Exit Function
                
                Linea = Linea + 1
            
            End If
            
            
            
            
            
            'La cancelacion
            SQL = Linea & ",'" & CtaCancelacion & "','" & Format(RS!codfaccl, "000000000") & "'," & vCP.condecli
            SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
            
            
            'Importe al DEBE
            SQL = SQL & TransformaComasPuntos(CStr(ImpDelVto)) & ",NULL,NULL,"
        
            'Contra partida
            SQL = SQL & "'" & Cuenta & "','CONTAB',0)"
            SQL = AmpRemesa & SQL
            If Not Ejecuta(SQL) Then Exit Function
            
            Linea = Linea + 1
                
        End If
        
        
        RS.MoveNext
    Wend
    RS.Close

    'Si tiene gastos
    If AgrupaCancelacion Then
    
        If Gastos > 0 Then
                'Tiene gastos
                Ampliacion = "Gastos en vtos. Rem: " & Codigo & "/" & Anyo
            
                            
                'SQL = Linea & ",'" & CtaBancoIngresos & "','" & Format(Rs!codfaccl, "000000000") & "'," & vCP.conhacli
                SQL = Linea & ",'" & CtaBancoIngresos & "','RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.conhacli
                SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
                
                
                'Importe al haber
                SQL = SQL & "NULL," & TransformaComasPuntos(CStr(Gastos)) & ","
                If RecuperaValor(CtaBanco, 3) <> "" Then
                    SQL = SQL & "'" & DevNombreSQL(RecuperaValor(CtaBanco, 3)) & "'"
                Else
                    SQL = SQL & "NULL"
                End If
                    
                'Contra partida
                SQL = SQL & ",'" & CtaCancelacion & "','CONTAB',0)"
                SQL = AmpRemesa & SQL
                If Not Ejecuta(SQL) Then Exit Function
                
                Linea = Linea + 1
            
        End If
        
        
            'La cancelacion
            Ampliacion = "Cancela remesa: " & Codigo & "/" & Anyo
            SQL = Linea & ",'" & CtaCancelacion & "','RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.condecli
            SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
            
            
            'Importe al DEBE
            ImpDelVto = Importe + Gastos
            SQL = SQL & TransformaComasPuntos(CStr(ImpDelVto)) & ",NULL,NULL,"
        
            'Contra partida
            SQL = SQL & "'" & Cuenta & "','CONTAB',0)"
            SQL = AmpRemesa & SQL
            If Not Ejecuta(SQL) Then Exit Function
            
            Linea = Linea + 1

    End If




    AmpRemesa = "F"    ' cancelada
    SQL = "UPDATE scobro SET"
    SQL = SQL & " siturem= '" & AmpRemesa
    SQL = SQL & "', ctabanc2= '" & RecuperaValor(CtaBanco, 1) & "'"
    SQL = SQL & " WHERE codrem=" & Codigo
    SQL = SQL & " and anyorem=" & Anyo
    Conn.Execute SQL

    SQL = "UPDATE remesas SET"
    SQL = SQL & " situacion= '" & AmpRemesa
    SQL = SQL & "' WHERE codigo=" & Codigo
    SQL = SQL & " and anyo=" & Anyo
    Conn.Execute SQL
    
    'Insertamos para pasar a hco
    InsertaTmpActualizar Mc.Contador, vCP.diaricli, FechaAbono
    
    'Todo OK
    RemesasCancelacionEfectos = True
    
    
ERemesa_CancelacionCliente:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    
    End If
    Set RS = Nothing
    Set Mc = Nothing
    Set vCP = Nothing
End Function



'*********************************************************************************
'
'   TALONES / PAGARES
'
'*********************************************************************************
'*********************************************************************************
'
'
'   LaOpcion:   0. Cancelar cliente
'
'   Mayo 2009.  Cambio.  La cancelacion la realiza en la recepcion de documentos
'
'DiarioConcepto:Llevara el diario y los conceptos al debe y al haber. NO cojera los de la stipforpa, si no de una window anterior
'              El cuarto pipe que lleva es si agrrupa en la cuenta puente
'                   es decir, en lugar de 43.1 contra 431.1
'                                         43.2 contra 431.1
'                           hacemos   43.1 y 43.2   contra la suma en 431.1
' Septiembre 2009
'           El quinto y sexto pipe llevaran, si necesario, cta dodne poner el benefic po perd del talon y si requiere cc

'### Noviembre 2014
' Si es contra una unica cuenta puente de talon / pagare, entonces para
' el concepto del esta pondremos:
'       o la contrapartida(nomacta)
'       o como esta, el numero de talon pagare



Public Function RemesasCancelacionTALONPAGARE_(Talones As Boolean, IdRecepcion As Integer, FechaAbono As Date, DiarioConcepto As String) As Boolean
'Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
'Dim Gastos As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim RS As ADODB.Recordset
Dim AmpRemesa As String
Dim CtaCancelacion As String
Dim Cuenta As String
Dim RaizCuentasCancelacion As String
Dim CuentaUnica As Boolean
Dim LCta As Integer
Dim Importeauxiliar As Currency
Dim AgrupaVtosPuente As Boolean
Dim CadenaAgrupaVtoPuente As String
Dim aux2 As String
Dim RequiereCCDiferencia As Boolean

Dim Obs As String
Dim TipForpa As Byte

    On Error GoTo ERemesa_CancelacionCliente2
    RemesasCancelacionTALONPAGARE_ = False
    

    If Talones Then
        'Sobre talones
        Cuenta = "taloncta"
    Else
        Cuenta = "pagarecta"
    End If
    RaizCuentasCancelacion = DevuelveDesdeBD(Cuenta, "paramtesor", "codigo", "1", "N")
    If RaizCuentasCancelacion = "" Then
        MsgBox "Error grave en configuración de  parámetros de tesorería. Falta cuenta cancelación", vbExclamation
        Exit Function
    End If
    
    LCta = Len(RaizCuentasCancelacion)
    CuentaUnica = LCta = vEmpresa.DigitosUltimoNivel
    
    
    'Comprobacion.  Para todos los efectos de la 43.... se cancelan con la 4310....
    '
    'Tendre que ver que existen estas cuentas
    Set RS = New ADODB.Recordset
    
    SQL = "DELETE FROM tmpcierre1 where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
        
    If Not CuentaUnica Then
        'Cancela contra subcuentas de cliente
        

        Ampliacion = ",CONCAT('" & RaizCuentasCancelacion & "',SUBSTRING(codmacta," & LCta + 1 & ")" & ")"
            
        SQL = "Select " & vUsu.Codigo & Ampliacion & " from scarecepdoc where codigo=" & IdRecepcion
        SQL = SQL & " GROUP BY codmacta"
        'INSERT
        SQL = "INSERT INTO tmpcierre1(codusu,cta) " & SQL
        Conn.Execute SQL
        
        'Ahora monto el select para ver que cuentas 430 no tienen la 4310
        SQL = "Select cta,codmacta from tmpcierre1 left join cuentas on tmpcierre1.cta=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
        SQL = SQL & " HAVING codmacta is null"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        Linea = 0
        While Not RS.EOF
            Linea = Linea + 1
            SQL = SQL & RS!Cta & "     "
            If Linea = 5 Then
                SQL = SQL & vbCrLf
                Linea = 0
            End If
            RS.MoveNext
        Wend
        RS.Close
        
        If SQL <> "" Then
            
            AmpRemesa = "CANCELACION remesa"
            
            SQL = "Cuentas " & AmpRemesa & ".  No existen las cuentas: " & vbCrLf & String(90, "-") & vbCrLf & SQL
            SQL = SQL & vbCrLf & "¿Desea crearlas?"
            Linea = 1
            If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                'Ha dicho que si desea crearlas
                
                Ampliacion = "CONCAT('" & RaizCuentasCancelacion & "',SUBSTRING(codmacta," & LCta + 1 & ")) "
            
                SQL = "Select codmacta," & Ampliacion & " from talones where codigo=" & IdRecepcion
                SQL = SQL & " and " & Ampliacion & " in "
                SQL = SQL & "(Select cta from tmpcierre1 left join cuentas on tmpcierre1.cta=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
                SQL = SQL & " AND codmacta is null)"
                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RS.EOF
                
                     SQL = "INSERT  IGNORE INTO cuentas(codmacta,nommacta ,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos) SELECT '"
                                ' CUenta puente
                     SQL = SQL & RS.Fields(1) & "',nommacta ,'S',0,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos from cuentas where codmacta = '"
                                'Cuenta en la scbro (codmacta)
                     SQL = SQL & RS.Fields(0) & "'"
                     Conn.Execute SQL
                     RS.MoveNext
                     
                Wend
                RS.Close
                Linea = 0
            End If
            If Linea = 1 Then GoTo ERemesa_CancelacionCliente2
        End If
        
    Else
        'Cancela contra UNA unica cuenta todos los vencimientos
        SQL = DevuelveDesdeBD("codmacta", "cuentas", "codmacta", RaizCuentasCancelacion, "T")
        If SQL = "" Then
            MsgBox "No existe la cuenta puente: " & RaizCuentasCancelacion, vbExclamation
            GoTo ERemesa_CancelacionCliente2
        End If
    End If

    
    'La forma de pago
    Set vCP = New Ctipoformapago
    If Talones Then
        SQL = CStr(vbTalon)
        Ampliacion = "Talones"
    Else
        SQL = CStr(vbPagare)
        Ampliacion = "Pagarés"
    End If
    If vCP.Leer(CInt(SQL)) = 1 Then GoTo ERemesa_CancelacionCliente2
    'Ahora fijo los valores que me ha dado el
    vCP.diaricli = RecuperaValor(DiarioConcepto, 1)
    vCP.condecli = RecuperaValor(DiarioConcepto, 2)
    vCP.conhacli = RecuperaValor(DiarioConcepto, 3)
    AgrupaVtosPuente = RecuperaValor(DiarioConcepto, 4) = 1
 '   AgrupaVtosPuente = AgrupaVtosPuente 'And CuentaUnica
    Set Mc = New Contadores
    
    
    If Mc.ConseguirContador("0", FechaAbono <= vParam.fechafin, True) = 1 Then Exit Function
    
    
    'Insertamos la cabera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
    SQL = SQL & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ", '"
    SQL = SQL & "Cancelacion cliente"

    SQL = SQL & " NºRecepcion: " & IdRecepcion & "   " & Ampliacion & vbCrLf
    SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & vUsu.Nombre & "',"
    
    Obs = "ARICONTA 6: Cancelacion cliente NºRecepcion: " & IdRecepcion & "   " & Ampliacion & vbCrLf
    SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(Obs, "T") & ") ;"
    
    
    
    If Not Ejecuta(SQL) Then Exit Function
    
    
    
    
    Linea = 1
    Importe = 0
    'Gastos = 0
    
    vCP.descformapago = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)  'DEBE
    vCP.CadenaAuxiliar = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)   'DEBE
    
    
    SQL = "select cobros.*,l.importe,l.codigo,c.numeroref reftalonpag,c.banco  from (talones c inner join talones_facturas l on c.codigo = l.codigo) left join  cobros  on l.numserie=cobros.numserie and"
    SQL = SQL & " l.numfactu=cobros.numfactu and   l.fecfactu=cobros.fecfactu and l.numorden=cobros.numorden"
    SQL = SQL & " WHERE c.codigo= " & IdRecepcion
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Trozo comun
    AmpRemesa = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    AmpRemesa = AmpRemesa & "codmacta, numdocum, codconce, ampconce,timporteD,"
    AmpRemesa = AmpRemesa & " timporteH, codccost, ctacontr, idcontab, punteada, "
    AmpRemesa = AmpRemesa & " numserie, numfaccl, fecfactu, numorden, tipforpa, reftalonpag, bancotalonpag) "
    AmpRemesa = AmpRemesa & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & ","
    
    CadenaAgrupaVtoPuente = ""

    While Not RS.EOF
        Ampliacion = RS!NUmSerie   'SI DA ERROR ES QUE NO EXISTE mediante el left join
        
        
        
        'Neuvo dato para la ampliacion en la contabilizacion
        Ampliacion = " "
        Select Case vCP.amphacli
        Case 2, 4
            'La opcion Contrapartida BANCO NO vale ahora, pq no hay apunte a banco
            Ampliacion = DBLet(RS!reftalonpag, "T")
            If Ampliacion = "" Then Ampliacion = Ampliacion & Format(RS!FecVenci, "dd/mm/yyyy")
        Case 5
            Ampliacion = DBLet(RS!reftalonpag, "T")
            If Ampliacion = "" Then
                Ampliacion = RS!NUmSerie & "/" & RS!NumFactu
            Else
                Ampliacion = "Doc: " & Ampliacion
            End If
        Case Else
           If vCP.amphacli = 1 Then Ampliacion = vCP.siglas & " "
           Ampliacion = Ampliacion & RS!NUmSerie & "/" & RS!NumFactu
        End Select
        If Mid(Ampliacion, 1, 1) <> " " Then Ampliacion = " " & Ampliacion
        
        'Cancelacion
        If CuentaUnica Then
            Cuenta = RaizCuentasCancelacion
        Else
            Cuenta = RaizCuentasCancelacion & Mid(RS!codmacta, LCta + 1)
            
        End If
        CtaCancelacion = RS!codmacta
    
        
        
        
        'Si dice que agrupamos vto entonces NO
        If AgrupaVtosPuente Then
            If CadenaAgrupaVtoPuente = "" Then
                'Para insertarlo al final del proceso
                'Antes de ejecutar el sql(al final) substituiremos, la cadena
                ' la cadena ### por el importe total
                
                SQL = "1,'" & Cuenta & "','Nº" & Format(IdRecepcion, "0000000") & "'," & vCP.condecli
                
                
                'Noviembre 2014
                'si pone contrapartida, pondre la nommacta
                aux2 = ""
                If vCP.ampdecli = 4 Then aux2 = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", CtaCancelacion, "T")
                
                If aux2 = "" Then aux2 = Mid(vCP.descformapago & " " & DBLet(RS!reftalonpag, "T"), 1, 30)
                
                SQL = SQL & ",'" & DevNombreSQL(aux2) & "',"
                aux2 = ""
                'Importe al DEBE.
                SQL = SQL & "###,NULL,NULL,"
                'Contra partida
                SQL = SQL & "'" & CtaCancelacion & "','CONTAB',0,"
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ") "
                
                CadenaAgrupaVtoPuente = AmpRemesa & SQL
            End If
        End If
            
            
        
        
        'Crearemos el apnte y la contrapartida
        ' Es decir
        '   4310  contra 430
        '   430  contr   4310
        'Lineas de apuntes .
        
         
        'Cuenta
        SQL = Linea & ",'" & Cuenta & "','" & Format(RS!NumFactu, "000000000") & "'," & vCP.condecli
        
        
        'Noviembre 2014
        'Noviembre 2014
        'si pone contrapartida, pondre la nommacta
        aux2 = ""
        If vCP.ampdecli = 4 Then aux2 = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", CtaCancelacion, "T")
        If aux2 = "" Then aux2 = Mid(vCP.descformapago & Ampliacion, 1, 30)
        SQL = SQL & ",'" & DevNombreSQL(aux2) & "',"
        
        
        
        
        Importe = Importe + RS!Importe
        'Gastos = Gastos + DBLet(Rs!Gastos, "N")
        
        
        'Importe VA alhaber del cliente, al debe de la cancelacion
        SQL = SQL & TransformaComasPuntos(RS!Importe) & ",NULL,NULL,"
    
        'Contra partida
        SQL = SQL & "'" & CtaCancelacion & "','CONTAB',0,"
        
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ") "
        
        
        
        
        SQL = AmpRemesa & SQL
        If Not AgrupaVtosPuente Then
            If Not Ejecuta(SQL) Then Exit Function
        End If
        Linea = Linea + 1 'Siempre suma mas uno
        
        
        'La contrapartida
        SQL = Linea & ",'" & CtaCancelacion & "','" & Format(RS!NumFactu, "000000000") & "'," & vCP.conhacli
        SQL = SQL & ",'" & DevNombreSQL(Mid(vCP.CadenaAuxiliar & Ampliacion, 1, 30)) & "',"
        
        
        '
        SQL = SQL & "NULL," & TransformaComasPuntos(RS!Importe) & ",NULL,"
    
        'Contra partida
        SQL = SQL & "'" & Cuenta & "','CONTAB',0,"
        
        TipForpa = DevuelveDesdeBD("tipforpa", "formapago", "codforpa", RS!codforpa, "N")

        
        SQL = SQL & DBSet(RS!NUmSerie, "T") & "," & DBSet(RS!NumFactu, "N") & "," & DBSet(RS!FecFactu, "F") & "," & DBSet(RS!numorden, "N") & ","
        SQL = SQL & DBSet(TipForpa, "T") & "," & DBSet(RS!reftalonpag, "T") & "," & DBSet(RS!Banco, "T") & ") "
        
        SQL = AmpRemesa & SQL
        
        If Not Ejecuta(SQL) Then Exit Function
        Linea = Linea + 1
            
        '
        RS.MoveNext
    Wend
    RS.Close



    
    SQL = "Select importe,codmacta,numeroref from talones where codigo = " & IdRecepcion
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then Err.Raise 513, , "No se ha encontrado documento: " & IdRecepcion
    Importeauxiliar = RS!Importe
    Cuenta = RS!codmacta
    Ampliacion = DevNombreSQL(RS!numeroref)
    RS.Close


    If Importe <> Importeauxiliar Then
    
        CtaCancelacion = RecuperaValor(DiarioConcepto, 5)
        If CtaCancelacion = "" Then Err.Raise 513, , "Cuenta beneficios/pérdidas  NO espeficicada"
        
        'Hemos llegado a aqui.
        'Veremos si hay diferencia entre la suma de importe y el importe del documento.
        '
        Importeauxiliar = Importeauxiliar - Importe
        If Len(Ampliacion) > 10 Then Ampliacion = Right(Ampliacion, 10)
        
        SQL = Linea & ",'" & CtaCancelacion & "','Nº" & Format(IdRecepcion, "00000000") & "'," & vCP.condecli
        
        'Ampliacion
        If Talones Then
            aux2 = " Tal nº: " & Ampliacion
        Else
            aux2 = " Pag. nº: " & Ampliacion
        End If
        SQL = SQL & ",'" & DevNombreSQL(Mid(vCP.descformapago & aux2, 1, 30)) & "',"

        
        If Importeauxiliar < 0 Then
            'NEgativo. Va en positivo al otro lado
            SQL = SQL & TransformaComasPuntos(Abs(Importeauxiliar)) & ",NULL,"
        Else
            SQL = SQL & "NULL," & TransformaComasPuntos(CStr(Importeauxiliar)) & ","
        End If
                
        'Centro de coste
        RequiereCCDiferencia = False
        If vParam.autocoste Then
            aux2 = Mid(CtaCancelacion, 1, 1)
            If aux2 = "6" Or aux2 = "7" Then RequiereCCDiferencia = True
        End If
        If RequiereCCDiferencia Then
            CtaCancelacion = UCase(RecuperaValor(DiarioConcepto, 6))
            If CtaCancelacion = "" Then Err.Raise 513, , "Centro de coste  NO espeficicado"
            CtaCancelacion = "'" & CtaCancelacion & "'"
        Else
             CtaCancelacion = "NULL"
        End If
        SQL = SQL & CtaCancelacion
        
        'Contra partida
        If CuentaUnica Then
            Cuenta = "'" & RaizCuentasCancelacion & "'"
        Else
            Cuenta = "NULL"
        End If
        
        
        SQL = SQL & "," & Cuenta & ",'CONTAB',0,"
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ") "
        
        
        
        
        SQL = AmpRemesa & SQL
        
        If Not Ejecuta(SQL) Then Exit Function
        Linea = Linea + 1
        
        
        If AgrupaVtosPuente Then
            'Modificamos el importe final por si esta agrupando vencimientos
            Importe = Importeauxiliar + Importe
        Else
                'creamos la contrapartida para que  cuadre el asiento
                
                If Len(Ampliacion) > 10 Then Ampliacion = Right(Ampliacion, 10)
                
                SQL = Linea & "," & Cuenta & ",'Nº" & Format(IdRecepcion, "00000000") & "'," & vCP.conhacli
                
                 If Talones Then
                    aux2 = " Tal nº: " & Ampliacion
                Else
                    aux2 = " Pag. nº: " & Ampliacion
                End If
                SQL = SQL & ",'" & DevNombreSQL(Mid(vCP.CadenaAuxiliar & aux2, 1, 30)) & "',"
                
                If Importeauxiliar > 0 Then
                    'NEgativo. Va en positivo al otro lado
                    SQL = SQL & TransformaComasPuntos(CStr(Importeauxiliar)) & ",NULL,"
                Else
                    SQL = SQL & "NULL," & TransformaComasPuntos(Abs(Importeauxiliar)) & ","
                End If
                        
                'Centro de coste
                RequiereCCDiferencia = False
                If vParam.autocoste Then
                    aux2 = Mid(Cuenta, 2, 1)  'pq lleva una comilla
                    If aux2 = "6" Or aux2 = "7" Then RequiereCCDiferencia = True
                End If
                If RequiereCCDiferencia Then
                    CtaCancelacion = UCase(RecuperaValor(DiarioConcepto, 6))
                    If CtaCancelacion = "" Then Err.Raise 513, , "Centro de coste  NO espeficicado"
                    CtaCancelacion = "'" & CtaCancelacion & "'"
                Else
                     CtaCancelacion = "NULL"
                End If
                SQL = SQL & CtaCancelacion
                
                'Contra partida
                CtaCancelacion = RecuperaValor(DiarioConcepto, 5)
                SQL = SQL & ",'" & CtaCancelacion & "','CONTAB',0,"
                
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ") "
        
                SQL = AmpRemesa & SQL
                
                If Not Ejecuta(SQL) Then Exit Function
                Linea = Linea + 1
            End If
                
    End If
    
    If AgrupaVtosPuente Then
        'Tenmos que reemplazar
        'en CadenaAgrupaVtoPuente    ###:importe
        SQL = TransformaComasPuntos(CStr(Importe))
        SQL = Replace(CadenaAgrupaVtoPuente, "###", SQL)
        Conn.Execute SQL
    End If


    AmpRemesa = "F"    ' cancelada
    
    SQL = "UPDATE talones SET contabilizada = 1"
    SQL = SQL & " WHERE codigo = " & IdRecepcion
    
    Conn.Execute SQL

    
    'Insertamos para pasar a hco
    InsertaTmpActualizar Mc.Contador, vCP.diaricli, FechaAbono
    
    'Todo OK
    RemesasCancelacionTALONPAGARE_ = True
    
    
ERemesa_CancelacionCliente2:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
   
    End If
    Set RS = Nothing
    Set Mc = Nothing
    Set vCP = Nothing
End Function




'Devolvera
'       0. ADelante, pero no actualizamos pq no habra nada
'       1    "  y acutalizamos
'       2    Errores
Public Function RemesasEliminarVtosRem2(Codigo As Integer, Anyo As Integer, FechaAbono As Date, ByRef FP As Ctipoformapago, AgrupaCancelacion As Boolean) As Byte
'Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
Dim SQL As String
Dim Ampliacion As String
Dim RS As ADODB.Recordset
Dim AmpRemesa As String
Dim CtaCancelacion As String
Dim Cuenta As String
Dim RaizCuentasCancelacion As String
Dim CuentaUnica As Boolean
Dim LCta As Integer
Dim ImporteRiesgoBanco As Currency
Dim Dias1 As Integer 'Si no llega al limite
Dim Dias2 As Integer 'Si pasa del limite
Dim J As Integer
Dim F As Date
Dim CuentaPuente As Boolean
Dim BorrarEfecto As Boolean
Dim ImporteVto As Currency
Dim EstaAgrupandoVtos As Boolean
Dim CadenaAgrupacion As String
Dim ParaLineasDocumentosRecibidos As String
Dim DiferenciasImportes As String
Dim TipoRemesa As Byte
Dim NumeroEfectos As Integer
Dim Eliminados As Integer


    TipoRemesa = 1 'FALTA QUITAR todo lo k no se utilice
    On Error GoTo ERemesa_Elivto
    RemesasEliminarVtosRem2 = 2
    
    
    CuentaPuente = False
    EstaAgrupandoVtos = False
    If TipoRemesa > 1 Then
        'Sobre talones
        'NO DEBERIA HABER ENTRADO EN ESTA OPCION
        MsgBox "Error. Opcion eliminar vto incorrecta. Avise soporte tecnico", vbCritical
        Exit Function
    End If
        
    'Efectos. Viene de cancelacion
    SQL = "ctaefectcomerciales"
    Cuenta = "RemesaCancelacion"


    
    If CuentaPuente Then
        RaizCuentasCancelacion = DevuelveDesdeBD(Cuenta, "paramtesor", "codigo", "1", "N", SQL)
        If RaizCuentasCancelacion = "" Then
            MsgBox "Error grave en configuracion de  parametros de tesoreria. Falta cuenta cancelacion", vbExclamation
            Exit Function
        End If
        
        '<>"" Siginifca que lleva cuenta de efectos comerciales descontados.
        '      con lo cual la cuenta de cancelacion es la que pone aqui.
        '      Y sera a ultimo nivel o raiz, igual que la otra
        If SQL <> "" Then RaizCuentasCancelacion = SQL
        
        
        LCta = Len(RaizCuentasCancelacion)
        CuentaUnica = LCta = vEmpresa.DigitosUltimoNivel
        If TipoRemesa = 1 Then EstaAgrupandoVtos = CuentaUnica And AgrupaCancelacion
    End If
            
    Set RS = New ADODB.Recordset
    
    
    'Datos bancos. Importe maximo para dias 1, dias2 si no llega
    If TipoRemesa = 3 Then
        'Sobre talones
        Cuenta = "100000000,talondias,talondias"
    ElseIf TipoRemesa = 2 Then
        Cuenta = "100000000,pagaredias,pagaredias"
    Else
        'Efectos.
        Cuenta = "remesariesgo,remesadiasmenor,remesadiasmayor"
    End If
        
    SQL = "select ctaefectosdesc," & Cuenta & " from remesas r,ctabancaria b where r.codmacta=b.codmacta and codigo=" & Codigo & " AND anyo = " & Anyo
    CtaCancelacion = ""
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    SQL = ""
    If Not RS.EOF Then
        
        If Not IsNull(RS.Fields(0)) Then
            CtaCancelacion = RS.Fields(0)
        Else
            If CuentaPuente Then SQL = "Cuenta efectos descontados"
        End If
        
        If IsNull(RS.Fields(2)) Then
            SQL = SQL & "Dias eliminacion"
        Else
            Dias1 = RS.Fields(2)
            Dias2 = RS.Fields(3)
        End If
        
        
        
            'Esto solo puede pasar en los efectos
            If IsNull(RS.Fields(1)) Then
                SQL = SQL & " Importe riesgo efectos"
            Else
                ImporteRiesgoBanco = RS.Fields(1)
            End If
        
        
        Dias2 = DBLet(RS.Fields(3), "N")
    End If
    RS.Close
    If SQL <> "" Then
       
            'Si lleva cuenta puente, falta configurar
            MsgBox "Falta configurar: " & SQL, vbExclamation
            GoTo ERemesa_Elivto

    End If


    'La forma de pago
    If CuentaPuente Then
            
            If TipoRemesa = 3 Then
                'SQL = CStr(vbTalon)
                Ampliacion = "Talones"
            ElseIf TipoRemesa = 2 Then
                'SQL = CStr(vbPagare)
                Ampliacion = "Pagarés"
            Else
                'SQL = CStr(vbTipoPagoRemesa)
                Ampliacion = "Efectos"
            End If
            
            
            
            Set Mc = New Contadores
            
            
            If Mc.ConseguirContador("0", FechaAbono <= vParam.fechafin, True) = 1 Then GoTo ERemesa_Elivto
            
        
        
            'Insertamos la cabera
            SQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) VALUES ("
            SQL = SQL & FP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador
            SQL = SQL & ", 1, NULL, 'Eliminacion de efectos remesa: " & Codigo & " / " & Anyo & "   " & Ampliacion & vbCrLf
            SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & vUsu.Nombre & "');"
            If Not Ejecuta(SQL) Then Exit Function
        
    
        
    
            Linea = 1
            Importe = 0

    
            FP.descformapago = DevuelveDesdeBD("nomconce", "conceptos", "codconce", FP.conhacli)
    
    End If
    
    
    
  
    'Si es talon pagare el importe lo coje de las lineas de vto, NO de la scobro
    If TipoRemesa > 1 Then
        'TALON PAGARE
        SQL = "select scobro.*,l.importe,l.numserie vto from   slirecepdoc l left join  scobro  on l.numserie=scobro.numserie and"
        SQL = SQL & " l.numfaccl=scobro.codfaccl and   l.fecfaccl=scobro.fecfaccl and l.numvenci=scobro.numorden"
        
    Else
        SQL = "Select * from scobro "
    End If
    SQL = SQL & " where codrem=" & Codigo & " AND anyorem = " & Anyo
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Si no lleva cuenta puente, NO contbiliza nada
    If CuentaPuente Then
        AmpRemesa = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, "
        AmpRemesa = AmpRemesa & "codmacta, numdocum, codconce, ampconce,timporteD,"
        AmpRemesa = AmpRemesa & " timporteH, codccost, ctacontr, idcontab, punteada) "
        AmpRemesa = AmpRemesa & "VALUES (" & FP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & ","
    End If
    CadenaAgrupacion = ""
    NumRegElim = 0 'Total registro
    NumeroEfectos = 0
    Eliminados = 0
    While Not RS.EOF
        
        J = DateDiff("d", RS!FecVenci, FechaAbono)
        If RS!ImpVenci > ImporteRiesgoBanco Then
            J = J - Dias2
        Else
            J = J - Dias1
        End If
        NumRegElim = NumRegElim + 1
        NumeroEfectos = NumeroEfectos + 1
        If J > 0 Then
        
            If CuentaPuente Then
        
        
                'Han pasado mas dias de los que poner en paraemtros. Podremos borrar el efecto
                'Ampliacion
                Ampliacion = FP.descformapago & " "
                   
                   
                If EstaAgrupandoVtos Then
                   If CadenaAgrupacion = "" Then
                        'Creo el la linea para insertar
                        
                        CadenaAgrupacion = Linea & ",'" & RaizCuentasCancelacion & "','RE" & Format(Codigo, "0000") & Anyo & "'," & FP.conhacli
                        CadenaAgrupacion = CadenaAgrupacion & ",'" & DevNombreSQL(Mid(Ampliacion & "Rem: " & Codigo & "-" & Anyo, 1, 30)) & "',NULL,@@@@@@"
                
                        CadenaAgrupacion = CadenaAgrupacion & ",NULL,'" & CtaCancelacion & "','CONTAB',0)"
                        'Luego reemplazare @@@@@@ por el importe total
                    End If
                            
                End If
                
                'Neuvo dato para la ampliacion en la contabilizacion
                Select Case FP.amphacli
                Case 2, 4
                    'La opcion Contrapartida BANCO NO vale ahora, pq no hay apunte a banco
                    Ampliacion = Ampliacion & Format(RS!FecVenci, "dd/mm/yyyy")
                    
                Case Else
                   If FP.amphacli = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
                   Ampliacion = Ampliacion & RS!NUmSerie & "/" & RS!codfaccl
                End Select
                
        

                Cuenta = RaizCuentasCancelacion
                If Not CuentaUnica Then Cuenta = Cuenta & Mid(RS!codmacta, LCta + 1)
                
            
            
             
                'Cuenta
                SQL = Linea & ",'" & Cuenta & "','" & Format(RS!codfaccl, "000000000") & "'," & FP.conhacli
                SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',NULL,"
                If TipoRemesa = 1 Then
                    ImporteVto = RS!ImpVenci
                    
                Else
                    'Talones pagares, podria ser que no pagara todo.
                    'El rs esta con left join sobre slirecdoc
                    'Si es NULL vto significa que esta en scobro y no en slirecdoc. Algo esta mal.
                    If IsNull(RS!Vto) Then
                        If MsgBox("Vencimiento no existe en rececpcion documentos. ¿Continuar?", vbYesNo) = vbNo Then
                            'esto generara un error
                            SQL = RS!Vto
                        Else
                            ImporteVto = RS!impcobro
                        End If
                    Else
                        ImporteVto = RS!Importe
                    End If
                End If
                
                Importe = Importe + ImporteVto
                'Importe VA alhaber de
                SQL = SQL & TransformaComasPuntos(CStr(ImporteVto)) & ",NULL,"
            
                'Contra partida
                SQL = SQL & "'" & CtaCancelacion & "','CONTAB',0)"
                SQL = AmpRemesa & SQL
                If Not EstaAgrupandoVtos Then
                    If Not Ejecuta(SQL) Then Exit Function
                End If
            Else
                If TipoRemesa > 1 Then
                    ImporteVto = RS!Importe
                Else
                    ImporteVto = RS!ImpVenci
                End If
            End If
            Linea = Linea + 1
            

            
            'Me cargo el efecto y si tuviera devoluciones
            'Para talones/pagares podria darse el caso que NO todo el importe es
            'es el que ha sido pagado. Entonces procederemos de otra forma
            BorrarEfecto = True
            If TipoRemesa <> 1 Then
                'TALONES REMESAS
                
                'Compruebo que el total impcobrado es el total +gastos
                
                If RS!ImpVenci + DBLet(RS!Gastos, "N") > ImporteVto Then BorrarEfecto = False
                    
                    
            End If
            'Comun a borrar y updatear
            SQL = " WHERE numserie='" & RS!NUmSerie & "' AND codfaccl = " & RS!codfaccl
            SQL = SQL & " AND fecfaccl = '" & Format(RS!fecfaccl, FormatoFecha) & "' AND numorden = " & RS!numorden
            If BorrarEfecto Then
                
                Conn.Execute "DELETE FROM scobro " & SQL
                Conn.Execute "DELETE FROM sefecdev " & SQL
                Eliminados = Eliminados + 1
            Else
                'Se trata de actualizar
                'Vamos a quitar la marca de remesado. En importe pondremos el importe que habia y en el campo observaciones
                'indicaremos que ya ha sido remesado y por cuanto importe
                
                'Campo obs..servaciones guardara los datos antiguos
                Ampliacion = "Vto: " & Format(RS!ImpVenci, FormatoImporte)
                If DBLet(RS!Gastos, "N") > 0 Then Ampliacion = Ampliacion & "/ Gastos " & Format(RS!Gastos, FormatoImporte)
                'Fecha del ultimo cobro
                Ampliacion = Ampliacion & "   Ultimo cobro: " & Format(RS!fecultco, "dd/mm/yyyy") & "  " & Format(RS!impcobro, FormatoImporte)
                'Lo meto en la observacion
                Ampliacion = "obs = '" & Ampliacion & "',"
                
                
                'Agosto 2009
                'Como el vto esta en slirecepdoc NO hace falta cponer esto
                'Ampliacion = Ampliacion & "fecultco = NULL, impcobro=NULL,"
                Ampliacion = Ampliacion & "codrem=NULL,Tiporem=NULL,Anyorem=NULL,siturem=NULL"
                
                
                'Los gastos los pondre a null
                Ampliacion = Ampliacion & ",gastos=NULL"
                'ImporteVto = Rs!impvenci + DBLet(Rs!Gastos, "N") - Rs!impcobro
                'Ampliacion = Ampliacion & ",impvenci = " & TransformaComasPuntos(CStr(ImporteVto))
                
                'Raferencia talon/pagare tb
                Ampliacion = Ampliacion & ",reftalonpag=NULL,recedocu=0"
                
                ParaLineasDocumentosRecibidos = SQL
                
                SQL = "UPDATE scobro SET " & Ampliacion & SQL
                Ampliacion = ""
                Conn.Execute SQL
                
                'Realmente para tipo=1 NO deberia llegar aquin
                If TipoRemesa <> 1 Then
                        SQL = "UPDATE slirecepdoc SET numserie="" """ & ParaLineasDocumentosRecibidos
                        SQL = Replace(SQL, "numorden", "numvenci")
                        SQL = Replace(SQL, "codfaccl", "numfaccl")
                        If Not Ejecuta(SQL) Then
                            SQL = "Error actualizando tabla lineas documentos recibidos"
                            MsgBox SQL, vbExclamation
                        End If
                End If
                
            End If
        End If
            
        RS.MoveNext
    Wend
    RS.Close

    
        If EstaAgrupandoVtos Then
            If CadenaAgrupacion <> "" Then
                'OK inserto el total
                Ampliacion = TransformaComasPuntos(CStr(Importe))
                SQL = Replace(CadenaAgrupacion, "@@@@@@", Ampliacion)
                Conn.Execute AmpRemesa & SQL
                Linea = 2 'La uno es
            End If
        End If
    
        If Linea > 1 Then
            'Hago el contrapunte
            If CuentaPuente Then
                Ampliacion = FP.descformapago & " Re: " & Codigo & " - " & Anyo
                SQL = "RE" & Format(Codigo, "0000") & Format(Anyo, "0000")
                SQL = Linea & ",'" & CtaCancelacion & "','" & SQL & "'," & FP.conhacli
                SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
            
            
                'Importe al DEBE
                SQL = SQL & TransformaComasPuntos(CStr(Importe)) & ",NULL,NULL,"
        
                'Contra partida
                SQL = SQL & "NULL,'CONTAB',0)"
                SQL = AmpRemesa & SQL
            
            
  
                Conn.Execute SQL
            Else
                Linea = Linea + 1 'Para que despues no de distinto el numero de efectos eliminados
            End If
        Else
            
            
            'Muestro un mesaje diciendo que Ningun vto ha sido eliminado. No deberia pasar pero por si acaso
            'compruebo que tenga vtos
            If NumRegElim > 0 Then MsgBox "No se ha podido eliminar ningun vencimiento de la remesa " & Codigo & " / " & Anyo, vbInformation
            RemesasEliminarVtosRem2 = 0
            
            If CuentaPuente Then
                SQL = "DELETE FROM cabapu  WHERE numdiari =" & FP.diaricli
                SQL = SQL & " and fechaent = '" & Format(FechaAbono, FormatoFecha) & "' and numasien = " & Mc.Contador
                Conn.Execute SQL
            End If
        End If


    'Si la hemos borrado toda, o no....
    Linea = Linea - 1 'Empieza en uno, luego el total vtos eliminados es linea-1
                      'En numregelim tengo los vtos totales de la remesa
                      'Si queda alguno o no, haremos unas cosas u otras
    If Eliminados < NumeroEfectos Then
        AmpRemesa = "Y"
    
        'QUEDA ALGUNO
        SQL = "UPDATE scobro SET"
        SQL = SQL & " siturem= 'Y'"
        SQL = SQL & " WHERE codrem=" & Codigo
        SQL = SQL & " and anyorem=" & Anyo
           
        Conn.Execute SQL
    Else
        AmpRemesa = "Z"  'TOdos eliminados
    End If
    SQL = "UPDATE remesas SET"
    SQL = SQL & " situacion= '" & AmpRemesa
    SQL = SQL & "' WHERE codigo=" & Codigo
    SQL = SQL & " and anyo=" & Anyo
    Conn.Execute SQL
    
    'Insertamos para pasar a hco
    If CuentaPuente Then InsertaTmpActualizar Mc.Contador, FP.diaricli, FechaAbono
    
    'Todo OK
    If NumRegElim > 0 Then
        RemesasEliminarVtosRem2 = 1
        'Para que no actualice el apunte , ya que no se ha creado
        If Not CuentaPuente Then RemesasEliminarVtosRem2 = 0
    End If
    
ERemesa_Elivto:
    If Err.Number <> 0 Then
        
        MuestraError Err.Number, Err.Description
        RemesasEliminarVtosRem2 = 2
    End If
    Set RS = Nothing
    Set Mc = Nothing

End Function



















'TALONES PAGARES
Public Function RemesasEliminarVtosTalonesPagares(TipoRemesa As Byte, Codigo As Integer, Anyo As Integer, FechaAbono As Date, ByRef FP As Ctipoformapago, AgrupaCancelacion_ As Boolean) As Byte
'Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
Dim SQL As String
Dim Ampliacion As String
Dim RS As ADODB.Recordset
Dim AmpRemesa As String
Dim CtaCancelacion As String
Dim Cuenta As String
Dim RaizCuentasCancelacion As String
Dim CuentaUnica As Boolean
Dim LCta As Integer
Dim Dias1 As Integer 'Si no llega al limite
Dim J As Integer
Dim F As Date
Dim CuentaPuente As Boolean
Dim BorrarEfecto As Boolean
Dim ImporteVto As Currency
Dim EstaAgrupandoVtos As Boolean
Dim CadenaAgrupacion As String
Dim ParaLineasDocumentosRecibidos As String
Dim vId As Integer
Dim ImporteDocumento As Currency
Dim SumasImportesDocumentos As Currency

Dim EliminaEnRecepcionDocumentos As String

    On Error GoTo ERemesa_Elivto
    RemesasEliminarVtosTalonesPagares = 2
    
    
    CuentaPuente = False
    EstaAgrupandoVtos = False
    If TipoRemesa = 3 Then
        'Sobre talones
        Cuenta = "taloncta"
        CuentaPuente = vParamT.PagaresCtaPuente
    ElseIf TipoRemesa = 2 Then
        CuentaPuente = vParamT.TalonesCtaPuente
        Cuenta = "pagarecta"
    Else
        'Efectos. Viene de cancelacion
        Cuenta = "RemesaCancelacion"
    End If
    
    
    If CuentaPuente Then
        RaizCuentasCancelacion = DevuelveDesdeBD(Cuenta, "paramtesor", "codigo", "1", "N")
        If RaizCuentasCancelacion = "" Then
            MsgBox "Error grave en configuracion de  parametros de tesoreria. Falta cuenta cancelacion", vbExclamation
            Exit Function
        End If
        
        LCta = Len(RaizCuentasCancelacion)
        CuentaUnica = LCta = vEmpresa.DigitosUltimoNivel
      
        EstaAgrupandoVtos = AgrupaCancelacion_
       
    End If
            
    Set RS = New ADODB.Recordset
    
    EliminaEnRecepcionDocumentos = ""
    'Datos bancos. Importe maximo para dias 1, dias2 si no llega
    If TipoRemesa = 3 Then
        'Sobre talones
        Cuenta = "talondias"
    Else
        Cuenta = "pagaredias"
    End If
        
    SQL = "select ctaefectosdesc," & Cuenta & " from remesas r,ctabancaria b where r.codmacta=b.codmacta and codigo=" & Codigo & " AND anyo = " & Anyo
    CtaCancelacion = ""
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    SQL = ""
    If Not RS.EOF Then
        
        If Not IsNull(RS.Fields(0)) Then
            CtaCancelacion = RS.Fields(0)
        Else
            If CuentaPuente Then SQL = "Cuenta efectos descontados"
        End If
        
        If IsNull(RS.Fields(1)) Then
            SQL = SQL & "Dias eliminacion"
        Else
            Dias1 = RS.Fields(1)
        End If

    End If
    RS.Close
    If SQL <> "" Then
        MsgBox "Falta configurar: " & SQL, vbExclamation
        GoTo ERemesa_Elivto
    End If


    'La forma de pago
    If CuentaPuente Then
            
            If TipoRemesa = 3 Then
                'SQL = CStr(vbTalon)
                Ampliacion = "Talones"
            ElseIf TipoRemesa = 2 Then
                'SQL = CStr(vbPagare)
                Ampliacion = "Pagarés"
            Else
                'SQL = CStr(vbTipoPagoRemesa)
                Ampliacion = "Efectos"
            End If
            
            
            
            Set Mc = New Contadores
            
            
            If Mc.ConseguirContador("0", FechaAbono <= vParam.fechafin, True) = 1 Then GoTo ERemesa_Elivto
            
        
        
            'Insertamos la cabera
            SQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) VALUES ("
            SQL = SQL & FP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador
            SQL = SQL & ", 1, NULL, 'Eliminacion de efectos remesa: " & Codigo & " / " & Anyo & "   " & Ampliacion & vbCrLf
            SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & vUsu.Nombre & "');"
            If Not Ejecuta(SQL) Then Exit Function
        
    
        
    
            Linea = 1
            Importe = 0

    
            FP.descformapago = DevuelveDesdeBD("nomconce", "conceptos", "codconce", FP.conhacli)
            FP.CadenaAuxiliar = DevuelveDesdeBD("nomconce", "conceptos", "codconce", FP.condecli)
    End If
    
    
    
  
    'Si es talon pagare el importe lo coje de las lineas de vto, NO de la scobro

    SQL = "select scobro.*,l.importe,l.numserie vto,id from   slirecepdoc l left join  scobro  on l.numserie=scobro.numserie and"
    SQL = SQL & " l.numfaccl=scobro.codfaccl and   l.fecfaccl=scobro.fecfaccl and l.numvenci=scobro.numorden"
    SQL = SQL & " where codrem=" & Codigo & " AND anyorem = " & Anyo
    SQL = SQL & " ORDER BY ID" 'Para ir comprobando documento por documento si
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Si no lleva cuenta puente, NO contbiliza nada
    If CuentaPuente Then
        AmpRemesa = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, "
        AmpRemesa = AmpRemesa & "codmacta, numdocum, codconce, ampconce,timporteD,"
        AmpRemesa = AmpRemesa & " timporteH, codccost, ctacontr, idcontab, punteada) "
        AmpRemesa = AmpRemesa & "VALUES (" & FP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & ","
    End If
    CadenaAgrupacion = ""
    NumRegElim = 0 'Total registro
    SumasImportesDocumentos = 0
    EliminaEnRecepcionDocumentos = "|"
    While Not RS.EOF
        J = DateDiff("d", RS!FecVenci, FechaAbono)
        J = J - Dias1
        NumRegElim = NumRegElim + 1
        If J > 0 Then
            
                    'AHora me guardo, si procede, el id, asi luego vere si puedo eliminar
                    'en recepcion de documentos
                    If vParamT.EliminaRecibidosRiesgo Then
                        SQL = "|" & RS!Id & "|"
                        If InStr(1, EliminaEnRecepcionDocumentos, SQL) = 0 Then EliminaEnRecepcionDocumentos = EliminaEnRecepcionDocumentos & RS!Id & "|"
                    End If
            
            
            
            
            If CuentaPuente Then
                If vId <> RS!Id Then
                    'Ha cambiado de documento
                    If vId > 0 Then
                        'Conseguimos importe documento
                        Cuenta = "codmacta"
                        Ampliacion = DevuelveDesdeBD("importe", "scarecepdoc", "codigo", CStr(vId), "N", Cuenta)
                        ImporteVto = CCur(Ampliacion)
                        'Comprobamos con el importe parcial.
                        If ImporteVto <> ImporteDocumento Then
                            'Ha habido difernecias

                        
                            'Y si no agrupamos
                            If EstaAgrupandoVtos Then
                                'Metemos uno ajustando importe
                        
                        
                            End If 'de agrupando
                        End If '<>importe
                    End If 'ID>0
                    'Inicializamos valores
                    vId = RS!Id
                    SumasImportesDocumentos = SumasImportesDocumentos + ImporteDocumento
                    ImporteDocumento = 0
                    
                    

            
            

                    
                    
                End If 'rs!id <>ID
                    
                'Han pasado mas dias de los que poner en paraemtros. Podremos borrar el efecto
                'Ampliacion
                Ampliacion = FP.descformapago & " "
                   
                   
                If EstaAgrupandoVtos Then
                   If CadenaAgrupacion = "" Then
                        'Creo el la linea para insertar
                        If Not CuentaUnica Then
                            Cuenta = RaizCuentasCancelacion
                            Cuenta = Cuenta & Mid(RS!codmacta, LCta + 1)
                        Else
                            Cuenta = RaizCuentasCancelacion
                        End If
                        CadenaAgrupacion = Linea & ",'" & Cuenta & "','RE" & Format(Codigo, "0000") & Anyo & "'," & FP.conhacli
                        CadenaAgrupacion = CadenaAgrupacion & ",'" & DevNombreSQL(Mid(Ampliacion & "Rem: " & Codigo & "-" & Anyo, 1, 30)) & "',NULL,@@@@@@"
                
                        CadenaAgrupacion = CadenaAgrupacion & ",NULL,'" & CtaCancelacion & "','CONTAB',0)"
                        'Luego reemplazare @@@@@@ por el importe total
                    End If

                End If
                
                'Neuvo dato para la ampliacion en la contabilizacion
                Select Case FP.amphacli
                Case 2, 4
                    'La opcion Contrapartida BANCO NO vale ahora, pq no hay apunte a banco
                    Ampliacion = Ampliacion & Format(RS!FecVenci, "dd/mm/yyyy")
                    
                Case Else
                   If FP.amphacli = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
                   Ampliacion = Ampliacion & RS!NUmSerie & "/" & RS!codfaccl
                End Select
                
                
                Cuenta = RaizCuentasCancelacion
                If Not CuentaUnica Then Cuenta = Cuenta & Mid(RS!codmacta, LCta + 1)
                
            
            
             
                'Cuenta
                SQL = Linea & ",'" & Cuenta & "','" & Format(RS!codfaccl, "000000000") & "'," & FP.conhacli
                SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',NULL,"
                If TipoRemesa = 1 Then
                    ImporteVto = RS!ImpVenci
                    
                Else
                    'Talones pagares, podria ser que no pagara todo.
                    'El rs esta con left join sobre slirecdoc
                    'Si es NULL vto significa que esta en scobro y no en slirecdoc. Algo esta mal.
                    If IsNull(RS!Vto) Then
                        If MsgBox("Vencimiento no existe en rececpcion documentos. ¿Continuar?", vbYesNo) = vbNo Then
                            'esto generara un error
                            SQL = RS!Vto
                        Else
                            ImporteVto = RS!impcobro
                        End If
                    Else
                        ImporteVto = RS!Importe
                    End If
                End If
                
                Importe = Importe + ImporteVto
                ImporteDocumento = ImporteDocumento + ImporteVto
                
                'Importe VA alhaber de
                SQL = SQL & TransformaComasPuntos(CStr(ImporteVto)) & ",NULL,"
            
                'Contra partida
                SQL = SQL & "'" & CtaCancelacion & "','CONTAB',0)"
                SQL = AmpRemesa & SQL
                If Not EstaAgrupandoVtos Then
                    If Not Ejecuta(SQL) Then Exit Function
                End If
            Else
                ImporteVto = RS!Importe
            End If
            Linea = Linea + 1
            

            
            'Me cargo el efecto y si tuviera devoluciones
            'Para talones/pagares podria darse el caso que NO todo el importe es
            'es el que ha sido pagado. Entonces procederemos de otra forma
            BorrarEfecto = True
            If TipoRemesa <> 1 Then
                'TALONES REMESAS
                
                'Compruebo que el total impcobrado es el total +gastos
                
                If RS!ImpVenci + DBLet(RS!Gastos, "N") > ImporteVto Then BorrarEfecto = False
                    
                    
            End If
            'Comun a borrar y updatear
            SQL = " WHERE numserie='" & RS!NUmSerie & "' AND codfaccl = " & RS!codfaccl
            SQL = SQL & " AND fecfaccl = '" & Format(RS!fecfaccl, FormatoFecha) & "' AND numorden = " & RS!numorden
            If BorrarEfecto Then
                
                Conn.Execute "DELETE FROM scobro " & SQL
                Conn.Execute "DELETE FROM sefecdev " & SQL
                Debug.Print ""
            Else
                'Se trata de actualizar
                'Vamos a quitar la marca de remesado. En importe pondremos el importe que habia y en el campo observaciones
                'indicaremos que ya ha sido remesado y por cuanto importe
                
                'Campo obs..servaciones guardara los datos antiguos
                Ampliacion = "Vto: " & Format(RS!ImpVenci, FormatoImporte)
                If DBLet(RS!Gastos, "N") > 0 Then Ampliacion = Ampliacion & "/ Gastos " & Format(RS!Gastos, FormatoImporte)
                'Fecha del ultimo cobro
                Ampliacion = Ampliacion & "   Ultimo cobro: " & Format(RS!fecultco, "dd/mm/yyyy") & "  " & Format(RS!impcobro, FormatoImporte)
                'Lo meto en la observacion
                Ampliacion = "obs = '" & Ampliacion & "',"
                
                
                'Agosto 2009
                'Como el vto esta en slirecepdoc NO hace falta cponer esto
                'Ampliacion = Ampliacion & "fecultco = NULL, impcobro=NULL,"
                Ampliacion = Ampliacion & "codrem=NULL,Tiporem=NULL,Anyorem=NULL,siturem=NULL"
                
                
                'Los gastos los pondre a null
                Ampliacion = Ampliacion & ",gastos=NULL"
                'ImporteVto = Rs!impvenci + DBLet(Rs!Gastos, "N") - Rs!impcobro
                'Ampliacion = Ampliacion & ",impvenci = " & TransformaComasPuntos(CStr(ImporteVto))
                
                'Raferencia talon/pagare tb
                Ampliacion = Ampliacion & ",reftalonpag=NULL,recedocu=0"
                
                ParaLineasDocumentosRecibidos = SQL
                
                SQL = "UPDATE scobro SET " & Ampliacion & SQL
                Ampliacion = ""
                Conn.Execute SQL
                
                
                        
                
                
                
                
                'Realmente para tipo=1 NO deberia llegar aquin
                If TipoRemesa <> 1 Then

         
                        SQL = "UPDATE slirecepdoc SET numserie="" """ & ParaLineasDocumentosRecibidos
                        SQL = Replace(SQL, "numorden", "numvenci")
                        SQL = Replace(SQL, "codfaccl", "numfaccl")
                        If Not Ejecuta(SQL) Then
                            SQL = "Error actualizando tabla lineas documentos recibidos"
                            MsgBox SQL, vbExclamation
                        End If
                End If
                
            End If  'Borrar efecto
        End If  'de             If CuentaPuente Then
            
        RS.MoveNext
    Wend
    RS.Close

    
    'Comprobamos que el importe del talon es el correcto
        If CuentaPuente And vId > 0 Then
            'Conseguimos importe documento
            Cuenta = "codmacta"
            Ampliacion = DevuelveDesdeBD("importe", "scarecepdoc", "codigo", CStr(vId), "N", Cuenta)
            ImporteVto = CCur(Ampliacion)
            
            If Not CuentaUnica Then
                Cuenta = RaizCuentasCancelacion & Mid(Cuenta, LCta + 1)
            Else
                Cuenta = RaizCuentasCancelacion
            End If
            'Comprobamos con el importe parcial.
            If ImporteVto <> ImporteDocumento Then
                ImporteVto = ImporteDocumento - ImporteVto
                
                
                Importe = Importe - ImporteVto
                'Ha habido difernecias
                'Y si no agrupamos
                If EstaAgrupandoVtos Then
                    'ya hemos cambiado el importe para los dos apuntes que
                    'quedan abajo uno ajustando importe
                    
                Else
                    'Creo una linea de ap
                    SQL = Linea & ",'" & Cuenta & "','" & Format(vId, "000000000") & "',"
                    
                    
                    If ImporteVto > 0 Then
                        'al debe o al haber
                        SQL = SQL & FP.condecli & ",'" & DevNombreSQL(FP.CadenaAuxiliar & " Elim. " & vId) & "'," & TransformaComasPuntos(CStr(ImporteVto)) & ",NULL,"
                    Else
                        SQL = SQL & FP.conhacli & ",'" & DevNombreSQL(FP.descformapago & " Elim." & vId) & "',NULL," & TransformaComasPuntos(Abs(ImporteVto)) & ","
                    End If
                    'Contra partida
                    SQL = SQL & "NULL,'" & CtaCancelacion & "','CONTAB',0)"
                    SQL = AmpRemesa & SQL
                    Ejecuta SQL
                    Linea = Linea + 1
                
                    
                End If 'EstaAgrupandoVtos
            End If  'ImporteVto <> ImporteDocumento
        End If  ' vId > 0

    
    
        If EstaAgrupandoVtos Then
            If CadenaAgrupacion <> "" Then
                'OK inserto el total
                Ampliacion = TransformaComasPuntos(CStr(Importe))
                SQL = Replace(CadenaAgrupacion, "@@@@@@", Ampliacion)
                Conn.Execute AmpRemesa & SQL
                Linea = 2 'La uno es
            End If
        End If
    
        If Linea > 1 Then
            'Hago el contrapunte
            If CuentaPuente Then
                Ampliacion = FP.descformapago & " Re: " & Codigo & " - " & Anyo
                SQL = "RE" & Format(Codigo, "0000") & Format(Anyo, "0000")
                SQL = Linea & ",'" & CtaCancelacion & "','" & SQL & "'," & FP.conhacli
                SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
            
            
                'Importe al DEBE
                SQL = SQL & TransformaComasPuntos(CStr(Importe)) & ",NULL,NULL,"
        
                'Contra partida
                If CuentaUnica Then
                    Cuenta = "'" & RaizCuentasCancelacion & "'"
                Else
                    If Len(Cuenta) = vEmpresa.DigitosUltimoNivel And IsNumeric(Cuenta) Then
                        'Dejo cuenta como esta
                        Cuenta = "'" & Cuenta & "'"
                    Else
                        Cuenta = "NULL"
                    End If
                End If
                SQL = SQL & Cuenta & ",'CONTAB',0)"
                SQL = AmpRemesa & SQL
            
            
  
                Conn.Execute SQL
            Else
                Linea = Linea + 1 'Para que despues no de distinto el numero de efectos eliminados
            End If
        Else
            
            
            'Muestro un mesaje diciendo que Ningun vto ha sido eliminado. No deberia pasar pero por si acaso
            'compruebo que tenga vtos
            If NumRegElim > 0 Then MsgBox "No se ha podido eliminar ningun vencimiento de la remesa " & Codigo & " / " & Anyo, vbInformation
            RemesasEliminarVtosTalonesPagares = 0
            
            If CuentaPuente Then
                SQL = "DELETE FROM cabapu  WHERE numdiari =" & FP.diaricli
                SQL = SQL & " and fechaent = '" & Format(FechaAbono, FormatoFecha) & "' and numasien = " & Mc.Contador
                Conn.Execute SQL
            End If
        End If


    'Si la hemos borrado toda, o no....
    Linea = Linea - 1 'Empieza en uno, luego el total vtos eliminados es linea-1
                      'En numregelim tengo los vtos totales de la remesa
                      'Si queda alguno o no, haremos unas cosas u otras
    If NumRegElim > Linea Then
        AmpRemesa = "Y"
    
        'QUEDA ALGUNO
        SQL = "UPDATE scobro SET"
        SQL = SQL & " siturem= 'Y'"
        SQL = SQL & " WHERE codrem=" & Codigo
        SQL = SQL & " and anyorem=" & Anyo
           
        Conn.Execute SQL
    Else
        AmpRemesa = "Z"  'TOdos eliminados
    End If
    SQL = "UPDATE remesas SET"
    SQL = SQL & " situacion= '" & AmpRemesa
    SQL = SQL & "' WHERE codigo=" & Codigo
    SQL = SQL & " and anyo=" & Anyo
    Conn.Execute SQL
    
    
    '-----------------------------------------------
    '-----------------------------------------------
    'Por ultimo.
    ' Si tiene la opcion de eliminar en documentos recibidos
    If vParamT.EliminaRecibidosRiesgo Then
        'QUito el preimer PIPE
        EliminaEnRecepcionDocumentos = Mid(EliminaEnRecepcionDocumentos, 2)
        While EliminaEnRecepcionDocumentos <> ""
            Linea = InStr(1, EliminaEnRecepcionDocumentos, "|")
            If Linea = 0 Then
                EliminaEnRecepcionDocumentos = ""
            Else
                AmpRemesa = Mid(EliminaEnRecepcionDocumentos, 1, Linea - 1)
                EliminaEnRecepcionDocumentos = Mid(EliminaEnRecepcionDocumentos, Linea + 1)
                'Ahora veo si toooodos los vtos asociados a ese ID estan eliminados
                SQL = "select scobro.codfaccl ,l.importe,l.numserie vto,id from   slirecepdoc l left join  scobro  on l.numserie=scobro.numserie and"
                SQL = SQL & " l.numfaccl=scobro.codfaccl and   l.fecfaccl=scobro.fecfaccl and l.numvenci=scobro.numorden"
                SQL = SQL & " where id = " & AmpRemesa
                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                vId = 0
                While vId = 0
                    If RS.EOF Then
                        vId = 1
                    Else
                        If Not IsNull(RS!codfaccl) Then
                            'Tiene un vencimento sin eliminar entodavia
                            vId = 1
                            AmpRemesa = ""
                        Else
                            RS.MoveNext
                        End If
                    End If
                Wend
                RS.Close
                If AmpRemesa <> "" Then
                    'OK. Borro
                    'Lineas
                    SQL = "Delete  from slirecepdoc WHERE id =" & AmpRemesa
                    EjecutarSQL SQL
                    'Cabeceras
                    SQL = "Delete  from scarecepdoc WHERE codigo =" & AmpRemesa
                    EjecutarSQL SQL
                    
                    
                End If
            End If
        Wend
    End If
    
    
    
    
    
    'Insertamos para pasar a hco
    If CuentaPuente Then InsertaTmpActualizar Mc.Contador, FP.diaricli, FechaAbono
    
    'Todo OK
    If NumRegElim > 0 Then
        RemesasEliminarVtosTalonesPagares = 1
        'Para que no actualice el apunte , ya que no se ha creado
        If Not CuentaPuente Then RemesasEliminarVtosTalonesPagares = 0
    End If
    
ERemesa_Elivto:
    If Err.Number <> 0 Then
        
        MuestraError Err.Number, Err.Description
        RemesasEliminarVtosTalonesPagares = 2
    End If
    Set RS = Nothing
    Set Mc = Nothing

End Function



'*********************************************************************************
'*********************************************************************************
'   Eliminar TALON PAGARE contabilizado (contra ctas puente)
'
'
'DiarioConcepto:Llevara el diario y los conceptos al debe y al haber. NO cojera los de la stipforpa, si no de una window anterior
'              El cuarto pipe que lleva es si agrrupa en la cuenta puente
'                   es decir, en lugar de 43.1 contra 431.1
'                                         43.2 contra 431.1
'                           hacemos   43.1 y 43.2   contra la suma en 431.1
' Septiembre 2009
'           El quinto y sexto pipe llevaran, si necesario, cta dodne poner el benefic po perd del talon y si requiere cc
Public Function EliminarCancelacionTALONPAGARE(Talones As Boolean, IdRecepcion As Integer, FechaAbono As Date, DiarioConcepto As String) As Boolean
'Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
'Dim Gastos As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim RS As ADODB.Recordset
Dim AmpRemesa As String
Dim CtaCancelacion As String
Dim Cuenta As String
Dim RaizCuentasCancelacion As String
Dim CuentaUnica As Boolean
Dim LCta As Integer
Dim Importeauxiliar As Currency
Dim AgrupaVtosPuente As Boolean
Dim CadenaAgrupaVtoPuente As String
Dim aux2 As String
Dim RequiereCCDiferencia As Boolean

Dim Obs As String
Dim TipForpa As String


    On Error GoTo ERemesa_CancelacionCliente3
    EliminarCancelacionTALONPAGARE = False
    

    If Talones Then
        'Sobre talones
        Cuenta = "taloncta"
    Else
        Cuenta = "pagarecta"
    End If
    RaizCuentasCancelacion = DevuelveDesdeBD(Cuenta, "paramtesor", "codigo", "1", "N")
    If RaizCuentasCancelacion = "" Then
        MsgBox "Error grave en configuración de  parámetros de tesorería. Falta cuenta cancelación", vbExclamation
        Exit Function
    End If
    
    LCta = Len(RaizCuentasCancelacion)
    CuentaUnica = LCta = vEmpresa.DigitosUltimoNivel
    
    
    'Comprobacion.  Para todos los efectos de la 43.... se cancelan con la 4310....
    '
    'Tendre que ver que existen estas cuentas
    Set RS = New ADODB.Recordset
    
    SQL = "DELETE FROM tmpcierre1 where codusu = " & vUsu.Codigo
    Conn.Execute SQL
    
        
    If Not CuentaUnica Then
        'Cancela contra subcuentas de cliente
        

        Ampliacion = ",CONCAT('" & RaizCuentasCancelacion & "',SUBSTRING(codmacta," & LCta + 1 & ")" & ")"
            
        SQL = "Select " & vUsu.Codigo & Ampliacion & " from talones where codigo=" & IdRecepcion
        SQL = SQL & " GROUP BY codmacta"
        'INSERT
        SQL = "INSERT INTO tmpcierre1(codusu,cta) " & SQL
        Conn.Execute SQL
        
        'Ahora monto el select para ver que cuentas 430 no tienen la 4310
        SQL = "Select cta,codmacta from tmpcierre1 left join cuentas on tmpcierre1.cta=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
        SQL = SQL & " HAVING codmacta is null"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        Linea = 0
        While Not RS.EOF
            Linea = Linea + 1
            SQL = SQL & RS!Cta & "     "
            If Linea = 5 Then
                SQL = SQL & vbCrLf
                Linea = 0
            End If
            RS.MoveNext
        Wend
        RS.Close
        
        If SQL <> "" Then
            
            AmpRemesa = "CANCELACION contab"
            
            SQL = "Cuentas " & AmpRemesa & ".  No existen las cuentas: " & vbCrLf & String(90, "-") & vbCrLf & SQL
            SQL = SQL & vbCrLf & "¿Desea crearlas?"
            Linea = 1
            If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
                'Ha dicho que si desea crearlas
                
                Ampliacion = "CONCAT('" & RaizCuentasCancelacion & "',SUBSTRING(codmacta," & LCta + 1 & ")) "
            
                SQL = "Select codmacta," & Ampliacion & " from talones where codigo=" & IdRecepcion
                SQL = SQL & " and " & Ampliacion & " in "
                SQL = SQL & "(Select cta from tmpcierre1 left join cuentas on tmpcierre1.cta=cuentas.codmacta WHERE codusu = " & vUsu.Codigo
                SQL = SQL & " AND codmacta is null)"
                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not RS.EOF
                
                     SQL = "INSERT  IGNORE INTO cuentas(codmacta,nommacta ,apudirec,model347,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos) SELECT '"
                                ' CUenta puente
                     SQL = SQL & RS.Fields(1) & "',nommacta ,'S',0,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos from cuentas where codmacta = '"
                                'Cuenta en la scbro (codmacta)
                     SQL = SQL & RS.Fields(0) & "'"
                     Conn.Execute SQL
                     RS.MoveNext
                     
                Wend
                RS.Close
                Linea = 0
            End If
            If Linea = 1 Then GoTo ERemesa_CancelacionCliente3
        End If
        
    Else
        'Cancela contra UNA unica cuenta todos los vencimientos
        SQL = DevuelveDesdeBD("codmacta", "cuentas", "codmacta", RaizCuentasCancelacion, "T")
        If SQL = "" Then
            MsgBox "No existe la cuenta puente: " & RaizCuentasCancelacion, vbExclamation
            GoTo ERemesa_CancelacionCliente3
        End If
    End If

    
    'La forma de pago
    Set vCP = New Ctipoformapago
    If Talones Then
        SQL = CStr(vbTalon)
        Ampliacion = "Talones"
    Else
        SQL = CStr(vbPagare)
        Ampliacion = "Pagarés"
    End If
    If vCP.Leer(CInt(SQL)) = 1 Then GoTo ERemesa_CancelacionCliente3
    'Ahora fijo los valores que me ha dado el
    vCP.diaricli = RecuperaValor(DiarioConcepto, 1)
    'En la contabilizacion
    'vCP.condecli = RecuperaValor(DiarioConcepto, 2)
    'vCP.conhacli = RecuperaValor(DiarioConcepto, 3)
    'En la eliminacion
    vCP.conhacli = RecuperaValor(DiarioConcepto, 2)
    vCP.condecli = RecuperaValor(DiarioConcepto, 3)
    AgrupaVtosPuente = RecuperaValor(DiarioConcepto, 4) = 1
 
 
 
    Set Mc = New Contadores
    
    
    If Mc.ConseguirContador("0", FechaAbono <= vParam.fechafin, True) = 1 Then Exit Function
    
    
    
    'Insertamos la cabera
    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion ) VALUES ("
    SQL = SQL & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ", '"
    SQL = SQL & "Eliminar recepcion documentos contabilizada(cancelada )"

    SQL = SQL & " NºRecepcion: " & IdRecepcion & "   " & Ampliacion & vbCrLf
    SQL = SQL & "Generado el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & vUsu.Nombre & "',"
    
    Obs = "ARICONTA 6: Eliminar recepción documentos contabilizada: " & vbCrLf & " NºRecepcion: " & IdRecepcion & "   " & Ampliacion & vbCrLf
    SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(Obs, "T") & ");"
    
    If Not Ejecuta(SQL) Then Exit Function
    
    Linea = 1
    Importe = 0
    'Gastos = 0
    
    vCP.descformapago = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)  'DEBE
    vCP.CadenaAuxiliar = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)   'DEBE
    
    
    SQL = "select cobros.*,l.importe,l.codigo, c.numeroref reftalonpag, c.banco from  (talones c inner join  talones_facturas l on c.codigo = l.codigo)  left join  cobros  on l.numserie=cobros.numserie and"
    SQL = SQL & " l.numfactu=cobros.numfactu and   l.fecfactu=cobros.fecfactu and l.numorden=cobros.numorden"
    SQL = SQL & " WHERE l.codigo= " & IdRecepcion
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Trozo comun
    AmpRemesa = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    AmpRemesa = AmpRemesa & "codmacta, numdocum, codconce, ampconce,timporteD,"
    AmpRemesa = AmpRemesa & " timporteH, codccost, ctacontr, idcontab, punteada, "
    AmpRemesa = AmpRemesa & " numserie, numfaccl, fecfactu, numorden, tipforpa, reftalonpag, bancotalonpag) "
    AmpRemesa = AmpRemesa & "VALUES (" & vCP.diaricli & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & ","
    
    CadenaAgrupaVtoPuente = ""

    While Not RS.EOF
        Ampliacion = RS!NUmSerie   'SI DA ERROR ES QUE NO EXISTE mediante el left join
        
        
        'Neuvo dato para la ampliacion en la contabilizacion
        Ampliacion = " "
        Select Case vCP.ampdecli
        Case 2, 4
            'La opcion Contrapartida BANCO NO vale ahora, pq no hay apunte a banco
            Ampliacion = DBLet(RS!reftalonpag, "T")
            If Ampliacion = "" Then Ampliacion = Ampliacion & Format(RS!FecVenci, "dd/mm/yyyy")
        Case 5
            Ampliacion = DBLet(RS!reftalonpag, "T")
            If Ampliacion = "" Then
                Ampliacion = RS!NUmSerie & "/" & RS!NumFactu
            Else
                Ampliacion = "Doc: " & Ampliacion
            End If
        Case Else
           If vCP.ampdecli = 1 Then Ampliacion = vCP.siglas & " "
           Ampliacion = Ampliacion & RS!NUmSerie & "/" & RS!NumFactu
        End Select
        If Mid(Ampliacion, 1, 1) <> " " Then Ampliacion = " " & Ampliacion
        
        'Cancelacion
        If CuentaUnica Then
            Cuenta = RaizCuentasCancelacion
        Else
            Cuenta = RaizCuentasCancelacion & Mid(RS!codmacta, LCta + 1)
            
        End If
        CtaCancelacion = RS!codmacta
    
        
        'Si dice que agrupamos vto entonces NO
        If AgrupaVtosPuente Then
            If CadenaAgrupaVtoPuente = "" Then
                'Para insertarlo al final del proceso
                'Antes de ejecutar el sql(al final) substituiremos, la cadena
                ' la cadena ### por el importe total
                
                SQL = "1,'" & Cuenta & "','Nº" & Format(IdRecepcion, "0000000") & "'," & vCP.condecli
                
                SQL = SQL & ",'" & DevNombreSQL(Mid(vCP.descformapago & " " & DBLet(RS!reftalonpag, "T"), 1, 30)) & "',"
                'Importe al HABER.
                SQL = SQL & "NULL,###,NULL,"
                'Contra partida
                SQL = SQL & "'" & CtaCancelacion & "','CONTAB',0,"
                SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
                
                CadenaAgrupaVtoPuente = AmpRemesa & SQL
            End If
        End If
        
        
        'Crearemos el apnte y la contrapartida
        ' Es decir
        '   4310  contra 430
        '   430  contr   4310
        'Lineas de apuntes .
        
         
        'Cuenta
        SQL = Linea & ",'" & Cuenta & "','" & Format(RS!NumFactu, "000000000") & "'," & vCP.condecli
        SQL = SQL & ",'" & DevNombreSQL(Mid(vCP.descformapago & Ampliacion, 1, 30)) & "',"
        
        
        
        Importe = Importe + RS!Importe
        'Gastos = Gastos + DBLet(Rs!Gastos, "N")
        
        
        'Importe VA alhaber del cliente, al debe de la cancelacion
        SQL = SQL & "NULL," & TransformaComasPuntos(RS!Importe) & ",NULL,"
    
        'Contra partida
        SQL = SQL & "'" & CtaCancelacion & "','CONTAB',0,"
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
        
        
        
        SQL = AmpRemesa & SQL
        If Not AgrupaVtosPuente Then
            If Not Ejecuta(SQL) Then Exit Function
        End If
        Linea = Linea + 1 'Siempre suma mas uno
        
        
        'La contrapartida
        SQL = Linea & ",'" & CtaCancelacion & "','" & Format(RS!NumFactu, "000000000") & "'," & vCP.conhacli
        SQL = SQL & ",'" & DevNombreSQL(Mid(vCP.CadenaAuxiliar & Ampliacion, 1, 30)) & "',"
        
        
        '
        SQL = SQL & TransformaComasPuntos(RS!Importe) & ",NULL,NULL,"
    
        'Contra partida
        SQL = SQL & "'" & Cuenta & "','CONTAB',0,"
        TipForpa = DevuelveDesdeBD("tipforpa", "formapago", "codforpa", RS!codforpa, "N")
        
        SQL = SQL & DBSet(RS!NUmSerie, "T") & "," & DBSet(RS!NumFactu, "N") & "," & DBSet(RS!FecFactu, "F") & "," & DBSet(RS!numorden, "N") & ","
        SQL = SQL & DBSet(TipForpa, "T") & "," & DBSet(RS!reftalonpag, "T") & "," & DBSet(RS!Banco, "T") & ")"

        SQL = AmpRemesa & SQL
        
        If Not Ejecuta(SQL) Then Exit Function
        Linea = Linea + 1
            
        
        RS.MoveNext
    Wend
    RS.Close


    
    SQL = "Select importe,codmacta,numeroref from talones where codigo = " & IdRecepcion
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then Err.Raise 513, , "No se ha encontrado documento: " & IdRecepcion
    Importeauxiliar = RS!Importe
    Cuenta = RS!codmacta
    Ampliacion = DevNombreSQL(RS!numeroref)
    RS.Close


    If Importe <> Importeauxiliar Then
    
        CtaCancelacion = RecuperaValor(DiarioConcepto, 5)
        If CtaCancelacion = "" Then Err.Raise 513, , "Cuenta beneficios/pérdidas  NO espeficicada"
        
        'Hemos llegado a aqui.
        'Veremos si hay diferencia entre la suma de importe y el importe del documento.
        '
        Importeauxiliar = Importeauxiliar - Importe
        If Len(Ampliacion) > 10 Then Ampliacion = Right(Ampliacion, 10)
        
        SQL = Linea & ",'" & CtaCancelacion & "','Nº" & Format(IdRecepcion, "00000000") & "'," & vCP.conhacli
        
        'Ampliacion
        If Talones Then
            aux2 = " Tal nº: " & Ampliacion
        Else
            aux2 = " Pag. nº: " & Ampliacion
        End If
        SQL = SQL & ",'" & DevNombreSQL(Mid(vCP.descformapago & aux2, 1, 30)) & "',"

        
        If Importeauxiliar < 0 Then
            'NEgativo. Va en positivo al otro lado
            SQL = SQL & "NULL," & TransformaComasPuntos(Abs(Importeauxiliar)) & ","
        Else
            SQL = SQL & TransformaComasPuntos(CStr(Importeauxiliar)) & ",NULL,"
        End If
                
        'Centro de coste
        RequiereCCDiferencia = False
        If vParam.autocoste Then
            aux2 = Mid(CtaCancelacion, 1, 1)
            If aux2 = "6" Or aux2 = "7" Then RequiereCCDiferencia = True
        End If
        If RequiereCCDiferencia Then
            CtaCancelacion = UCase(RecuperaValor(DiarioConcepto, 6))
            If CtaCancelacion = "" Then Err.Raise 513, , "Centro de coste  NO espeficicado"
            CtaCancelacion = "'" & CtaCancelacion & "'"
        Else
             CtaCancelacion = "NULL"
        End If
        SQL = SQL & CtaCancelacion
        
        'Contra partida
        If CuentaUnica Then
            Cuenta = "'" & RaizCuentasCancelacion & "'"
        Else
            Cuenta = "NULL"
        End If
        
        
        SQL = SQL & "," & Cuenta & ",'CONTAB',0,"
        
        SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
        
        
        
        SQL = AmpRemesa & SQL
        
        If Not Ejecuta(SQL) Then Exit Function
        Linea = Linea + 1
        
        
        If AgrupaVtosPuente Then
            'Modificamos el importe final por si esta agrupando vencimientos
            Importe = Importeauxiliar + Importe
        Else
                'creamos la contrapartida para que  cuadre el asiento
            
                If Len(Ampliacion) > 10 Then Ampliacion = Right(Ampliacion, 10)
                
                SQL = Linea & "," & Cuenta & ",'Nº" & Format(IdRecepcion, "00000000") & "'," & vCP.conhacli
                
                 If Talones Then
                    aux2 = " Tal nº: " & Ampliacion
                Else
                    aux2 = " Pag. nº: " & Ampliacion
                End If
                SQL = SQL & ",'" & DevNombreSQL(Mid(vCP.CadenaAuxiliar & aux2, 1, 30)) & "',"
        
                
                If Importeauxiliar > 0 Then
                    'NEgativo. Va en positivo al otro lado
                    SQL = SQL & TransformaComasPuntos(CStr(Importeauxiliar)) & ",NULL,"
                Else
                    SQL = SQL & "NULL," & TransformaComasPuntos(Abs(Importeauxiliar)) & ","
                End If
                        
                'Centro de coste
                RequiereCCDiferencia = False
                If vParam.autocoste Then
                    aux2 = Mid(Cuenta, 2, 1)  'pq lleva una comilla
                    If aux2 = "6" Or aux2 = "7" Then RequiereCCDiferencia = True
                End If
                If RequiereCCDiferencia Then
                    CtaCancelacion = UCase(RecuperaValor(DiarioConcepto, 6))
                    If CtaCancelacion = "" Then Err.Raise 513, , "Centro de coste  NO espeficicado"
                    CtaCancelacion = "'" & CtaCancelacion & "'"
                Else
                     CtaCancelacion = "NULL"
                End If
                SQL = SQL & CtaCancelacion
                
                'Contra partida
                CtaCancelacion = RecuperaValor(DiarioConcepto, 5)
                SQL = SQL & ",'" & CtaCancelacion & "','CONTAB',0,"
                
                '###FALTA1
                
                SQL = AmpRemesa & SQL
                
                If Not Ejecuta(SQL) Then Exit Function
                Linea = Linea + 1
            End If
                
    End If
    
    If AgrupaVtosPuente Then
        'Tenmos que reemplazar
        'en CadenaAgrupaVtoPuente    ###:importe
        SQL = TransformaComasPuntos(CStr(Importe))
        SQL = Replace(CadenaAgrupaVtoPuente, "###", SQL)
        Conn.Execute SQL
    End If


    AmpRemesa = "F"    ' cancelada
    




    
    'Insertamos para pasar a hco
    InsertaTmpActualizar Mc.Contador, vCP.diaricli, FechaAbono
    
    'Todo OK
    EliminarCancelacionTALONPAGARE = True
    
    
ERemesa_CancelacionCliente3:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    
    End If
    Set RS = Nothing
    Set Mc = Nothing
    Set vCP = Nothing
End Function







'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'
'
'   Contabilizacion especial N19.
'   Genera tantos apuntes como fechas vto haya que sera la fecha del asie
'
'
'
'
'   Solo Recibo bancario, norma 19, si ctas puente
'
'------------------------------------------------------------------------
'------------------------------------------------------------------------
'------------------------------------------------------------------------


Public Function ContabNorma19PorFechaVto(Codigo As Integer, Anyo As Integer, CtaBanco As String) As Boolean
Dim Cuenta As String
Dim Mc As Contadores
Dim Linea As Integer
Dim Importe As Currency
Dim Gastos As Currency
Dim vCP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim RS As ADODB.Recordset
Dim AmpRemesa As String
'Dim CtaParametros As String
'Dim Cuenta As String
'
'
Dim ImpoAux As Currency


Dim ColFechas As Collection  'Cada una de las fechas de vencimiento de la remesa
Dim NF As Integer
Dim FecAsto As Date

    On Error GoTo ECon
    
    ContabNorma19PorFechaVto = False

    'La forma de pago
    Set vCP = New Ctipoformapago
    If vCP.Leer(vbTipoPagoRemesa) = 1 Then GoTo ECon
    
    Set RS = New ADODB.Recordset
    Set ColFechas = New Collection
    
    
    SQL = "Select fecvenci from cobros WHERE codrem=" & Codigo & " AND anyorem = " & Anyo & " GROUP BY fecvenci ORDER By fecvenci"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        SQL = RS.Fields(0)
        ColFechas.Add SQL
        RS.MoveNext
    Wend
    RS.Close
    If ColFechas.Count = 0 Then Err.Raise 513, "No hay vencimientos(n19)"
    
    
    For NF = 1 To ColFechas.Count
        FecAsto = CDate(ColFechas.Item(NF))
        
        Set Mc = New Contadores
    
    
        If Mc.ConseguirContador("0", FecAsto <= vParam.fechafin, True) = 1 Then Exit Function
    
    
        'Insertamos la cabera
        SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion,desdeaplicacion) VALUES ("
        SQL = SQL & vCP.diaricli & ",'" & Format(FecAsto, FormatoFecha) & "'," & Mc.Contador
        SQL = SQL & ", '"
        SQL = SQL & "Abono remesa: " & Codigo & " / " & Anyo & "       N19" & vbCrLf
        SQL = SQL & "Proceso: " & NF & " / " & ColFechas.Count & vbCrLf & "',"
        'SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy") & " por " & vUsu.Nombre & "');"
        SQL = SQL & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Abono remesa')"
        If Not Ejecuta(SQL) Then Exit Function
        
        Linea = 1
        Importe = 0
        Gastos = 0
        
        'La ampliacion para el banco
        AmpRemesa = ""
        SQL = "Select * from remesas WHERE codigo=" & Codigo & " AND anyo = " & Anyo

        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        'NO puede ser EOF
        
        
        If Not IsNull(RS!Descripcion) Then AmpRemesa = RS!Descripcion
        
        
        If AmpRemesa = "" Then
            AmpRemesa = " Remesa: " & Codigo & "/" & Anyo
        Else
            AmpRemesa = " " & AmpRemesa
        End If
        
        RS.Close
        
        'AHORA Febrero 2009
        '572 contra  5208  Efectos descontados
        '-------------------------------------
        SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
        SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
        SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
        SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FecAsto, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"
    
        Gastos = 0
        
        Importe = 0
        SQL = "Select * from cobros WHERE codrem=" & Codigo & " AND anyorem = " & Anyo
        'y por vencimiento
        SQL = SQL & " AND fecvenci = '" & Format(FecAsto, FormatoFecha) & "'"
        RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RS.EOF
            'Banco contra cliente
            'La linea del banco
            SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
            SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
            SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada, numserie,numfaccl,fecfactu,numorden,tipforpa, tiporem,codrem,anyorem) "
            SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FecAsto, FormatoFecha) & "'," & Mc.Contador & "," & Linea & ",'"
        
            'Cuenta
            SQL = SQL & RS!codmacta & "','" & RS!NUmSerie & Format(RS!NumFactu, "0000000") & "'," & vCP.conhacli
            
            
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
            Ampliacion = Ampliacion & " "
                   
            'Neuvo dato para la ampliacion en la contabilizacion
            Select Case vCP.amphacli
            Case 2
               Ampliacion = Ampliacion & Format(RS!FecVenci, "dd/mm/yyyy")
            Case 4
                'Contrapartida BANCO
                Cuenta = RecuperaValor(CtaBanco, 1)
                Cuenta = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", Cuenta, "T")
                Ampliacion = Ampliacion & AmpRemesa
            Case Else
               If vCP.amphacli = 1 Then Ampliacion = Ampliacion & vCP.siglas & " "
               Ampliacion = Ampliacion & RS!NUmSerie & Format(RS!NumFactu, "0000000")
            End Select
            SQL = SQL & ",'" & DevNombreSQL(Mid(Ampliacion, 1, 30)) & "',"
            
            Importe = Importe + RS!ImpVenci
                
            Gastos = Gastos + DBLet(RS!Gastos, "N")
            
            ' timporteH, codccost, ctacontr, idcontab, punteada
            'Importe
            SQL = SQL & "NULL," & TransformaComasPuntos(RS!ImpVenci) & ",NULL,"
        
            If vCP.ctrdecli = 1 Then
                SQL = SQL & "'" & RecuperaValor(CtaBanco, 1) & "',"
            Else
                SQL = SQL & "NULL,"
            End If
            SQL = SQL & "'COBROS',0,"
            
            'los datos de la factura (solo en el apunte del cliente)
            Dim TipForpa As Byte
            TipForpa = DevuelveDesdeBD("tipforpa", "formapago", "codforpa", RS!codforpa, "N")
            
            SQL = SQL & DBSet(RS!NUmSerie, "T") & "," & DBSet(RS!NumFactu, "N") & "," & DBSet(RS!FecFactu, "F") & "," & DBSet(RS!numorden, "N") & "," & DBSet(TipForpa, "N") & ","
            SQL = SQL & "1," & DBSet(Codigo, "N") & "," & DBSet(Anyo, "N") & ")"
                
            
            If Not Ejecuta(SQL) Then Exit Function
            
            Linea = Linea + 1
            RS.MoveNext
        Wend
        RS.Close
        
        
        'La linea del banco
        SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
        SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
        SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) "
        SQL = SQL & "VALUES (" & vCP.diaricli & ",'" & Format(FecAsto, FormatoFecha) & "'," & Mc.Contador & ","
    
        
        'Gastos de los recibos.
        'Si tiene alguno de los efectos remesados gastos
        If Gastos > 0 Then
            Linea = Linea + 1
            Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.conhacli)
            Ampliacion = "RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.conhacli & ",'" & Ampliacion & " " & Codigo & "/" & Anyo & "'"
    
    
    
            Ampliacion = Linea & ",'" & RecuperaValor(CtaBanco, 2) & "','" & Ampliacion & ",NULL,"
            Ampliacion = Ampliacion & TransformaComasPuntos(CStr(Gastos)) & ","
    
          
            Ampliacion = Ampliacion & "NULL"
           
            Ampliacion = Ampliacion & ",NULL,'COBROS',0)"
            Ampliacion = SQL & Ampliacion
            If Not Ejecuta(Ampliacion) Then Exit Function
            Linea = Linea + 1
        End If
        
      
       
        
        ImpoAux = Importe + Gastos
        
        
        Ampliacion = DevuelveDesdeBD("nomconce", "conceptos", "codconce", vCP.condecli)
        Ampliacion = Ampliacion & AmpRemesa
        Ampliacion = Linea & ",'" & RecuperaValor(CtaBanco, 1) & "','RE" & Format(Codigo, "0000") & Format(Anyo, "0000") & "'," & vCP.condecli & ",'" & Ampliacion & "',"
        Ampliacion = Ampliacion & TransformaComasPuntos(CStr(ImpoAux)) & ",NULL,NULL,"
        
        If vCP.ctrdecli = 0 Then
            Ampliacion = Ampliacion & "NULL"
        Else
    
            Ampliacion = Ampliacion & "NULL"
    
        End If
        Ampliacion = Ampliacion & ",'COBROS',0)"
        Ampliacion = SQL & Ampliacion
        If Not Ejecuta(Ampliacion) Then Exit Function
        
        'Insertamos para pasar a hco
        InsertaTmpActualizar Mc.Contador, vCP.diaricli, FecAsto
        
        
        
        
        'Estamos recorriendo por fechas
        Set Mc = Nothing
   Next NF
        
        
    'AHora actualizamos los efectos.
    SQL = "UPDATE cobros SET"
    SQL = SQL & " siturem= 'Q'"
    SQL = SQL & ", situacion = 1 "
    SQL = SQL & " WHERE codrem=" & Codigo
    SQL = SQL & " and anyorem=" & Anyo
    Conn.Execute SQL

    'Todo OK
    ContabNorma19PorFechaVto = True
ECon:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    
    End If
    Set RS = Nothing
    Set Mc = Nothing
    Set vCP = Nothing
    Set ColFechas = Nothing
End Function

