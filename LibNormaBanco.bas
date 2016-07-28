Attribute VB_Name = "LinNormasBanco"
Option Explicit

    Dim NF As Integer
    Dim Registro As String
    Dim SQL As String

    Dim AuxD As String
    Private NumeroTransferencia As Integer


Public Function FrmtStr(Campo As String, Longitud As Integer) As String
    FrmtStr = Mid(Trim(Campo) & Space(Longitud), 1, Longitud)
End Function


'DATOSEXTRA  :
' 1: SUFIJOEM
' 2: TEXTO ORDENANTE
' Nuevo parametro:  Si el banco emite o no  (BancoEmiteDocumento)

'MODIFICACION 20 JUNIO 2012
'------------------------------
'  Si llevamos: vParam.Norma19xFechaVto presentara un fichero con varios ordenantes
' ENE 2014.
'  SEPA. Campo 17. Identifacador deudor. Si grabo BIC o CIF para las EMPRESAS. Particulares siempre NIF

'OCT 2015
'   Si lleva F.Cobro significa que van todos a esa fecha. Si es "" es que es fec vencimientos
Public Function GrabarDisketteNorma19(NomFichero As String, Remesa As String, FecPre As String, DatosExtra As String, TipoReferenciaCliente As Byte, FecCobro2 As String, BancoEmiteDocumento As Boolean, SepaEmpresasGraboNIF As Boolean, N19_15 As Boolean, FormatoXML As Boolean, esAnticipoCredito As Boolean) As Boolean

    
    If vParamT.NuevasNormasSEPA Then
                
        'GrabarDisketteNorma19 = GrabarDisketteNorma19SEPA(NomFichero, Remesa, FecPre, DatosExtra, TipoReferenciaCliente, FecCobro, BancoEmiteDocumento)
        GrabarDisketteNorma19 = GrabarFicheroNorma19SEPA(NomFichero, Remesa, FecPre, TipoReferenciaCliente, RecuperaValor(DatosExtra, 1), FecCobro2, SepaEmpresasGraboNIF, N19_15, FormatoXML, esAnticipoCredito)
    End If
End Function














Private Function HayKImprimirOpcionales() As Boolean
Dim I As Integer
Dim C As String

    On Error GoTo EImprimirOpcionales
    HayKImprimirOpcionales = False
    
    'Compruebo los cuatro primeros
    I = 0

    If Not IsNull(miRsAux.Fields!text41csb) Then I = I + 1
    If Not IsNull(miRsAux.Fields!text42csb) Then I = I + 1
    If Not IsNull(miRsAux.Fields!text43csb) Then I = I + 1
        
    If I > 0 Then HayKImprimirOpcionales = True
        
    

    

    Exit Function
EImprimirOpcionales:
    Err.Clear



End Function




Private Function ImprimeOpcionales(N19 As Boolean, Valores As String, Registro As Integer, ByRef ValorEnOpcionalesVar As Boolean) As String
Dim C As String
Dim J As Integer
Dim N As Integer
    ImprimeOpcionales = ""
    ValorEnOpcionalesVar = False
    If N19 Then
        ImprimeOpcionales = "56" & CStr(80 + Registro)
    End If
    ImprimeOpcionales = ImprimeOpcionales & Valores
    N = 0
    For J = 1 To 3
        C = "text" & (Registro + 3) & CStr(J) & "csb"
        C = DBLet(miRsAux.Fields(C), "T")
        If C <> "" Then N = N + 1
        C = FrmtStr(C, 40)
        ImprimeOpcionales = ImprimeOpcionales & C
    Next J
    ImprimeOpcionales = Mid(ImprimeOpcionales & Space(60), 1, 162)
    ValorEnOpcionalesVar = N > 0
End Function





Private Function comprobarCuentasBancariasRecibos(Remesa As String) As Boolean
Dim CC As String
On Error GoTo EcomprobarCuentasBancariasRecibos

    comprobarCuentasBancariasRecibos = False

    SQL = "select * from cobros where codrem = " & RecuperaValor(Remesa, 1)
    SQL = SQL & " AND anyorem=" & RecuperaValor(Remesa, 2)
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    NF = 0
    While Not miRsAux.EOF

        If DBLet(miRsAux!IBAN, "T") = "" Or Len(DBLet(miRsAux!IBAN, "T")) <> 24 Then
            SQL = ""
        Else
            SQL = "D"
        End If

    
        If SQL = "" Then
             Registro = Registro & miRsAux!codmacta & " - " & miRsAux!NUmSerie & "/" & miRsAux!NumFactu & "-" & miRsAux!numorden
             If NF < 2 Then
                Registro = Registro & "         "
                NF = NF + 1
             Else
                Registro = Registro & vbCrLf
                NF = 0
            End If
    
        End If
    
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Registro <> "" Then
        SQL = "Los siguientes vencimientos no tienen la cuenta bancaria con todos los datos." & vbCrLf & Registro
        MsgBox SQL, vbExclamation
        Exit Function
    End If
    
    
    'Si llega aqui es que todos tienen DATOS
    SQL = "select iban from cobros where codrem = " & RecuperaValor(Remesa, 1)
    SQL = SQL & " AND anyorem=" & RecuperaValor(Remesa, 2)
    SQL = SQL & " GROUP BY iban "
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    While Not miRsAux.EOF
                SQL = Mid(miRsAux!IBAN, 5, 4) ' Código de entidad receptora
                SQL = SQL & Mid(miRsAux!IBAN, 9, 4) ' Código de oficina receptora
                
                SQL = SQL & Mid(miRsAux!IBAN, 15, 10) ' Código de cuenta
                
                CC = Mid(miRsAux!IBAN, 13, 2) ' Dígitos de control
                
                'Este lo mando.
                SQL = CodigoDeControl(SQL)
                If SQL <> CC Then
                    
                    SQL = " - " & Mid(miRsAux!IBAN, 13, 2) & "- " & Mid(miRsAux!IBAN, 15, 10) & " --> CC. correcto:" & SQL
                    SQL = Mid(miRsAux!IBAN, 5, 4) & " - " & Mid(miRsAux!IBAN, 9, 4) & SQL
                    Registro = Registro & SQL & vbCrLf
                End If
                miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Registro <> "" Then
        SQL = "Las siguientes cuentas no son correctas.:" & vbCrLf & Registro
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    
    
    If vParamT.NuevasNormasSEPA Then
        'Si continuar y esta bien, veremos si todas los bancos tienen BIC asociado
        Registro = ""
        SQL = "select mid(cobros.iban,5,4) codbanco,bics.entidad from cobros left join bics on mid(cobros.iban,5,4)=bics.entidad WHERE "
        SQL = SQL & " codrem = " & RecuperaValor(Remesa, 1)
        SQL = SQL & " AND anyorem=" & RecuperaValor(Remesa, 2) & " group by 1"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Registro = ""
        While Not miRsAux.EOF
            If IsNull(miRsAux!Entidad) Then Registro = Registro & "/    " & miRsAux!codbanco & "    "
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If Registro <> "" Then
            Registro = Mid(Registro, 2) & vbCrLf & vbCrLf & "¿Continuar?"
            SQL = "Las siguientes bancos no tiene BIC asocidado:" & vbCrLf & vbCrLf & Registro
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Function
        End If
        
        
    End If
    
    
    
    comprobarCuentasBancariasRecibos = True
    Exit Function
EcomprobarCuentasBancariasRecibos:
    MuestraError Err.Number, "comprobar Cuentas Bancarias Recibos"
End Function

'La norma 19 acepta como identificador del "cliente" el campo referencia en la BD
'Con lo cual comporbaremos que no esta en blanco
Private Function ComprobarCampoReferenciaRemesaNorma19(Remesa As String) As Boolean
    ComprobarCampoReferenciaRemesaNorma19 = False
    SQL = "select codmacta,NUmSerie,numfactu,numorden,referencia from cobros where codrem = " & RecuperaValor(Remesa, 1)
    SQL = SQL & " AND anyorem=" & RecuperaValor(Remesa, 2) & " ORDER BY codmacta"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    SQL = ""
    NF = 0
    While Not miRsAux.EOF
        If DBLet(miRsAux!Referencia, "T") = "" Then
            Registro = Registro & miRsAux!codmacta & " - " & miRsAux!NUmSerie & "/" & miRsAux!NumFactu & "-" & miRsAux!numorden & vbCrLf
            NF = NF + 1
        Else
            If Len(miRsAux!Referencia) > 12 Then SQL = SQL & miRsAux!codmacta & " - " & miRsAux!NUmSerie & "/" & miRsAux!NumFactu & "-" & miRsAux!numorden & "(" & miRsAux!Referencia & ")" & vbCrLf
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If NF > 0 Then
        Registro = "Referencias vacias: " & NF & vbCrLf & vbCrLf & Registro
        MsgBox Registro, vbExclamation
    Else
        If SQL <> "" Then
            Registro = "Longitud referencia incorrecta: " & vbCrLf & vbCrLf & SQL
            Registro = Registro & vbCrLf & "¿Continuar?"
            If MsgBox(Registro, vbQuestion + vbYesNo) = vbNo Then Exit Function
        End If
        ComprobarCampoReferenciaRemesaNorma19 = True
    End If
End Function


'Modificacion noviembre 2012
'El fichero(en alzira) viene en formato WRI
'es decir el salto de linea no es el mismo. Por lo tanto
' input nf,cad  solo le UN registro con toda la informacion
' Preprocesaremos el fichero.
'  0.- Abrir
'  1.- Leer linea y apuntar a siguiente
'  2.- Preguntar si es ultima linea
'  3.- Cerrar coolee0ction
Private Sub ProcesoFicheroDevolucion(OptProces As Byte, ByRef LinFichero As Collection)
Dim B As Boolean
    'No pongo on error Que salte en el SUB ProcesaCabeceraFicheroDevolucion

    Select Case OptProces
    Case 0
        'Abrir el fichero y cargar el objeto COLLECTION
        NF = FreeFile
        Open Registro For Input As #NF
        Line Input #NF, Registro
        Set LinFichero = New Collection
        
        
        'Veremos que tipo de fichero es Normal. Ni lleva saltos de linea ni lleva vbcr ni vblf
        B = InStr(1, Registro, vbCrLf) > 0
        If B > 0 Then
            SQL = vbCrLf 'separaremos por este
        Else
            B = InStr(1, Registro, vbCr) > 0
            If B Then
                SQL = vbCr
            Else
                B = InStr(1, Registro, vbLf)
                If B Then SQL = vbLf
            End If
        End If
        
        If Not B Then
            'Normal.
            LinFichero.Add Registro
            While Not EOF(NF)
                Line Input #NF, Registro
                LinFichero.Add Registro
            Wend
        Else
            'El fichero NO va separado correctamente(tipo alzira nuevo WRI)
            Do
                NumRegElim = InStr(1, Registro, SQL)
                If NumRegElim = 0 Then
                    'NO DEBERIA PASAR
                    MsgBox "Preproceso fichero banco. Numregelim=0.  Avise soporte tecnico", vbExclamation
                Else

                    LinFichero.Add Mid(Registro, 1, NumRegElim - 1)
                    NumRegElim = NumRegElim + Len(SQL)
                    Registro = Mid(Registro, NumRegElim)  'quito el separador
                End If
                    
            Loop Until Registro = ""
        
        End If
        Close #NF
        NF = 1 'Puntero a la linea en question
        
    Case 1
        'Recorrer el COLLECTION
        'Damos la linea y movemos a la siguiente
        If NF <= LinFichero.Count Then
            Registro = LinFichero(NF)
            NF = NF + 1
        Else
            Err.Raise 513, "Sobrepasaod vector"
        End If
    Case 2
        'reutilizamos variables
        If NF > LinFichero.Count Then
            Registro = "Si"
        Else
            Registro = ""
        End If
    Case 4
        'Cerrar
        Set LinFichero = Nothing
    End Select

End Sub


'---------------------------------------------------------------------
'  DEVOLUCION FICHERO

Public Sub ProcesaCabeceraFicheroDevolucion(Fichero As String, ByRef Remesa As String)
Dim aux2 As String  'Para buscar los vencimientos
Dim FinLecturaLineas As Boolean
Dim TodoOk As Boolean
Dim ErroresVto As String
Dim Cuantos As Integer
Dim Bien As Integer
Dim LinDelFichero As Collection
Dim EsFormatoAntiguoDevolucion As Boolean

    On Error GoTo EDevRemesa
    Remesa = ""
    
    EsFormatoAntiguoDevolucion = Dir(App.Path & "\DevRecAnt.dat") <> ""
    
    
    'ANTES nov 2012
    '
    'nf = FreeFile
    'Open Fichero For Input As #nf
    Registro = Fichero 'para no pasr mas variables al proceso
    ProcesoFicheroDevolucion 0, LinDelFichero 'abrir el fichero y volcarlo sobre un Collection
    
    'Proceso la primera linea. A veriguare a que norma pertenece
    ' y hallare la remesa
    'Line Input #nf, Registro
    ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
    
    'Comproamos ciertas cosas
    SQL = "Linea 1 vacia"
    If Registro <> "" Then
        
        'NIF
        SQL = Mid(Registro, 5, 9)
        
        'Tiene valor
        If Len(Registro) <> 162 Then
            SQL = "Longitud linea incorrecta(162)"
        Else
            'Noviembre 2012
            'en lugar de 5190 comprobamos que sea 519
            If Mid(Registro, 1, 3) <> "519" Then
                SQL = "Cadena control incorrecta(519)"
            Else
                SQL = ""
            End If
        End If
    End If
    
    If SQL = "" Then
    
        'Segunda LINEA.
        'Line Input #nf, Registro
        ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
        
        SQL = "Linea 2 vacia"
        If Registro <> "" Then
            
            'NIF
            SQL = Mid(Registro, 5, 9)
            
            
            'Tiene valor
            If Len(Registro) <> 162 Then
                SQL = "Longitud linea incorrecta(162)"
            Else
                'En lugar de 5390 comprobamos por 539
                If Mid(Registro, 1, 3) <> "539" Then
                    SQL = "Cadena control incorrecta(539)"
                Else
                    
                    SQL = "Falta linea 569"
                    Remesa = ""
                    Do
                        ProcesoFicheroDevolucion 2, LinDelFichero  'vemos si es ultima linea
                        
                        If Registro <> "" Then
                            SQL = "FIN LINEAS. No se ha encontrado linea: 569"
                            Remesa = "NO"
                        Else
                            'Line Input #nf, Registro
                            ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
                            
                            'BUsco la linea:
                            '5690
                            If Registro <> "" Then
                                'Nov 2012   En lugar de 5690 comprobamos 569
                                If Mid(Registro, 1, 3) = "569" Then
                                    SQL = ""
                                    Remesa = "NO"
                                End If
                            End If
                        End If
                        
                    Loop Until Remesa <> ""
                    Remesa = ""
                    
                    If SQL = "" Then
                        'VAMOS BIEN. Veremos si a partir de los datos del recibo nos dan la remesa
                        'Para ello bucaremos en registro, la cadena que contiene los datos
                        'del vencimiento
                        'Registro=
                        '5690B97230080000970000100066COSTURATEX,  S.L.                       007207779700001000660000022516311205A020574911Fac
                        '5690F46024196009242820002250DAVID MONTAGUD CARRASCO                 318871052428200022500000010187                FRA 2731591 GASOLINERA ALZICOOP         1

                        Set miRsAux = New ADODB.Recordset
                        ErroresVto = ""
                        FinLecturaLineas = False
                        Cuantos = 0
                        Bien = 0
                        Do
                            
                            If Mid(Registro, 1, 3) = "569" Then
                                'Los vtos vienen en estas lineas
                                Cuantos = Cuantos + 1
                                Registro = Mid(Registro, 99, 17)
                                SQL = "Select codrem,anyorem,siturem from cobros where fecfactu='20" & Mid(Registro, 5, 2) & "-" & Mid(Registro, 3, 2) & "-" & Mid(Registro, 1, 2)
                                aux2 = SQL
                                
                                'Problemas en alzira
                                'If Not IsNumeric(Mid(Registro, 17, 1)) Then
                                'Sept 2013
                                If Not EsFormatoAntiguoDevolucion Then
                                    SQL = SQL & "' AND numserie like '" & Trim(Mid(Registro, 7, 1)) & "%' AND numfactu = " & Val(Mid(Registro, 9, 7)) & " AND numorden=" & Mid(Registro, 16, 1)
                                    'Problema en herbelca. El numero de vto NO viene con la factura
                                    aux2 = aux2 & "' AND numserie like '" & Trim(Mid(Registro, 7, 1)) & "%' AND numfactu = " & Val(Mid(Registro, 9, 8))
                                    
                                Else
                                    'El vencimiento si que es el 17
                                    SQL = SQL & "' AND numserie like '" & Trim(Mid(Registro, 7, 1)) & "%' AND numfactu = " & Val(Mid(Registro, 10, 7)) & " AND numorden=" & Mid(Registro, 17, 1)
                                    aux2 = aux2 & "' AND numserie like '" & Trim(Mid(Registro, 7, 1)) & "%' AND numfactu = " & Val(Mid(Registro, 10, 8))
                                    
                                End If
                                
                                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                                TodoOk = False
                                SQL = "Vencimiento no encontrado: " & Registro
                                If Not miRsAux.EOF Then
                                    If IsNull(miRsAux!CodRem) Then
                                        SQL = "Vencimiento sin Remesa: " & Registro
                                    Else
                                        SQL = miRsAux!CodRem & "|" & miRsAux!AnyoRem & "|·"
                                        
                                        If InStr(1, Remesa, SQL) = 0 Then Remesa = Remesa & SQL
                                        SQL = ""
                                        TodoOk = True
                                    End If
                                End If
                                miRsAux.Close
                                
                                
                                If Not TodoOk Then
                                    'Los busco sin Numorden
                                    miRsAux.Open aux2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                                    If Not miRsAux.EOF Then
                                        If IsNull(miRsAux!CodRem) Then
                                            SQL = "Vencimiento sin Remesa: " & Registro
                                        Else
                                            SQL = miRsAux!CodRem & "|" & miRsAux!AnyoRem & "|·"
                                            
                                            If InStr(1, Remesa, SQL) = 0 Then Remesa = Remesa & SQL
                                            SQL = ""
                                            TodoOk = True
                                        End If
                                    End If
                                    miRsAux.Close
                                
                                End If
                                
                                
                                
                                If SQL <> "" Then
                                    ErroresVto = ErroresVto & vbCrLf & SQL
                                Else
                                    Bien = Bien + 1
                                End If
                            Else
                                'La linea no empieza por 569
                                'veremos los totales
                                
                                If Mid(Registro, 1, 3) = "599" Then
                                    'TOTAL TOTAL
                                    SQL = Mid(Registro, 105, 10)
                                    If Val(SQL) <> Cuantos Then ErroresVto = "Fichero: " & SQL & "   Leidos" & Cuantos & vbCrLf & ErroresVto & vbCrLf & SQL
                                End If
                            End If
                            
                            'Siguiente linea
                            ProcesoFicheroDevolucion 2, LinDelFichero  'vemos si es ultima linea
                            
                            If Registro <> "" Then
                                FinLecturaLineas = True
                            Else
                                'Line Input #nf, Registro
                                ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
                            End If
                            
                        Loop Until FinLecturaLineas
                        
                        If Cuantos <> Bien Then ErroresVto = ErroresVto & vbCrLf & "Total: " & Cuantos & "   Correctos:" & Bien
                        
                        SQL = ErroresVto
                        Set miRsAux = Nothing
                    
                    End If
                End If  'Control SEGUNDA LINEA
        
        
            End If
        End If
    
    End If  'DE SEGUNDA LINEA
    
    ProcesoFicheroDevolucion 3, LinDelFichero
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
    Else
        'Remesa = Mid(Registro, 1, 4) & "|" & Mid(Registro, 5) & "|"
        
        
        'Ahora comprobaremos que para cada remesa  veremos si existe y si la situacion es la contabilizadxa
        SQL = Remesa
        Registro = "" 'Cadena de error de situacion remesas
        Set miRsAux = New ADODB.Recordset
        Do
            Cuantos = InStr(1, SQL, "·")
            If Cuantos = 0 Then
                SQL = ""
            Else
                aux2 = Mid(SQL, 1, Cuantos - 1)
                SQL = Mid(SQL, Cuantos + 1)
                
                
                'En aux2 tendre codrem|anñorem|
                aux2 = RecuperaValor(aux2, 1) & " AND anyo = " & RecuperaValor(aux2, 2)
                aux2 = "Select situacion from remesas where codigo = " & aux2
                miRsAux.Open aux2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If miRsAux.EOF Then
                    aux2 = "-No se encuentra remesa"
                Else
                    'Si que esta.
                    'Situacion
                    If CStr(miRsAux!Situacion) <> "Q" Then
                        aux2 = "- Situacion incorrecta : " & miRsAux!Situacion
                    Else
                        aux2 = "" 'TODO OK
                    End If
                End If
            
                If aux2 <> "" Then
                    aux2 = aux2 & " ->" & Mid(miRsAux.Source, InStr(1, UCase(miRsAux.Source), " WHERE ") + 7)
                    aux2 = Replace(aux2, " AND ", " ")
                    aux2 = Replace(aux2, "anyo", "año")
                    Registro = Registro & vbCrLf & aux2
                End If
                miRsAux.Close
            End If
        Loop Until SQL = ""
        Set miRsAux = Nothing
        
        
        If Registro <> "" Then
            Registro = "Error remesas " & vbCrLf & String(30, "=") & Registro
            MsgBox Registro, vbExclamation
            
            'Pongo REMESA=""
            Remesa = "" 'para que no continue el preoceso de devolucion
        End If
        
    End If
    
    Exit Sub
EDevRemesa:
    Remesa = ""
    MuestraError Err.Number, "Procesando fichero devolucion"
End Sub




Public Sub ProcesaLineasFicheroDevolucion(Fichero As String, ByRef Listado As Collection, ByRef EsSepa As Boolean)
Dim Registro As String
Dim SumaComprobacion As Currency
Dim impo As Currency
Dim Fin As Boolean
Dim B As Boolean
Dim AUX As String
Dim C2 As String
Dim bol As Boolean

    On Error GoTo EDevRemesa1
  
    
    
    

    NF = FreeFile
    Open Fichero For Input As #NF
    
    'Las dos primeras son el encabezado.
    ' Noviembre 2012. Hay que comprobar que si vienen todo en una linea o NO
    Line Input #NF, Registro
    
    
    B = InStr(1, Registro, vbCrLf) > 0
    If B > 0 Then
        AUX = vbCrLf 'separaremos por este
    Else
        B = InStr(1, Registro, vbCr) > 0
        If B Then
            AUX = vbCr
        Else
            B = InStr(1, Registro, vbLf)
            If B Then AUX = vbLf
        End If
    End If
    
    EsSepa = False
    If Mid(Registro, 1, 4) = "2119" Then EsSepa = True
        
    
    
    If B Then
        'TRAE TODO en una unica linea. Separaremos por el vbcr o vbcrlf
        Do
                NumRegElim = InStr(1, Registro, AUX)
                If NumRegElim = 0 Then
                    
                Else

                    SQL = Mid(Registro, 1, NumRegElim - 1)
                    NumRegElim = NumRegElim + Len(AUX)
                    Registro = Mid(Registro, NumRegElim)  'quito el separador
                    
                    
                   
                    
                    
                    If EsSepa Then
                        C2 = Mid(SQL, 1, 2)
                        If C2 = "23" Then
                            impo = Val(Mid(SQL, 89, 11)) / 100
                            SumaComprobacion = SumaComprobacion + impo
                            
                            'Cuestion 2
                            'Datos identifictivos del vencimiento
                            SQL = Mid(SQL, 21, 35)
                            Listado.Add SQL
                            SQL = ""
                        Else
                            If C2 = "99" Then 'antes 5990
                                Fin = True
                                impo = Val(Mid(SQL, 3, 17)) / 100
                            Else
                                SQL = ""
                            End If
                        End If
                    Else
                        C2 = Mid(SQL, 1, 3)
                        If C2 = "569" Then
                            impo = Val(Mid(SQL, 89, 10)) / 100
                            SumaComprobacion = SumaComprobacion + impo
                            
                            'Cuestion 2
                            'Datos identifictivos del vencimiento
                            SQL = Mid(SQL, 89, 27)
                            Listado.Add SQL
                            SQL = ""
                        Else
                            If C2 = "599" Then 'antes 5990
                                Fin = True
                                impo = Val(Mid(SQL, 89, 10)) / 100
                            Else
                                SQL = ""
                            End If
                        End If
                    
                    End If
                    
                End If
                    
        Loop Until Registro = ""
            
        'Cerramos y salimos
        Close #NF
        Exit Sub
    End If
    
    Line Input #NF, Registro
    
    'Ahora empezamos
    SumaComprobacion = 0
    Fin = False
    SQL = ""
    Do
        Line Input #NF, Registro
        If Registro <> "" Then
         
            SQL = Mid(Registro, 1, 3)
            
            If EsSepa Then
                bol = Mid(Registro, 1, 4) = "2319"
            Else
                bol = SQL = "569"
            End If
            If bol Then
                'Registro normal de devolucion
                '1... 68 carcaater
                '5690B972300800003169816315  RUANO MORENO, VICENTE                   "
                '69 .. 162
                '3082140015316981631500000350890047080000004708Fact. 2059121 31/12/2005 Tarj   9434    1
                
                'Cuestion 1:
                'Importe: 0000035089 desde la poscion  hasta la posicion
                If EsSepa Then
                    impo = Val(Mid(Registro, 89, 11)) / 100
                Else
                    impo = Val(Mid(Registro, 89, 10)) / 100
                End If
                SumaComprobacion = SumaComprobacion + impo
                
                'Cuestion 2
                'Datos identifictivos del vencimiento
                If EsSepa Then
                    SQL = Mid(Registro, 21, 35)
                Else
                    SQL = Mid(Registro, 89, 27)
                End If
                Listado.Add SQL
                SQL = ""
            Else
                
                If EsSepa Then
                    bol = Mid(Registro, 1, 2) = "99"
                Else
                    bol = SQL = "599"
                End If
                    
                If bol Then
                    Fin = True
                    If EsSepa Then
                        impo = Val(Mid(Registro, 3, 17)) / 100
                    Else
                        impo = Val(Mid(Registro, 89, 10)) / 100
                    End If
                Else
                    SQL = ""
                End If
            End If
        End If
        If EOF(NF) Then Fin = True
    Loop Until Fin
    Close #NF
    
    If SQL = "" Then
        MsgBox "No se ha leido la linea final fichero", vbExclamation
        Set Listado = Nothing
    Else
        'OK salimos
        If impo <> SumaComprobacion Then
            SQL = "Error leyendo importes. ¿Desea continuar con los datos obtenidos?"
            If MsgBox(SQL, vbExclamation) = vbNo Then Set Listado = Nothing
        End If
    End If
    
    
    Exit Sub
EDevRemesa1:
    MuestraError Err.Number, "Lineas devolucion"
End Sub


'------ aqui aqui aqui


        


'******************************************************************************************************************
'******************************************************************************************************************
'******************************************************************************************************************
'******************************************************************************************************************
'
'       Normas 34 y 68
'
'******************************************************************************************************************
'******************************************************************************************************************
'******************************************************************************************************************
'******************************************************************************************************************

'----------------------------------------------------------------------
'  Copia fichero generado bajo
'Public Sub CopiarFicheroNorma43(Es34 As Boolean, Destino As String)
Public Sub CopiarFicheroNormaBancaria(TipoFichero As Byte, Destino As String)
    
    'If Not CopiarEnDisquette(True, 3) Then
        AuxD = Destino
        'CopiarEnDisquette False, 0, Es34 'A disco
        CopiarEnDisquette TipoFichero
        
End Sub
'Private Function CopiarEnDisquette(A_disquetera As Boolean, Intentos As Byte, Es34 As Boolean) As Boolean
'TipoFichero
'   0- norma 34
'   1- N8
'   2- Caixa confirming
Private Function CopiarEnDisquette(TipoFichero As Byte) As Boolean
Dim I As Integer
Dim Cad As String

On Error Resume Next

    CopiarEnDisquette = False
    
 '   If A_disquetera Then
 '       For I = 1 To Intentos
 '           Cad = "Introduzca un disco vacio. (" & I & ")"
 '           MsgBox Cad, vbInformation
 '           FileCopy App.Path & "\norma34.txt", "a:\norma34.txt"
 '           If Err.Number <> 0 Then
 '               MuestraError Err.Number, "Copiar En Disquette"
 '           Else
 '               CopiarEnDisquette = True
 '               Exit For
 '           End If
 '       Next I
 '   Else
        If AuxD = "" Then
            Cad = Format(Now, "ddmmyyhhnn")
            Cad = App.Path & "\" & Cad & ".txt"
        Else
            Cad = AuxD
        End If
        'If Es34 Then
        '    FileCopy App.Path & "\norma34.txt", Cad
        'Else
        '    FileCopy App.Path & "\norma68.txt", Cad
        'End If
        Select Case TipoFichero
        Case 0
            FileCopy App.Path & "\norma34.txt", Cad
        Case 1
            FileCopy App.Path & "\norma34.txt", Cad
        Case 2
            If vParamT.PagosConfirmingCaixa Then
                FileCopy App.Path & "\normaCaixa.txt", Cad
            Else
                FileCopy App.Path & "\norma68.txt", Cad
            End If
            
        End Select
        If Err.Number <> 0 Then
            MsgBox "Error creando copia fichero. Consulte soporte técnico." & vbCrLf & Err.Description, vbCritical
        Else
            MsgBox "El fichero esta guardado como: " & Cad, vbInformation
        End If
            
    'End If
End Function

Public Function GeneraFicheroNorma34(CIF As String, Fecha As Date, CuentaPropia As String, ConceptoTransferencia As String, vNumeroTransferencia As Integer, ByRef ConceptoTr As String, Pagos As Boolean) As Boolean
    
    
    If vParamT.NuevasNormasSEPA Then
        If vParamT.NormasFormatoXML Then
            GeneraFicheroNorma34 = GeneraFicheroNorma34SEPA_XML(CIF, Fecha, CuentaPropia, CLng(vNumeroTransferencia), Pagos, ConceptoTransferencia)
        Else
            GeneraFicheroNorma34 = GeneraFicheroNorma34SEPA(CIF, Fecha, CuentaPropia, CLng(vNumeroTransferencia), Pagos, ConceptoTransferencia)
        End If
    Else
        
        GeneraFicheroNorma34 = GeneraFicheroNorma34_(CIF, Fecha, CuentaPropia, ConceptoTransferencia, vNumeroTransferencia, ConceptoTr, Pagos)
    End If

End Function


Public Function GeneraFicheroNorma34SEPA_XML(CIF As String, Fecha As Date, CuentaPropia2 As String, NumeroTransferencia As Long, Pagos As Boolean, ConceptoTr As String) As Boolean
Dim Regs As Integer
Dim Importe As Currency
Dim Im As Currency
Dim Cad As String
Dim AUX As String
Dim SufijoOEM As String
Dim NFic As Integer
Dim EsPersonaJuridica2 As Boolean

    On Error GoTo EGen3
    GeneraFicheroNorma34SEPA_XML = False
    
    NFic = -1
    
    
    'Cargamos la cuenta
    Cad = "Select * from bancos where codmacta='" & CuentaPropia2 & "'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        Cad = ""
    Else
        If IsNull(miRsAux!IBAN) Then
            Cad = ""
        Else
            SufijoOEM = "000" ''Sufijo3414
            Cad = miRsAux!IBAN
            If DBLet(miRsAux!Sufijo3414, "T") <> "" Then SufijoOEM = Right("000" & miRsAux!Sufijo3414, 3)
            CuentaPropia2 = Cad
        End If
        
        
    End If
    miRsAux.Close
  
    If Cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia2, vbExclamation
        Exit Function
    End If
    
    NFic = FreeFile
    Open App.Path & "\norma34.txt" For Output As NFic
    
    
    Print #NFic, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    Print #NFic, "<Document xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""urn:iso:std:iso:20022:tech:xsd:pain.001.001.03"">"
    Print #NFic, "<CstmrCdtTrfInitn>"
    Print #NFic, "   <GrpHdr>"
    Cad = "TRAN" & IIf(Pagos, "PAG", "ABO") & Format(NumeroTransferencia, "000000") & "F" & Format(Now, "yyyymmddThhnnss")
    Print #NFic, "      <MsgId>" & Cad & "</MsgId>"
    Print #NFic, "      <CreDtTm>" & Format(Now, "yyyy-mm-ddThh:nn:ss") & "</CreDtTm>"
    
    
    If Pagos Then
        AUX = "ImpEfect - coalesce(imppagad ,0)"
        Cad = "spagop"
    Else
        AUX = "abs(impvenci + coalesce(Gastos, 0) - coalesce(impcobro, 0))"
        Cad = "scobro"
    End If
    Cad = "Select count(*),sum(" & AUX & ") FROM " & Cad & " WHERE transfer = " & NumeroTransferencia
    AUX = "0|0|"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(1)) Then AUX = miRsAux.Fields(0) & "|" & Format(miRsAux.Fields(1), "#.00") & "|"
    End If
    miRsAux.Close
    
    Print #NFic, "      <NbOfTxs>" & RecuperaValor(AUX, 1) & "</NbOfTxs>"
    Print #NFic, "      <CtrlSum>" & TransformaComasPuntos(RecuperaValor(AUX, 2)) & "</CtrlSum>"
    Print #NFic, "      <InitgPty>"
    Print #NFic, "         <Nm>" & XML(vEmpresa.nomempre) & "</Nm>"
    Print #NFic, "         <Id>"
    Cad = Mid(CIF, 1, 1)
    
    EsPersonaJuridica2 = Not IsNumeric(Cad)

    
    
    
    Cad = "PrvtId"
    If EsPersonaJuridica2 Then Cad = "OrgId"
    
    Print #NFic, "           <" & Cad & ">"
    Print #NFic, "               <Othr>"
    Print #NFic, "                  <Id>" & CIF & SufijoOEM & "</Id>"
    Print #NFic, "               </Othr>"
    Print #NFic, "           </" & Cad & ">"
    
    Print #NFic, "         </Id>"
    Print #NFic, "      </InitgPty>"
    Print #NFic, "   </GrpHdr>"

    Print #NFic, "   <PmtInf>"
    
    Print #NFic, "      <PmtInfId>" & Format(Now, "yyyymmddhhnnss") & CIF & "</PmtInfId>"
    Print #NFic, "      <PmtMtd>TRF</PmtMtd>"
    Print #NFic, "      <ReqdExctnDt>" & Format(Fecha, "yyyy-mm-dd") & "</ReqdExctnDt>"
    Print #NFic, "      <Dbtr>"
    
     'Nombre
    miRsAux.Open "Select siglasvia ,direccion ,numero ,codpobla,pobempre,provempre,provincia from empresa2"
    Cad = Cad & FrmtStr(vEmpresa.nomempre, 70)
    If miRsAux.EOF Then Err.Raise 513, , "Error obteniendo datos empresa(empresa2)"
    
    Print #NFic, "         <Nm>" & XML(vEmpresa.nomempre) & "</Nm>"
    Print #NFic, "         <PstlAdr>"
    Print #NFic, "            <Ctry>ES</Ctry>"

    Cad = DBLet(miRsAux!siglasvia, "T") & " " & miRsAux!Direccion & " " & DBLet(miRsAux!numero, "T") & " "
    Cad = Cad & Trim(DBLet(miRsAux!CodPobla, "T") & " " & miRsAux!pobempre) & " "
    Cad = Cad & DBLet(miRsAux!provincia, "T")
    miRsAux.Close
    Print #NFic, "            <AdrLine>" & XML(Trim(Cad)) & "</AdrLine>"
    
    Print #NFic, "         </PstlAdr>"
    Print #NFic, "         <Id>"
    
    AUX = "PrvtId"
    If EsPersonaJuridica2 Then AUX = "OrgId"
   
    
    Print #NFic, "            <" & AUX & ">"
    
    Print #NFic, "               <Othr>"
    Print #NFic, "                  <Id>" & CIF & SufijoOEM & "</Id>"
    Print #NFic, "               </Othr>"
    Print #NFic, "            </" & AUX & ">"
    Print #NFic, "         </Id>"
    Print #NFic, "    </Dbtr>"
    
    
    Print #NFic, "    <DbtrAcct>"
    Print #NFic, "       <Id>"
    Print #NFic, "          <IBAN>" & Trim(CuentaPropia2) & "</IBAN>"
    Print #NFic, "       </Id>"
    Print #NFic, "       <Ccy>EUR</Ccy>"
    Print #NFic, "    </DbtrAcct>"
    Print #NFic, "    <DbtrAgt>"
    Print #NFic, "       <FinInstnId>"
    
    Cad = Mid(CuentaPropia2, 5, 4)
    Cad = DevuelveDesdeBD("bic", "bics", "entidad", Cad)
    Print #NFic, "          <BIC>" & Trim(Cad) & "</BIC>"
    Print #NFic, "       </FinInstnId>"
    Print #NFic, "    </DbtrAgt>"
    
    
    
    
    'Para ello abrimos la tabla tmpNorma34
    If Pagos Then
        Cad = "Select spagop.*,nommacta,dirdatos,codposta,dirdatos,desprovi,pais,cuentas.despobla,bic,nifdatos from spagop"
        Cad = Cad & " left join sbic on spagop.entidad=sbic.entidad INNER JOIN cuentas ON"
        Cad = Cad & " codmacta=ctaprove WHERE transfer =" & NumeroTransferencia
    Else
        'ABONOS
         '
        Cad = "Select mid(cobros.iban,5,4) as entidad,mid(cobros.iban,9,4) as oficina,mid(cobro.iban,15,10) cuentaba,mid(cobros.iban,13,2) as CC,cobros.iban"
        Cad = Cad & ",nomclien nommacta,dirdatos,codposta,despobla,impvenci,cobros.codmacta,codpais,Gastos,impcobro,desprovi"
        Cad = Cad & " ,NUmSerie,numfactu,fecfactu,numorden,text33csb,text41csb,bic,nifclien nifdatos from cobros"
        Cad = Cad & " LEFT JOIN bics on mid(cobros.iban,5,4)=bics.entidad "
        Cad = Cad & " WHERE transfer =" & NumeroTransferencia
    End If
    miRsAux.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    Regs = 0
    While Not miRsAux.EOF
        Print #NFic, "   <CdtTrfTxInf>"
        Print #NFic, "      <PmtId>"
        
         
        If Pagos Then
            'numfactu fecfactu numorden
            AUX = FrmtStr(miRsAux!NumFactu, 10)
            AUX = AUX & Format(miRsAux!FecFactu, "yyyymmdd") & Format(miRsAux!numorden, "000")
        
        Else
            'fecfaccl
            AUX = FrmtStr(miRsAux!NUmSerie, 3) & Format(miRsAux!NumFactu, "00000000")
            AUX = AUX & Format(miRsAux!FecFactu, "yyyymmdd") & Format(miRsAux!numorden, "000")
        End If
        
        Print #NFic, "         <EndToEndId>" & AUX & "</EndToEndId>"
        Print #NFic, "      </PmtId>"
        Print #NFic, "      <PmtTpInf>"
        If Pagos Then
            Im = DBLet(miRsAux!imppagad, "N")
            Im = miRsAux!ImpEfect - Im
            AUX = miRsAux!ctaprove

        Else
            Im = Abs(miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N")) - DBLet(miRsAux!impcobro, "N")
            AUX = miRsAux!codmacta
        End If
        
        'Persona fisica o juridica
        Cad = Mid(miRsAux!nifdatos, 1, 1)
        EsPersonaJuridica2 = Not IsNumeric(Cad)
        'Como da problemas Cajamar, siempre ponemos Perosna juridica. Veremos
        EsPersonaJuridica2 = True
        
        
        Importe = Importe + Im
        Regs = Regs + 1
        
        Print #NFic, "          <SvcLvl><Cd>SEPA</Cd></SvcLvl>"
        'Print #NFic, "          <LclInstrm><Cd>SDCL</Cd></LclInstrm>"
        If ConceptoTr = "1" Then
            AUX = "SALA"
        ElseIf ConceptoTr = "0" Then
            AUX = "PENS"
        Else
            AUX = "TRAD"
        End If
        Print #NFic, "          <CtgyPurp><Cd>" & AUX & "</Cd></CtgyPurp>"
        Print #NFic, "       </PmtTpInf>"
        Print #NFic, "       <Amt>"
        Cad = Format(Im, "#.00")
        Print #NFic, "          <InstdAmt Ccy=""EUR"">" & TransformaComasPuntos(Cad) & "</InstdAmt>"
        Print #NFic, "       </Amt>"
        Print #NFic, "       <CdtrAgt>"
        Print #NFic, "          <FinInstnId>"
        Cad = DBLet(miRsAux!BIC, "T")
        If Cad = "" Then Err.Raise 513, , "No existe BIC: " & miRsAux!Nommacta & vbCrLf & "Entidad: " & miRsAux!Entidad
        Print #NFic, "             <BIC>" & DBLet(miRsAux!BIC, "T") & "</BIC>"
        Print #NFic, "          </FinInstnId>"
        Print #NFic, "       </CdtrAgt>"
        Print #NFic, "       <Cdtr>"
        Print #NFic, "          <Nm>" & XML(miRsAux!Nommacta) & "</Nm>"
        
        
        'Como cajamar da problemas, lo quitamos para todos
        'Print #NFic, "          <PstlAdr>"
        '
        'Cad = "ES"
        'If Not IsNull(miRsAux!PAIS) Then Cad = Mid(miRsAux!PAIS, 1, 2)
        'Print #NFic, "              <Ctry>" & Cad & "</Ctry>"
        '
        'If Not IsNull(miRsAux!dirdatos) Then Print #NFic, "              <AdrLine>" & XML(miRsAux!dirdatos) & "</AdrLine>"
        'Cad = XML(Trim(DBLet(miRsAux!codposta, "T") & " " & DBLet(miRsAux!despobla, "T")))
        'If Cad <> "" Then Print #NFic, "              <AdrLine>" & Cad & "</AdrLine>"
        'If Not IsNull(miRsAux!desprovi) Then Print #NFic, "              <AdrLine>" & XML(miRsAux!desprovi) & "</AdrLine>"
        'Print #NFic, "           </PstlAdr>"
        
        
        
        Print #NFic, "           <Id>"
        AUX = "PrvtId"
        If EsPersonaJuridica2 Then AUX = "OrgId"
      
        Print #NFic, "               <" & AUX & ">"
        Print #NFic, "                  <Othr>"
        Print #NFic, "                     <Id>" & miRsAux!nifdatos & "</Id>"
        'Da problemas.... con Cajamar
        'Print #NFic, "                     <Issr>NIF</Issr>"
        Print #NFic, "                  </Othr>"
        Print #NFic, "               </" & AUX & ">"
        Print #NFic, "           </Id>"
        Print #NFic, "        </Cdtr>"
        Print #NFic, "        <CdtrAcct>"
        Print #NFic, "           <Id>"
        Print #NFic, "              <IBAN>" & IBAN_Destino(False) & "</IBAN>"
        Print #NFic, "           </Id>"
        Print #NFic, "        </CdtrAcct>"
        Print #NFic, "      <Purp>"
        
        If ConceptoTr = "1" Then
            AUX = "SALA"
        ElseIf ConceptoTr = "0" Then
            AUX = "PENS"
        Else
            AUX = "TRAD"
        End If
        
        Print #NFic, "         <Cd>" & AUX & "</Cd>"
        Print #NFic, "      </Purp>"
        Print #NFic, "      <RmtInf>"
        'Print #NFic, "      <Ustrd>ESTE ES EL CONCEPTO, POR TANTO NO SE SI SERA EL TEXTO QUE IRA DONDE TIENE QUE IR, O EN OTRO LADAO... A SABER. LO QUE ESTA CLARO ES QUE VA.</Ustrd>
        
        If Pagos Then
            ''`text1csb` `text2csb`
            AUX = DBLet(miRsAux!text1csb, "T") & " " & DBLet(miRsAux!text2csb, "T")
        Else
            '`text33csb` `text41csb`
            AUX = DBLet(miRsAux!text33csb, "T") & " " & DBLet(miRsAux!text41csb, "T")
        End If
        If Trim(AUX) = "" Then AUX = miRsAux!Nommacta
        Print #NFic, "         <Ustrd>" & XML(Trim(AUX)) & "</Ustrd>"
        Print #NFic, "      </RmtInf>"
        Print #NFic, "   </CdtTrfTxInf>"
 
       
    
            
        miRsAux.MoveNext
    Wend
    Print #NFic, "   </PmtInf>"
    Print #NFic, "</CstmrCdtTrfInitn></Document>"
    
    
    miRsAux.Close
    Set miRsAux = Nothing
    Close (NFic)
    NFic = -1
    If Regs > 0 Then GeneraFicheroNorma34SEPA_XML = True
    Exit Function
EGen3:
    MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    If NFic > 0 Then Close (NFic)
End Function

Public Function comprobarCuentasBancariasPagos(Transferencia As String, Anyo As String, Pagos As Boolean) As Boolean
Dim CC As String
Dim IBAN As String
On Error GoTo EcomprobarCuentasBancariasPagos

    comprobarCuentasBancariasPagos = False
    If Pagos Then
        SQL = "select * from pagos where transfer = " & Transferencia
    Else
        'ABONOS
        SQL = "Select * "
        SQL = SQL & " FROM cobros where transfer=" & Transferencia
        SQL = SQL & " and anyorem = " & DBSet(Anyo, "N")
    End If
    
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    NF = 0
    While Not miRsAux.EOF

        If DBLet(miRsAux!IBAN, "T") = "" Or Len(DBLet(miRsAux!IBAN, "T")) <> 24 Then
            SQL = ""
        Else
            SQL = "D"
        End If

    
        If SQL = "" Then
             Registro = Registro & miRsAux!codmacta & " - " & miRsAux!NUmSerie & "/" & miRsAux!NumFactu & "-" & miRsAux!numorden
             If NF < 2 Then
                Registro = Registro & "         "
                NF = NF + 1
             Else
                Registro = Registro & vbCrLf
                NF = 0
            End If
    
        End If
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If Registro <> "" Then
        SQL = "Los siguientes vencimientos no tienen la cuenta bancaria con todos los datos." & vbCrLf & Registro
        MsgBox SQL, vbExclamation
        Exit Function
    End If
    
    
    'Si llega aqui es que todos tienen DATOS
    If Pagos Then
        SQL = "select iban from pagos where transfer = " & Transferencia
        SQL = SQL & " GROUP BY mid(iban,5,4),mid(iban,9,4),mid(iban,15,10),mid(iban,13,2)"
    Else
        SQL = "SELECT iban"
        SQL = SQL & " FROM cobros where transfer=" & Transferencia
        SQL = SQL & " GROUP BY mid(iban,5,4),mid(iban,9,4),mid(iban,15,10),mid(iban,13,2)"
    End If
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    While Not miRsAux.EOF
                SQL = Mid(miRsAux!IBAN, 5, 4) ' Código de entidad receptora
                SQL = SQL & Mid(miRsAux!IBAN, 9, 4) ' Código de oficina receptora
                
                SQL = SQL & Mid(miRsAux!IBAN, 15, 10) ' Código de cuenta
                
                CC = Mid(miRsAux!IBAN, 13, 2) ' Dígitos de control
                
                'Este lo mando.
                IBAN = Mid(SQL, 1, 8) & CC & Mid(SQL, 9)
                
                SQL = CodigoDeControl(SQL)
                If SQL <> CC Then
                    
                    SQL = " - " & Mid(miRsAux!IBAN, 13, 2) & "- " & Mid(miRsAux!IBAN, 15, 10) & " --> CC. correcto:" & SQL
                    SQL = Mid(miRsAux!IBAN, 5, 4) & " - " & Mid(miRsAux!IBAN, 9, 4) & SQL
                    Registro = Registro & SQL & vbCrLf
                End If
                
                
                'Noviembre 2013
                'IBAN
                If vParam.NuevasNormasSEPA Then
                        SQL = "ES"
                        If DBLet(miRsAux!IBAN, "T") <> "" Then SQL = Mid(miRsAux!IBAN, 1, 2)
                    
                
                        If Not DevuelveIBAN2(SQL, IBAN, IBAN) Then
                            
                            SQL = "Error calculo"
                        Else
                            SQL = SQL & IBAN
                            If Mid(DBLet(miRsAux!IBAN, "T"), 1, 4) <> SQL Then
                                SQL = "Error IBAN. Calculado " & SQL & " / " & Mid(DBLet(miRsAux!IBAN, "T"), 1, 4)
                            Else
                                'OK
                                SQL = ""
                            End If
                        End If
                        
                        If SQL <> "" Then
                            SQL = SQL & " - " & Mid(miRsAux!IBAN, 13, 2) & "- " & Mid(miRsAux!IBAN, 15, 10) & " --> CC. correcto:" & SQL
                            SQL = Mid(miRsAux!IBAN, 5, 4) & " - " & Mid(miRsAux!IBAN, 9, 4) & SQL
                            Registro = Registro & "Error obteniendo IBAN: " & SQL & vbCrLf
                        End If
                End If
                
                
                miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Registro <> "" Then
        SQL = "Generando diskette." & vbCrLf & vbCrLf
        SQL = SQL & "Las siguientes cuentas no son correctas.:" & vbCrLf & Registro
        SQL = SQL & vbCrLf & "¿Desea continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    comprobarCuentasBancariasPagos = True
    Exit Function
EcomprobarCuentasBancariasPagos:
    MuestraError Err.Number, "comprobar Cuentas Bancarias pagos"
End Function



Private Function RellenaABlancos(CADENA As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim Cad As String
    
    Cad = Space(Longitud)
    If PorLaDerecha Then
        Cad = CADENA & Cad
        RellenaABlancos = Left(Cad, Longitud)
    Else
        Cad = Cad & CADENA
        RellenaABlancos = Right(Cad, Longitud)
    End If
    
End Function



Private Function RellenaAceros(CADENA As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim Cad As String
    
    Cad = Mid("00000000000000000000", 1, Longitud)
    If PorLaDerecha Then
        Cad = CADENA & Cad
        RellenaAceros = Left(Cad, Longitud)
    Else
        Cad = Cad & CADENA
        RellenaAceros = Right(Cad, Longitud)
    End If
    
End Function





'******************************************************************************************************************
'******************************************************************************************************************
'
'       Genera fichero CAIXACONFIRMING
'
'Cuenta propia tendra empipados entidad|sucursal|cc|cuenta|
Public Function GeneraFicheroCaixaConfirming(CIF As String, Fecha As Date, CuentaPropia As String, vNumeroTransferencia As Integer, ByVal ConceptoTr_ As String) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Im As Currency
Dim RS As ADODB.Recordset
Dim AUX As String
Dim Cad As String


    On Error GoTo EGen
    GeneraFicheroCaixaConfirming = False
    
    NumeroTransferencia = vNumeroTransferencia
    
    
    'Cargamos la cuenta
    Cad = "Select * from ctabancaria where codmacta='" & CuentaPropia & "'"
    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    AUX = Right("    " & CIF, 9)
    AUX = Mid(CIF & Space(10), 1, 9)
    If RS.EOF Then
        Cad = ""
    Else
        If IsNull(RS!Entidad) Then
            Cad = ""
        Else
            
            CodigoOrdenante = Format(RS!Entidad, "0000") & Format(DBLet(RS!oficina, "N"), "0000") & Format(DBLet(RS!Control, "N"), "00") & Format(DBLet(RS!CtaBanco, "T"), "0000000000")
            
            If Not DevuelveIBAN2("ES", CodigoOrdenante, Cad) Then Cad = ""
            CuentaPropia = "ES" & Cad & CodigoOrdenante
                        
            'Esta variable NO se utiliza. La cojo "prestada"
            'Guardare el numero de contrato de CAIXACONFIRMING
            ' Sera, un char de 14
            ' Si no pone nada sera oficnacuenta  Total 14 posiciones
            ConceptoTr_ = Trim(DBLet(RS!CaixaConfirming, "T"))
            If ConceptoTr_ = "" Then ConceptoTr_ = Mid(CodigoOrdenante, 5, 4) & Mid(CodigoOrdenante, 11, 10)
            
            '                ENTIDAD
            ConceptoTr_ = Mid(CodigoOrdenante, 1, 4) & ConceptoTr_
        End If
        
        
    End If
    RS.Close
    Set RS = Nothing
    If Cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\normaCaixa.txt" For Output As #NFich
    
    
    
    
    
    'Codigo ordenante
    '---------------------------------------------------
    'Si el banco tiene puesto si ID de norma34 entonces
    'la pongo aquin. Lo he cargado antes sobre la variable AUX
    CodigoOrdenante = Left(AUX & "          ", 10)  'CIF EMPRESA
  
    Set RS = New ADODB.Recordset
    
    'CABECERA
    'UNo
    AUX = "0156" & CodigoOrdenante & Space(12) & "001" & Format(Fecha, "ddmmyy") & Space(6)
    AUX = AUX & ConceptoTr_ & "1" & "EUR" & Space(9)   'Ya esta. Ya he utlizado la variable ConceptoTr_. Nada mas
    Print #NFich, AUX
    'Nombre
    AUX = "0156" & CodigoOrdenante & Space(12) & "002" & FrmtStr(vEmpresa.nomempre, 36) & Space(7)
    Print #NFich, AUX
    
    'Registros obligatorios  3 4
    AUX = "Select pobempre, provempre from empresa2"
    RS.Open AUX, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUEDE SER EOF
    For Regs = 0 To 1
        AUX = "0156" & CodigoOrdenante & Space(12) & Format(Regs + 3, "000") & FrmtStr(DBLet(RS.Fields(Regs), "T"), 36) & Space(7)
        Print #NFich, AUX
    Next
    RS.Close
    
    
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma34
    
    AUX = "Select spagop.*,nommacta,dirdatos,codposta,dirdatos,despobla,nifdatos,razosoci,desprovi,pais from spagop,cuentas"
    AUX = AUX & " where codmacta=ctaprove and transfer =" & NumeroTransferencia
    RS.Open AUX, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If RS.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not RS.EOF
                '*********************************************************
                'Suposicion 1,. TODOS son nacionales
                '*********************************************************
                Im = DBLet(RS!imppagad, "N")
                Im = RS!ImpEfect - Im
                AUX = RellenaABlancos(RS!nifdatos, True, 12)
                
                    
                'Reg 010
                AUX = "0656" & CodigoOrdenante & AUX & "010"
                AUX = AUX & RellenaAceros(CStr(Im * 100), False, 12)
                AUX = AUX & FrmtStr(DBLet(RS!Entidad, "N"), 4) & FrmtStr(DBLet(RS!oficina, "N"), 4)
                AUX = AUX & FrmtStr(DBLet(RS!cuentaba, "N"), 10) & "1" & "9" & "  " & FrmtStr(DBLet(RS!CC, "N"), 2)
                AUX = AUX & "N" & "C" & "EUR  "
                Print #NFich, AUX
                
        
           
           
                
                'OBligaorio 011   Nombre
                AUX = RellenaABlancos(RS!nifdatos, True, 12)
                AUX = "0656" & CodigoOrdenante & AUX & "011"
                AUX = AUX & FrmtStr(DBLet(RS!razosoci, "T"), 36) & Space(7)
                Print #NFich, AUX
           
                'OBligaorio 012   direccion
                AUX = RellenaABlancos(RS!nifdatos, True, 12)
                AUX = "0656" & CodigoOrdenante & AUX & "012"
                AUX = AUX & FrmtStr(DBLet(RS!dirdatos, "T"), 36) & Space(7)
                Print #NFich, AUX
           
                'OBligaorio 014   cpos provi
                AUX = RellenaABlancos(RS!nifdatos, True, 12)
                AUX = "0656" & CodigoOrdenante & AUX & "014"
                AUX = AUX & FrmtStr(DBLet(RS!codposta, "N"), 5) & FrmtStr(DBLet(RS!desPobla, "T"), 31) & Space(7)
                Print #NFich, AUX
                
                'OBligaorio 016   ID factura
                AUX = RellenaABlancos(RS!nifdatos, True, 12)
                AUX = "0656" & CodigoOrdenante & AUX & "016"
                AUX = AUX & "T" & Format(RS!FecFactu, "ddmmyy") & FrmtStr(RS!NumFactu, 15) & Format(RS!fecefect, "ddmmyy") & Space(15)
                Print #NFich, AUX
           
                 
        
               'Totales
               Importe = Importe + Im
               Regs = Regs + 1
               RS.MoveNext
        Wend
        'Imprimimos totales
        AUX = "08" & "56"
        AUX = AUX & CodigoOrdenante    'llevara tb la ID del socio
        AUX = AUX & Space(15)
        AUX = AUX & RellenaAceros(CStr(Int(Round(Importe * 100, 2))), False, 12)
        AUX = AUX & RellenaAceros(CStr((Regs)), False, 8)
        AUX = AUX & RellenaAceros(CStr((Regs * 5) + 4 + 1), False, 10)    '4 de cabecera + uno de totales
        AUX = RellenaABlancos(AUX, True, 72)
        Print #NFich, AUX
        
        
    End If
    RS.Close
    Set RS = Nothing
    Close (NFich)
    If Regs > 0 Then
        GeneraFicheroCaixaConfirming = True
    Else
        MsgBox "No se han leido registros en la tabla de pagos", vbExclamation
    End If
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description

End Function








'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'
'
'
'
'            SSSSSS         EEEEEEEE             PPPPPPP                 A
'           SS              EE                   PP     P               A A
'            SS             EE                   PP     P              A   A
'              SSS          EEEEEEEE             PPPPPPP              AAAAAAA
'                SS         EE                   PP                  A       A
'               SS          EE                   PP                 A         A
'           SSSSS           EEEEEEEE             PP                A           A
'
'
'
'
'
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
'************************************************************************************************
', DatosExtra As String
' N19Punto19  -> True.  19.14
'             -> False. 19.44


'SEPA XML:   Hay un modulo donde genera el fichero. Las comprobaciones iniciales son las mismas
'             para ambos modulos
'
' En funcion del parametro llamara a uno u a otro

'Si viene FECHACOBRO es que todos los vencimientos van a esa FECHA
'       si no , cada vto lleva su fecha

Private Function GrabarFicheroNorma19SEPA(NomFichero As String, Remesa As String, FecPre As String, TipoReferenciaCliente As Byte, Sufijo As String, FechaCobro As String, SEPA_EmpresasGraboNIF As Boolean, Norma19_15 As Boolean, FormatoXML As Boolean, esAnticipoCredito As Boolean) As Boolean
Dim B As Boolean
    '-- Genera_Remesa: Esta función genera la remesa indicada, en el fichero correspondiente

    Dim DatosBanco As String  'oficina,sucursla,cta, sufijo
    Dim NifEmpresa_ As String
    
    '-- Primero comprobamos que la remesa no haya sido enviada ya
    SQL = "SELECT * FROM remesas,bancos WHERE codigo = " & RecuperaValor(Remesa, 1)
    SQL = SQL & " AND anyo = " & RecuperaValor(Remesa, 2) & " AND remesas.codmacta = bancos.codmacta "
    
    Set miRsAux = New ADODB.Recordset
    DatosBanco = ""
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        If miRsAux!Situacion >= "C" Then
            MsgBox "La remesa ya figura como enviada", vbCritical
            
        Else
            'Cargo algunos de los datos de la remesa
            DatosBanco = miRsAux!IBAN
            
             'En datos extra dejo el CONCEPTO PPAL
             'DatosExtra = RecuperaValor(DatosExtra, 2)
        End If
    Else
        MsgBox "La remesa solicitada no existe", vbCritical
    End If
    miRsAux.Close
    
    If DatosBanco = "" Then Exit Function
    
    If Not comprobarCuentasBancariasRecibos(Remesa) Then Exit Function




    'Si es el campo referencia del fichero de cobros, entonces hay que comprobar que es obligado
    If TipoReferenciaCliente = 2 Then
        'Campo REFERENCAI como identificador
        If Not ComprobarCampoReferenciaRemesaNorma19(Remesa) Then Exit Function
    End If


    'Ahora cargare el NIF y la empresa
    SQL = "Select * from empresa2"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NifEmpresa_ = ""
    If Not miRsAux.EOF Then
        NifEmpresa_ = DBLet(miRsAux!nifempre, "T")
    End If
    miRsAux.Close
    If NifEmpresa_ = "" Then
        MsgBox "Datos empresa MAL configurados", vbExclamation
        Exit Function
    End If
    
    'Desde aqui, cada norma sigue su camino, generando un fichero al final
    
    If FormatoXML Then
        B = GrabarDisketteNorma19_SEPA_XML(NomFichero, Remesa, FecPre, TipoReferenciaCliente, Sufijo, FechaCobro, SEPA_EmpresasGraboNIF, Norma19_15, DatosBanco, NifEmpresa_, esAnticipoCredito)
    End If
    GrabarFicheroNorma19SEPA = B
End Function





'miRsAux no lo paso pq es GLOBAL
'TipoRegistro
'   0: Cabecera deudor
'   1. Total deudor/FECHA
'   2. Total deudor
'   3. Total general
Private Sub ImprimiSEPA_ProveedorFecha2(TipoRegistro As Byte, IdDeudorAcreedor As String, Fecha As Date, Registros003 As Integer, Suma As Currency, NumeroLineasTotalesSinCabceraPresentador As Integer, IdNorma As String)
Dim Cad As String

    Select Case TipoRegistro
    Case 0
        'Cabecera de ACREEDOR-FECHA
        Cad = "02" & IdNorma & "002"   '19143-> Podria ser 19154 ver pdf
        Cad = Cad & IdDeudorAcreedor
        
        'Fecha cobro
        Cad = Cad & Format(miRsAux!FecVenci, "yyyymmdd")
        
        'Nomprove
        Cad = Cad & DatosBasicosDelAcreedor
        'EN SQL llevamos el IBAN completo del acredor, es decir, de la empresa presentardora que le deben los deudores
        Cad = Cad & SQL & Space(10)  'El iban son 24 y dejan hasta 34 psociones
        '
        Cad = Cad & Space(301)
        
    Case 1
        'total x fecha -deudor
        Cad = "04"
        Cad = Cad & IdDeudorAcreedor

        'Fecha cobro
        Cad = Cad & Format(Fecha, "yyyymmdd")

        Cad = Cad & Right(String(17, "0") & (Suma * 100), 17) ' Suma total de registros
        Cad = Cad & Format(Registros003, "00000000")
        Cad = Cad & Format(NumeroLineasTotalesSinCabceraPresentador + 2, "0000000000") ' +cabecera y pie
        Cad = Cad & FrmtStr(" ", 520) ' LIBRE

        
        
    Case 2
        'total deudor
        Cad = "05"
        Cad = Cad & IdDeudorAcreedor

        Cad = Cad & Right(String(17, "0") & (Suma * 100), 17) ' Suma total de registros
        Cad = Cad & Format(Registros003, "00000000")
        Cad = Cad & Format(NumeroLineasTotalesSinCabceraPresentador + 2, "0000000000") '
        Cad = Cad & FrmtStr(" ", 528) ' LIBRE
      
    Case 3
        'total general
        Cad = "99"
        Cad = Cad & Right(String(17, "0") & (Suma * 100), 17) ' Suma total de registros
        Cad = Cad & Format(Registros003, "00000000")
        Cad = Cad & Format(NumeroLineasTotalesSinCabceraPresentador + 2, "0000000000") ' +cabecera y pie
        Cad = Cad & FrmtStr(" ", 563) ' LIBRE
      
    End Select
        
    Print #NF, Cad
        
        
End Sub

' AT-09.   70 + 50 + 50 + 40 +2
Private Function DatosBasicosDelDeudor() As String
        DatosBasicosDelDeudor = FrmtStr(miRsAux!Nommacta, 70)
        'dirdatos,codposta,despobla,pais desprovi
        DatosBasicosDelDeudor = DatosBasicosDelDeudor & FrmtStr(DBLet(miRsAux!dirdatos, "T"), 50)
        DatosBasicosDelDeudor = DatosBasicosDelDeudor & FrmtStr(Trim(DBLet(miRsAux!codposta, "T") & " " & DBLet(miRsAux!desPobla, "T")), 50)
        DatosBasicosDelDeudor = DatosBasicosDelDeudor & FrmtStr(DBLet(miRsAux!desProvi, "T"), 40)
        
        If IsNull(miRsAux!PAIS) Then
            DatosBasicosDelDeudor = DatosBasicosDelDeudor & "ES"
        Else
            DatosBasicosDelDeudor = DatosBasicosDelDeudor & Mid(miRsAux!PAIS, 1, 2)
        End If
End Function


'NUestros datos basicos
' AT-09.   70 + 50 + 50 + 40 +2
Private Function DatosBasicosDelAcreedor() As String
Dim RN As ADODB.Recordset

        'NO PUEDE SER EOF
        Set RN = New ADODB.Recordset
        RN.Open "Select * from empresa2", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText


        'siglasvia direccion  numero puerta  codpos poblacion provincia


        DatosBasicosDelAcreedor = FrmtStr(vEmpresa.nomempre, 70)
        'dirdatos,codposta,despobla,pais desprovi
        DatosBasicosDelAcreedor = DatosBasicosDelAcreedor & FrmtStr(Trim(DBLet(RN!siglasvia, "T") & " " & DBLet(RN!Direccion, "T") & ", " & DBLet(RN!numero, "T") & " " & DBLet(RN!puerta, "T")), 50)
        DatosBasicosDelAcreedor = DatosBasicosDelAcreedor & FrmtStr(Trim(DBLet(RN!codpos, "T") & " " & DBLet(RN!Poblacion, "T")), 50)
        DatosBasicosDelAcreedor = DatosBasicosDelAcreedor & FrmtStr(DBLet(RN!provincia, "T"), 40)
         
        DatosBasicosDelAcreedor = DatosBasicosDelAcreedor & "ES"
        
        
        RN.Close
        Set RN = Nothing
End Function





Private Sub ImprimeEnXML(Anidacion As Byte, Fich As Integer, Etiqueta As String)

End Sub











'---------------------------------------------------------------------
'  DEVOLUCION FICHERO  SEPA
'---------------------------
Public Sub ProcesaCabeceraFicheroDevolucionSEPA(Fichero As String, ByRef Remesa As String)
Dim aux2 As String  'Para buscar los vencimientos
Dim FinLecturaLineas As Boolean
Dim TodoOk As Boolean
Dim ErroresVto As String
Dim Cuantos As Integer
Dim Bien As Integer
Dim LinDelFichero As Collection


    On Error GoTo eProcesaCabeceraFicheroDevolucionSEPA
    Remesa = ""
    
    
    
    
    'ANTES nov 2012
    '
    'nf = FreeFile
    'Open Fichero For Input As #nf
    Registro = Fichero 'para no pasr mas variables al proceso
    ProcesoFicheroDevolucion 0, LinDelFichero 'abrir el fichero y volcarlo sobre un Collection
    
    'Proceso la primera linea. A veriguare a que norma pertenece
    ' y hallare la remesa
    'Line Input #nf, Registro
    ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
    
    'Comproamos ciertas cosas
    SQL = "Linea 1 vacia"
    If Registro <> "" Then
        
        
        
        'Tiene valor
        If Len(Registro) <> 600 Then
            SQL = "Longitud linea incorrecta(600)"
        Else
            'Febrero 2014
            'Devolucion:2119
            'Rechazo:   1119
            'Antes: Mid(Registro, 1, 4) <> "2119"
            
            If Mid(Registro, 2, 3) <> "119" Then
                SQL = "Cadena control incorrecta(?119)"
            Else
                SQL = ""
            End If
        End If
    End If
    
    If SQL = "" Then
    
        'Segunda LINEA.
        'Line Input #nf, Registro
        ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
        
        SQL = "Linea 2 vacia"
        If Registro <> "" Then
            
           
            
            
            'Tiene valor
            If Len(Registro) <> 600 Then
                SQL = "Longitud linea incorrecta(600)"
            Else
                'Devolucion:2219
                'Rechazo:   1119
                'Antes: Mid(Registro, 1, 4) <> "2119"
                
                If Mid(Registro, 2, 3) <> "219" Then
                    SQL = "Cadena control incorrecta(?219)"
                Else
                    
                    SQL = "Falta linea 2319"  'la que lleva los vtos
                    Remesa = ""
                    Do
                        ProcesoFicheroDevolucion 2, LinDelFichero  'vemos si es ultima linea
                        
                        If Registro <> "" Then
                            SQL = "FIN LINEAS. No se ha encontrado linea: 2319"
                            Remesa = "NO"
                        Else
                            'Line Input #nf, Registro
                            ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
                            
                            'BUsco la linea:
                            '5690
                            If Registro <> "" Then
                                '2319  Lleva los vtos
                                '1319 en devoluciones
                                If Mid(Registro, 2, 3) = "319" Then
                                    SQL = ""
                                    Remesa = "NO"
                                End If
                            End If
                        End If
                        
                    Loop Until Remesa <> ""
                    Remesa = ""
                    
                    If SQL = "" Then
                        'VAMOS BIEN. Veremos si a partir de los datos del recibo nos dan la remesa
                        'Para ello bucaremos en registro, la cadena que contiene los datos
                        'del vencimiento
                        'Registro=
                        '2319143003430000061 M  0330047820131201001   430000061 M  0330047820131201001
                        'sigue arriba RCURTRAD0000001210020091031CCRIES2AXXXCOANNA, COOP. V.                                                      CAMINO HONDO, 1                                   46820                                                                                     ES1IF46024493                          F46024493                          AES1830820134930330000488          TRADFACTURA: M-3300478 de Fecha 01 dic 2013                                                                                                     MD0120131230
                        Set miRsAux = New ADODB.Recordset
                        ErroresVto = ""
                        FinLecturaLineas = False
                        Cuantos = 0
                        Bien = 0
                        Do
                            'Devolucion:2319
                            'Rechazo:   1319
                            'Antes: Mid(Registro, 1, 4) <> "2119"
            
                            If Mid(Registro, 2, 3) = "319" Then
                                'Los vtos vienen en estas lineas
                                Cuantos = Cuantos + 1
                                Registro = Mid(Registro, 21, 35)
                                'M  0330047820131201001
                                SQL = "Select codrem,anyorem,siturem from cobros where fecfactu='" & Mid(Registro, 12, 4) & "-" & Mid(Registro, 16, 2) & "-" & Mid(Registro, 18, 2)
                                
                                SQL = SQL & "' AND numserie = '" & Trim(Mid(Registro, 1, 3)) & "' AND numfactu = " & Val(Mid(Registro, 4, 8)) & " AND numorden=" & Mid(Registro, 20, 3)
                                
                                
                                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                                TodoOk = False
                                SQL = "Vencimiento no encontrado: " & Registro
                                If Not miRsAux.EOF Then
                                    If IsNull(miRsAux!CodRem) Then
                                        SQL = "Vencimiento sin Remesa: " & Registro
                                    Else
                                        SQL = miRsAux!CodRem & "|" & miRsAux!AnyoRem & "|·"
                                        
                                        If InStr(1, Remesa, SQL) = 0 Then Remesa = Remesa & SQL
                                        SQL = ""
                                        TodoOk = True
                                    End If
                                End If
                                miRsAux.Close
                                
                               
                                
                                
                                
                                If SQL <> "" Then
                                    ErroresVto = ErroresVto & vbCrLf & SQL
                                Else
                                    Bien = Bien + 1
                                End If
                            Else
                                'La linea no empieza por 569
                                'veremos los totales
                                
                                If Mid(Registro, 1, 2) = "99" Then
                                    'TOTAL TOTAL
                                    SQL = Mid(Registro, 20, 8)
                                    If Val(SQL) <> Cuantos Then ErroresVto = "Fichero: " & SQL & "   Leidos" & Cuantos & vbCrLf & ErroresVto & vbCrLf & SQL
                                End If
                            End If
                            
                            'Siguiente linea
                            ProcesoFicheroDevolucion 2, LinDelFichero  'vemos si es ultima linea
                            
                            If Registro <> "" Then
                                FinLecturaLineas = True
                            Else
                                'Line Input #nf, Registro
                                ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
                            End If
                            
                        Loop Until FinLecturaLineas
                        
                        If Cuantos <> Bien Then ErroresVto = ErroresVto & vbCrLf & "Total: " & Cuantos & "   Correctos:" & Bien
                        
                        SQL = ErroresVto
                        Set miRsAux = Nothing
                    
                    End If
                End If  'Control SEGUNDA LINEA
        
        
            End If
        End If
    
    End If  'DE SEGUNDA LINEA
    
    ProcesoFicheroDevolucion 4, LinDelFichero
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
    Else
        'Remesa = Mid(Registro, 1, 4) & "|" & Mid(Registro, 5) & "|"
        
        
        'Ahora comprobaremos que para cada remesa  veremos si existe y si la situacion es la contabilizadxa
        SQL = Remesa
        Registro = "" 'Cadena de error de situacion remesas
        Set miRsAux = New ADODB.Recordset
        Do
            Cuantos = InStr(1, SQL, "·")
            If Cuantos = 0 Then
                SQL = ""
            Else
                aux2 = Mid(SQL, 1, Cuantos - 1)
                SQL = Mid(SQL, Cuantos + 1)
                
                
                'En aux2 tendre codrem|anñorem|
                aux2 = RecuperaValor(aux2, 1) & " AND anyo = " & RecuperaValor(aux2, 2)
                aux2 = "Select situacion from remesas where codigo = " & aux2
                miRsAux.Open aux2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If miRsAux.EOF Then
                    aux2 = "-No se encuentra remesa"
                Else
                    'Si que esta.
                    'Situacion
                    If CStr(miRsAux!Situacion) <> "Q" And CStr(miRsAux!Situacion) <> "Y" Then
                        aux2 = "- Situacion incorrecta : " & miRsAux!Situacion
                    Else
                        aux2 = "" 'TODO OK
                    End If
                End If
            
                If aux2 <> "" Then
                    aux2 = aux2 & " ->" & Mid(miRsAux.Source, InStr(1, UCase(miRsAux.Source), " WHERE ") + 7)
                    aux2 = Replace(aux2, " AND ", " ")
                    aux2 = Replace(aux2, "anyo", "año")
                    Registro = Registro & vbCrLf & aux2
                End If
                miRsAux.Close
            End If
        Loop Until SQL = ""
        Set miRsAux = Nothing
        
        
        If Registro <> "" Then
            Registro = "Error remesas " & vbCrLf & String(30, "=") & Registro
            MsgBox Registro, vbExclamation
            
            'Pongo REMESA=""
            Remesa = "" 'para que no continue el preoceso de devolucion
        End If
        
    End If
    
    Exit Sub
eProcesaCabeceraFicheroDevolucionSEPA:
    Remesa = ""
    MuestraError Err.Number, "Procesando fichero devolucion SEPA"
End Sub




Public Function EsFicheroDevolucionSEPA2(elpath As String) As Byte
Dim NF As Integer

    On Error GoTo eEsFicheroDevolucionSEPA
    EsFicheroDevolucionSEPA2 = 0   'N19 Antiquisima      1.- SEPA txt    2 SEPA xml
    NF = FreeFile
    Open elpath For Input As #NF
    If Not EOF(NF) Then
        Line Input #NF, SQL
        If SQL <> "" Then
            '                 DEVOLUCION                RECHAZO
            If LCase(Mid(SQL, 1, 5)) = "<?xml" Then
                EsFicheroDevolucionSEPA2 = 2
            Else
                If Mid(SQL, 1, 2) = "21" Or Mid(SQL, 1, 2) = "11" Then
                    EsFicheroDevolucionSEPA2 = 1
                Else
                    EsFicheroDevolucionSEPA2 = 0
                End If
            End If
        End If
    End If
    Close #NF
eEsFicheroDevolucionSEPA:
    Err.Clear
End Function
