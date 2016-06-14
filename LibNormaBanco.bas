Attribute VB_Name = "LinNormasBanco"
Option Explicit

    Dim NF As Integer
    Dim Registro As String
    Dim Sql As String

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
Public Function GrabarDisketteNorma19(NomFichero As String, Remesa As String, FecPre As String, DatosExtra As String, TipoReferenciaCliente As Byte, FecCobro2 As String, BancoEmiteDocumento As Boolean, SepaEmpresasGraboNIF As Boolean, N19_15 As Boolean, FormatoXML As Boolean) As Boolean

    
    If vParamT.NuevasNormasSEPA Then
                
        'GrabarDisketteNorma19 = GrabarDisketteNorma19SEPA(NomFichero, Remesa, FecPre, DatosExtra, TipoReferenciaCliente, FecCobro, BancoEmiteDocumento)
        GrabarDisketteNorma19 = GrabarFicheroNorma19SEPA(NomFichero, Remesa, FecPre, TipoReferenciaCliente, RecuperaValor(DatosExtra, 1), FecCobro2, SepaEmpresasGraboNIF, N19_15, FormatoXML)
    End If
End Function


Private Function GrabarDisketteNorma19NORMAL(NomFichero As String, Remesa As String, FecPre As String, DatosExtra As String, TipoReferenciaCliente As Byte, FecCobro As Date, BancoEmiteDocumento As Boolean) As Boolean
Dim ValorEnOpcionales As Boolean
    '-- Genera_Remesa: Esta función genera la remesa indicada, en el fichero correspondiente
    Dim mAux As String
    Dim SumaImportes As Currency
    Dim SumReg As Integer
    Dim SumTotal As Integer
    Dim DatosBanco As String  'oficina,sucursla,cta, sufijo
    Dim vSufijo As String
    Dim NifEmpresa As String
    Dim ImprimeOpc As Boolean
    Dim ValoresOpcionales As String
    Dim ImpEfe As Currency
    Dim J As Integer
    
    Dim msgSerie As String
    
    On Error GoTo Err_Remesa
    
    
    '-- Primero comprobamos que la remesa no haya sido enviada ya
    Sql = "SELECT * FROM remesas,ctabancaria WHERE codigo = " & RecuperaValor(Remesa, 1)
    Sql = Sql & " AND anyo = " & RecuperaValor(Remesa, 2) & " AND remesas.codmacta = ctabancaria.codmacta "
    
    Set miRsAux = New ADODB.Recordset
    DatosBanco = ""
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        If miRsAux!Situacion >= "C" Then
            MsgBox "La remesa ya figura como enviada", vbCritical
            
        Else
            'Cargo algunos de los datos de la remesa
            DatosBanco = Format(miRsAux!Entidad, "0000") & "|" & Format(miRsAux!Oficina, "0000") & "|" & Format(miRsAux!Control, "00") & "|" & Format(miRsAux!CtaBanco, "0000000000") & "|"
            vSufijo = RecuperaValor(DatosExtra, 1)
            If Trim(vSufijo) = "" Then vSufijo = Mid(miRsAux!sufijoem & "   ", 1, 3)
             'En datos extra dejo el CONCEPTO PPAL
             DatosExtra = RecuperaValor(DatosExtra, 2)
        End If
    Else
        MsgBox "La remesa solicitada no existe", vbCritical
    End If
    miRsAux.Close
    
    If DatosBanco = "" Then Exit Function
    
    If Not comprobarCuentasBancariasRecibos(Remesa) Then Exit Function
    
    If TipoReferenciaCliente = 3 Then
        'Campo REFERENCAI como identificador
        If Not ComprobarCampoReferenciaRemesaNorma19(Remesa) Then Exit Function
    End If
    
    
    'Ahora cargare el NIF y la empresa
    Sql = "Select * from empresa2"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NifEmpresa = ""
    If Not miRsAux.EOF Then
        NifEmpresa = DBLet(miRsAux!nifempre, "T")
    End If
    miRsAux.Close
    If NifEmpresa = "" Then
        MsgBox "Datos empresa MAL configurados", vbExclamation
        Exit Function
    End If
    
    
    '-- Abrir el fichero a enviar
    NF = FreeFile()
    Open NomFichero For Output As #NF
    
    Sql = "select  scobro.*,nommacta,nifdatos from scobro,cuentas where "
    Sql = Sql & " scobro.codmacta = cuentas.codmacta "
    Sql = Sql & " AND codrem = " & RecuperaValor(Remesa, 1)
    Sql = Sql & " AND anyorem=" & RecuperaValor(Remesa, 2)
    
    
    'EL ORDEN QUE QUERAMOS
    msgSerie = ""
    Remesa = RecuperaValor(Remesa, 1)
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        'Rs.MoveFirst
     
    
'        '-- Registro 5180
        Registro = "5180"
        Registro = Registro & FrmtStr(NifEmpresa, 9)   '-- Alinea NIF
        Registro = Registro & FrmtStr(vSufijo, 3) ' Sufijo
        Registro = Registro & Format(FecPre, "ddmmyy") ' Fecha de presentación
        Registro = Registro & FrmtStr(" ", 6) ' LIBRE
        Registro = Registro & FrmtStr(DatosExtra, 40)   ' Nombre del cliente presentador
        Registro = Registro & FrmtStr(" ", 20) ' LIBRE
        Registro = Registro & RecuperaValor(DatosBanco, 1)
        Registro = Registro & RecuperaValor(DatosBanco, 2)  ' Código de oficina receptora
        'IDENDIFICADOR DE REMESA
        '12 caracteres
        Registro = Registro & "RE" & Format(vEmpresa.codempre, "00") & Format(miRsAux!CodRem, "0000") & Format(miRsAux!AnyoRem, "0000")
        Registro = Registro & FrmtStr(" ", 54) ' LIBRE
        SumTotal = SumTotal + 1
        Print #NF, Registro
       

        
        
        '-- Registro 5380
        Registro = "5380"
        Registro = Registro & FrmtStr(NifEmpresa, 9)  '-- Alinea NIF
        Registro = Registro & FrmtStr(vSufijo, 3) ' Sufijo
        Registro = Registro & Format(FecPre, "ddmmyy") ' Fecha de confección del soporte
        Registro = Registro & Format(FecCobro, "ddmmyy") ' Fecha de cargo de recibos'
        Registro = Registro & FrmtStr(DatosExtra, 40)  ' Nombre del cliente presentador
        Registro = Registro & RecuperaValor(DatosBanco, 1) ' Código de entidad receptora
        Registro = Registro & RecuperaValor(DatosBanco, 2) ' Código de oficina receptora
        Registro = Registro & RecuperaValor(DatosBanco, 3) 'Dígitos de control
        Registro = Registro & RecuperaValor(DatosBanco, 4) ' Código de cuenta
        Registro = Registro & FrmtStr(" ", 8) ' LIBRE
        If BancoEmiteDocumento Then
            Registro = Registro & "01" ' Fijo 01   Procedimiento UNO. Emite documento
        Else
            Registro = Registro & "02" '    Procedimiento DOS. NO, repito NO, emite documento
        End If
        'Nuevo 24 Febrero 2006
        Registro = Registro & "RE" & Format(vEmpresa.codempre, "00") & Format(miRsAux!CodRem, "0000") & Format(miRsAux!AnyoRem, "0000")
        
        Registro = Mid(Registro & Space(100), 1, 162)
        SumTotal = SumTotal + 1
        Print #NF, Registro
        '-- Leemos secuencialmente las líneas de remesa
        While Not miRsAux.EOF
            'Tenemos k ver si imprimimos los opcionales
            If BancoEmiteDocumento Then
                ImprimeOpc = HayKImprimirOpcionales
            Else
                ImprimeOpc = False
            End If
        
        

            
            ValoresOpcionales = FrmtStr(NifEmpresa, 9)  '-- Alinea NIF
            ValoresOpcionales = ValoresOpcionales & FrmtStr(vSufijo, 3) ' Sufijo
            'Segun sea lo que quiera el cliente que le ponga como referencia
            'Opcion nueva: 3   Quiere el campo referencia de scobro
            Select Case TipoReferenciaCliente
            Case 1
                'ALZIRA. La referencia final de 12 es el ctan bancaria del cli + su CC
                    Registro = Format(miRsAux!digcontr, "00") ' Dígitos de control
                    Registro = Registro & Format(miRsAux!Cuentaba, "0000000000") ' Código de cuenta
            Case 2
                'NIF
                Registro = DBLet(miRsAux!nifdatos, "T")
                If Registro = "" Then Registro = miRsAux!codmacta
                Registro = Mid(Registro & Space(12), 1, 12)
                
            Case 3
                'Referencia en el VTO. No es Nula
                Registro = Space(12) & miRsAux!referencia
                Registro = Right(Registro, 12)
            Case Else
                'Antes
                'Registro = miRsAux!NUmSerie & Format(miRsAux!codfaccl, "0000000000") & Format(miRsAux!numorden, "0")
                Registro = miRsAux!codmacta
                Registro = Right("0000000000" & Registro, 12)
            End Select
            Registro = Mid(Registro, 1, 12)
            ValoresOpcionales = ValoresOpcionales & Registro
            
            'Registro = Registro & FrmtStr(NifEmpresa, 9)  '-- Alinea NIF
            'Registro = Registro & FrmtStr(vSufijo, 3) ' Sufijo
            'Registro = Registro & FrmtStr(miRsAux!NUmSerie, 3) & FrmtStr(miRsAux!codfaccl, 7) & "-" & FrmtStr(miRsAux!numorden, 1)
            
            
            '-- Registro 5680
            Registro = "5680"
            Registro = Registro & ValoresOpcionales
            
            
            
            Registro = Registro & FrmtStr(DevNombreSQL(miRsAux!Nommacta), 40)
            Registro = Registro & Format(miRsAux!codbanco, "0000") ' Código de entidad receptora
            Registro = Registro & Format(miRsAux!codsucur, "0000") ' Código de oficina receptora
            Registro = Registro & Format(miRsAux!digcontr, "00") ' Dígitos de control
            Registro = Registro & Format(miRsAux!Cuentaba, "0000000000") ' Código de cuenta
            
            ImpEfe = DBLet(miRsAux!Gastos, "N")
            ImpEfe = miRsAux!ImpVenci + ImpEfe
            Registro = Registro & Format(ImpEfe * 100, String(10, "0")) ' Importe
            
            
            Registro = Registro & Format(miRsAux!fecfaccl, "ddmmyy")  ' Identificador de domiciliación
            'Antes sept 2011
            'Registro = Registro & miRsAux!NUmSerie & FrmtStr(Format(miRsAux!codfaccl, "00000000"), 8) & Format(miRsAux!numorden, "0")
            
            'Oct 2011
            'PROBLEMA GRANDE
            'Tenemos 10 caracteres para identificar el vto
            'DOS para la serie, 7 para la factura y 1 para el vto
            
            mAux = Mid(miRsAux!NUmSerie & "  ", 1, 2)
            mAux = mAux & Format(miRsAux!codfaccl, "0000000")
            Registro = Registro & mAux & Format(miRsAux!numorden, "0")
            
            'Registro = Registro & FrmtStr(mAux, 10) ' Identificador de devolución
            mAux = DBLet(miRsAux!text33csb, "T")
            If mAux = "" Then mAux = "FACTURA: " & miRsAux!NUmSerie & "-" & miRsAux!codfaccl & " de Fecha " & Format(miRsAux!fecfaccl, "dd/mm/yyyy")
            
            Registro = Registro & FrmtStr(mAux, 40) ' Primer Concepto
            Registro = Registro & Format(miRsAux!FecVenci, "ddmmyy") & "  "
            Print #NF, Registro
            SumReg = SumReg + 1
            SumTotal = SumTotal + 1
            SumaImportes = SumaImportes + ImpEfe
            
            If ImprimeOpc Then
                For J = 1 To 5
                    Registro = ImprimeOpcionales(True, ValoresOpcionales, J, ValorEnOpcionales)
                    If ValorEnOpcionales Then
                        Print #NF, Registro
                        SumTotal = SumTotal + 1
                    End If
                Next J
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        
        '-- Registro 5880
        Registro = "5880"
        Registro = Registro & FrmtStr(NifEmpresa, 9)  '-- Alinea NIF
        Registro = Registro & FrmtStr(vSufijo, 3) ' Sufijo
        Registro = Registro & FrmtStr(" ", 72) ' LIBRE
        Registro = Registro & Format(SumaImportes * 100, String(10, "0")) ' Suma de importes
        Registro = Registro & FrmtStr(" ", 6) ' LIBRE
        Registro = Registro & Format(SumReg, String(10, "0")) ' Suma de registros 0680
        Registro = Registro & Format(SumTotal, String(10, "0")) ' Suma total de registros
        Registro = Registro & FrmtStr(" ", 38) ' LIBRE
        SumTotal = SumTotal + 1
        Print #NF, Registro
        '-- Registro 5980
        Registro = "5980"
        Registro = Registro & FrmtStr(NifEmpresa, 9)  '-- Alinea NIF
        Registro = Registro & FrmtStr(vSufijo, 3) ' Sufijo
        Registro = Registro & FrmtStr(" ", 52) ' LIBRE
        Registro = Registro & "0001" ' Suma de ordenantes (siempre es uno)
        Registro = Registro & FrmtStr(" ", 16) ' LIBRE
        Registro = Registro & Format(SumaImportes * 100, String(10, "0")) ' Suma de importes
        Registro = Registro & FrmtStr(" ", 6) ' LIBRE
        Registro = Registro & Format(SumReg, String(10, "0"))  ' Suma de registros 0680
        SumTotal = SumTotal + 1
        Registro = Registro & Format(SumTotal, String(10, "0")) ' Suma total de registros
        Registro = Registro & FrmtStr(" ", 38) ' LIBRE
        Print #NF, Registro
    End If
    Close #NF
    If SumTotal > 0 Then GrabarDisketteNorma19NORMAL = True
    Exit Function
Err_Remesa:
    MsgBox "Err: " & Err.Number & vbCrLf & _
        Err.Description, vbCritical, "Grabación del diskette de Remesa"
        
End Function











'Agruparemos para cada FECHA
Private Function GrabarDisketteNorma19FECHAS(NomFichero As String, Remesa2 As String, FecPre As String, DatosExtra As String, TipoReferenciaCliente_ As Byte, FecCobro As Date, BancoEmiteDocumento As Boolean) As Boolean
Dim ValorEnOpcionales As Boolean
    '-- Genera_Remesa: Esta función genera la remesa indicada, en el fichero correspondiente
    Dim mAux As String
    Dim SumaImportes As Currency
    Dim SumReg As Integer
    Dim SumTotal As Integer
    Dim DatosBanco As String  'oficina,sucursla,cta, sufijo
    Dim vSufijo As String
    Dim NifEmpresa As String
    Dim ImprimeOpc As Boolean
    Dim ValoresOpcionales As String
    Dim ImpEfe As Currency
    Dim J As Integer
    
    Dim msgSerie As String
    
    Dim ColFechas As Collection
    Dim Z As Integer
    Dim FCargo As Date
    Dim TotalFecha As Currency
    Dim NumFec As Integer

    
    On Error GoTo Err_Remesa
    
    
    
    
    
    '-- Primero comprobamos que la remesa no haya sido enviada ya
    Sql = "SELECT * FROM remesas,ctabancaria WHERE codigo = " & RecuperaValor(Remesa2, 1)
    Sql = Sql & " AND anyo = " & RecuperaValor(Remesa2, 2) & " AND remesas.codmacta = ctabancaria.codmacta "
    
    Set miRsAux = New ADODB.Recordset
    DatosBanco = ""
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        If miRsAux!Situacion >= "C" Then
            MsgBox "La remesa ya figura como enviada", vbCritical
            
        Else
            'Cargo algunos de los datos de la remesa
            DatosBanco = Format(miRsAux!Entidad, "0000") & "|" & Format(miRsAux!Oficina, "0000") & "|" & Format(miRsAux!Control, "00") & "|" & Format(miRsAux!CtaBanco, "0000000000") & "|"
            vSufijo = RecuperaValor(DatosExtra, 1)
            If Trim(vSufijo) = "" Then vSufijo = Mid(miRsAux!sufijoem & "   ", 1, 3)
             'En datos extra dejo el CONCEPTO PPAL
             DatosExtra = RecuperaValor(DatosExtra, 2)
        End If
    Else
        MsgBox "La remesa solicitada no existe", vbCritical
    End If
    miRsAux.Close
    
    If DatosBanco = "" Then Exit Function
    
    If Not comprobarCuentasBancariasRecibos(Remesa2) Then Exit Function
    
    If TipoReferenciaCliente_ = 3 Then
        'Campo REFERENCAI como identificador
        If Not ComprobarCampoReferenciaRemesaNorma19(Remesa2) Then Exit Function
    End If
    
    
    
    
    
    
    'Ahora cargare el NIF y la empresa
    Sql = "Select * from empresa2"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NifEmpresa = ""
    If Not miRsAux.EOF Then
        NifEmpresa = DBLet(miRsAux!nifempre, "T")
    End If
    miRsAux.Close
    If NifEmpresa = "" Then
        MsgBox "Datos empresa MAL configurados", vbExclamation
        Exit Function
    End If
    
    
    'Vamos a ver las fechas de presentacion
    Sql = "SELECT fecvenci FROM scobro WHERE codrem = " & RecuperaValor(Remesa2, 1)
    Sql = Sql & " AND anyorem = " & RecuperaValor(Remesa2, 2) & " GROUP BY 1"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set ColFechas = New Collection
    While Not miRsAux.EOF
        ColFechas.Add CStr(miRsAux!FecVenci)
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If ColFechas.Count = 0 Then
        MsgBox "ninguna fecha encontrada", vbExclamation
        Exit Function
    End If
    
    
    
    '-- Abrir el fichero a enviar
    NF = FreeFile()
    Open NomFichero For Output As #NF
    
    
    
    
        'Primer registro
        '-- Registro 5180
        Registro = "5180"
        Registro = Registro & FrmtStr(NifEmpresa, 9)   '-- Alinea NIF
        Registro = Registro & FrmtStr(vSufijo, 3) ' Sufijo
        Registro = Registro & Format(FecPre, "ddmmyy") ' Fecha de presentación
        Registro = Registro & FrmtStr(" ", 6) ' LIBRE
        Registro = Registro & FrmtStr(DatosExtra, 40)   ' Nombre del cliente presentador
        Registro = Registro & FrmtStr(" ", 20) ' LIBRE
        Registro = Registro & RecuperaValor(DatosBanco, 1)
        Registro = Registro & RecuperaValor(DatosBanco, 2)  ' Código de oficina receptora
        'IDENDIFICADOR DE REMESA
        '12 caracteres
        Registro = Registro & "RE" & Format(vEmpresa.codempre, "00") & Format(RecuperaValor(Remesa2, 1), "0000") & Format(RecuperaValor(Remesa2, 2), "0000")
        Registro = Registro & FrmtStr(" ", 54) ' LIBRE
        SumTotal = SumTotal + 1
        Print #NF, Registro
       

    
    
    For Z = 1 To ColFechas.Count
        NumFec = 0
        FCargo = CDate(ColFechas.Item(Z))
        TotalFecha = 0
        
    
        Sql = "select  scobro.*,nommacta,nifdatos from scobro,cuentas where "
        Sql = Sql & " scobro.codmacta = cuentas.codmacta "
        Sql = Sql & " AND codrem = " & RecuperaValor(Remesa2, 1)
        Sql = Sql & " AND anyorem=" & RecuperaValor(Remesa2, 2)
        Sql = Sql & " AND fecvenci='" & Format(FCargo, FormatoFecha) & "'"
    
    
    
        'EL ORDEN QUE QUERAMOS
        msgSerie = ""
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
     
   
        
        
        '-- Registro 5380
        Registro = "5380"
        Registro = Registro & FrmtStr(NifEmpresa, 9)  '-- Alinea NIF
        Registro = Registro & FrmtStr(vSufijo, 3) ' Sufijo
        Registro = Registro & Format(FecPre, "ddmmyy") ' Fecha de confección del soporte
        
        'Registro = Registro & Format(FecCobro, "ddmmyy") ' Fecha de cargo de recibos'
        Registro = Registro & Format(FCargo, "ddmmyy") ' Fecha de cargo de recibos'
        
        Registro = Registro & FrmtStr(DatosExtra, 40)  ' Nombre del cliente presentador
        Registro = Registro & RecuperaValor(DatosBanco, 1) ' Código de entidad receptora
        Registro = Registro & RecuperaValor(DatosBanco, 2) ' Código de oficina receptora
        Registro = Registro & RecuperaValor(DatosBanco, 3) 'Dígitos de control
        Registro = Registro & RecuperaValor(DatosBanco, 4) ' Código de cuenta
        Registro = Registro & FrmtStr(" ", 8) ' LIBRE
        If BancoEmiteDocumento Then
            Registro = Registro & "01" ' Fijo 01   Procedimiento UNO. Emite documento
        Else
            Registro = Registro & "02" '    Procedimiento DOS. NO, repito NO, emite documento
        End If
        'Nuevo 24 Febrero 2006
        Registro = Registro & "RE" & Format(vEmpresa.codempre, "00") & Format(miRsAux!CodRem, "0000") & Format(miRsAux!AnyoRem, "0000")
        
        Registro = Mid(Registro & Space(100), 1, 162)
        SumTotal = SumTotal + 1
        Print #NF, Registro
        '-- Leemos secuencialmente las líneas de remesa
        While Not miRsAux.EOF
            'Tenemos k ver si imprimimos los opcionales
            If BancoEmiteDocumento Then
                ImprimeOpc = HayKImprimirOpcionales
            Else
                ImprimeOpc = False
            End If
        
        

            
            ValoresOpcionales = FrmtStr(NifEmpresa, 9)  '-- Alinea NIF
            ValoresOpcionales = ValoresOpcionales & FrmtStr(vSufijo, 3) ' Sufijo
            'Segun sea lo que quiera el cliente que le ponga como referencia
            'Opcion nueva: 3   Quiere el campo referencia de scobro
            Select Case TipoReferenciaCliente_
            Case 1
                'ALZIRA. La referencia final de 12 es el ctan bancaria del cli + su CC
                    Registro = Format(miRsAux!digcontr, "00") ' Dígitos de control
                    Registro = Registro & Format(miRsAux!Cuentaba, "0000000000") ' Código de cuenta
            Case 2
                'NIF
                Registro = DBLet(miRsAux!nifdatos, "T")
                If Registro = "" Then Registro = miRsAux!codmacta
                Registro = Mid(Registro & Space(12), 1, 12)
                
            Case 3
                'Referencia en el VTO. No es Nula
                Registro = Space(12) & miRsAux!referencia
                Registro = Right(Registro, 12)
            Case Else
                'Antes
                'Registro = miRsAux!NUmSerie & Format(miRsAux!codfaccl, "0000000000") & Format(miRsAux!numorden, "0")
                Registro = miRsAux!codmacta
                Registro = Right("0000000000" & Registro, 12)
            End Select
            Registro = Mid(Registro, 1, 12)
            ValoresOpcionales = ValoresOpcionales & Registro
            
            'Registro = Registro & FrmtStr(NifEmpresa, 9)  '-- Alinea NIF
            'Registro = Registro & FrmtStr(vSufijo, 3) ' Sufijo
            'Registro = Registro & FrmtStr(miRsAux!NUmSerie, 3) & FrmtStr(miRsAux!codfaccl, 7) & "-" & FrmtStr(miRsAux!numorden, 1)
            
            
            '-- Registro 5680
            Registro = "5680"
            Registro = Registro & ValoresOpcionales
            
            
            
            Registro = Registro & FrmtStr(DevNombreSQL(miRsAux!Nommacta), 40)
            Registro = Registro & Format(miRsAux!codbanco, "0000") ' Código de entidad receptora
            Registro = Registro & Format(miRsAux!codsucur, "0000") ' Código de oficina receptora
            Registro = Registro & Format(miRsAux!digcontr, "00") ' Dígitos de control
            Registro = Registro & Format(miRsAux!Cuentaba, "0000000000") ' Código de cuenta
            
            ImpEfe = DBLet(miRsAux!Gastos, "N")
            ImpEfe = miRsAux!ImpVenci + ImpEfe
            Registro = Registro & Format(ImpEfe * 100, String(10, "0")) ' Importe
            
            
            Registro = Registro & Format(miRsAux!fecfaccl, "ddmmyy")  ' Identificador de domiciliación
        
            
            mAux = Mid(miRsAux!NUmSerie & "  ", 1, 2)
            mAux = mAux & Format(miRsAux!codfaccl, "0000000")
            Registro = Registro & mAux & Format(miRsAux!numorden, "0")
            
            'Registro = Registro & FrmtStr(mAux, 10) ' Identificador de devolución
            mAux = DBLet(miRsAux!text33csb, "T")
            If mAux = "" Then mAux = "FACTURA: " & miRsAux!NUmSerie & "-" & miRsAux!codfaccl & " de Fecha " & Format(miRsAux!fecfaccl, "dd/mm/yyyy")
            
            Registro = Registro & FrmtStr(mAux, 40) ' Primer Concepto
            Registro = Registro & Format(miRsAux!FecVenci, "ddmmyy") & "  "
            Print #NF, Registro
            SumReg = SumReg + 1
            SumTotal = SumTotal + 1
            SumaImportes = SumaImportes + ImpEfe
            TotalFecha = TotalFecha + ImpEfe
            NumFec = NumFec + 1
            
            If ImprimeOpc Then
                For J = 1 To 5
                    Registro = ImprimeOpcionales(True, ValoresOpcionales, J, ValorEnOpcionales)
                    If ValorEnOpcionales Then
                        Print #NF, Registro
                        SumTotal = SumTotal + 1
                    End If
                Next J
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
        
        '-- Registro 5880
        Registro = "5880"
        Registro = Registro & FrmtStr(NifEmpresa, 9)  '-- Alinea NIF
        Registro = Registro & FrmtStr(vSufijo, 3) ' Sufijo
        Registro = Registro & FrmtStr(" ", 72) ' LIBRE
        Registro = Registro & Format(TotalFecha * 100, String(10, "0")) ' Suma de importes
        Registro = Registro & FrmtStr(" ", 6) ' LIBRE
        Registro = Registro & Format(NumFec, String(10, "0")) ' Suma de registros 0680
        Registro = Registro & Format(NumFec + 2, String(10, "0")) ' Suma + cabcera y pie
        Registro = Registro & FrmtStr(" ", 38) ' LIBRE
        SumTotal = SumTotal + 1
        Print #NF, Registro
      
    Next
    
      'TOTAL TOTAL
      '-- Registro 5980
        Registro = "5980"
        Registro = Registro & FrmtStr(NifEmpresa, 9)  '-- Alinea NIF
        Registro = Registro & FrmtStr(vSufijo, 3) ' Sufijo
        Registro = Registro & FrmtStr(" ", 52) ' LIBRE
        
        'Total ordenantes cambia ahora
        'Registro = Registro & "0001" ' Suma de ordenantes (siempre es uno)
        Registro = Registro & Format(ColFechas.Count, "0000")
        
        
        Registro = Registro & FrmtStr(" ", 16) ' LIBRE
        Registro = Registro & Format(SumaImportes * 100, String(10, "0")) ' Suma de importes
        Registro = Registro & FrmtStr(" ", 6) ' LIBRE
        Registro = Registro & Format(SumReg, String(10, "0"))  ' Suma de registros 0680
        SumTotal = SumTotal + 1
        Registro = Registro & Format(SumTotal, String(10, "0")) ' Suma total de registros
        Registro = Registro & FrmtStr(" ", 38) ' LIBRE
        Print #NF, Registro
    
    
    Set ColFechas = Nothing
    Set miRsAux = Nothing
    Close #NF
    If SumTotal > 0 Then GrabarDisketteNorma19FECHAS = True
    Exit Function
Err_Remesa:
    MsgBox "Err: " & Err.Number & vbCrLf & _
        Err.Description, vbCritical, "Grabación del diskette de Remesa"
        
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

    Sql = "select * from cobros where codrem = " & RecuperaValor(Remesa, 1)
    Sql = Sql & " AND anyorem=" & RecuperaValor(Remesa, 2)
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    NF = 0
    While Not miRsAux.EOF

        If IsNull(miRsAux!Entidad) Or IsNull(miRsAux!Control) Then
            'Ya esta mal
            Sql = ""
        Else
            If IsNull(miRsAux!Cuentaba) Or IsNull(miRsAux!Control) Then
                'mal tb
                Sql = ""
            Else
                'TIENE DATOS
                Sql = "D"
            End If
        End If
    
        If Sql = "" Then
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
        Sql = "Los siguientes vencimientos no tienen la cuenta bancaria con todos los datos." & vbCrLf & Registro
        MsgBox Sql, vbExclamation
        Exit Function
    End If
    
    
    'Si llega aqui es que todos tienen DATOS
    Sql = "select entidad,oficina,cuentaba,control from cobros where codrem = " & RecuperaValor(Remesa, 1)
    Sql = Sql & " AND anyorem=" & RecuperaValor(Remesa, 2)
    Sql = Sql & " GROUP BY entidad,oficina,cuentaba,control"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    While Not miRsAux.EOF
                Sql = Format(miRsAux!Entidad, "0000")  ' Código de entidad receptora
                Sql = Sql & Format(miRsAux!Oficina, "0000") ' Código de oficina receptora
                
                Sql = Sql & Format(miRsAux!Cuentaba, "0000000000") ' Código de cuenta
                
                CC = Format(miRsAux!Control, "00") ' Dígitos de control
                
                'Este lo mando.
                Sql = CodigoDeControl(Sql)
                If Sql <> CC Then
                    
                    Sql = " - " & Format(miRsAux!Control, "00") & "- " & Format(miRsAux!Cuentaba, "0000000000") & " --> CC. correcto:" & Sql
                    Sql = Format(miRsAux!entridad, "0000") & " - " & Format(miRsAux!Oficina, "0000") & Sql
                    Registro = Registro & Sql & vbCrLf
                End If
                miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Registro <> "" Then
        Sql = "Las siguientes cuentas no son correctas.:" & vbCrLf & Registro
        If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    
    
    If vParamT.NuevasNormasSEPA Then
        'Si continuar y esta bien, veremos si todas los bancos tienen BIC asociado
        Registro = ""
        Sql = "select cobros.entidad codbanco,bics.entidad from cobros left join bics on cobros.entidad=bics.entidad WHERE "
        Sql = Sql & " codrem = " & RecuperaValor(Remesa, 1)
        Sql = Sql & " AND anyorem=" & RecuperaValor(Remesa, 2) & " group by 1"
        miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Registro = ""
        While Not miRsAux.EOF
            If IsNull(miRsAux!Entidad) Then Registro = Registro & "/    " & miRsAux!codbanco & "    "
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If Registro <> "" Then
            Registro = Mid(Registro, 2) & vbCrLf & vbCrLf & "¿Continuar?"
            Sql = "Las siguientes bancos no tiene BIC asocidado:" & vbCrLf & vbCrLf & Registro
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Function
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
    Sql = "select codmacta,NUmSerie,numfactu,numorden,referencia from cobros where codrem = " & RecuperaValor(Remesa, 1)
    Sql = Sql & " AND anyorem=" & RecuperaValor(Remesa, 2) & " ORDER BY codmacta"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    Sql = ""
    NF = 0
    While Not miRsAux.EOF
        If DBLet(miRsAux!referencia, "T") = "" Then
            Registro = Registro & miRsAux!codmacta & " - " & miRsAux!NUmSerie & "/" & miRsAux!NumFactu & "-" & miRsAux!numorden & vbCrLf
            NF = NF + 1
        Else
            If Len(miRsAux!referencia) > 12 Then Sql = Sql & miRsAux!codmacta & " - " & miRsAux!NUmSerie & "/" & miRsAux!NumFactu & "-" & miRsAux!numorden & "(" & miRsAux!referencia & ")" & vbCrLf
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If NF > 0 Then
        Registro = "Referencias vacias: " & NF & vbCrLf & vbCrLf & Registro
        MsgBox Registro, vbExclamation
    Else
        If Sql <> "" Then
            Registro = "Longitud referencia incorrecta: " & vbCrLf & vbCrLf & Sql
            Registro = Registro & vbCrLf & "¿Continuar?"
            If MsgBox(Registro, vbQuestion + vbYesNo) = vbNo Then Exit Function
        End If
        ComprobarCampoReferenciaRemesaNorma19 = True
    End If
End Function



Public Function comprobarCuentasBancariasPagos(Transferencia As String, Pagos As Boolean) As Boolean
Dim CC As String
Dim IBAN As String
On Error GoTo EcomprobarCuentasBancariasPagos

    comprobarCuentasBancariasPagos = False
    If Pagos Then
        Sql = "select * from spagop where transfer = " & Transferencia
    Else
        'ABONOS
        'numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci,
        'ctabanc1, codbanco, codsucur, digcontr, cuentaba,
        'ctabanc2, fecultco, impcobro, emitdocum, recedocu, contdocu,
        'ultimareclamacion, agente, departamento, codrem, anyorem, siturem, gastos,
        'Devuelto, situacionjuri, noremesar, obs, transfer)
        Sql = "Select numserie, codfaccl, fecfaccl, numorden, codmacta as ctaprove, "
        Sql = Sql & "codbanco as entidad,codsucur as oficina,cuentaba,digcontr as CC"
        Sql = Sql & " FROM scobro where transfer=" & Transferencia
        
    End If
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    NF = 0
    While Not miRsAux.EOF

        If DBLet(IsNull(miRsAux!Entidad), "T") = "" Or DBLet(miRsAux!Oficina, "T") = "" Then
            'Ya esta mal
            Sql = ""
        Else
            If IsNull(miRsAux!Cuentaba) Or IsNull(miRsAux!CC) Then
                'mal tb
                Sql = ""
            Else
                'TIENE DATOS
                Sql = "D"
            End If
        End If
    
        If Sql = "" Then
             If Pagos Then
                Registro = Registro & miRsAux!ctaprove & " - " & miRsAux!NumFactu & " : " & miRsAux!FecFactu & "-" & miRsAux!numorden
             Else
                Registro = Registro & miRsAux!ctaprove & " - " & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000000") & " : " & miRsAux!fecfaccl & "-" & miRsAux!numorden
             End If
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
        Sql = "Los siguientes vencimientos no tienen la cuenta bancaria con todos los datos." & vbCrLf & Registro
        MsgBox Sql, vbExclamation
        Exit Function
    End If
    
    
    'Si llega aqui es que todos tienen DATOS
    If Pagos Then
        Sql = "select entidad,oficina,cuentaba,cc,iban from spagop where transfer = " & Transferencia
        Sql = Sql & " GROUP BY entidad,oficina,cuentaba,cc"
    Else
        Sql = "SELECT codbanco as entidad,codsucur as oficina,cuentaba,digcontr as CC,iban"
        Sql = Sql & " FROM scobro where transfer=" & Transferencia
        Sql = Sql & " GROUP BY entidad,oficina,cuentaba,cc"
    End If
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    While Not miRsAux.EOF
                Sql = Format(miRsAux!Entidad, "0000")  ' Código de entidad receptora
                Sql = Sql & Format(miRsAux!Oficina, "0000") ' Código de oficina receptora
                
                Sql = Sql & Format(miRsAux!Cuentaba, "0000000000") ' Código de cuenta
                
                CC = Format(miRsAux!CC, "00") ' Dígitos de control
                
                'Este lo mando.
                IBAN = Mid(Sql, 1, 8) & CC & Mid(Sql, 9)
                
                Sql = CodigoDeControl(Sql)
                If Sql <> CC Then
                    
                    Sql = " - " & Format(miRsAux!CC, "00") & "- " & Format(miRsAux!Cuentaba, "0000000000") & " --> CC. correcto:" & Sql
                    Sql = Format(miRsAux!Entidad, "0000") & " - " & Format(miRsAux!Oficina, "0000") & Sql
                    Registro = Registro & Sql & vbCrLf
                End If
                
                
                'Noviembre 2013
                'IBAN
                If vParamT.NuevasNormasSEPA Then
                        Sql = "ES"
                        If DBLet(miRsAux!IBAN, "T") <> "" Then Sql = Mid(miRsAux!IBAN, 1, 2)
                    
                
                        If Not DevuelveIBAN2(Sql, IBAN, IBAN) Then
                            
                            Sql = "Error calculo"
                        Else
                            Sql = Sql & IBAN
                            If DBLet(miRsAux!IBAN, "T") <> Sql Then
                                Sql = "Error IBAN. Calculado " & Sql & " / " & DBLet(miRsAux!IBAN, "T")
                            Else
                                'OK
                                Sql = ""
                            End If
                        End If
                        
                        If Sql <> "" Then
                            Sql = Sql & " - " & Format(miRsAux!CC, "00") & "- " & Format(miRsAux!Cuentaba, "0000000000")
                            Sql = Format(miRsAux!Entidad, "0000") & " - " & Format(miRsAux!Oficina, "0000") & Sql
                            Registro = Registro & "Error obteniendo IBAN: " & Sql & vbCrLf
                        End If
                End If
                
                
                miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If Registro <> "" Then
        Sql = "Generando diskette." & vbCrLf & vbCrLf
        Sql = Sql & "Las siguientes cuentas no son correctas.:" & vbCrLf & Registro
        Sql = Sql & vbCrLf & "¿Desea continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    comprobarCuentasBancariasPagos = True
    Exit Function
EcomprobarCuentasBancariasPagos:
    MuestraError Err.Number, "comprobar Cuentas Bancarias pagos"
End Function

'Public Function CodigoDeControl(ByVal strBanOfiCuenta As String) As String
'
'Dim conPesos
'Dim lngPrimerCodigo As Long, lngSegundoCodigo As Long
'Dim I As Long, J As Long
'conPesos = "06030709100508040201"
'J = 1
'lngPrimerCodigo = 0
'lngSegundoCodigo = 0
'
'' Banco(4) + Oficina(4) nos dará el primer dígito de control
'For I = 8 To 1 Step -1
'  lngPrimerCodigo = lngPrimerCodigo + (Mid$(strBanOfiCuenta, I, 1) * Mid$(conPesos, J, 2))
'  J = J + 2
'Next I
'
'J = 1 ' reiniciar el contador de pesos
'
'' Número de cuenta nos dará el segundo digito de control
'For I = 18 To 9 Step -1
'  lngSegundoCodigo = lngSegundoCodigo + (Mid$(strBanOfiCuenta, I, 1) * Mid$(conPesos, J, 2))
'  J = J + 2
'Next I
'
'
'' ajustar el primer dígito de control
'lngPrimerCodigo = 11 - (lngPrimerCodigo Mod 11)
'If lngPrimerCodigo = 11 Then
'    lngPrimerCodigo = 0
'ElseIf lngPrimerCodigo = 10 Then
'    lngPrimerCodigo = 1
'End If
'
'
'' ajustar el segundo dígito de control
'lngSegundoCodigo = 11 - (lngSegundoCodigo Mod 11)
'If lngSegundoCodigo = 11 Then
'    lngSegundoCodigo = 0
'ElseIf lngSegundoCodigo = 10 Then
'    lngSegundoCodigo = 1
'End If
'
'' convertirlos en cadenas y concatenarlos
'CodigoDeControl = Format(lngPrimerCodigo) & Format(lngSegundoCodigo)
'
'End Function
'

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
            Sql = vbCrLf 'separaremos por este
        Else
            B = InStr(1, Registro, vbCr) > 0
            If B Then
                Sql = vbCr
            Else
                B = InStr(1, Registro, vbLf)
                If B Then Sql = vbLf
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
                NumRegElim = InStr(1, Registro, Sql)
                If NumRegElim = 0 Then
                    'NO DEBERIA PASAR
                    MsgBox "Preproceso fichero banco. Numregelim=0.  Avise soporte tecnico", vbExclamation
                Else

                    LinFichero.Add Mid(Registro, 1, NumRegElim - 1)
                    NumRegElim = NumRegElim + Len(Sql)
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
    Sql = "Linea 1 vacia"
    If Registro <> "" Then
        
        'NIF
        Sql = Mid(Registro, 5, 9)
        
        'Tiene valor
        If Len(Registro) <> 162 Then
            Sql = "Longitud linea incorrecta(162)"
        Else
            'Noviembre 2012
            'en lugar de 5190 comprobamos que sea 519
            If Mid(Registro, 1, 3) <> "519" Then
                Sql = "Cadena control incorrecta(519)"
            Else
                Sql = ""
            End If
        End If
    End If
    
    If Sql = "" Then
    
        'Segunda LINEA.
        'Line Input #nf, Registro
        ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
        
        Sql = "Linea 2 vacia"
        If Registro <> "" Then
            
            'NIF
            Sql = Mid(Registro, 5, 9)
            
            
            'Tiene valor
            If Len(Registro) <> 162 Then
                Sql = "Longitud linea incorrecta(162)"
            Else
                'En lugar de 5390 comprobamos por 539
                If Mid(Registro, 1, 3) <> "539" Then
                    Sql = "Cadena control incorrecta(539)"
                Else
                    
                    Sql = "Falta linea 569"
                    Remesa = ""
                    Do
                        ProcesoFicheroDevolucion 2, LinDelFichero  'vemos si es ultima linea
                        
                        If Registro <> "" Then
                            Sql = "FIN LINEAS. No se ha encontrado linea: 569"
                            Remesa = "NO"
                        Else
                            'Line Input #nf, Registro
                            ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
                            
                            'BUsco la linea:
                            '5690
                            If Registro <> "" Then
                                'Nov 2012   En lugar de 5690 comprobamos 569
                                If Mid(Registro, 1, 3) = "569" Then
                                    Sql = ""
                                    Remesa = "NO"
                                End If
                            End If
                        End If
                        
                    Loop Until Remesa <> ""
                    Remesa = ""
                    
                    If Sql = "" Then
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
                                Sql = "Select codrem,anyorem,siturem from cobros where fecfactu='20" & Mid(Registro, 5, 2) & "-" & Mid(Registro, 3, 2) & "-" & Mid(Registro, 1, 2)
                                aux2 = Sql
                                
                                'Problemas en alzira
                                'If Not IsNumeric(Mid(Registro, 17, 1)) Then
                                'Sept 2013
                                If Not EsFormatoAntiguoDevolucion Then
                                    Sql = Sql & "' AND numserie like '" & Trim(Mid(Registro, 7, 1)) & "%' AND numfactu = " & Val(Mid(Registro, 9, 7)) & " AND numorden=" & Mid(Registro, 16, 1)
                                    'Problema en herbelca. El numero de vto NO viene con la factura
                                    aux2 = aux2 & "' AND numserie like '" & Trim(Mid(Registro, 7, 1)) & "%' AND numfactu = " & Val(Mid(Registro, 9, 8))
                                    
                                Else
                                    'El vencimiento si que es el 17
                                    Sql = Sql & "' AND numserie like '" & Trim(Mid(Registro, 7, 1)) & "%' AND numfactu = " & Val(Mid(Registro, 10, 7)) & " AND numorden=" & Mid(Registro, 17, 1)
                                    aux2 = aux2 & "' AND numserie like '" & Trim(Mid(Registro, 7, 1)) & "%' AND numfactu = " & Val(Mid(Registro, 10, 8))
                                    
                                End If
                                
                                miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                                TodoOk = False
                                Sql = "Vencimiento no encontrado: " & Registro
                                If Not miRsAux.EOF Then
                                    If IsNull(miRsAux!CodRem) Then
                                        Sql = "Vencimiento sin Remesa: " & Registro
                                    Else
                                        Sql = miRsAux!CodRem & "|" & miRsAux!AnyoRem & "|·"
                                        
                                        If InStr(1, Remesa, Sql) = 0 Then Remesa = Remesa & Sql
                                        Sql = ""
                                        TodoOk = True
                                    End If
                                End If
                                miRsAux.Close
                                
                                
                                If Not TodoOk Then
                                    'Los busco sin Numorden
                                    miRsAux.Open aux2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                                    If Not miRsAux.EOF Then
                                        If IsNull(miRsAux!CodRem) Then
                                            Sql = "Vencimiento sin Remesa: " & Registro
                                        Else
                                            Sql = miRsAux!CodRem & "|" & miRsAux!AnyoRem & "|·"
                                            
                                            If InStr(1, Remesa, Sql) = 0 Then Remesa = Remesa & Sql
                                            Sql = ""
                                            TodoOk = True
                                        End If
                                    End If
                                    miRsAux.Close
                                
                                End If
                                
                                
                                
                                If Sql <> "" Then
                                    ErroresVto = ErroresVto & vbCrLf & Sql
                                Else
                                    Bien = Bien + 1
                                End If
                            Else
                                'La linea no empieza por 569
                                'veremos los totales
                                
                                If Mid(Registro, 1, 3) = "599" Then
                                    'TOTAL TOTAL
                                    Sql = Mid(Registro, 105, 10)
                                    If Val(Sql) <> Cuantos Then ErroresVto = "Fichero: " & Sql & "   Leidos" & Cuantos & vbCrLf & ErroresVto & vbCrLf & Sql
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
                        
                        Sql = ErroresVto
                        Set miRsAux = Nothing
                    
                    End If
                End If  'Control SEGUNDA LINEA
        
        
            End If
        End If
    
    End If  'DE SEGUNDA LINEA
    
    ProcesoFicheroDevolucion 3, LinDelFichero
    If Sql <> "" Then
        MsgBox Sql, vbExclamation
    Else
        'Remesa = Mid(Registro, 1, 4) & "|" & Mid(Registro, 5) & "|"
        
        
        'Ahora comprobaremos que para cada remesa  veremos si existe y si la situacion es la contabilizadxa
        Sql = Remesa
        Registro = "" 'Cadena de error de situacion remesas
        Set miRsAux = New ADODB.Recordset
        Do
            Cuantos = InStr(1, Sql, "·")
            If Cuantos = 0 Then
                Sql = ""
            Else
                aux2 = Mid(Sql, 1, Cuantos - 1)
                Sql = Mid(Sql, Cuantos + 1)
                
                
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
        Loop Until Sql = ""
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
Dim Aux As String
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
        Aux = vbCrLf 'separaremos por este
    Else
        B = InStr(1, Registro, vbCr) > 0
        If B Then
            Aux = vbCr
        Else
            B = InStr(1, Registro, vbLf)
            If B Then Aux = vbLf
        End If
    End If
    
    EsSepa = False
    If Mid(Registro, 1, 4) = "2119" Then EsSepa = True
        
    
    
    If B Then
        'TRAE TODO en una unica linea. Separaremos por el vbcr o vbcrlf
        Do
                NumRegElim = InStr(1, Registro, Aux)
                If NumRegElim = 0 Then
                    
                Else

                    Sql = Mid(Registro, 1, NumRegElim - 1)
                    NumRegElim = NumRegElim + Len(Aux)
                    Registro = Mid(Registro, NumRegElim)  'quito el separador
                    
                    
                   
                    
                    
                    If EsSepa Then
                        C2 = Mid(Sql, 1, 2)
                        If C2 = "23" Then
                            impo = Val(Mid(Sql, 89, 11)) / 100
                            SumaComprobacion = SumaComprobacion + impo
                            
                            'Cuestion 2
                            'Datos identifictivos del vencimiento
                            Sql = Mid(Sql, 21, 35)
                            Listado.Add Sql
                            Sql = ""
                        Else
                            If C2 = "99" Then 'antes 5990
                                Fin = True
                                impo = Val(Mid(Sql, 3, 17)) / 100
                            Else
                                Sql = ""
                            End If
                        End If
                    Else
                        C2 = Mid(Sql, 1, 3)
                        If C2 = "569" Then
                            impo = Val(Mid(Sql, 89, 10)) / 100
                            SumaComprobacion = SumaComprobacion + impo
                            
                            'Cuestion 2
                            'Datos identifictivos del vencimiento
                            Sql = Mid(Sql, 89, 27)
                            Listado.Add Sql
                            Sql = ""
                        Else
                            If C2 = "599" Then 'antes 5990
                                Fin = True
                                impo = Val(Mid(Sql, 89, 10)) / 100
                            Else
                                Sql = ""
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
    Sql = ""
    Do
        Line Input #NF, Registro
        If Registro <> "" Then
         
            Sql = Mid(Registro, 1, 3)
            
            If EsSepa Then
                bol = Mid(Registro, 1, 4) = "2319"
            Else
                bol = Sql = "569"
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
                    Sql = Mid(Registro, 21, 35)
                Else
                    Sql = Mid(Registro, 89, 27)
                End If
                Listado.Add Sql
                Sql = ""
            Else
                
                If EsSepa Then
                    bol = Mid(Registro, 1, 2) = "99"
                Else
                    bol = Sql = "599"
                End If
                    
                If bol Then
                    Fin = True
                    If EsSepa Then
                        impo = Val(Mid(Registro, 3, 17)) / 100
                    Else
                        impo = Val(Mid(Registro, 89, 10)) / 100
                    End If
                Else
                    Sql = ""
                End If
            End If
        End If
        If EOF(NF) Then Fin = True
    Loop Until Fin
    Close #NF
    
    If Sql = "" Then
        MsgBox "No se ha leido la linea final fichero", vbExclamation
        Set Listado = Nothing
    Else
        'OK salimos
        If impo <> SumaComprobacion Then
            Sql = "Error leyendo importes. ¿Desea continuar con los datos obtenidos?"
            If MsgBox(Sql, vbExclamation) = vbNo Then Set Listado = Nothing
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
Dim cad As String

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
            cad = Format(Now, "ddmmyyhhnn")
            cad = App.Path & "\" & cad & ".txt"
        Else
            cad = AuxD
        End If
        'If Es34 Then
        '    FileCopy App.Path & "\norma34.txt", Cad
        'Else
        '    FileCopy App.Path & "\norma68.txt", Cad
        'End If
        Select Case TipoFichero
        Case 0
            FileCopy App.Path & "\norma34.txt", cad
        Case 1
            FileCopy App.Path & "\norma34.txt", cad
        Case 2
            If vParamT.PagosConfirmingCaixa Then
                FileCopy App.Path & "\normaCaixa.txt", cad
            Else
                FileCopy App.Path & "\norma68.txt", cad
            End If
            
        End Select
        If Err.Number <> 0 Then
            MsgBox "Error creando copia fichero. Consulte soporte técnico." & vbCrLf & Err.Description, vbCritical
        Else
            MsgBox "El fichero esta guardado como: " & cad, vbInformation
        End If
            
    'End If
End Function







Private Function RellenaABlancos(CADENA As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim cad As String
    
    cad = Space(Longitud)
    If PorLaDerecha Then
        cad = CADENA & cad
        RellenaABlancos = Left(cad, Longitud)
    Else
        cad = cad & CADENA
        RellenaABlancos = Right(cad, Longitud)
    End If
    
End Function



Private Function RellenaAceros(CADENA As String, PorLaDerecha As Boolean, Longitud As Integer) As String
Dim cad As String
    
    cad = Mid("00000000000000000000", 1, Longitud)
    If PorLaDerecha Then
        cad = CADENA & cad
        RellenaAceros = Left(cad, Longitud)
    Else
        cad = cad & CADENA
        RellenaAceros = Right(cad, Longitud)
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
Dim Aux As String
Dim cad As String


    On Error GoTo EGen
    GeneraFicheroCaixaConfirming = False
    
    NumeroTransferencia = vNumeroTransferencia
    
    
    'Cargamos la cuenta
    cad = "Select * from ctabancaria where codmacta='" & CuentaPropia & "'"
    Set RS = New ADODB.Recordset
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = Right("    " & CIF, 9)
    Aux = Mid(CIF & Space(10), 1, 9)
    If RS.EOF Then
        cad = ""
    Else
        If IsNull(RS!Entidad) Then
            cad = ""
        Else
            
            CodigoOrdenante = Format(RS!Entidad, "0000") & Format(DBLet(RS!Oficina, "N"), "0000") & Format(DBLet(RS!Control, "N"), "00") & Format(DBLet(RS!CtaBanco, "T"), "0000000000")
            
            If Not DevuelveIBAN2("ES", CodigoOrdenante, cad) Then cad = ""
            CuentaPropia = "ES" & cad & CodigoOrdenante
                        
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
    If cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\normaCaixa.txt" For Output As #NFich
    
    
    
    
    
    'Codigo ordenante
    '---------------------------------------------------
    'Si el banco tiene puesto si ID de norma34 entonces
    'la pongo aquin. Lo he cargado antes sobre la variable AUX
    CodigoOrdenante = Left(Aux & "          ", 10)  'CIF EMPRESA
  
    Set RS = New ADODB.Recordset
    
    'CABECERA
    'UNo
    Aux = "0156" & CodigoOrdenante & Space(12) & "001" & Format(Fecha, "ddmmyy") & Space(6)
    Aux = Aux & ConceptoTr_ & "1" & "EUR" & Space(9)   'Ya esta. Ya he utlizado la variable ConceptoTr_. Nada mas
    Print #NFich, Aux
    'Nombre
    Aux = "0156" & CodigoOrdenante & Space(12) & "002" & FrmtStr(vEmpresa.nomempre, 36) & Space(7)
    Print #NFich, Aux
    
    'Registros obligatorios  3 4
    Aux = "Select pobempre, provempre from empresa2"
    RS.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'NO PUEDE SER EOF
    For Regs = 0 To 1
        Aux = "0156" & CodigoOrdenante & Space(12) & Format(Regs + 3, "000") & FrmtStr(DBLet(RS.Fields(Regs), "T"), 36) & Space(7)
        Print #NFich, Aux
    Next
    RS.Close
    
    
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma34
    
    Aux = "Select spagop.*,nommacta,dirdatos,codposta,dirdatos,despobla,nifdatos,razosoci,desprovi,pais from spagop,cuentas"
    Aux = Aux & " where codmacta=ctaprove and transfer =" & NumeroTransferencia
    RS.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
                Aux = RellenaABlancos(RS!nifdatos, True, 12)
                
                    
                'Reg 010
                Aux = "0656" & CodigoOrdenante & Aux & "010"
                Aux = Aux & RellenaAceros(CStr(Im * 100), False, 12)
                Aux = Aux & FrmtStr(DBLet(RS!Entidad, "N"), 4) & FrmtStr(DBLet(RS!Oficina, "N"), 4)
                Aux = Aux & FrmtStr(DBLet(RS!Cuentaba, "N"), 10) & "1" & "9" & "  " & FrmtStr(DBLet(RS!CC, "N"), 2)
                Aux = Aux & "N" & "C" & "EUR  "
                Print #NFich, Aux
                
        
           
           
                
                'OBligaorio 011   Nombre
                Aux = RellenaABlancos(RS!nifdatos, True, 12)
                Aux = "0656" & CodigoOrdenante & Aux & "011"
                Aux = Aux & FrmtStr(DBLet(RS!razosoci, "T"), 36) & Space(7)
                Print #NFich, Aux
           
                'OBligaorio 012   direccion
                Aux = RellenaABlancos(RS!nifdatos, True, 12)
                Aux = "0656" & CodigoOrdenante & Aux & "012"
                Aux = Aux & FrmtStr(DBLet(RS!dirdatos, "T"), 36) & Space(7)
                Print #NFich, Aux
           
                'OBligaorio 014   cpos provi
                Aux = RellenaABlancos(RS!nifdatos, True, 12)
                Aux = "0656" & CodigoOrdenante & Aux & "014"
                Aux = Aux & FrmtStr(DBLet(RS!codposta, "N"), 5) & FrmtStr(DBLet(RS!desPobla, "T"), 31) & Space(7)
                Print #NFich, Aux
                
                'OBligaorio 016   ID factura
                Aux = RellenaABlancos(RS!nifdatos, True, 12)
                Aux = "0656" & CodigoOrdenante & Aux & "016"
                Aux = Aux & "T" & Format(RS!FecFactu, "ddmmyy") & FrmtStr(RS!NumFactu, 15) & Format(RS!Fecefect, "ddmmyy") & Space(15)
                Print #NFich, Aux
           
                 
        
               'Totales
               Importe = Importe + Im
               Regs = Regs + 1
               RS.MoveNext
        Wend
        'Imprimimos totales
        Aux = "08" & "56"
        Aux = Aux & CodigoOrdenante    'llevara tb la ID del socio
        Aux = Aux & Space(15)
        Aux = Aux & RellenaAceros(CStr(Int(Round(Importe * 100, 2))), False, 12)
        Aux = Aux & RellenaAceros(CStr((Regs)), False, 8)
        Aux = Aux & RellenaAceros(CStr((Regs * 5) + 4 + 1), False, 10)    '4 de cabecera + uno de totales
        Aux = RellenaABlancos(Aux, True, 72)
        Print #NFich, Aux
        
        
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

Private Function GrabarFicheroNorma19SEPA(NomFichero As String, Remesa As String, FecPre As String, TipoReferenciaCliente As Byte, Sufijo As String, FechaCobro As String, SEPA_EmpresasGraboNIF As Boolean, Norma19_15 As Boolean, FormatoXML As Boolean) As Boolean
Dim B As Boolean
    '-- Genera_Remesa: Esta función genera la remesa indicada, en el fichero correspondiente

    Dim DatosBanco As String  'oficina,sucursla,cta, sufijo
    Dim NifEmpresa_ As String
    
    '-- Primero comprobamos que la remesa no haya sido enviada ya
    Sql = "SELECT * FROM remesas,bancos WHERE codigo = " & RecuperaValor(Remesa, 1)
    Sql = Sql & " AND anyo = " & RecuperaValor(Remesa, 2) & " AND remesas.codmacta = bancos.codmacta "
    
    Set miRsAux = New ADODB.Recordset
    DatosBanco = ""
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
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
    Sql = "Select * from empresa2"
    miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        B = GrabarDisketteNorma19_SEPA_XML(NomFichero, Remesa, FecPre, TipoReferenciaCliente, Sufijo, FechaCobro, SEPA_EmpresasGraboNIF, Norma19_15, DatosBanco, NifEmpresa_)
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
Dim cad As String

    Select Case TipoRegistro
    Case 0
        'Cabecera de ACREEDOR-FECHA
        cad = "02" & IdNorma & "002"   '19143-> Podria ser 19154 ver pdf
        cad = cad & IdDeudorAcreedor
        
        'Fecha cobro
        cad = cad & Format(miRsAux!FecVenci, "yyyymmdd")
        
        'Nomprove
        cad = cad & DatosBasicosDelAcreedor
        'EN SQL llevamos el IBAN completo del acredor, es decir, de la empresa presentardora que le deben los deudores
        cad = cad & Sql & Space(10)  'El iban son 24 y dejan hasta 34 psociones
        '
        cad = cad & Space(301)
        
    Case 1
        'total x fecha -deudor
        cad = "04"
        cad = cad & IdDeudorAcreedor

        'Fecha cobro
        cad = cad & Format(Fecha, "yyyymmdd")

        cad = cad & Right(String(17, "0") & (Suma * 100), 17) ' Suma total de registros
        cad = cad & Format(Registros003, "00000000")
        cad = cad & Format(NumeroLineasTotalesSinCabceraPresentador + 2, "0000000000") ' +cabecera y pie
        cad = cad & FrmtStr(" ", 520) ' LIBRE

        
        
    Case 2
        'total deudor
        cad = "05"
        cad = cad & IdDeudorAcreedor

        cad = cad & Right(String(17, "0") & (Suma * 100), 17) ' Suma total de registros
        cad = cad & Format(Registros003, "00000000")
        cad = cad & Format(NumeroLineasTotalesSinCabceraPresentador + 2, "0000000000") '
        cad = cad & FrmtStr(" ", 528) ' LIBRE
      
    Case 3
        'total general
        cad = "99"
        cad = cad & Right(String(17, "0") & (Suma * 100), 17) ' Suma total de registros
        cad = cad & Format(Registros003, "00000000")
        cad = cad & Format(NumeroLineasTotalesSinCabceraPresentador + 2, "0000000000") ' +cabecera y pie
        cad = cad & FrmtStr(" ", 563) ' LIBRE
      
    End Select
        
    Print #NF, cad
        
        
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


Private Function IBAN_Destino(Cobros As Boolean) As String
    If Cobros Then
        IBAN_Destino = FrmtStr(DBLet(miRsAux!IBAN, "T"), 4) ' ES00
        IBAN_Destino = IBAN_Destino & Format(miRsAux!codbanco, "0000") ' Código de entidad receptora
        IBAN_Destino = IBAN_Destino & Format(miRsAux!codsucur, "0000") ' Código de oficina receptora
        IBAN_Destino = IBAN_Destino & Format(miRsAux!digcontr, "00") ' Dígitos de control
        IBAN_Destino = IBAN_Destino & Format(miRsAux!Cuentaba, "0000000000") ' Código de cuenta
    Else
        
        'entidad oficina CC cuentaba
        IBAN_Destino = FrmtStr(DBLet(miRsAux!IBAN, "T"), 4) ' ES00
        IBAN_Destino = IBAN_Destino & Format(miRsAux!Entidad, "0000") ' Código de entidad receptora
        IBAN_Destino = IBAN_Destino & Format(miRsAux!Oficina, "0000") ' Código de oficina receptora
        IBAN_Destino = IBAN_Destino & Format(miRsAux!CC, "00") ' Dígitos de control
        IBAN_Destino = IBAN_Destino & Format(miRsAux!Cuentaba, "0000000000") ' Código de cuenta
    End If
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
    Sql = "Linea 1 vacia"
    If Registro <> "" Then
        
        
        
        'Tiene valor
        If Len(Registro) <> 600 Then
            Sql = "Longitud linea incorrecta(600)"
        Else
            'Febrero 2014
            'Devolucion:2119
            'Rechazo:   1119
            'Antes: Mid(Registro, 1, 4) <> "2119"
            
            If Mid(Registro, 2, 3) <> "119" Then
                Sql = "Cadena control incorrecta(?119)"
            Else
                Sql = ""
            End If
        End If
    End If
    
    If Sql = "" Then
    
        'Segunda LINEA.
        'Line Input #nf, Registro
        ProcesoFicheroDevolucion 1, LinDelFichero  'leo la linea y apunto a la siguiente
        
        Sql = "Linea 2 vacia"
        If Registro <> "" Then
            
           
            
            
            'Tiene valor
            If Len(Registro) <> 600 Then
                Sql = "Longitud linea incorrecta(600)"
            Else
                'Devolucion:2219
                'Rechazo:   1119
                'Antes: Mid(Registro, 1, 4) <> "2119"
                
                If Mid(Registro, 2, 3) <> "219" Then
                    Sql = "Cadena control incorrecta(?219)"
                Else
                    
                    Sql = "Falta linea 2319"  'la que lleva los vtos
                    Remesa = ""
                    Do
                        ProcesoFicheroDevolucion 2, LinDelFichero  'vemos si es ultima linea
                        
                        If Registro <> "" Then
                            Sql = "FIN LINEAS. No se ha encontrado linea: 2319"
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
                                    Sql = ""
                                    Remesa = "NO"
                                End If
                            End If
                        End If
                        
                    Loop Until Remesa <> ""
                    Remesa = ""
                    
                    If Sql = "" Then
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
                                Sql = "Select codrem,anyorem,siturem from cobros where fecfactu='" & Mid(Registro, 12, 4) & "-" & Mid(Registro, 16, 2) & "-" & Mid(Registro, 18, 2)
                                
                                Sql = Sql & "' AND numserie = '" & Trim(Mid(Registro, 1, 3)) & "' AND numfactu = " & Val(Mid(Registro, 4, 8)) & " AND numorden=" & Mid(Registro, 20, 3)
                                
                                
                                miRsAux.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                                TodoOk = False
                                Sql = "Vencimiento no encontrado: " & Registro
                                If Not miRsAux.EOF Then
                                    If IsNull(miRsAux!CodRem) Then
                                        Sql = "Vencimiento sin Remesa: " & Registro
                                    Else
                                        Sql = miRsAux!CodRem & "|" & miRsAux!AnyoRem & "|·"
                                        
                                        If InStr(1, Remesa, Sql) = 0 Then Remesa = Remesa & Sql
                                        Sql = ""
                                        TodoOk = True
                                    End If
                                End If
                                miRsAux.Close
                                
                               
                                
                                
                                
                                If Sql <> "" Then
                                    ErroresVto = ErroresVto & vbCrLf & Sql
                                Else
                                    Bien = Bien + 1
                                End If
                            Else
                                'La linea no empieza por 569
                                'veremos los totales
                                
                                If Mid(Registro, 1, 2) = "99" Then
                                    'TOTAL TOTAL
                                    Sql = Mid(Registro, 20, 8)
                                    If Val(Sql) <> Cuantos Then ErroresVto = "Fichero: " & Sql & "   Leidos" & Cuantos & vbCrLf & ErroresVto & vbCrLf & Sql
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
                        
                        Sql = ErroresVto
                        Set miRsAux = Nothing
                    
                    End If
                End If  'Control SEGUNDA LINEA
        
        
            End If
        End If
    
    End If  'DE SEGUNDA LINEA
    
    ProcesoFicheroDevolucion 4, LinDelFichero
    If Sql <> "" Then
        MsgBox Sql, vbExclamation
    Else
        'Remesa = Mid(Registro, 1, 4) & "|" & Mid(Registro, 5) & "|"
        
        
        'Ahora comprobaremos que para cada remesa  veremos si existe y si la situacion es la contabilizadxa
        Sql = Remesa
        Registro = "" 'Cadena de error de situacion remesas
        Set miRsAux = New ADODB.Recordset
        Do
            Cuantos = InStr(1, Sql, "·")
            If Cuantos = 0 Then
                Sql = ""
            Else
                aux2 = Mid(Sql, 1, Cuantos - 1)
                Sql = Mid(Sql, Cuantos + 1)
                
                
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
        Loop Until Sql = ""
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
        Line Input #NF, Sql
        If Sql <> "" Then
            '                 DEVOLUCION                RECHAZO
            If LCase(Mid(Sql, 1, 5)) = "<?xml" Then
                EsFicheroDevolucionSEPA2 = 2
            Else
                If Mid(Sql, 1, 2) = "21" Or Mid(Sql, 1, 2) = "11" Then
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
