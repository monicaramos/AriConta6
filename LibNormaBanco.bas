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
    SQL = "SELECT * FROM remesas,ctabancaria WHERE codigo = " & RecuperaValor(Remesa, 1)
    SQL = SQL & " AND anyo = " & RecuperaValor(Remesa, 2) & " AND remesas.codmacta = ctabancaria.codmacta "
    
    Set miRsAux = New ADODB.Recordset
    DatosBanco = ""
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
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
    SQL = "Select * from empresa2"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
    
    SQL = "select  scobro.*,nommacta,nifdatos from scobro,cuentas where "
    SQL = SQL & " scobro.codmacta = cuentas.codmacta "
    SQL = SQL & " AND codrem = " & RecuperaValor(Remesa, 1)
    SQL = SQL & " AND anyorem=" & RecuperaValor(Remesa, 2)
    
    
    'EL ORDEN QUE QUERAMOS
    msgSerie = ""
    Remesa = RecuperaValor(Remesa, 1)
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
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
    SQL = "SELECT * FROM remesas,ctabancaria WHERE codigo = " & RecuperaValor(Remesa2, 1)
    SQL = SQL & " AND anyo = " & RecuperaValor(Remesa2, 2) & " AND remesas.codmacta = ctabancaria.codmacta "
    
    Set miRsAux = New ADODB.Recordset
    DatosBanco = ""
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
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
    SQL = "Select * from empresa2"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
    SQL = "SELECT fecvenci FROM scobro WHERE codrem = " & RecuperaValor(Remesa2, 1)
    SQL = SQL & " AND anyorem = " & RecuperaValor(Remesa2, 2) & " GROUP BY 1"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
        
    
        SQL = "select  scobro.*,nommacta,nifdatos from scobro,cuentas where "
        SQL = SQL & " scobro.codmacta = cuentas.codmacta "
        SQL = SQL & " AND codrem = " & RecuperaValor(Remesa2, 1)
        SQL = SQL & " AND anyorem=" & RecuperaValor(Remesa2, 2)
        SQL = SQL & " AND fecvenci='" & Format(FCargo, FormatoFecha) & "'"
    
    
    
        'EL ORDEN QUE QUERAMOS
        msgSerie = ""
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
     
   
        
        
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




'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
'               32 , 32 , 32 , 32 , 32
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------

Public Function GrabarDisketteNorma32(NomFichero As String, Remesa As String, FecPre As String) As Boolean
    '-- Genera_Remesa: Esta función genera la remesa indicada, en el fichero correspondiente
    Dim mAux As String
    Dim SumaImportes As Currency
    Dim SumTotal As Integer
    Dim DatosBanco As String  'oficina,sucursla,cta, sufijo
    Dim vSufijo As String
    Dim NifEmpresa As String
    Dim DatosEmpresa As String
    Dim ImpEfect As Currency
    On Error GoTo Err_Remesa
    
    
    '-- Primero comprobamos que la remesa no haya sido enviada ya
    SQL = "SELECT * FROM remesas,ctabancaria WHERE codigo = " & RecuperaValor(Remesa, 1)
    SQL = SQL & " AND anyo = " & RecuperaValor(Remesa, 2) & " AND remesas.codmacta = ctabancaria.codmacta "
    
    Set miRsAux = New ADODB.Recordset
    DatosBanco = ""
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        If miRsAux!Situacion >= "C" Then
            MsgBox "La remesa ya figura como enviada", vbCritical
        Else
            'Cargo algunos de los datos de la remesa
            DatosBanco = Format(miRsAux!Entidad, "0000") & "|" & Format(miRsAux!Oficina, "0000") & "|" & Format(miRsAux!Control, "00") & "|" & Format(miRsAux!CtaBanco, "0000000000") & "|"
            DatosBanco = DatosBanco & DBLet(miRsAux!idcedente, "T") & "|"
            vSufijo = Right("000" & RecuperaValor(Remesa, 3), 3)
        End If
    Else
        MsgBox "La remesa solicitada no existe", vbCritical
    End If
    miRsAux.Close
    
    If DatosBanco = "" Then Exit Function
    
    If Not comprobarCuentasBancariasRecibos(Remesa) Then Exit Function
    
'Ahora cargare el NIF y la empresa
    SQL = "Select * from empresa2"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NifEmpresa = ""
    DatosEmpresa = ""
    If Not miRsAux.EOF Then
        NifEmpresa = DBLet(miRsAux!nifempre, "T")
        If Not IsNull(miRsAux!Poblacion) And Not IsNull(miRsAux!codpos) Then DatosEmpresa = miRsAux!Poblacion & "|" & miRsAux!codpos & "|"
    End If
    miRsAux.Close
    If NifEmpresa = "" Or DatosEmpresa = "" Then
        MsgBox "Datos empresa MAL configurados:" & NifEmpresa & " - " & DatosEmpresa, vbExclamation
        Exit Function
    End If
    
        
    '-- Abrir el fichero a enviar
    NF = FreeFile()
    Open NomFichero For Output As #NF
    
    SQL = "select  scobro.*,cuentas.* from scobro,cuentas where "
    SQL = SQL & " scobro.codmacta=cuentas.codmacta AND "
    SQL = SQL & " codrem=" & RecuperaValor(Remesa, 1) & " AND anyorem=" & RecuperaValor(Remesa, 2)
    
    
    'EL ORDEN QUE QUERAMOS
    Remesa = RecuperaValor(Remesa, 1)
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        'Rs.MoveFirst
     
    
'        '-- Registro CABCERA FICHERO
        Registro = "0265  "
        Registro = Registro & Format(FecPre, "ddmmyy") & "0001"
        Registro = Registro & FrmtStr(" ", 12)
        Registro = Registro & FrmtStr(" ", 23)
        Registro = Registro & RecuperaValor(DatosBanco, 1) '
        Registro = Registro & RecuperaValor(DatosBanco, 2) '
        Registro = Registro & FrmtStr(vEmpresa.nomempre & " " & NifEmpresa, 61)    '-- Alinea NIF
        Registro = Registro & FrmtStr(" ", 45) ' LIBRE
        Print #NF, Mid(Registro, 1, 162)
       

        
        
        '-- Registro Cabecera REMESA
        Registro = "1165  "
        Registro = Registro & Format(FecPre, "ddmmyy") ' Fecha de confección del soporte
        Registro = Registro & "0001"
        Registro = Registro & FrmtStr(" ", 12) ' LIBRE
        
        Registro = Registro & FrmtStr(RecuperaValor(DatosBanco, 5), 15)   ' LIBRE
        Registro = Registro & "1"
        Registro = Registro & FrmtStr(" ", 21) ' LIBRE
        
                
        SQL = RecuperaValor(DatosBanco, 1)  'Dígitos de control
        SQL = SQL & RecuperaValor(DatosBanco, 2)  ' Código de cuenta
        SQL = SQL & RecuperaValor(DatosBanco, 3)  ' Código de cuenta
        SQL = SQL & RecuperaValor(DatosBanco, 4)  ' Código de cuenta
        Registro = Registro & SQL & SQL & SQL & " " & FrmtStr(" ", 24)
        Registro = Registro & FrmtStr(" ", 20) ' LIBRE
        
        Print #NF, Mid(Registro, 1, 162)
        SumTotal = 0
        '-- Leemos secuencialmente las líneas de remesa
        While Not miRsAux.EOF
            '-- PRIMER REGISTRO
            Registro = "2565  "
            
            SQL = miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000000") & "/" & miRsAux!numorden
            Registro = Registro & FrmtStr(SQL, 15)
            Registro = Registro & "0001"
            Registro = Registro & FrmtStr(RecuperaValor(DatosEmpresa, 2), 2)
            Registro = Registro & FrmtStr("", 9)
            Registro = Registro & FrmtStr(RecuperaValor(DatosEmpresa, 1), 20)
            
            ImpEfect = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N")
            
            Registro = Registro & FrmtStr(Format(ImpEfect * 100, String(9, "0")), 9) ' Importe
            Registro = Registro & FrmtStr("", 15)
            Registro = Registro & FrmtStr(Format(miRsAux!FecVenci, "ddmmyy"), 6) ' fecha
            Registro = Registro & FrmtStr("", 55)
            Registro = Mid(Registro, 1, 162)
            Print #NF, Registro
            
            
            
            '-------- SEGUNDO
            Registro = "2665  "
            Registro = Registro & FrmtStr(SQL, 15)
            Registro = Registro & "  2"
            Registro = Registro & FrmtStr(Format(FecPre, "ddmmyy"), 6) ' fecha
            Registro = Registro & "20"
            
            Registro = Registro & Format(miRsAux!codbanco, "0000") ' Código de entidad receptora
            Registro = Registro & Format(miRsAux!codsucur, "0000") ' Código de oficina receptora
            Registro = Registro & FrmtStr(miRsAux!digcontr, 2) ' Dígitos de control
            Registro = Registro & Format(miRsAux!Cuentaba, String(10, "0")) ' Código de cuenta
            Registro = Registro & FrmtStr(vEmpresa.nomempre, 24)
            Registro = Registro & FrmtStr(miRsAux!Nommacta, 34)
            
            mAux = "FACTURA : " & miRsAux!NUmSerie & "-" & Format(miRsAux!codfaccl, "00000000") & " VTO : " & miRsAux!fecfaccl
            Registro = Registro & FrmtStr(mAux, 40) ' Primer Concepto
            Registro = Registro & FrmtStr(" ", 20) ' LIBRE
            Registro = Mid(Registro, 1, 162)
            Print #NF, Registro
            
            
            
            'Registro 3
            Registro = "2765  "
            Registro = Registro & FrmtStr(SQL, 15)
            Registro = Registro & "  "
            Registro = Registro & FrmtStr(DBLet(miRsAux!desPobla, "T"), 34)
            Registro = Registro & FrmtStr(DBLet(miRsAux!codposta, "T"), 5)
            Registro = Registro & FrmtStr(DBLet(miRsAux!desPobla, "T"), 20)
            Registro = Registro & FrmtStr(DBLet(miRsAux!codposta, "T"), 2)
            Registro = Registro & FrmtStr(" ", 120) ' LIBRE
            Registro = Mid(Registro, 1, 162)
            Print #NF, Registro
            
            
            
            
            SumTotal = SumTotal + 1
            SumaImportes = SumaImportes + ImpEfect
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        'Fin remesa
        Registro = "7165  "
        Registro = Registro & FrmtStr(Format(FecPre, "ddmmyy"), 6) ' fecha
        Registro = Registro & "0001"
        Registro = Registro & String(70, " ")
        Registro = Mid(Registro, 1, 75)
        Registro = Registro & Format(SumaImportes * 100, String(10, "0")) ' Suma de importes
        Registro = Registro & Format((SumTotal * 3) + 2, String(7, "0")) ' Suma de registros
        Registro = Registro & Format(SumTotal, String(5, "0")) ' Suma total de registros
        Registro = Registro & FrmtStr(" ", 65) ' LIBRE
        Registro = Mid(Registro, 1, 162)
        Print #NF, Registro
        
        
        '-- FIN FICHER0
        Registro = "9865  "
        Registro = Registro & String(70, " ")
        Registro = Mid(Registro, 1, 75)
        Registro = Registro & Format(SumaImportes * 100, String(10, "0")) ' Suma de importes
        Registro = Registro & String(50, " ")
        Registro = Mid(Registro, 1, 126)
        Registro = Registro & "00001"
        Registro = Registro & Format((SumTotal * 3) + 4, String(7, "0")) ' Suma de registros
        Registro = Registro & Format(SumTotal, String(5, "0")) ' Suma total de registros
        Registro = Registro & FrmtStr(" ", 58) ' LIBRE
        Registro = Mid(Registro, 1, 162)
        Print #NF, Registro
        
        
    End If
    Close #NF
    GrabarDisketteNorma32 = True
    Exit Function
Err_Remesa:
    MsgBox "Err: " & Err.Number & vbCrLf & _
        Err.Description, vbCritical, "Grabación del diskette de Remesa"
        
End Function





'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
'               58 58 58 58 58 58 58 58 58
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------

Public Function GrabarDisketteNorma58(NomFichero As String, Remesa As String, FecPre As String, DatosExtra As String, TipoReferenciaCliente As Byte, FecCobro As Date) As Boolean
    '-- Genera_Remesa: Esta función genera la remesa indicada, en el fichero correspondiente
    Dim mAux As String
    Dim SumaImportes As Currency
    Dim SumTotal As Integer
    Dim DatosBanco As String  'oficina,sucursla,cta, sufijo
    Dim vSufijo As String
    Dim NifEmpresa As String
    Dim DatosEmpresa As String
    Dim ImpEfect As Currency
    
    Dim IdenUnico As String
    
    On Error GoTo Err_Remesa
    
    GrabarDisketteNorma58 = False
    
    '-- Primero comprobamos que la remesa no haya sido enviada ya
    SQL = "SELECT * FROM remesas,ctabancaria WHERE codigo = " & RecuperaValor(Remesa, 1)
    SQL = SQL & " AND anyo = " & RecuperaValor(Remesa, 2) & " AND remesas.codmacta = ctabancaria.codmacta "
    
    Set miRsAux = New ADODB.Recordset
    DatosBanco = ""
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        If miRsAux!Situacion >= "C" Then
            MsgBox "La remesa ya figura como enviada", vbCritical
        Else
            'Cargo algunos de los datos de la remesa
            DatosBanco = Format(miRsAux!Entidad, "0000") & "|" & Format(miRsAux!Oficina, "0000") & "|" & Format(miRsAux!Control, "00") & "|" & Format(miRsAux!CtaBanco, "0000000000") & "|"
            DatosBanco = DatosBanco & DBLet(miRsAux!idcedente, "T") & "|"
            vSufijo = RecuperaValor(DatosExtra, 1)
            If Trim(vSufijo) = "" Then vSufijo = Mid(miRsAux!sufijoem & "000", 1, 3)
             'En datos extra dejo el CONCEPTO PPAL
             DatosExtra = RecuperaValor(DatosExtra, 2)
            
            
        End If
    Else
        MsgBox "La remesa solicitada no existe", vbCritical
    End If
    miRsAux.Close
    
    If DatosBanco = "" Then Exit Function
    
    If Not comprobarCuentasBancariasRecibos(Remesa) Then Exit Function
    
'Ahora cargare el NIF y la empresa
    SQL = "Select * from empresa2"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NifEmpresa = ""
    DatosEmpresa = ""
    If Not miRsAux.EOF Then
        DatosEmpresa = DBLet(miRsAux!nifempre, "T")
        NifEmpresa = Mid(DatosEmpresa, 1, 1) 'PQ es empresa
        DatosEmpresa = Right("00000000" & Mid(DatosEmpresa, 2), 8)
        NifEmpresa = NifEmpresa & DatosEmpresa
        
        DatosEmpresa = ""
        If Not IsNull(miRsAux!Poblacion) And Not IsNull(miRsAux!codpos) Then DatosEmpresa = miRsAux!Poblacion & "|" & miRsAux!codpos & "|"
    End If
    miRsAux.Close
    If NifEmpresa = "" Or DatosEmpresa = "" Then
        MsgBox "Datos empresa MAL configurados:" & NifEmpresa & " - " & DatosEmpresa, vbExclamation
        Exit Function
    End If
    
         
    '-- Abrir el fichero a enviar
    NF = FreeFile()
    Open NomFichero For Output As #NF
    
    SQL = "select  scobro.*,cuentas.* from scobro,cuentas where "
    SQL = "select  scobro.*,nommacta,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos from scobro,cuentas where  "
    SQL = SQL & " scobro.codmacta=cuentas.codmacta AND "
    SQL = SQL & " codrem=" & RecuperaValor(Remesa, 1) & " AND anyorem=" & RecuperaValor(Remesa, 2)
    
    
    'EL ORDEN QUE QUERAMOS
    Remesa = RecuperaValor(Remesa, 1)
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        'Rs.MoveFirst
     
    
'        '-- Registro CABCERA FICHERO
        Registro = "5170"
        Registro = Registro & NifEmpresa
        Registro = Registro & vSufijo
        Registro = Registro & Format(FecPre, "ddmmyy")
        Registro = Registro & FrmtStr(" ", 6)
        Registro = Registro & FrmtStr(DatosExtra, 40)
        Registro = Registro & FrmtStr(" ", 20)
        Registro = Registro & RecuperaValor(DatosBanco, 1) '
        Registro = Registro & RecuperaValor(DatosBanco, 2) '
        'Registro = Registro & FrmtStr(" ", 12)
        Registro = Registro & "RE" & Format(vEmpresa.codempre, "00") & Format(miRsAux!CodRem, "0000") & Format(miRsAux!AnyoRem, "0000")
        Registro = Registro & FrmtStr(RecuperaValor(DatosBanco, 5), 11)
        Registro = FrmtStr(Registro, 162)
        Print #NF, Registro
        

        
        
        '-- Registro Cabecera REMESA
        Registro = "5370"
        Registro = Registro & NifEmpresa
        Registro = Registro & vSufijo
        Registro = Registro & Format(FecPre, "ddmmyy") ' Fecha de confección del soporte
        Registro = Registro & FrmtStr(" ", 6)
        Registro = Registro & FrmtStr(DatosExtra, 40)
        Registro = Registro & RecuperaValor(DatosBanco, 1) '
        Registro = Registro & RecuperaValor(DatosBanco, 2) '
        Registro = Registro & RecuperaValor(DatosBanco, 3) '
        Registro = Registro & RecuperaValor(DatosBanco, 4) '
        Registro = Registro & FrmtStr(" ", 8)
        Registro = Registro & "06"
        Registro = Registro & FrmtStr(" ", 10)
        Registro = Registro & FrmtStr(RecuperaValor(DatosBanco, 5), 11)
        Registro = Registro & FrmtStr(" ", 31)
        Registro = Registro & "00000000"
        Registro = FrmtStr(Registro, 162)

        
        Print #NF, Registro
        SumTotal = 0
        '-- Leemos secuencialmente las líneas de remesa
        While Not miRsAux.EOF
        
'            Select Case TipoReferenciaCliente
'            Case 1
'                'ALZIRA. La referencia final de 12 es el ctan bancaria del cli + su CC
'                    Registro = Format(miRsAux!digcontr, "00") ' Dígitos de control
'                    Registro = Registro & Format(miRsAux!cuentaba, "0000000000") ' Código de cuenta
'            Case 2
'                'NIF
'                Registro = DBLet(miRsAux!Nifdatos, "T")
'                If Registro = "" Then Registro = miRsAux!codmacta
'                Registro = Mid(Registro & Space(12), 1, 12)
'
'            Case Else
'                Registro = miRsAux!NUmSerie & Format(miRsAux!codfaccl, "0000000000") & Format(miRsAux!numorden, "0")
'            End Select
        
        
        
        
        
        
            '-- PRIMER REGISTRO
            Registro = "5670"
            Registro = Registro & NifEmpresa     'Trozo comun
            Registro = Registro & vSufijo
            
            'ANTES
            'SQL = miRsAux!NUmSerie & Format(miRsAux!codfaccl, "000000000") & Format(miRsAux!numorden, "0")
            'AHora
            'IDENTIFICADOR UNICO.  No puede repetirse
            SQL = Format(miRsAux!codfaccl, "0000000")
            SQL = Right(SQL, 7)
            
            SQL = Mid(miRsAux!NUmSerie & "  ", 1, 2) & SQL & Format(miRsAux!numorden, "0")
            SQL = SQL & Right(CStr(miRsAux!fecfaccl), 2) 'El año
            IdenUnico = SQL
            
            Registro = Registro & FrmtStr(IdenUnico, 12)
            Registro = Registro & FrmtStr(miRsAux!Nommacta, 40)
            
            Registro = Registro & Format(miRsAux!codbanco, "0000") ' Código de entidad receptora
            Registro = Registro & Format(miRsAux!codsucur, "0000") ' Código de oficina receptora
            Registro = Registro & FrmtStr(miRsAux!digcontr, 2) ' Dígitos de control
            Registro = Registro & Format(miRsAux!Cuentaba, String(10, "0"))  ' Código de cuenta
            
            ImpEfect = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N")
            
            Registro = Registro & FrmtStr(Format(ImpEfect * 100, String(10, "0")), 10) ' Importe
            'Codigo cliente, los 6 ultimos
            'IDENTIFICACION
            'Registro = Registro & Right(miRsAux!codmacta, 6)
            'Registro = Registro & FrmtStr(SQL, 10)
            Registro = Registro & Format(miRsAux!fecfaccl, "ddmmyy")
            Registro = Registro & FrmtStr(SQL, 10)
            
            If IsNull(miRsAux!text33csb) Then
                Registro = Registro & FrmtStr("FACT: " & SQL & " -FECHA: " & Format(miRsAux!fecfaccl, "dd/mm/yyyy"), 40)
            Else
                Registro = Registro & FrmtStr(miRsAux!text33csb, 40)
            End If
            Registro = Registro & FrmtStr(Format(miRsAux!FecVenci, "ddmmyy"), 6) ' fecha
            Registro = Registro & "  "
            Registro = Mid(Registro, 1, 162)
            Print #NF, Registro
            
            
            
            
            '-------- SEGUNDO    ->>> NOOOO lo grabo
            'NOOOO lo grabo NOOOO lo grabo  NOOOO lo grabo   NOOOO lo grabo
            'NOOOO lo grabo NOOOO lo grabo  NOOOO lo grabo   NOOOO lo grabo
            Registro = "5671"
            Registro = Registro & NifEmpresa     'Trozo comun
            Registro = Registro & vSufijo
            
            Registro = Registro & FrmtStr(SQL, 15)
            Registro = Registro & "  2"
            Registro = Registro & FrmtStr(Format(FecPre, "ddmmyy"), 6) ' fecha
            Registro = Registro & "20"
            
            Registro = Registro & Format(miRsAux!codbanco, "0000") ' Código de entidad receptora
            Registro = Registro & Format(miRsAux!codsucur, "0000") ' Código de oficina receptora
            Registro = Registro & FrmtStr(miRsAux!digcontr, 2) ' Dígitos de control
            Registro = Registro & Format(miRsAux!Cuentaba, String(10, "0")) ' Código de cuenta
            Registro = Registro & FrmtStr(vEmpresa.nomempre, 24)
            Registro = Registro & FrmtStr(miRsAux!Nommacta, 34)
            
            mAux = "FACTURA : " & miRsAux!NUmSerie & "-" & Format(miRsAux!codfaccl, "00000000") & " VTO : " & miRsAux!fecfaccl
            Registro = Registro & FrmtStr(mAux, 40) ' Primer Concepto
            Registro = Registro & FrmtStr(" ", 20) ' LIBRE
            Registro = Mid(Registro, 1, 162)
           ' Print #NF, Registro
            
            
            
            'Registro ULTIMO , el 0
            Registro = "5676"
            Registro = Registro & NifEmpresa     'Trozo comun
            Registro = Registro & vSufijo
            
            Registro = Registro & FrmtStr(IdenUnico, 12)
            Registro = Registro & FrmtStr(DBLet(miRsAux!dirdatos, "T"), 40)
            Registro = Registro & FrmtStr(DBLet(miRsAux!desPobla, "T"), 35)
            Registro = Registro & FrmtStr(DBLet(miRsAux!codposta, "T"), 5)
            
            'Marzo2015
            Registro = Registro & FrmtStr(DBLet(miRsAux!desProvi, "T"), 38)
            Registro = Registro & FrmtStr(DBLet(miRsAux!codposta, "T"), 2)
            
            'Marzo 2015. Fecha presentacion, NO fecha vencimiento
            'Registro = Registro & FrmtStr(Format(miRsAux!fecvenci, "ddmmyy"), 6) ' fecha
            Registro = Registro & FrmtStr(Format(FecPre, "ddmmyy"), 6) ' fecha
            Registro = Registro & FrmtStr(" ", 8) ' LIBRE
            Registro = Mid(Registro, 1, 162)
            Print #NF, Registro    'Marzo 2015. Herbelca. SI lo grabamos
            
            
            
            
            
            
            
            
            SumTotal = SumTotal + 1   '1   Marzo 2015. Hemos puesto a imprimir una linea pas. Por cada vto son dos lineas
            
            
            'SumaImportes = SumaImportes + miRsAux!impvenci
            SumaImportes = SumaImportes + ImpEfect
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        'Fin remesa
        Registro = "5870"
        Registro = Registro & NifEmpresa     'Trozo comun
        Registro = Registro & vSufijo
        Registro = Registro & FrmtStr(" ", 72)
        Registro = Registro & Format(SumaImportes * 100, String(10, "0")) ' Suma de importes
        Registro = Registro & FrmtStr(" ", 6)
        Registro = Registro & Format(SumTotal, String(10, "0")) ' Suma total de registros
        Registro = Registro & Format((SumTotal * 2) + 2, String(10, "0")) ' Suma de registros
        Registro = Registro & FrmtStr(" ", 100) ' LIBRE
        Registro = Mid(Registro, 1, 162)
        Print #NF, Registro
        
        
        '-- FIN FICHER0
        Registro = "5970"
        Registro = Registro & NifEmpresa     'Trozo comun
        Registro = Registro & vSufijo
        Registro = Registro & FrmtStr(" ", 52)
        Registro = Registro & FrmtStr("0001 ", 20)
        
        Registro = Registro & Format(SumaImportes * 100, String(10, "0")) ' Suma de importes
        Registro = Registro & FrmtStr(" ", 6)
        Registro = Registro & Format(SumTotal, String(10, "0")) ' Suma total de registros
        Registro = Registro & Format((SumTotal * 2) + 4, String(10, "0")) ' Suma de registros
        
        
        Print #NF, Registro

    End If
    Close #NF
    
    GrabarDisketteNorma58 = True
    Exit Function
Err_Remesa:
    MsgBox "Err: " & Err.Number & vbCrLf & _
        Err.Description, vbCritical, "Grabación del diskette de Remesa"
        
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

        If IsNull(miRsAux!Entidad) Or IsNull(miRsAux!Control) Then
            'Ya esta mal
            SQL = ""
        Else
            If IsNull(miRsAux!Cuentaba) Or IsNull(miRsAux!Control) Then
                'mal tb
                SQL = ""
            Else
                'TIENE DATOS
                SQL = "D"
            End If
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
    SQL = "select entidad,oficina,cuentaba,control from cobros where codrem = " & RecuperaValor(Remesa, 1)
    SQL = SQL & " AND anyorem=" & RecuperaValor(Remesa, 2)
    SQL = SQL & " GROUP BY entidad,oficina,cuentaba,control"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    While Not miRsAux.EOF
                SQL = Format(miRsAux!Entidad, "0000")  ' Código de entidad receptora
                SQL = SQL & Format(miRsAux!Oficina, "0000") ' Código de oficina receptora
                
                SQL = SQL & Format(miRsAux!Cuentaba, "0000000000") ' Código de cuenta
                
                CC = Format(miRsAux!Control, "00") ' Dígitos de control
                
                'Este lo mando.
                SQL = CodigoDeControl(SQL)
                If SQL <> CC Then
                    
                    SQL = " - " & Format(miRsAux!Control, "00") & "- " & Format(miRsAux!Cuentaba, "0000000000") & " --> CC. correcto:" & SQL
                    SQL = Format(miRsAux!entridad, "0000") & " - " & Format(miRsAux!Oficina, "0000") & SQL
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
        SQL = "select cobros.entidad codbanco,bics.entidad from cobros left join bics on cobros.entidad=bics.entidad WHERE "
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
        If DBLet(miRsAux!referencia, "T") = "" Then
            Registro = Registro & miRsAux!codmacta & " - " & miRsAux!NUmSerie & "/" & miRsAux!NumFactu & "-" & miRsAux!numorden & vbCrLf
            NF = NF + 1
        Else
            If Len(miRsAux!referencia) > 12 Then SQL = SQL & miRsAux!codmacta & " - " & miRsAux!NUmSerie & "/" & miRsAux!NumFactu & "-" & miRsAux!numorden & "(" & miRsAux!referencia & ")" & vbCrLf
            
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



Public Function comprobarCuentasBancariasPagos(Transferencia As String, Pagos As Boolean) As Boolean
Dim CC As String
Dim IBAN As String
On Error GoTo EcomprobarCuentasBancariasPagos

    comprobarCuentasBancariasPagos = False
    If Pagos Then
        SQL = "select * from spagop where transfer = " & Transferencia
    Else
        'ABONOS
        'numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci,
        'ctabanc1, codbanco, codsucur, digcontr, cuentaba,
        'ctabanc2, fecultco, impcobro, emitdocum, recedocu, contdocu,
        'ultimareclamacion, agente, departamento, codrem, anyorem, siturem, gastos,
        'Devuelto, situacionjuri, noremesar, obs, transfer)
        SQL = "Select numserie, codfaccl, fecfaccl, numorden, codmacta as ctaprove, "
        SQL = SQL & "codbanco as entidad,codsucur as oficina,cuentaba,digcontr as CC"
        SQL = SQL & " FROM scobro where transfer=" & Transferencia
        
    End If
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    NF = 0
    While Not miRsAux.EOF

        If DBLet(IsNull(miRsAux!Entidad), "T") = "" Or DBLet(miRsAux!Oficina, "T") = "" Then
            'Ya esta mal
            SQL = ""
        Else
            If IsNull(miRsAux!Cuentaba) Or IsNull(miRsAux!CC) Then
                'mal tb
                SQL = ""
            Else
                'TIENE DATOS
                SQL = "D"
            End If
        End If
    
        If SQL = "" Then
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
        SQL = "Los siguientes vencimientos no tienen la cuenta bancaria con todos los datos." & vbCrLf & Registro
        MsgBox SQL, vbExclamation
        Exit Function
    End If
    
    
    'Si llega aqui es que todos tienen DATOS
    If Pagos Then
        SQL = "select entidad,oficina,cuentaba,cc,iban from spagop where transfer = " & Transferencia
        SQL = SQL & " GROUP BY entidad,oficina,cuentaba,cc"
    Else
        SQL = "SELECT codbanco as entidad,codsucur as oficina,cuentaba,digcontr as CC,iban"
        SQL = SQL & " FROM scobro where transfer=" & Transferencia
        SQL = SQL & " GROUP BY entidad,oficina,cuentaba,cc"
    End If
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Registro = ""
    While Not miRsAux.EOF
                SQL = Format(miRsAux!Entidad, "0000")  ' Código de entidad receptora
                SQL = SQL & Format(miRsAux!Oficina, "0000") ' Código de oficina receptora
                
                SQL = SQL & Format(miRsAux!Cuentaba, "0000000000") ' Código de cuenta
                
                CC = Format(miRsAux!CC, "00") ' Dígitos de control
                
                'Este lo mando.
                IBAN = Mid(SQL, 1, 8) & CC & Mid(SQL, 9)
                
                SQL = CodigoDeControl(SQL)
                If SQL <> CC Then
                    
                    SQL = " - " & Format(miRsAux!CC, "00") & "- " & Format(miRsAux!Cuentaba, "0000000000") & " --> CC. correcto:" & SQL
                    SQL = Format(miRsAux!Entidad, "0000") & " - " & Format(miRsAux!Oficina, "0000") & SQL
                    Registro = Registro & SQL & vbCrLf
                End If
                
                
                'Noviembre 2013
                'IBAN
                If vParamT.NuevasNormasSEPA Then
                        SQL = "ES"
                        If DBLet(miRsAux!IBAN, "T") <> "" Then SQL = Mid(miRsAux!IBAN, 1, 2)
                    
                
                        If Not DevuelveIBAN2(SQL, IBAN, IBAN) Then
                            
                            SQL = "Error calculo"
                        Else
                            SQL = SQL & IBAN
                            If DBLet(miRsAux!IBAN, "T") <> SQL Then
                                SQL = "Error IBAN. Calculado " & SQL & " / " & DBLet(miRsAux!IBAN, "T")
                            Else
                                'OK
                                SQL = ""
                            End If
                        End If
                        
                        If SQL <> "" Then
                            SQL = SQL & " - " & Format(miRsAux!CC, "00") & "- " & Format(miRsAux!Cuentaba, "0000000000")
                            SQL = Format(miRsAux!Entidad, "0000") & " - " & Format(miRsAux!Oficina, "0000") & SQL
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
Dim AUX2 As String  'Para buscar los vencimientos
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
                                AUX2 = SQL
                                
                                'Problemas en alzira
                                'If Not IsNumeric(Mid(Registro, 17, 1)) Then
                                'Sept 2013
                                If Not EsFormatoAntiguoDevolucion Then
                                    SQL = SQL & "' AND numserie like '" & Trim(Mid(Registro, 7, 1)) & "%' AND numfactu = " & Val(Mid(Registro, 9, 7)) & " AND numorden=" & Mid(Registro, 16, 1)
                                    'Problema en herbelca. El numero de vto NO viene con la factura
                                    AUX2 = AUX2 & "' AND numserie like '" & Trim(Mid(Registro, 7, 1)) & "%' AND numfactu = " & Val(Mid(Registro, 9, 8))
                                    
                                Else
                                    'El vencimiento si que es el 17
                                    SQL = SQL & "' AND numserie like '" & Trim(Mid(Registro, 7, 1)) & "%' AND numfactu = " & Val(Mid(Registro, 10, 7)) & " AND numorden=" & Mid(Registro, 17, 1)
                                    AUX2 = AUX2 & "' AND numserie like '" & Trim(Mid(Registro, 7, 1)) & "%' AND numfactu = " & Val(Mid(Registro, 10, 8))
                                    
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
                                    miRsAux.Open AUX2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
                AUX2 = Mid(SQL, 1, Cuantos - 1)
                SQL = Mid(SQL, Cuantos + 1)
                
                
                'En aux2 tendre codrem|anñorem|
                AUX2 = RecuperaValor(AUX2, 1) & " AND anyo = " & RecuperaValor(AUX2, 2)
                AUX2 = "Select situacion from remesas where codigo = " & AUX2
                miRsAux.Open AUX2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If miRsAux.EOF Then
                    AUX2 = "-No se encuentra remesa"
                Else
                    'Si que esta.
                    'Situacion
                    If CStr(miRsAux!Situacion) <> "Q" Then
                        AUX2 = "- Situacion incorrecta : " & miRsAux!Situacion
                    Else
                        AUX2 = "" 'TODO OK
                    End If
                End If
            
                If AUX2 <> "" Then
                    AUX2 = AUX2 & " ->" & Mid(miRsAux.Source, InStr(1, UCase(miRsAux.Source), " WHERE ") + 7)
                    AUX2 = Replace(AUX2, " AND ", " ")
                    AUX2 = Replace(AUX2, "anyo", "año")
                    Registro = Registro & vbCrLf & AUX2
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

                    SQL = Mid(Registro, 1, NumRegElim - 1)
                    NumRegElim = NumRegElim + Len(Aux)
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



'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'Cuenta propia tendra empipados entidad|sucursal|cc|cuenta|
Private Function GeneraFicheroNorma34_(CIF As String, Fecha As Date, CuentaPropia As String, ConceptoTransferencia As String, vNumeroTransferencia As Integer, ByRef ConceptoTr As String, Pagos As Boolean) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Im As Currency
Dim RS As ADODB.Recordset
Dim Aux As String
Dim cad As String


    On Error GoTo EGen
    GeneraFicheroNorma34_ = False
    
    NumeroTransferencia = vNumeroTransferencia
    
    
    'Cargamos la cuenta
    cad = "Select * from ctabancaria where codmacta='" & CuentaPropia & "'"
    Set RS = New ADODB.Recordset
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Aux = Right("    " & CIF, 10)
    If RS.EOF Then
        cad = ""
    Else
        If IsNull(RS!Entidad) Then
            cad = ""
        Else
            cad = Format(RS!Entidad, "0000") & "|" & Format(DBLet(RS!Oficina, "T"), "0000") & "|" & DBLet(RS!Control, "T") & "|" & Format(DBLet(RS!CtaBanco, "T"), "0000000000") & "|"
            CuentaPropia = cad
        End If
        
        
        'Identificador norma bancaria
        If Not IsNull(RS!idnorma34) Then Aux = RS!idnorma34
    End If
    RS.Close
    Set RS = Nothing
    If cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\norma34.txt" For Output As #NFich
    
    
    
    
    
    'Codigo ordenante
    '---------------------------------------------------
    'Si el banco tiene puesto si ID de norma34 entonces
    'la pongo aquin. Lo he cargado antes sobre la variable AUX
    CodigoOrdenante = Left(Aux & "          ", 10)  'CIF EMPRESA
    
    
    'CABECERA
    Cabecera1 NFich, CodigoOrdenante, Fecha, CuentaPropia, cad
    Cabecera2 NFich, CodigoOrdenante, cad
    Cabecera3 NFich, CodigoOrdenante, cad
    Cabecera4 NFich, CodigoOrdenante, cad
    
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma34
    Set RS = New ADODB.Recordset
    If Pagos Then
        Aux = "Select spagop.*,nommacta,dirdatos,codposta,dirdatos,despobla from spagop,cuentas"
        Aux = Aux & " where codmacta=ctaprove and transfer =" & NumeroTransferencia
    Else
        'ABONOS
         '
        Aux = "Select scobro.codbanco as entidad,scobro.codsucur as oficina,scobro.cuentaba,scobro.digcontr as CC"
        Aux = Aux & ",nommacta,dirdatos,codposta,dirdatos,despobla,impvenci,scobro.codmacta from scobro,cuentas"
        Aux = Aux & " where cuentas.codmacta=scobro.codmacta and transfer =" & NumeroTransferencia
    End If
    RS.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If RS.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not RS.EOF
            If Pagos Then
                Im = DBLet(RS!imppagad, "N")
                Im = RS!ImpEfect - Im
                Aux = RellenaAceros(RS!ctaprove, False, 12)
            
            Else
                Im = Abs(RS!ImpVenci)
                Aux = RellenaAceros(RS!codmacta, False, 12)
            End If
            
            'Cad = "06"
            'Cad = Cad & "56"
            'Cad = Cad & " "
            'Aux = "06" & "56" & " " & CodigoOrdenante & Aux  'Ordenante y socio juntos
        
            Aux = "06" & "56" & CodigoOrdenante & Aux   'Ordenante y socio juntos
        
            Linea1 NFich, Aux, RS, Im, cad, ConceptoTransferencia
            Linea2 NFich, Aux, RS, cad
            Linea3 NFich, Aux, RS, cad
            Linea4 NFich, Aux, RS, cad
            Linea5 NFich, Aux, RS, cad
            Linea6 NFich, Aux, RS, cad, ConceptoTr, Pagos
            If Pagos Then Linea7 NFich, Aux, RS, cad
        
        
        
        
            Importe = Importe + Im
            Regs = Regs + 1
            RS.MoveNext
        Wend
        'Imprimimos totales
        Totales NFich, CodigoOrdenante, Importe, Regs, cad, Pagos
    End If
    RS.Close
    Set RS = Nothing
    Close (NFich)
    If Regs > 0 Then GeneraFicheroNorma34_ = True
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description

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



'Private Sub Cabecera1(NF As Integer,ByRef CodOrde As String)
'Dim Cad As String
'
'End Sub

Private Sub Cabecera1(NF As Integer, ByRef CodOrde As String, Fecha As Date, Cta As String, ByRef cad As String)

    cad = "03"
    cad = cad & "56"
    'cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "001"
    cad = cad & Format(Now, "ddmmyy")
    cad = cad & Format(Fecha, "ddmmyy")
    'Cuenta bancaria
    cad = cad & RecuperaValor(Cta, 1)
    cad = cad & RecuperaValor(Cta, 2)
    cad = cad & RecuperaValor(Cta, 4)
    cad = cad & "0"  'Sin relacion
    cad = cad & "   " & RecuperaValor(Cta, 3)  'Digito de control bancario
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub



Private Sub Cabecera2(NF As Integer, ByRef CodOrde As String, ByRef cad As String)
    cad = "03"
    cad = cad & "56"
    'cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "002"
    
    cad = cad & RellenaABlancos(vEmpresa.nomempre, True, 30)   'Nombre empresa
  
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Cabecera3(NF As Integer, ByRef CodOrde As String, ByRef cad As String)
    cad = "03"
    cad = cad & "56"
    'cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "003"
    
    
    AuxD = DevuelveDesdeBD("direccion", "empresa2", "codigo", 1, "N")
    cad = cad & RellenaABlancos(AuxD, True, 30)   'Nombre empresa
    cad = cad & RellenaABlancos("", True, 30)   'Nombre empresa
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub



Private Sub Cabecera4(NF As Integer, ByRef CodOrde As String, ByRef cad As String)

    cad = "03"
    cad = cad & "56"
    'cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "004"
    
    AuxD = DevuelveDesdeBD("codpos", "empresa2", "codigo", 1, "N")
    cad = cad & RellenaABlancos(AuxD, False, 5)
    cad = cad & " "
    AuxD = DevuelveDesdeBD("provincia", "empresa2", "codigo", 1, "N")
    cad = cad & RellenaABlancos(AuxD, True, 30)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub



'ConceptoTransferencia
'1.- Abono nomina
'9.- Transferencia ordinaria
Private Sub Linea1(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Importe1 As Currency, ByRef cad As String, vConceptoTransferencia As String)


   
    '
    cad = CodOrde   'llevara tb la ID del socio
    cad = cad & "010"
    cad = cad & RellenaAceros(CStr(Round(Importe1, 2) * 100), False, 12)
    
    cad = cad & RellenaAceros(CStr(RS1!Entidad), False, 4)     'Entidad
    cad = cad & RellenaAceros(CStr(RS1!Oficina), False, 4)   'Sucur
    cad = cad & RellenaAceros(CStr(RS1!Cuentaba), False, 10)  'Cta
    cad = cad & "1" & vConceptoTransferencia
    cad = cad & "  "
    cad = cad & RellenaAceros(CStr(RS1!CC), False, 2)  'CC
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea2(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "011"
    cad = cad & RellenaABlancos(RS1!Nommacta, False, 36)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea3(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "012"
    cad = cad & RellenaABlancos(DBLet(RS1!dirdatos, "T"), False, 36)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea4(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "013"
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea5(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "014"
    cad = cad & RellenaABlancos(DBLet(RS1!codposta, "T"), False, 5) & " "
    cad = cad & RellenaABlancos(DBLet(RS1!desPobla, "T"), False, 30)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea6(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String, ByRef ConceptoT As String, Pagos As Boolean)
Dim Aux As String

    Aux = ConceptoT
    If Pagos Then
        'Tiene dos campos para las descripcion. Si no tiene nada pondre la descripcion de la transferencia
        Aux = Trim(DBLet(RS1!text1csb, "T"))
        If Aux = "" Then Aux = ConceptoT
    End If

    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "016"
    cad = cad & RellenaABlancos(Aux, False, 35)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub


Private Sub Linea7(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)


    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "017"
    cad = cad & RellenaABlancos(DBLet(RS1!text2csb, "T"), False, 35)
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub




Private Sub Totales(NF As Integer, ByRef CodOrde As String, Total As Currency, Registros As Integer, ByRef cad As String, Pagos As Boolean)
    cad = "08" & "56"
    cad = cad & CodOrde    'llevara tb la ID del socio
    cad = cad & Space(15)
    cad = cad & RellenaAceros(CStr(Int(Round(Total * 100, 2))), False, 12)
    cad = cad & RellenaAceros(CStr(Registros), False, 8)
    If Pagos Then
        cad = cad & RellenaAceros(CStr((Registros * 7) + 4 + 1), False, 10)
    Else
        cad = cad & RellenaAceros(CStr((Registros * 6) + 4 + 1), False, 10)
    End If
    cad = RellenaABlancos(cad, True, 72)
    Print #NF, cad
End Sub











'******************************************************************************************************************
'******************************************************************************************************************
'
'       Genera fichero NORMA 68
'
'Cuenta propia tendra empipados entidad|sucursal|cc|cuenta|
Public Function GeneraFicheroNorma68(CIF As String, Fecha As Date, CuentaPropia As String, vNumeroTransferencia As Integer, ByRef ConceptoTr As String) As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim Im As Currency
Dim RS As ADODB.Recordset
Dim Aux As String
Dim cad As String
Dim PagosJuntos As Boolean

    On Error GoTo EGen
    GeneraFicheroNorma68 = False
    
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
            
        End If
        
        
    End If
    RS.Close
    Set RS = Nothing
    If cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia, vbExclamation
        Exit Function
    End If
    
    NFich = FreeFile
    Open App.Path & "\norma68.txt" For Output As #NFich
    
    
    
    
    
    'Codigo ordenante
    '---------------------------------------------------
    'Si el banco tiene puesto si ID de norma34 entonces
    'la pongo aquin. Lo he cargado antes sobre la variable AUX
    CodigoOrdenante = Left(Aux & "          ", 9)  'CIF EMPRESA
    CodigoOrdenante = CodigoOrdenante & "000" 'el sufijo
    
    'CABECERA
    Cabecera1_68 NFich, CodigoOrdenante, Fecha, CuentaPropia, cad
   
    Aux = DevuelveDesdeBD("conceptotrans", "stransfer", "codigo", CStr(vNumeroTransferencia))
    PagosJuntos = Aux = "1"
    
    
    'Imprimimos las lineas
    'Para ello abrimos la tabla tmpNorma34
    Set RS = New ADODB.Recordset
    Aux = "Select spagop.*,nommacta,dirdatos,codposta,dirdatos,despobla,nifdatos,razosoci,desprovi,pais from spagop,cuentas"
    Aux = Aux & " where codmacta=ctaprove and transfer =" & NumeroTransferencia
    RS.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Importe = 0
    If RS.EOF Then
        'No hayningun registro
        
    Else
        Regs = 0
        While Not RS.EOF
            
                Im = DBLet(RS!imppagad, "N")
                Im = RS!ImpEfect - Im
                Aux = RellenaABlancos(RS!nifdatos, True, 12)
            

            
            
            Aux = "06" & "59" & CodigoOrdenante & Aux   'Ordenante y nifprove
        
            Linea1_68 NFich, Aux, RS, cad
            Linea2_68 NFich, Aux, RS, cad
            Linea3_68 NFich, Aux, RS, cad
            Linea4_68 NFich, Aux, RS, cad
            'Antes
            'Linea5_68 NFich, AUX, RS, Cad, Fecha, Im
            'Ahora en funcion de si los queremos todos juntos o cada uno a su vto
            Linea5_68 NFich, Aux, RS, cad, IIf(PagosJuntos, Fecha, RS!fecefect), Im
            
            
            Linea6_68 NFich, Aux, RS, Im, cad, ConceptoTr
            'If Pagos Then Linea7 NFich, Aux, RS, Cad
        
        
        
        
            Importe = Importe + Im
            Regs = Regs + 1
            RS.MoveNext
        Wend
        'Imprimimos totales
        Totales68 NFich, CodigoOrdenante, Importe, Regs, cad
    End If
    RS.Close
    Set RS = Nothing
    Close (NFich)
    If Regs > 0 Then
        GeneraFicheroNorma68 = True
    Else
        MsgBox "No se han leido registros en la tabala de pagos", vbExclamation
    End If
    Exit Function
EGen:
    MuestraError Err.Number, Err.Description

End Function





Private Sub Cabecera1_68(NF As Integer, ByRef CodOrde As String, Fecha As Date, IBAN As String, ByRef cad As String)

    cad = "03"
    cad = cad & "59"
    'cad = cad & " "
    cad = cad & CodOrde
    cad = cad & Space(12) & "001"
    
    cad = cad & Format(Fecha, "ddmmyy")
    
    'Cuenta bancaria
    cad = cad & Space(9)
    cad = cad & IBAN
    cad = RellenaABlancos(cad, True, 100)
    cad = Mid(cad, 1, 100)
    Print #NF, cad
End Sub







Private Sub Linea1_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "010"
    If IsNull(RS1!razosoci) Then
        cad = cad & RellenaABlancos(RS1!Nommacta, True, 40)
    Else
        cad = cad & RellenaABlancos(RS1!razosoci, True, 40)
    End If
    cad = RellenaABlancos(cad, True, 100)
    cad = Mid(cad, 1, 100)
    Print #NF, cad
End Sub


Private Sub Linea2_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "011"
    cad = cad & RellenaABlancos(DBLet(RS1!dirdatos, "T"), True, 45)
    cad = RellenaABlancos(cad, True, 100)
    cad = Mid(cad, 1, 100)
    Print #NF, cad
End Sub





Private Sub Linea3_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "012"
    cad = cad & RellenaABlancos(DBLet(RS1!codposta, "T"), True, 5) & " "
    cad = cad & RellenaABlancos(DBLet(RS1!desPobla, "T"), True, 40)
    cad = RellenaABlancos(cad, True, 100)
    cad = Mid(cad, 1, 100)
    Print #NF, cad
End Sub

Private Sub Linea4_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "013"
    'De mommento pongo balancos, ya que es para extranjero
    'Cad = Cad & RellenaABlancos(DBLet(RS1!codposta, "T"), False, 5) & " "
    cad = cad & "     "
    cad = cad & RellenaABlancos(DBLet(RS1!desProvi, "T"), True, 30)   'desprovi,pais
    cad = cad & RellenaABlancos(DBLet(RS1!PAIS, "T"), True, 20)   'desprovi,pais
    cad = RellenaABlancos(cad, True, 100)
    cad = Mid(cad, 1, 100)
    Print #NF, cad
End Sub

' Febrero 2016.
' En la cabecera llevamos si queremos todos los pagos a una fecha o cada uno en su vencimiento
' con lo cual aqui siempre enviaremos el valor fecha que ya llevara uno u otro
Private Sub Linea5_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef cad As String, ByRef Fechapag As Date, ByRef Importe1 As Currency)
    cad = CodOrde    'llevara tb la ID del socio
    cad = cad & "014"

    cad = cad & "00000000" 'Numero de pago domiciliado
    
    cad = cad & Format(Fechapag, "ddmmyyyy")
    'Cad = Cad & Format(RS1!Fecefect, "ddmmyyyy") 'fecha vencimiento de cada recibo   'YA VIENEN CARGADA en fecha doc lo que corresponda
   
    cad = cad & RellenaAceros(CStr(Round(Importe1, 2) * 100), False, 12)
    cad = cad & "0" 'presentacion
    cad = cad & "ES1" 'presentacion
    cad = RellenaABlancos(cad, True, 100)
    cad = Mid(cad, 1, 99) & "1"
    Print #NF, cad
End Sub


Private Sub Linea6_68(NF As Integer, ByRef CodOrde As String, ByRef RS1 As ADODB.Recordset, ByRef Importe1 As Currency, ByRef cad As String, vConceptoTransferencia As String)


   
    '
    cad = CodOrde   'llevara tb la ID del socio
    cad = cad & "015"
    cad = cad & "00000000" 'Numero de pago domiciliado
    cad = cad & RellenaABlancos(RS1!NumFactu, False, 12)
    cad = cad & Format(RS1!FecFactu, "ddmmyyyy") 'fecha fac

    cad = cad & RellenaAceros(CStr(Round(Importe1, 2) * 100), False, 12)
    
    cad = cad & "H"
    'Cad = Cad & RellenaABlancos(vConceptoTransferencia, False, 26)
    cad = cad & "ADJUNTAMOS PAGO FACTURA     "
    cad = RellenaABlancos(cad, True, 100)
    cad = Mid(cad, 1, 100)
    Print #NF, cad
End Sub



Private Sub Totales68(NF As Integer, ByRef CodOrde As String, Total As Currency, Registros As Integer, ByRef cad As String)
    cad = "08" & "59"
    cad = cad & CodOrde    'llevara tb la ID del socio
    cad = cad & Space(15)
    cad = cad & RellenaAceros(CStr(Int(Round(Total * 100, 2))), False, 12)
    'Cad = Cad & RellenaAceros(CStr(Registros), False, 8)
    cad = cad & RellenaAceros(CStr((Registros * 6) + 1 + 1), False, 10)
    cad = RellenaABlancos(cad, True, 100)
    Print #NF, cad
End Sub





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
                Aux = Aux & "T" & Format(RS!FecFactu, "ddmmyy") & FrmtStr(RS!NumFactu, 15) & Format(RS!fecefect, "ddmmyy") & Space(15)
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
        B = GrabarDisketteNorma19_SEPA_XML(NomFichero, Remesa, FecPre, TipoReferenciaCliente, Sufijo, FechaCobro, SEPA_EmpresasGraboNIF, Norma19_15, DatosBanco, NifEmpresa_)
    Else
        B = GrabarDisketteNorma19SEPA(NomFichero, Remesa, FecPre, TipoReferenciaCliente, Sufijo, FechaCobro, SEPA_EmpresasGraboNIF, Norma19_15, DatosBanco, NifEmpresa_)
    End If
    GrabarFicheroNorma19SEPA = B
End Function



'Este es el que estaba, y como estaba
Private Function GrabarDisketteNorma19SEPA(NomFichero As String, Remesa As String, FecPre As String, TipoReferenciaCliente As Byte, Sufijo As String, FechaCobro As String, SEPA_EmpresasGraboNIF As Boolean, Norma19_15 As Boolean, DatosCtaBanco As String, NifEmpresa As String) As Boolean

Dim ValorEnOpcionales As Boolean
   

    Dim ImpEfe As Currency

    '
    Dim IdDeudor As String
    Dim Cuenta As String
    Dim Fecha As Date
    
    
    Dim NRegistros003(2) As Integer  'registros003 para fecha, proveedor , total
    Dim NumLineas(2) As Integer   'fecha, proveedor , total
    Dim Totales(2) As Currency    'x Fecha,proveedor,total
    
    Dim EsPersonaJuridica As Boolean
    
    Dim J As Integer
    Dim IdNorma As String  '1914 o 1915
    
    On Error GoTo Err_Remesa19sepa


    
    If Norma19_15 Then
        IdNorma = "19154"
    Else
        IdNorma = "19143"
    End If
    
    '-- Abrir el fichero a enviar
    NF = FreeFile()
    Open NomFichero For Output As #NF
    
    SQL = "select  numserie,codfaccl,fecfaccl,numorden,scobro.codmacta,codrem,anyorem,Tiporem,"
    'Por parametro
    'If vParam.Norma19xFechaVto Then
    '    SQL = SQL & " fecvenci"
    'Else
    '    SQL = SQL & "'" & Format(FecCobro, FormatoFecha) & "'"
    'End If
    If FechaCobro = "" Then
        SQL = SQL & " fecvenci"
    Else
        SQL = SQL & "'" & Format(FechaCobro, FormatoFecha) & "'"
    End If
    
    
    
    SQL = SQL & " as fecvenci,impvenci,ctabanc1,codbanco"
    SQL = SQL & ",codsucur,digcontr,scobro.cuentaba,text33csb,text41csb,gastos,scobro.iban"
    SQL = SQL & ",nommacta,nifdatos,dirdatos,codposta,despobla, desprovi,pais,bic,referencia  from scobro"
    SQL = SQL & "  left join sbic on codbanco=entidad inner join cuentas on "
    SQL = SQL & " scobro.codmacta = cuentas.codmacta WHERE "
    SQL = SQL & " codrem = " & RecuperaValor(Remesa, 1)
    SQL = SQL & " AND anyorem=" & RecuperaValor(Remesa, 2)
    
    
    'sepa
    SQL = SQL & " order by  fecvenci,nifdatos,scobro.codmacta"
    
    Remesa = RecuperaValor(Remesa, 1)
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        
        
     
    
        'Cabecera de presentador
        Registro = "01" & IdNorma & "001"   '19143-> Podria ser 19154 ver pdf
        
        SQL = Trim(NifEmpresa) + "ES00"   'Identificacion acreedor
        SQL = CadenaTextoMod97(SQL)
        'Si no es dos digitos es un mensaje de error
        If Len(SQL) <> 2 Then Err.Raise 513, , SQL
        SQL = "ES" & SQL & Sufijo & NifEmpresa
        Registro = Registro & FrmtStr(SQL, 35)
        
        
        
        
        SQL = vEmpresa.nomempre
        Registro = Registro & FrmtStr(SQL, 70)
        
        SQL = Format(FecPre, "yyyymmdd")
        Registro = Registro & FrmtStr(SQL, 8)
        
        
        SQL = "PRE" & Format(Now, "yyyymmddhhnnss")
        'Los milisegundos
        SQL = SQL & Format((Timer - Int(Timer)) * 10000, "0000") & "0"
        'Idententificacion propia
        '   tiporem,codrem,anyorem
        SQL = SQL & "RE" & miRsAux!Tiporem & Format(miRsAux!CodRem, "000000") & Format(miRsAux!AnyoRem, "0000")
        Registro = Registro & FrmtStr(SQL, 35)
        
        
        'Idenficacion entidad/oficina receptora
        Registro = Registro & RecuperaValor(DatosCtaBanco, 1) ' Código de entidad receptora
        Registro = Registro & RecuperaValor(DatosCtaBanco, 2) ' Código de oficina receptora
        Registro = Registro & Space(434) 'LIBRE
        
        
        Print #NF, Registro
       

        'VA por cliente / fecha
        '
        '----------------
        Cuenta = ""
        

        
        'Variables
        'Inicializamos el total general
        NRegistros003(2) = 0: NumLineas(2) = 0: Totales(2) = 0
        While Not miRsAux.EOF
            
            'NumLineas(0)  cuantos registros para una fecha
            'NumLineas(1)  cuantas fechas por proveedor
            
            
            'If Cuenta <> miRsAux!nifdatos Then
            If Cuenta = "" Then   'SOLO IMPRIMIREMOS UNO
                'Cuenta NUEVA
                'If Cuenta <> "" Then
                '
                '    'Cerramos la fecha primero
                '    ImprimiSEPA_ProveedorFecha2 1, IdDeudor, Fecha, NRegistros003(0), Totales(0), NumLineas(0) 'pq me da lo mismo
                '    NumLineas(1) = NumLineas(1) + 1
                '    'Ahora el proveedor
                '    ImprimiSEPA_ProveedorFecha2 2, IdDeudor, miRsAux!fecvenci, NRegistros003(1), Totales(1), NumLineas(1)
                '
                '    NumLineas(2) = NumLineas(2) + 2  'Total lineas. Dos mas pq una es del total fec y otra del total prov
                'End If
            
                'Cuenta = miRsAux!nifdatos
                Cuenta = NifEmpresa
                Fecha = miRsAux!FecVenci
                'Resitros03 para la fecha y el proveedor
                'Total para fecha y proveedor
                For J = 0 To 1
                    NRegistros003(J) = 0: NumLineas(J) = 0: Totales(J) = 0
                Next
                
                'Me guardo el identeficador del deudor
                'SQL = Trim(miRsAux!nifdatos) + "ES00"
                SQL = NifEmpresa + "ES00"
                SQL = CadenaTextoMod97(SQL)
                
                'Si no es dos digitos es un mensaje de error
                If Len(SQL) <> 2 Then Err.Raise 513, , SQL
                'IdDeudor = Mid("ES" & SQL & "000" & miRsAux!nifdatos & Space(25), 1, 35)
                IdDeudor = Mid("ES" & SQL & Sufijo & NifEmpresa & Space(25), 1, 35)
                
                'Si que suma para las lineas totales
                'es global. Para no pasar los datos de la cuenta
                
                SQL = RecuperaValor(DatosCtaBanco, 5)
                For J = 1 To 4
                    SQL = SQL & RecuperaValor(DatosCtaBanco, J)
                Next
                ImprimiSEPA_ProveedorFecha2 0, IdDeudor, Fecha, 0, 0, 0, IdNorma
                NumLineas(2) = NumLineas(2) + 1
                
                
            End If
            
            If Fecha <> miRsAux!FecVenci Then
                
                'FIN fecha
                'Guardamos el totalfechaproveedor anterior
                'NRegistros003(J)  NumLineas(J)  Totales(J) = 0
                ImprimiSEPA_ProveedorFecha2 1, IdDeudor, Fecha, NRegistros003(0), Totales(0), NumLineas(0), IdNorma 'pq me da lo mismo
                
                
                
                'Cabcera nueva fecha
                
                SQL = RecuperaValor(DatosCtaBanco, 5)
                For J = 1 To 4
                    SQL = SQL & RecuperaValor(DatosCtaBanco, J)
                Next

                
                Fecha = miRsAux!FecVenci
                ImprimiSEPA_ProveedorFecha2 0, IdDeudor, Fecha, 0, 0, 0, IdNorma
                NumLineas(1) = NumLineas(1) + 1
                NumLineas(2) = NumLineas(2) + 1
                
                
                'Reseteamos por fecha
                NRegistros003(0) = 0: NumLineas(0) = 0: Totales(0) = 0
                NumLineas(1) = NumLineas(1) + 1
                NumLineas(2) = NumLineas(2) + 1
                
                
            End If
            
            
            'Sumamos el total y una linea mas para cada uno
            ImpEfe = DBLet(miRsAux!Gastos, "N")
            ImpEfe = miRsAux!ImpVenci + ImpEfe
            For J = 0 To 2
                Totales(J) = Totales(J) + ImpEfe
                NRegistros003(J) = NRegistros003(J) + 1
                NumLineas(J) = NumLineas(J) + 1
            Next
            
            'Restro 1º INDIVIDUAL
            Registro = "03" & IdNorma & "003"   '19143-> Podria ser 19154 ver pdf
            
            'Referencia del adeudo
            
            SQL = FrmtStr(miRsAux!codmacta, 10) & FrmtStr(miRsAux!NUmSerie, 3) & Format(miRsAux!codfaccl, "00000000")
            SQL = SQL & Format(miRsAux!fecfaccl, "yyyymmdd") & Format(miRsAux!numorden, "000")
            Registro = Registro & FrmtStr(SQL, 35)
            
                
            
            'Referencia unica del mandato
            'Esta comentado Registro = Registro & FrmtStr(SQL, 35)  'Si tuviera en el campo 9 va la fecha
        
         
            'Opcion nueva: 3   Quiere el campo referencia de scobro
            Select Case TipoReferenciaCliente
            Case 1
                'ALZIRA. La referencia final de 12 es el ctan bancaria del cli + su CC
                SQL = Format(miRsAux!digcontr, "00") ' Dígitos de control
                SQL = SQL & Format(miRsAux!Cuentaba, "0000000000") ' Código de cuenta
            Case 2
                'NIF
                SQL = DBLet(miRsAux!nifdatos, "T")
                If SQL = "" Then SQL = miRsAux!codmacta
             
                
            Case 3
                'Referencia en el VTO. No es Nula
                SQL = DBLet(miRsAux!referencia, "T")
                If SQL = "" Then SQL = miRsAux!codmacta
            Case Else
                
                SQL = miRsAux!codmacta
                
            End Select
            Registro = Registro & FrmtStr(SQL, 35)
            
            
        
            'Tipo de adeudo
            Registro = Registro & "RCUR"   'FNAL  FRST OFF RCUR

            'Categoria de adeudo Habra que ponerlo en la pantalla  AT-59
            Registro = Registro & "TRAD"
            
            'Importe efecto
            Registro = Registro & Format(ImpEfe * 100, String(11, "0")) ' Importe
        
            'Ver PDF. (AT-25)
            Registro = Registro & "20091031"   '31-10-2009,
        
            'aqui pone que va el BIC 11pos
            SQL = DBLet(miRsAux!BIC, "T")
            Registro = Registro & FrmtStr(SQL, 11)
        
            'Datos deudor 70+50+50+40+2
            Registro = Registro & DatosBasicosDelDeudor
            
            'Tipo identificador deudor.  Persona fisica (2) o juridica (1)
            SQL = Mid(miRsAux!nifdatos, 1, 1)
            EsPersonaJuridica = Not IsNumeric(SQL)
            
            SQL = "2"
            If EsPersonaJuridica Then SQL = "1"
            Registro = Registro & SQL
            
            
            'Campo 17. Identifiacion del deudor
            If EsPersonaJuridica Then
                
                '15/01/2014  ->Sepa_EmpresasGraboNIF
                'Metemos una A y luego el BIC
                'Metmos una I o el NIF
                If SEPA_EmpresasGraboNIF Then
                    SQL = "I"
                    Registro = Registro & FrmtStr(SQL & miRsAux!nifdatos, 36)
                Else
                    SQL = "A"
                    Registro = Registro & FrmtStr(SQL & DBLet(miRsAux!BIC, "T"), 36)
                End If
            Else
                SQL = "J"
                Registro = Registro & FrmtStr(SQL & miRsAux!nifdatos, 36)
            End If
            

            
            
            'Campo 18
            SQL = miRsAux!nifdatos
            If Not EsPersonaJuridica Then SQL = ""   'Si no es juridica NO se graba nada, si es, lo mismo que el anterior
            Registro = Registro & FrmtStr(SQL, 35)
            
            'Campo 19
            Registro = Registro & "A"  'IBAN a piñon la a
            'Canpo 20
            Registro = Registro & IBAN_Destino(True) & Space(10) 'reserva 34
            'Campo 21. Proposito del adeudo. Tablas PDF
            Registro = Registro & "TRAD"
            'Campo 22. Concepto. Los 140 carcateres del txtcsb3
            'JUNIO 2014. Lo hacia todo junto
            'SQL = DBLet(miRsAux!text33csb, "T") & DBLet(miRsAux!text41csb, "T")
            'Registro = Registro & FrmtStr(SQL, 140)
            
            SQL = DBLet(miRsAux!text33csb, "T")
            Registro = Registro & FrmtStr(SQL, 80)
            SQL = DBLet(miRsAux!text41csb, "T")
            Registro = Registro & FrmtStr(SQL, 60)
            
            
            'Libre
            Registro = Registro & Space(19)
            
            Print #NF, Registro
            
            'Siguiente
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        
        'Falta imprimir el total fecha, total proveedor y total general
        ImprimiSEPA_ProveedorFecha2 1, IdDeudor, Fecha, NRegistros003(0), Totales(0), NumLineas(0), IdNorma 'pq me da lo mismo
        NumLineas(1) = NumLineas(1) + 1
        'Ahora el proveedor
        ImprimiSEPA_ProveedorFecha2 2, IdDeudor, Fecha, NRegistros003(1), Totales(1), NumLineas(1), IdNorma
        NumLineas(2) = NumLineas(2) + 2  'Total lineas. Dos mas pq una es del total fec y otra del total prov
        'TOTAL GENERAL
        ImprimiSEPA_ProveedorFecha2 3, IdDeudor, Fecha, NRegistros003(2), Totales(2), NumLineas(2), IdNorma
              
    End If  'De EOF
    Close #NF
    If NRegistros003(2) > 0 Then GrabarDisketteNorma19SEPA = True
    Exit Function
Err_Remesa19sepa:
    MsgBox "Err: " & Err.Number & vbCrLf & _
        Err.Description, vbCritical, "Grabación del diskette de Remesa SEPA"
        

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
        cad = cad & SQL & Space(10)  'El iban son 24 y dejan hasta 34 psociones
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





'-----------------  Norma 34
Private Function GeneraFicheroNorma34SEPA(CIF As String, Fecha As Date, CuentaPropia2 As String, NumeroTransferencia As Long, Pagos As Boolean, ConceptoTr As String) As Boolean
Dim Regs As Integer
Dim Importe As Currency
Dim Im As Currency
Dim cad As String
Dim Aux As String
Dim SufijoOEM As String

    On Error GoTo EGen2
    GeneraFicheroNorma34SEPA = False
    

    
    
    'Cargamos la cuenta
    cad = "Select * from ctabancaria where codmacta='" & CuentaPropia2 & "'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        cad = ""
    Else
        If IsNull(miRsAux!Entidad) Then
            cad = ""
        Else
            SufijoOEM = "000" ''Sufijo3414
            cad = miRsAux!IBAN & Format(miRsAux!Entidad, "0000") & Format(DBLet(miRsAux!Oficina, "T"), "0000") & DBLet(miRsAux!Control, "T") & Format(DBLet(miRsAux!CtaBanco, "T"), "0000000000")
            If DBLet(miRsAux!Sufijo3414, "T") <> "" Then SufijoOEM = Right("000" & miRsAux!Sufijo3414, 3)
            CuentaPropia2 = cad
        End If
        
        
       
    End If
    miRsAux.Close
  
    If cad = "" Then
        MsgBox "Error leyendo datos para: " & CuentaPropia2, vbExclamation
        Exit Function
    End If
    
    NF = FreeFile
    Open App.Path & "\norma34.txt" For Output As #NF
    
    
    
    'SEPA
    '1.- Cabecera ordenante
    '------------------------------------------------------------------------
    cad = "01" & "ORD" & "34145" & "001" & CIF
        
    'sufijo (Tenemos el OEM, que se utiliza para las otras normas antiguas
    cad = cad & SufijoOEM
    cad = cad & Format(Now, "yyyymmdd")
    cad = cad & Format(Fecha, "yyyymmdd")
    cad = cad & "A" 'IBAN
     
    'EL IBAN propiamente
    cad = cad & FrmtStr(CuentaPropia2, 34)
    cad = cad & "0" 'Cargo por cada operacion
    
    'Nombre
    miRsAux.Open "Select siglasvia ,direccion ,numero ,codpobla,pobempre,provempre,provincia from empresa2"
    cad = cad & FrmtStr(vEmpresa.nomempre, 70)
    If miRsAux.EOF Then
        cad = cad & FrmtStr("", 140)
    Else
        'Direccion
        cad = cad & FrmtStr(Trim(DBLet(miRsAux!siglasvia, "T") & " " & miRsAux!Direccion & " " & DBLet(miRsAux!numero, "T")), 50)
        cad = cad & FrmtStr(Trim(DBLet(miRsAux!CodPobla, "T") & " " & miRsAux!pobempre), 50)
        cad = cad & FrmtStr(DBLet(miRsAux!provincia, "T"), 40)
    End If
    miRsAux.Close
    
    'Pais y libre
    cad = cad & "ES" & FrmtStr("", 311)
    Print #NF, cad
  
  
  
    '2.- Registro cabecera TRANSFERENCIA
    '------------------------------------------------------------------------
    cad = "02" & "SCT" & "34145" & CIF
        
    'sufijo (Tenemos el OEM, que se utiliza para las otras normas
    cad = cad & SufijoOEM
    cad = cad & FrmtStr("", 578)
    Print #NF, cad
    
    
    
    'Para ello abrimos la tabla tmpNorma34
    If Pagos Then
        cad = "Select spagop.*,nommacta,dirdatos,codposta,dirdatos,desprovi,pais,cuentas.despobla,bic from spagop"
        cad = cad & " left join sbic on spagop.entidad=sbic.entidad INNER JOIN cuentas ON"
        cad = cad & " codmacta=ctaprove WHERE transfer =" & NumeroTransferencia
    Else
        'ABONOS
         '
        cad = "Select scobro.codbanco as entidad,scobro.codsucur as oficina,scobro.cuentaba,scobro.digcontr as CC,scobro.iban"
        cad = cad & ",nommacta,dirdatos,codposta,despobla,impvenci,scobro.codmacta,pais,Gastos,impcobro,desprovi"
        cad = cad & " ,NUmSerie,codfaccl,fecfaccl,numorden,text33csb,text41csb,bic from scobro"
        cad = cad & " LEFT JOIN sbic on scobro.codbanco=sbic.entidad INNER JOIN cuentas ON"
        cad = cad & " cuentas.codmacta=scobro.codmacta WHERE transfer =" & NumeroTransferencia
    End If
    miRsAux.Open cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    If Not miRsAux.EOF Then
        cad = "#"
        While Not miRsAux.EOF
            If IsNull(miRsAux!BIC) Then
                If InStr(1, cad, "#" & miRsAux!Entidad & "#") = 0 Then cad = cad & miRsAux!Entidad & "#"
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.MoveFirst
        
        
        If Len(cad) > 1 Then
            cad = Mid(cad, 2)
            cad = Mid(cad, 1, Len(cad) - 1)
            cad = Replace(cad, "#", "   /   ")
            cad = "Bancos sin BIC asignado:" & vbCrLf & cad & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then
                miRsAux.Close
                Close (NF)
                Exit Function
            End If
        End If
        
    End If
    
    
    Regs = 0
    Importe = 0
    If miRsAux.EOF Then
        'No hayningun registro

    Else
        While Not miRsAux.EOF
            If Pagos Then
                Im = DBLet(miRsAux!imppagad, "N")
                Im = miRsAux!ImpEfect - Im
                Aux = miRsAux!ctaprove

            Else
                Im = Abs(miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N")) - DBLet(miRsAux!impcobro, "N")
                Aux = miRsAux!codmacta
            End If
            Aux = FrmtStr(Aux, 10)
            Importe = Importe + Im
            Regs = Regs + 1
            
            'Campo 1,2,3
            cad = "03" & "SCT" & "34145" & "002"
            
            'Campo 5 . Referencia del ordenante
            If Pagos Then
                'numfactu fecfactu numorden
                Aux = Aux & FrmtStr(miRsAux!NumFactu, 10)
                Aux = Aux & Format(miRsAux!FecFactu, "yyyymmdd") & Format(miRsAux!numorden, "000")
            
            Else
                'fecfaccl
                Aux = Aux & FrmtStr(miRsAux!NUmSerie, 3) & Format(miRsAux!codfaccl, "00000000")
                Aux = Aux & Format(miRsAux!fecfaccl, "yyyymmdd") & Format(miRsAux!numorden, "000")
            End If
            cad = cad & FrmtStr(Aux, 35)
            
            'Campo 6
            cad = cad & "A"
            
            'IBAN
            cad = cad & FrmtStr(IBAN_Destino(False), 34)
            
            'Campo8 Importe
            cad = cad & Format(Im * 100, String(11, "0")) ' Importe
            
            'Campo9
            cad = cad & "3" 'gastos compartidos
            'Campo 10
            cad = cad & FrmtStr(DBLet(miRsAux!BIC, "T"), 11) 'BIC

            'nommacta,dirdatos,codposta,dirdatos,despobla,impvenci,scobro.codmacta
            'Datos Basicos del beneficiario
            cad = cad & DatosBasicosDelDeudor
            
            'Campo16 ID del pago. Concepto
            If Pagos Then
                ''`text1csb` `text2csb`
                Aux = DBLet(miRsAux!text1csb, "T") & DBLet(miRsAux!text2csb, "T")
            Else
                '`text33csb` `text41csb`
                Aux = DBLet(miRsAux!text33csb, "T") & DBLet(miRsAux!text41csb, "T")
            End If
            cad = cad & FrmtStr(Aux, 140)
            
            'Campo17
            cad = cad & FrmtStr("", 35)  'Reservado
            
            'Campo18  campo19
            
           
            
            If ConceptoTr = "1" Then
                cad = cad & "SALASALA"
            ElseIf ConceptoTr = "0" Then
                cad = cad & "PENSPENS"
            Else
                cad = cad & "TRADTRAD"
            End If
            
           
            
            cad = cad & FrmtStr("", 99)  'libre
            
            Print #NF, cad
            
            miRsAux.MoveNext
        Wend
        
    
        'TOTALES
        '----------------------------------
        'Total trasnferencia SEPA
        'Campo 1,2
        cad = "04" & "SCT"
        
        'Campo3 Importe total
        cad = cad & Format(Importe * 100, String(17, "0")) ' Importe
        cad = cad & Format(Regs, String(8, "0")) ' Importe
        'Total registros son
        'Reg(numreo de adeudos + 1 reg01 + un reg02 + reg04
        cad = cad & Format(Regs + 2, String(10, "0")) ' Importe   '2014-01-29  HABIA un reg + 3
        cad = cad & FrmtStr("", 560)  'libre
        Print #NF, cad
        
        'Total general
        cad = "99" & "ORD"
        
        'Campo3 Importe total
        cad = cad & Format(Importe * 100, String(17, "0")) ' Importe
        cad = cad & Format(Regs, String(8, "0")) ' Importe
        
        'Igual que arriba as uno
        'Reg(numreo de adeudos + 1 reg01 + un reg02 + reg04  +1
        cad = cad & Format(Regs + 4, String(10, "0")) ' Importe
        cad = cad & FrmtStr("", 560)  'libre
        Print #NF, cad
        
        
        
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    Close (NF)
    If Regs > 0 Then GeneraFicheroNorma34SEPA = True
    Exit Function
EGen2:
    MuestraError Err.Number, Err.Description

End Function






'---------------------------------------------------------------------
'  DEVOLUCION FICHERO  SEPA
'---------------------------
Public Sub ProcesaCabeceraFicheroDevolucionSEPA(Fichero As String, ByRef Remesa As String)
Dim AUX2 As String  'Para buscar los vencimientos
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
                AUX2 = Mid(SQL, 1, Cuantos - 1)
                SQL = Mid(SQL, Cuantos + 1)
                
                
                'En aux2 tendre codrem|anñorem|
                AUX2 = RecuperaValor(AUX2, 1) & " AND anyo = " & RecuperaValor(AUX2, 2)
                AUX2 = "Select situacion from remesas where codigo = " & AUX2
                miRsAux.Open AUX2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If miRsAux.EOF Then
                    AUX2 = "-No se encuentra remesa"
                Else
                    'Si que esta.
                    'Situacion
                    If CStr(miRsAux!Situacion) <> "Q" And CStr(miRsAux!Situacion) <> "Y" Then
                        AUX2 = "- Situacion incorrecta : " & miRsAux!Situacion
                    Else
                        AUX2 = "" 'TODO OK
                    End If
                End If
            
                If AUX2 <> "" Then
                    AUX2 = AUX2 & " ->" & Mid(miRsAux.Source, InStr(1, UCase(miRsAux.Source), " WHERE ") + 7)
                    AUX2 = Replace(AUX2, " AND ", " ")
                    AUX2 = Replace(AUX2, "anyo", "año")
                    Registro = Registro & vbCrLf & AUX2
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
