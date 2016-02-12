Attribute VB_Name = "libNormaXML"
Option Explicit



'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'
'
'
' SEPA en XML
'
'
'
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////


Dim NFic As Integer   'Para no tener que pasarlo a todas las funciones

Private Function XML(CADENA As String) As String
Dim I As Integer
Dim AUX As String
Dim Le As String
Dim C As Integer
    'Carácter no permitido en XML  Representación ASCII
    '& (ampersand)          &amp;
    '< (menor que)          &lt;
    ' > (mayor que)         &gt;
    '“ (dobles comillas)    &quot;
    '' (apóstrofe)          &apos;
    
    'La ISO recomienda trabajar con los carcateres:
    'a b c d e f g h i j k l m n o p q r s t u v w x y z
    'A B C D E F G H I J K L M N O P Q R S T U V W X Y Z
    '0 1 2 3 4 5 6 7 8 9
    '/ - ? : ( ) . , ' +
    'Espacio
    AUX = ""
    For I = 1 To Len(CADENA)
        Le = Mid(CADENA, I, 1)
        C = Asc(Le)
        
        
        Select Case C
        Case 40 To 57
            'Caracteres permitidos y numeros
            
        Case 65 To 90
            'Letras mayusculas
            
        Case 97 To 122
            'Letras minusculas
            
        Case 32
            'espacio en balanco
            
        Case Else
            Le = " "
        End Select
        AUX = AUX & Le
    Next
    XML = AUX
End Function



Public Function GrabarDisketteNorma19_SEPA_XML(NomFichero As String, Remesa_ As String, FecPre As String, TipoReferenciaCliente As Byte, Sufijo As String, FechaCobro As String, SEPA_EmpresasGraboNIF As Boolean, Norma19_15 As Boolean, DatosBanco As String, NifEmpresa As String) As Boolean
    Dim ValorEnOpcionales As Boolean
    '-- Genera_Remesa: Esta función genera la remesa indicada, en el fichero correspondiente
    
    
    Dim SQL As String
    Dim ImpEfe As Currency

    '
    Dim IdDeudor As String
    Dim Cuenta As String
    Dim Fecha2 As Date
    Dim FinFecha As Boolean


    Dim EsPersonaJuridica As Boolean
    
    Dim J As Integer
    'Dim IdNorma As String  '1914 o 1915
    
    On Error GoTo Err_Remesa19sepa
    
    
    
    

    '-- Abrir el fichero a enviar
    NFic = FreeFile()
    Open NomFichero For Output As #NFic
    
    SQL = "select  numserie,codfaccl,fecfaccl,numorden,scobro.codmacta,codrem,anyorem,Tiporem,"
    
    'SEPTIEMBRE 2015
    'Todos van a la fecha de vencimiento
'    If vParam.Norma19xFechaVto Then
'        SQL = SQL & " fecvenci"
'    Else
'        SQL = SQL & "'" & Format(FecCobro, FormatoFecha) & "'"
'    End If
    'OCTUBRE. Si no indica fecha cobro, ira cada una con su vencimiento, si no la fecha de cobro
    
    If FechaCobro = "" Then
        SQL = SQL & " fecvenci"
    Else
        SQL = SQL & "'" & Format(FechaCobro, FormatoFecha) & "'"
    End If



    
    SQL = SQL & " as fecvenci,impvenci,ctabanc1,codbanco"
    SQL = SQL & ",codsucur,digcontr,scobro.cuentaba,text33csb,text41csb,gastos,scobro.iban"
    SQL = SQL & ",nommacta,nifdatos,dirdatos,codposta,despobla, desprovi,pais,bic,referencia,SEPA_Refere,SEPA_FecFirma  from scobro"
    SQL = SQL & "  left join sbic on codbanco=entidad inner join cuentas on "
    SQL = SQL & " scobro.codmacta = cuentas.codmacta WHERE "
    SQL = SQL & " codrem = " & RecuperaValor(Remesa_, 1)
    SQL = SQL & " AND anyorem=" & RecuperaValor(Remesa_, 2)
    
    
    'sepa
    SQL = SQL & " order by  fecvenci,nifdatos,scobro.codmacta"
    
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    SQL = ""
    If Not miRsAux.EOF Then
        
            'Encabezado
            Print #NFic, "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
            Print #NFic, "<Document xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""urn:iso:std:iso:20022:tech:xsd:pain.008.001.02"">"
            Print #NFic, "<CstmrDrctDbtInitn>"
                        
            Print #NFic, "<GrpHdr>"
            
            

            SQL = "PRE" & Format(Now, "yyyymmddhhnnss")
            'Los milisegundos
            SQL = SQL & Format((Timer - Int(Timer)) * 10000, "0000") & "0"
            'Idententificacion propia
            '   tiporem,codrem,anyorem
            SQL = SQL & "RE" & miRsAux!Tiporem & Format(miRsAux!CodRem, "000000") & Format(miRsAux!AnyoRem, "0000")
                    
            
            Print #NFic, "<MsgId>" & SQL & "</MsgId>"
            
            SQL = Format(Now, "yyyy-mm-dd") & "T" & Format(Now, "hh:mm:ss")   '<CreDtTm>2015-09-10T16:26:56</CreDtTm>
            Print #NFic, "   <CreDtTm>" & SQL & "</CreDtTm>"
            
            'Control sumatorio y numero de registro
            
            SQL = " codrem = " & RecuperaValor(Remesa_, 1) & " AND anyorem=" & RecuperaValor(Remesa_, 2) & " AND 1"
            SQL = DevuelveDesdeBD("concat(count(*),'|',sum(coalesce(gastos,0)+impvenci),'|')", "scobro", SQL, "1")
            Print #NFic, "   <NbOfTxs>" & RecuperaValor(SQL, 1) & "</NbOfTxs>"
            SQL = RecuperaValor(SQL, 2)
            Print #NFic, "   <CtrlSum>" & SQL & "</CtrlSum>"
            
            
            'Empezamos datos
            Print #NFic, "   <InitgPty>"
            Print #NFic, "     <Nm>" & XML(vEmpresa.nomempre) & "</Nm>"
            Print #NFic, "     <Id>"
                        
            'Tipo identificador deudor.  Persona fisica (2) o juridica (1)
            SQL = Mid(NifEmpresa, 1, 1)
            EsPersonaJuridica = Not IsNumeric(SQL)
            If EsPersonaJuridica Then
                Print #NFic, "        <OrgId>"
            Else
                Print #NFic, "        <PrvtId>"
            End If
            
            SQL = Trim(NifEmpresa) + "ES00"   'Identificacion acreedor
            SQL = CadenaTextoMod97(SQL)
            'Si no es dos digitos es un mensaje de error
            If Len(SQL) <> 2 Then Err.Raise 513, , SQL
            SQL = "ES" & SQL & Sufijo & NifEmpresa
            Print #NFic, "           <Othr>"
            Print #NFic, "              <Id>" & SQL & "</Id>"   'Ejemplo: ES3100024348588Y
            Print #NFic, "           </Othr>"
            
            If EsPersonaJuridica Then
                Print #NFic, "        </OrgId>"
            Else
                Print #NFic, "        </PrvtId>"
            End If
            
            
            Print #NFic, "      </Id>"
            Print #NFic, "   </InitgPty>"
            Print #NFic, "</GrpHdr>"
        
        
            
            
            Fecha2 = "01/01/1900"
            FinFecha = False
            While Not miRsAux.EOF
            
                'Informacion del PAGO.
                ' Se imprime una vez cada FECHA
                If Fecha2 <> miRsAux!fecvenci Then
                        
                        If Fecha2 > CDate("01/02/1900") Then Print #NFic, "</PmtInf>"
                        Fecha2 = miRsAux!fecvenci
                        
                        
                        'Previo envio vtos
                       Print #NFic, "<PmtInf>"

                        'SQL = "RE" & miRsAux!Tiporem & Format(miRsAux!CodRem, "000000") & Format(miRsAux!AnyoRem, "0000") & " " & Format(Fecha2, "dd/mm/yyyy")
                        SQL = "RE" & Format(miRsAux!CodRem, "00000") & Format(miRsAux!AnyoRem, "0000") & " " & Format(FecPre, "dd/mm/yy") & NifEmpresa
                        
                        Print #NFic, "   <PmtInfId>" & SQL & "</PmtInfId>"
                        Print #NFic, "   <PmtMtd>DD</PmtMtd>"             'DirectDebit
                        Print #NFic, "   <BtchBookg>false</BtchBookg>"    'True: un apunte por cada recib   False: Por el total
                        Print #NFic, "   <PmtTpInf>"
                        Print #NFic, "      <SvcLvl>"
                        Print #NFic, "          <Cd>SEPA</Cd>"
                        Print #NFic, "      </SvcLvl>"
                        Print #NFic, "      <LclInstrm>"
                        Print #NFic, "         <Cd>COR1</Cd>"   'CORE o COR1
                        Print #NFic, "      </LclInstrm>"
                        Print #NFic, "      <SeqTp>RCUR</SeqTp>"
                        Print #NFic, "      <CtgyPurp>"
                        Print #NFic, "         <Cd>TRAD</Cd>"
                        Print #NFic, "      </CtgyPurp>"
                        Print #NFic, "   </PmtTpInf>"
                        'Print #NFic, "   <ReqdColltnDt>" & Format(FecCobro, "yyyy-mm-dd") & "</ReqdColltnDt>"
                        Print #NFic, "   <ReqdColltnDt>" & Format(Fecha2, "yyyy-mm-dd") & "</ReqdColltnDt>"
                        Print #NFic, "   <Cdtr>"
                        Print #NFic, "      <Nm>" & XML(vEmpresa.nomempre) & "</Nm>"
                        Print #NFic, "      <PstlAdr>"
                        Print #NFic, "          <Ctry>ES</Ctry>"
                        SQL = "Direccion"
                        SQL = DBLet(miRsAux!dirdatos, "T")
                        If SQL <> "" Then Print #NFic, "          <AdrLine>" & XML(SQL) & "</AdrLine>"
                        Print #NFic, "      </PstlAdr>"
                        Print #NFic, "   </Cdtr>"
                        Print #NFic, "   <CdtrAcct>"
                        Print #NFic, "      <Id>"
                        'IBAN
                        SQL = RecuperaValor(DatosBanco, 5)
                        For J = 1 To 4
                            SQL = SQL & RecuperaValor(DatosBanco, J)
                        Next
            
                        Print #NFic, "         <IBAN>" & SQL & "</IBAN>"
                        Print #NFic, "      </Id>"
                        Print #NFic, "   </CdtrAcct>"
                        Print #NFic, "   <CdtrAgt>"
                        Print #NFic, "      <FinInstnId>"
                        SQL = RecuperaValor(DatosBanco, 1)
                        SQL = DevuelveDesdeBD("bic", "sbic", "entidad", SQL)
                        Print #NFic, "         <BIC>" & Trim(SQL) & "</BIC>"
                        Print #NFic, "      </FinInstnId>"
                        Print #NFic, "   </CdtrAgt>"
                        
                        Print #NFic, "   <CdtrSchmeId>"
                        Print #NFic, "       <Id>"
                        Print #NFic, "          <PrvtId>"
                        Print #NFic, "             <Othr>"
                        
                        SQL = Trim(NifEmpresa) + "ES00"   'Identificacion acreedor
                        SQL = CadenaTextoMod97(SQL)
                        'Si no es dos digitos es un mensaje de error
                        If Len(SQL) <> 2 Then Err.Raise 513, , SQL
                        SQL = "ES" & SQL & Sufijo & NifEmpresa
                        Print #NFic, "                 <Id>" & SQL & "</Id>"
                        Print #NFic, "                 <SchmeNm><Prtry>SEPA</Prtry></SchmeNm>"
                        Print #NFic, "             </Othr>"
                        Print #NFic, "          </PrvtId>"
                        Print #NFic, "       </Id>"
                        Print #NFic, "   </CdtrSchmeId>"
                End If
                
            
            
            
            
                'Tipo identificador deudor.  Persona fisica (2) o juridica (1)
                SQL = Mid(miRsAux!nifdatos, 1, 1)
                EsPersonaJuridica = Not IsNumeric(SQL)
                
                
                
            
            
                Print #NFic, "   <DrctDbtTxInf>"
                Print #NFic, "      <PmtId>"
                
                'Referencia del adeudo
                SQL = FrmtStr(miRsAux!codmacta, 10) & FrmtStr(miRsAux!NUmSerie, 3) & Format(miRsAux!codfaccl, "00000000")
                SQL = SQL & Format(miRsAux!fecfaccl, "yyyymmdd") & Format(miRsAux!numorden, "000")
                SQL = FrmtStr(SQL, 35)
                Print #NFic, "          <EndToEndId>" & SQL & "</EndToEndId>"
                Print #NFic, "      </PmtId>"
                
                
                ImpEfe = DBLet(miRsAux!Gastos, "N")
                ImpEfe = miRsAux!impvenci + ImpEfe
                Print #NFic, "      <InstdAmt Ccy=""EUR"">" & TransformaComasPuntos(CStr(ImpEfe)) & "</InstdAmt>"
                Print #NFic, "      <DrctDbtTx>"
                Print #NFic, "         <MndtRltdInf>"
                
                'Si la cuenta tiene ORDEN de mandato, coge este
                SQL = DBLet(miRsAux!SEPA_Refere, "T")
                If SQL = "" Then
                    Select Case TipoReferenciaCliente
                    Case 1
                        'ALZIRA. La referencia final de 12 es el ctan bancaria del cli + su CC
                        SQL = Format(miRsAux!digcontr, "00") ' Dígitos de control
                        SQL = SQL & Format(miRsAux!cuentaba, "0000000000") ' Código de cuenta
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
                End If
                Print #NFic, "            <MndtId>" & SQL & "</MndtId>"   'Orden de mandato
                
                'Si tiene fecha firma de mandato
                SQL = "2009-10-31"
                If Not IsNull(miRsAux!SEPA_FecFirma) Then SQL = Format(miRsAux!SEPA_FecFirma, "yyyy-mm-dd")
                Print #NFic, "            <DtOfSgntr>" & SQL & "</DtOfSgntr>"
                
                
                
                Print #NFic, "         </MndtRltdInf>"
                Print #NFic, "      </DrctDbtTx>"
                Print #NFic, "      <DbtrAgt>"
                Print #NFic, "         <FinInstnId>"
                SQL = FrmtStr(DBLet(miRsAux!BIC, "T"), 11)
                Print #NFic, "            <BIC>" & SQL & "</BIC>"
                Print #NFic, "         </FinInstnId>"
                Print #NFic, "      </DbtrAgt>"
                Print #NFic, "      <Dbtr>"
                
                Print #NFic, "         <Nm>" & XML(miRsAux!Nommacta) & "</Nm>"
                Print #NFic, "         <PstlAdr>"
                
                SQL = "ES"
                If Not IsNull(miRsAux!PAIS) Then SQL = Mid(miRsAux!PAIS, 1, 2)
                Print #NFic, "            <Ctry>" & SQL & "</Ctry>"
                
                
                If Not IsNull(miRsAux!dirdatos) Then Print #NFic, "              <AdrLine>" & XML(miRsAux!dirdatos) & "</AdrLine>"
                
                SQL = ""
                'SQL = XML(Trim(DBLet(miRsAux!codposta, "T") & " " & DBLet(miRsAux!despobla, "T")))
                'If SQL <> "" Then Print #NFic, "              <AdrLine>" & SQL & "</AdrLine>"If Not IsNull(miRsAux!desprovi) Then Print #NFic, "              <AdrLine>" & XML(miRsAux!desprovi) & "</AdrLine>"
                If DBLet(miRsAux!despobla, "T") = DBLet(miRsAux!desprovi, "N") Then
                    SQL = Trim(DBLet(miRsAux!codposta, "T") & "   " & DBLet(miRsAux!despobla, "T"))
                
                Else
                    SQL = Trim(DBLet(miRsAux!despobla, "T") & "   " & DBLet(miRsAux!codposta, "T"))
                    If Not IsNull(miRsAux!desprovi) Then SQL = SQL & "     " & miRsAux!desprovi
                End If
                If SQL <> "" Then Print #NFic, "              <AdrLine>" & XML(Mid(SQL, 1, 70)) & "</AdrLine>"
                
                
                
                Print #NFic, "         </PstlAdr>"
                Print #NFic, "         <Id>"
                Print #NFic, "            <PrvtId>"
                Print #NFic, "               <Othr>"
                
                
                'Opcion nueva: 3   Quiere el campo referencia de scobro
                Select Case TipoReferenciaCliente
                Case 1
                    'ALZIRA. La referencia final de 12 es el ctan bancaria del cli + su CC
                    SQL = Format(miRsAux!digcontr, "00") ' Dígitos de control
                    SQL = SQL & Format(miRsAux!cuentaba, "0000000000") ' Código de cuenta
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
                
                Print #NFic, "                   <Id>" & SQL & "</Id>"
                If TipoReferenciaCliente = 2 Then Print #NFic, "                   <Issr>NIF</Issr>"
                Print #NFic, "               </Othr>"
                Print #NFic, "            </PrvtId>"
                Print #NFic, "         </Id>"
                Print #NFic, "      </Dbtr>"
                Print #NFic, "      <DbtrAcct>"
                Print #NFic, "         <Id>"
                
                SQL = IBAN_Destino(True)   'Hay que poner TRUE aunque sea cobro
                Print #NFic, "            <IBAN>" & SQL & "</IBAN>"
                Print #NFic, "         </Id>"
                Print #NFic, "      </DbtrAcct>"
                Print #NFic, "      <Purp>"
                Print #NFic, "         <Cd>TRAD</Cd>"
                Print #NFic, "      </Purp>"
                Print #NFic, "      <RmtInf>"
                
                SQL = Trim(DBLet(miRsAux!text33csb, "T") & " " & FrmtStr(DBLet(miRsAux!text41csb, "T"), 60))
                If SQL = "" Then SQL = miRsAux!Nommacta
                Print #NFic, "         <Ustrd>" & XML(SQL) & "</Ustrd>"
                Print #NFic, "      </RmtInf>"
                Print #NFic, "   </DrctDbtTxInf>"
        
            
            
            'Siguiente
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        Set miRsAux = Nothing
        
              
              
        Print #NFic, "</PmtInf>"
        Print #NFic, "</CstmrDrctDbtInitn></Document>"
        
        
        GrabarDisketteNorma19_SEPA_XML = True
            
    End If  'De EOF
    Close #NFic
        
    
    
    
    Exit Function
Err_Remesa19sepa:
    MsgBox "Err: " & Err.Number & vbCrLf & _
        Err.Description, vbCritical, "Grabación del diskette de Remesa SEPA"
        

End Function







Private Function IBAN_Destino(Cobros As Boolean) As String
    If Cobros Then
        IBAN_Destino = FrmtStr(DBLet(miRsAux!IBAN, "T"), 4) ' ES00
        IBAN_Destino = IBAN_Destino & Format(miRsAux!codbanco, "0000") ' Código de entidad receptora
        IBAN_Destino = IBAN_Destino & Format(miRsAux!codsucur, "0000") ' Código de oficina receptora
        IBAN_Destino = IBAN_Destino & Format(miRsAux!digcontr, "00") ' Dígitos de control
        IBAN_Destino = IBAN_Destino & Format(miRsAux!cuentaba, "0000000000") ' Código de cuenta
    Else
        
        'entidad oficina CC cuentaba
        IBAN_Destino = FrmtStr(DBLet(miRsAux!IBAN, "T"), 4) ' ES00
        IBAN_Destino = IBAN_Destino & Format(miRsAux!Entidad, "0000") ' Código de entidad receptora
        IBAN_Destino = IBAN_Destino & Format(miRsAux!oficina, "0000") ' Código de oficina receptora
        IBAN_Destino = IBAN_Destino & Format(miRsAux!CC, "00") ' Dígitos de control
        IBAN_Destino = IBAN_Destino & Format(miRsAux!cuentaba, "0000000000") ' Código de cuenta
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
    Cad = "Select * from ctabancaria where codmacta='" & CuentaPropia2 & "'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        Cad = ""
    Else
        If IsNull(miRsAux!Entidad) Then
            Cad = ""
        Else
            SufijoOEM = "000" ''Sufijo3414
            Cad = miRsAux!IBAN & Format(miRsAux!Entidad, "0000") & Format(DBLet(miRsAux!oficina, "T"), "0000") & DBLet(miRsAux!Control, "T") & Format(DBLet(miRsAux!CtaBanco, "T"), "0000000000")
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

    Cad = DBLet(miRsAux!siglasvia, "T") & " " & miRsAux!direccion & " " & DBLet(miRsAux!numero, "T") & " "
    Cad = Cad & Trim(DBLet(miRsAux!codpobla, "T") & " " & miRsAux!pobempre) & " "
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
    Cad = DevuelveDesdeBD("bic", "sbic", "entidad", Cad)
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
        Cad = "Select scobro.codbanco as entidad,scobro.codsucur as oficina,scobro.cuentaba,scobro.digcontr as CC,scobro.iban"
        Cad = Cad & ",nommacta,dirdatos,codposta,despobla,impvenci,scobro.codmacta,pais,Gastos,impcobro,desprovi"
        Cad = Cad & " ,NUmSerie,codfaccl,fecfaccl,numorden,text33csb,text41csb,bic,nifdatos from scobro"
        Cad = Cad & " LEFT JOIN sbic on scobro.codbanco=sbic.entidad INNER JOIN cuentas ON"
        Cad = Cad & " cuentas.codmacta=scobro.codmacta WHERE transfer =" & NumeroTransferencia
    End If
    miRsAux.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    Regs = 0
    While Not miRsAux.EOF
        Print #NFic, "   <CdtTrfTxInf>"
        Print #NFic, "      <PmtId>"
        
         
        If Pagos Then
            'numfactu fecfactu numorden
            AUX = FrmtStr(miRsAux!numfactu, 10)
            AUX = AUX & Format(miRsAux!fecfactu, "yyyymmdd") & Format(miRsAux!numorden, "000")
        
        Else
            'fecfaccl
            AUX = FrmtStr(miRsAux!NUmSerie, 3) & Format(miRsAux!codfaccl, "00000000")
            AUX = AUX & Format(miRsAux!fecfaccl, "yyyymmdd") & Format(miRsAux!numorden, "000")
        End If
        
        Print #NFic, "         <EndToEndId>" & AUX & "</EndToEndId>"
        Print #NFic, "      </PmtId>"
        Print #NFic, "      <PmtTpInf>"
        If Pagos Then
            Im = DBLet(miRsAux!imppagad, "N")
            Im = miRsAux!ImpEfect - Im
            AUX = miRsAux!ctaprove

        Else
            Im = Abs(miRsAux!impvenci + DBLet(miRsAux!Gastos, "N")) - DBLet(miRsAux!impcobro, "N")
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
        Print #NFic, "          <InstdAmt Ccy=""EUR"">" & TransformaComasPuntos(CStr(Im)) & "</InstdAmt>"
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











'Devolucion SEPA
'
Public Sub ProcesaFicheroDevolucionSEPA_XML(Fichero As String, ByRef Remesa As String)
Dim aux2 As String  'Para buscar los vencimientos
Dim FinLecturaLineas As Boolean

Dim ErroresVto As String

Dim Posicion As Long
Dim L2 As Long
Dim SQL As String
Dim ContenidoFichero As String
Dim NF As Integer
Dim CadenaComprobacionDevueltos As String  'cuantos|importe|


    On Error GoTo eProcesaCabeceraFicheroDevolucionSEPA_XML
    Remesa = ""
    
    
    
   

    NF = FreeFile
    Open Fichero For Input As #NF
    ContenidoFichero = ""
    While Not EOF(NF)
        Line Input #NF, aux2
        ContenidoFichero = ContenidoFichero & aux2
    Wend
    Close #NF
    
    Set miRsAux = New ADODB.Recordset
    
    'Vamos a obtener el ID de la remesa  enviada
    ' Buscaremos la linea
    'Idententificacion propia  Ejemplo: <MsgId>PRE2015093012481641020RE10000802015</MsgId>  de donde RE mesa, 1 tipo 000080 Nº   ano;2015
    Posicion = PosicionEnFichero(1, ContenidoFichero, "<CstmrPmtStsRpt>")
    
    Posicion = PosicionEnFichero(Posicion, ContenidoFichero, "<OrgnlMsgId>")
    L2 = PosicionEnFichero(Posicion, ContenidoFichero, "</OrgnlMsgId>")
    
    aux2 = Mid(ContenidoFichero, Posicion, L2 - Posicion)
    aux2 = Mid(aux2, InStr(10, aux2, "RE") + 3) 'QUTIAMO EL RE y ye tipo RE1(ejemp)
    
    'Los 4 ultimos son año
    Remesa = Mid(aux2, 1, 6) & "|" & Mid(aux2, 7, 4) & "|"
    
    
    'Voy a buscar el numero total de vencimientos que devuelven y el importe total(comproabacion ultima
    Posicion = PosicionEnFichero(Posicion, ContenidoFichero, "<OrgnlPmtInfAndSts>")
    Posicion = PosicionEnFichero(Posicion, ContenidoFichero, "<OrgnlNbOfTxs>")
    '<OrgnlNbOfTxs>1</OrgnlNbOfTxs>
    L2 = PosicionEnFichero(Posicion, ContenidoFichero, "</OrgnlNbOfTxs>")
    CadenaComprobacionDevueltos = Mid(ContenidoFichero, Posicion, L2 - Posicion)
    
    '<OrgnlCtrlSum>5180.98</OrgnlCtrlSum>
    Posicion = PosicionEnFichero(Posicion, ContenidoFichero, "<OrgnlCtrlSum>")
    L2 = PosicionEnFichero(Posicion, ContenidoFichero, "</OrgnlCtrlSum>")
    CadenaComprobacionDevueltos = CadenaComprobacionDevueltos & Mid(ContenidoFichero, Posicion, L2 - Posicion)
            
    
    
    'Primera comprobacion. Existe la remesa obtenida
    
    
    'Vamos con los vtos  4300106840T  0001180220150925001

    Do
        Posicion = InStr(Posicion, ContenidoFichero, "<TxInfAndSts>")
        If Posicion > 0 Then
            
            'Si existe un grupo de registros TxInfAndSts, los de abjo deben existir SI o SI
            Posicion = PosicionEnFichero(Posicion, ContenidoFichero, "<OrgnlEndToEndId>")
            L2 = PosicionEnFichero(Posicion, ContenidoFichero, "</OrgnlEndToEndId>")
            aux2 = Mid(ContenidoFichero, Posicion, L2 - Posicion)
            
            'Id del recibo devuleto. Ejemplo
            '4300106840T  0001180220150925001
            ' asi es como se monta el el generador de remesa
            '           FrmtStr(miRsAux!codmacta, 10) & FrmtStr(miRsAux!NUmSerie, 3) & Format(miRsAux!codfaccl, "00000000")
            '           Format(miRsAux!fecfaccl, "yyyymmdd") & Format(miRsAux!numorden, "000")
            
            SQL = "Select codrem,anyorem,siturem from scobro where fecfaccl='" & Mid(aux2, 22, 4) & "-" & Mid(aux2, 26, 2) & "-" & Mid(aux2, 28, 2)
            SQL = SQL & "' AND numserie = '" & Trim(Mid(aux2, 11, 3)) & "' AND codfaccl = " & Val(Mid(aux2, 14, 8)) & " AND numorden=" & Mid(aux2, 30, 3)

            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = Mid(SQL, InStr(1, UCase(SQL), " WHERE ") + 7)
            SQL = Replace(SQL, "fecfaccl", "F.Fac:")
            SQL = Replace(SQL, "numserie", "Serie:")
            SQL = Replace(SQL, "codfaccl", "NºFac:")
            SQL = Replace(SQL, "numorden", "Ord:")
            SQL = Replace(SQL, "AND", ""): SQL = Replace(SQL, "=", "")
            SQL = "Vto no encontrado: " & Mid(SQL, InStr(1, UCase(SQL), " WHERE ") + 7)
            If Not miRsAux.EOF Then
                If IsNull(miRsAux!CodRem) Then
                    SQL = "Vencimiento sin Remesa: " & aux2
                Else
        
                    SQL = ""
                End If
            End If
            miRsAux.Close
            
            If SQL <> "" Then ErroresVto = ErroresVto & vbCrLf & SQL
            
            
            Posicion = InStr(Posicion, ContenidoFichero, "</TxInfAndSts>") + 11 'Para que pase al siguiente registro, si es que existe
            
        
        Else
           Posicion = Len(ContenidoFichero) + 1
        End If  'posicion>0 de OrgnlTxRef
        
        
    Loop Until Posicion > Len(ContenidoFichero)
    

    If ErroresVto <> "" Then
        MsgBox ErroresVto, vbExclamation
        Remesa = ""
    Else
        


    
        'En aux2 tendre codrem|anñorem|
        aux2 = RecuperaValor(Remesa, 1) & " AND anyo = " & RecuperaValor(Remesa, 2)
        aux2 = "Select situacion from remesas where codigo = " & aux2
        miRsAux.Open aux2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If miRsAux.EOF Then
            aux2 = "-No se encuentra remesa"
            
        Else
            'Si que esta.
            'Situacion
            If CStr(miRsAux!situacion) <> "Q" And CStr(miRsAux!situacion) <> "Y" Then
                aux2 = "- Situacion incorrecta : " & miRsAux!situacion
            Else
                aux2 = "" 'TODO OK
            End If
        End If

        If aux2 <> "" Then
            aux2 = aux2 & " ->" & Mid(miRsAux.Source, InStr(1, UCase(miRsAux.Source), " WHERE ") + 7)
            aux2 = Replace(aux2, " AND ", " ")
            aux2 = Replace(aux2, "anyo", "año")
            ErroresVto = ErroresVto & vbCrLf & aux2
        End If
        miRsAux.Close

    
    


        If ErroresVto <> "" Then
            aux2 = "Error remesas " & vbCrLf & String(30, "=") & ErroresVto
            MsgBox aux2, vbExclamation

            'Pongo REMESA=""
            Remesa = "" 'para que no continue el preoceso de devolucion
        End If

    End If
    Set miRsAux = Nothing
    Exit Sub
eProcesaCabeceraFicheroDevolucionSEPA_XML:
    Remesa = ""
    MuestraError Err.Number, "Procesando fichero devolucion SEPA XML" & Err.Description
    Set miRsAux = Nothing
End Sub

'Si no se encuentra lo que busco saltara un error
Private Function PosicionEnFichero(ByVal Inicio As Long, ContenidoDelFichero As String, QueBusco As String) As Long
    PosicionEnFichero = InStr(Inicio, ContenidoDelFichero, QueBusco)
    If PosicionEnFichero = 0 Then
        Err.Raise 513, "No se encuentra cadena: " & QueBusco
    Else
        If InStr(1, QueBusco, "</") Then
            'PosicionEnFichero = PosicionEnFichero - Len(QueBusco)
        Else
            PosicionEnFichero = PosicionEnFichero + Len(QueBusco)
        End If
    End If
        
End Function


'XML
Public Sub ProcesaLineasFicheroDevolucionXML(Fichero As String, ByRef Listado As Collection)
Dim NF As Integer
Dim ContenidoFichero As String
Dim Posicion As Long
Dim L2 As Long
Dim aux2 As String

    NF = FreeFile
    Open Fichero For Input As #NF
    ContenidoFichero = ""
    While Not EOF(NF)
        Line Input #NF, aux2
        ContenidoFichero = ContenidoFichero & aux2
    Wend
    Close #NF
    
   
    Posicion = 1
    Do
        Posicion = InStr(Posicion, ContenidoFichero, "<TxInfAndSts>")
        If Posicion > 0 Then
            
            'Si existe un grupo de registros TxInfAndSts, los de abjo deben existir SI o SI
            Posicion = PosicionEnFichero(Posicion, ContenidoFichero, "<OrgnlEndToEndId>")
            L2 = PosicionEnFichero(Posicion, ContenidoFichero, "</OrgnlEndToEndId>")
            aux2 = Mid(ContenidoFichero, Posicion, L2 - Posicion)
            
            'Id del recibo devuleto. Ejemplo
            '4300106840T  0001180220150925001
            ' asi es como se monta el el generador de remesa
            '           FrmtStr(miRsAux!codmacta, 10) & FrmtStr(miRsAux!NUmSerie, 3) & Format(miRsAux!codfaccl, "00000000")
            '           Format(miRsAux!fecfaccl, "yyyymmdd") & Format(miRsAux!numorden, "000")
            
            'Vamos a guardar en el col la linea en formato antiguo SEPA y asi no toco el programa
            'M  0330047820131201001   430000061
            aux2 = Mid(aux2, 11, 23) & "   " & Mid(aux2, 1, 10)
            Listado.Add aux2
            Posicion = InStr(Posicion, ContenidoFichero, "</TxInfAndSts>") + 11 'Para que pase al siguiente registro, si es que existe
            
        
        Else
           Posicion = Len(ContenidoFichero) + 1
        End If  'posicion>0 de OrgnlTxRef
        
    Loop Until Posicion > Len(ContenidoFichero)
    
End Sub
