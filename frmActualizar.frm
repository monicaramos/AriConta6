VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmActualizar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizar diario"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmActualizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frame1Asiento 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   4815
      Begin ComCtl2.Animation Animation1 
         Height          =   735
         Left            =   600
         TabIndex        =   4
         Top             =   1800
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1296
         _Version        =   327681
         FullWidth       =   241
         FullHeight      =   49
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label9 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   210
         TabIndex        =   5
         Top             =   120
         Width           =   4335
      End
      Begin VB.Label lblAsiento 
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "N� Asiento :"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmActualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public OpcionActualizar As Byte
    '1.- Actualizar 1 asiento
    '2.- Desactualiza pero NO insertes en apuntes
    '3.- Desactualizar asiento desde hco
    
    'Si el asiento es de una factura entonces NUMSERIE tendra "FRACLI" o "FRAPRO"
    ' con lo cual habra que poner su factura asociada a NULL
    
    '4.- Si es para enviar datos a impresora
    '5.- Actualiza mas de 1 asiento
    
    '6.- Integra 1 factura
    '7.- Elimina factura integrada . DesINTEGRA   . C L I E N T E S
    '8.- Integra 1 factura PROVEEDORES
    '9.- Elimina factura integrada . Desintegra. P R O V E E D O R E S
    
    '10 .- Integracion masiva facturas clientes
    '11 .- Integracion masiva facturas Proveedores
    
    
        
Public NumAsiento As Long
Public FechaAsiento As Date
Public NumDiari As Integer
Public NUmSerie As String
Public NumFac As Long
Public FechaAnterior As Date
Public Proveedor As String
Public FACTURA As String
Public FechaFactura As Date

Public DentroBeginTrans As Boolean

'Nuevo. 17 Cotubre de 2005
'-------------------------
'  Los clientes que facturan con mas de un diario, las facturas SIEMPRE
'  van al diaro de parametros, con lo cual ES una cagada
Public DiarioFacturas As Integer


Public SqlLog As String

Private WithEvents frmD As frmTiposDiario
Attribute frmD.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private Cuenta As String
Private ImporteD As Currency
Private ImporteH As Currency
Private CCost As String
'Y estas son privadas
Private Mes As Integer
Private Anyo As Integer
Dim Fecha As String  'TENDRA la fecha ya formateada en yyy-mm-dd
Dim PrimeraVez As Boolean
Dim Sql As String
Dim Rs As Recordset

Dim INC As Long

Dim NE As Integer
Dim ErroresAbiertos As Boolean
Dim NumErrores As Long

Dim ItmX As ListItem  'Para mostra errores masivos

Private Sub A�adeError(ByRef Mensaje As String)
On Error Resume Next
'Escribimos en el fichero
If Not ErroresAbiertos Then
    NE = FreeFile
    ErroresAbiertos = True
    Open App.Path & "\ErrActua.txt" For Output As NE
    If Err.Number <> 0 Then
        MsgBox " Error abriendo fichero errores", vbExclamation
        Err.Clear
    End If
End If
Print #NE, Mensaje
If Err.Number <> 0 Then
    Err.Clear
    NumErrores = -20000
Else
    NumErrores = NumErrores + 1
End If
End Sub



Private Function CadenaImporte(VaAlDebe As Boolean, ByRef Importe As Currency, ElImporteEsCero As Boolean) As String
Dim CadImporte As String

'Si va al debe, pero el importe es negativo entonces va al haber a no ser que la contabilidad admita importes negativos
    If Importe < 0 Then
        If Not vParam.abononeg Then
            VaAlDebe = Not VaAlDebe
            Importe = Abs(Importe)
        End If
    End If
    ElImporteEsCero = (Importe = 0)
    CadImporte = TransformaComasPuntos(CStr(Importe))
    If VaAlDebe Then
        CadenaImporte = CadImporte & ",NULL"
    Else
        CadenaImporte = "NULL," & CadImporte
    End If
End Function

Private Sub CargaProgres(Valor As Integer)
Me.ProgressBar1.Max = Valor
Me.ProgressBar1.Value = 0
End Sub



Private Sub IncrementaProgres(Veces As Integer)
On Error Resume Next
Me.ProgressBar1.Value = Me.ProgressBar1.Value + (Veces * INC)
If Err.Number <> 0 Then
    Err.Clear
    ProgressBar1.Value = 0
End If

End Sub





'Eliminar factura con asiento
Private Function EliminaFacturaConAsiento()
Dim Donde As String
Dim bol As Boolean
Dim LEtra As String
Dim Mc As Contadores
Dim Contabilizada As String

    On Error GoTo EEliminaFacturaConAsiento
    'Sabemos que
    'numasiento     --> N� aseinto
    'numfac         --> CODIGO FACTURA
    'NumDiari       --> ATENCION -> N� de diario, no como al integrar
    'FechaAsiento   --> Fecha asiento
    'NUmSerie       --> SERIE DE LA FACTURA  y el a�o (sep. con pipes)

    'Obtenemos el mes y el a�o
    Mes = Month(FechaAsiento)
    Anyo = Year(FechaAsiento)
    Fecha = Format(FechaAsiento, FormatoFecha)
    
    'Aqui bloquearemos
    Conn.BeginTrans
    
    'Eliminamos factura
    LEtra = RecuperaValor(NUmSerie, 1)
    If Me.OpcionActualizar = 7 Then
        '-------------------------------------------------------------
        '               C L I E N T E S
        '-------------------------------------------------------------
        Sql = " WHERE numserie = '" & LEtra & "'"
        Sql = Sql & " AND numfactu = " & NumFac
        Sql = Sql & " AND anofactu= " & RecuperaValor(NUmSerie, 2)
        'Las lineas
        Donde = "Linea factura"
        Cuenta = "DELETE from factcli_lineas " & Sql
        Conn.Execute Cuenta
        
        'totales de factura
        Donde = "Totales factura"
        Cuenta = "DELETE from factcli_totales " & Sql
        Conn.Execute Cuenta
        
        
        Contabilizada = "select count(*) from cobros where numserie = " & DBSet(LEtra, "T") & " and numfactu = " & NumFac & " and fecfactu = " & DBSet(FechaAsiento, "F") & " and impcobro <> 0 and not impcobro is null "
        
        If TotalRegistros(Contabilizada) <> 0 Then
            MsgBox "Hay cobros que ya se han efectuado. Revise cartera y contabilidad.", vbExclamation
        Else
            ' cobro de la factura
            Donde = "Cobro factura"
            
'            Cuenta = "DELETE from cobros_realizados where numserie = " & DBSet(LEtra, "T") & " and numfactu = " & NumFac & " and fecfactu = " & DBSet(FechaAsiento, "F")
'            Conn.Execute Cuenta
            
            
            Cuenta = "DELETE from cobros where numserie = " & DBSet(LEtra, "T") & " and numfactu = " & NumFac & " and fecfactu = " & DBSet(FechaAsiento, "F")
            Conn.Execute Cuenta
        End If
        
        'La factura
        Donde = "Cabecera factura"
        Cuenta = "DELETE from factcli " & Sql
        Conn.Execute Cuenta

    Else
        '-------------------------------------------------------------
        '       P R O V E E D O R E S
        '-------------------------------------------------------------
        Sql = " WHERE numserie = '" & LEtra & "'"
        Sql = Sql & " AND numregis = " & NumFac
        Sql = Sql & " AND anofactu= " & RecuperaValor(NUmSerie, 2)
        'Las lineas
        Donde = "Linea factura"
        Cuenta = "DELETE from factpro_lineas " & Sql
        Conn.Execute Cuenta
        
        'totales de factura
        Donde = "Totales factura"
        Cuenta = "DELETE from factpro_totales " & Sql
        Conn.Execute Cuenta
        
        Contabilizada = "select count(*) from pagos where numserie = " & DBSet(LEtra, "T") & " and codmacta = " & DBSet(Proveedor, "T") & " and numfactu = " & DBSet(FACTURA, "T") & " and fecfactu = " & DBSet(FechaFactura, "F") & " and imppagad <> 0 and not imppagad is null "
        
        If TotalRegistros(Contabilizada) <> 0 Then
            MsgBox "Hay pagos que ya se han efectuado. Revise cartera y contabilidad.", vbExclamation
        Else
            ' cobro de la factura
            Donde = "Pago factura"
            
            Cuenta = "DELETE from pagos where numserie = " & DBSet(LEtra, "T") & " and codmacta = " & DBSet(Proveedor, "T") & " and numfactu = " & DBSet(FACTURA, "T") & " and fecfactu = " & DBSet(FechaFactura, "F")
            Conn.Execute Cuenta
        End If
        
        'La factura
        Donde = "Cabecera factura"
        Cuenta = "DELETE from factpro " & Sql
        Conn.Execute Cuenta
        LEtra = RecuperaValor(NUmSerie, 1) '"1"
    End If

    bol = DesActualizaElASiento(Donde)

EEliminaFacturaConAsiento:
        If Err.Number <> 0 Then
            Sql = "Actualiza Asiento." & vbCrLf & "----------------------------" & vbCrLf
            Sql = Sql & Donde
            MuestraError Err.Number, Sql, Err.Description
            bol = False
        End If
        If bol Then
            Conn.CommitTrans
            
            'Intentamos devolver el contador
            If FechaAsiento >= vParam.fechaini Then
                Set Mc = New Contadores
                Mc.DevolverContador LEtra, (FechaAsiento <= vParam.fechafin), NumFac
                Set Mc = Nothing
            End If
            
            
            'INSERTO EN LOG
            Mes = 6
            
            
            If Me.OpcionActualizar <> 7 Then
                Mes = 10   'FRAPRO
                LEtra = ""
                
                vLog.Insertar 10, vUsu, SqlLog
            Else
                vLog.Insertar 6, vUsu, SqlLog
            End If
            
            
            EliminaFacturaConAsiento = True
            AlgunAsientoActualizado = True
        Else
            Conn.RollbackTrans
        End If
    
End Function



    






Private Sub Form_Activate()
Dim bol As Boolean
If PrimeraVez Then
    PrimeraVez = False
    Me.Refresh
    bol = False
    Select Case OpcionActualizar
    Case 1
        ActualizaAsiento
        bol = True
    Case 2, 3
        DesActualizaAsiento
        bol = True
    Case 6, 8
        'Integramos la factura (Dependera del opcion si es de clientes o de proveedores
        IntegraFactura
        bol = True
    Case 7, 9
         'Integramos la factura (Dependera del opcion si es de clientes o de proveedores
        EliminaFacturaConAsiento
        bol = True
        
        
    Case 16
        'Insertar Asiento en el hco
        
    End Select
    If bol Then Unload Me
End If
Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim B As Boolean

    Me.Icon = frmPpal.Icon


    ErroresAbiertos = False
    Limpiar Me
    PrimeraVez = True
    
    Select Case OpcionActualizar
    Case 1, 2, 3
        Label1.Caption = "N� Asiento"
        Me.lblAsiento.Caption = NumAsiento
        INC = 10  'Incremento para el proggress
        If OpcionActualizar = 1 Then
            Label9.Caption = "Actualizar"
        Else
            Label9.Caption = "Modi/Eliminar"
        End If
        'Tama�o
        Me.Height = 3000
        B = True
            
    Case 6, 7, 8, 9
        '// Estamos en Facturas
        Label1.Caption = "N� factura"
        If OpcionActualizar < 8 Then
            Label1.Caption = Label1.Caption & " Cliente"
        Else
            Label1.Caption = Label1.Caption & " Proveedor"
        End If
        Me.lblAsiento.Caption = NUmSerie & NumAsiento
        INC = 10  'Incremento para el proggress
        If OpcionActualizar = 6 Or OpcionActualizar = 8 Then
            Label9.Caption = "Integrar Factura"
        Else
            Label9.Caption = "Eliminar Factura"
        End If
        Me.Caption = "Actualizar facturas"
        'Tama�o
        Me.Height = 3315
        B = True
        
    End Select
    Me.frame1Asiento.Visible = B
    Me.Animation1.Visible = B
End Sub



Private Function IntegraFactura() As Boolean
Dim B As Boolean
Dim Donde As String
Dim vConta As Contadores

Dim TipoConce As String
On Error GoTo EIntegraFactura
    
    IntegraFactura = False
    
    If Not DentroBeginTrans Then Conn.BeginTrans
    Fecha = Format(FechaAsiento, FormatoFecha)
    
    
    'Vemos si estamos intentato forzar numero de asiento
    If NumAsiento > 0 Then
        'Primero que nada obtendremos el contador
        If AsientoExiste Then
            MsgBox "Ya existe el asiento con la numeraci�n: " & NumAsiento & " " & FechaAsiento & " " & NumDiari, vbExclamation
            'Vamoa al final del proceso de esta factura
            GoTo EIntegraFactura
        End If
    Else
        Donde = "Conseguir contador"
        Set vConta = New Contadores
        If vConta.ConseguirContador("0", (FechaAsiento <= vParam.fechafin), True) = 1 Then
            MsgBox "Error consiguiendo contador asiento", vbExclamation
            'Vamoa al final del proceso de esta factura
            GoTo EIntegraFactura
        End If
        
        If Not vConta.YaExisteContador((FechaAsiento <= vParam.fechafin), vParam.fechafin, (OpcionActualizar < 10)) Then
            If OpcionActualizar > 9 Then InsertaError "Error contadores asiento: " & vConta.Contador
            GoTo EIntegraFactura
        End If
        NumAsiento = vConta.Contador
        Set vConta = Nothing
    End If
    
    'Actualizamos los datos
    If OpcionActualizar = 6 Or OpcionActualizar = 10 Then
        B = IntegraLaFactura(Donde)
    Else
        B = IntegraLaFacturaProv(Donde)
    End If
    
EIntegraFactura:
    If Err.Number <> 0 Then
        If OpcionActualizar > 9 Then
            'Esta actualizando varias a la vez
            InsertaError Donde & " - " & Err.Description
        Else
            MuestraError Err.Number, "Integra factura(I)" & vbCrLf & Donde
        End If
        Err.Clear
        B = False
    End If
    If B Then
        
        If OpcionActualizar > 9 Then
            'Actualizando desde/hasta y ha ido bien. La meto al LOG
            vLog.AnyadeTextoDatosDes NUmSerie & Format(NumFac, "000000")
            'If OpcionActualizar = 10 Then
            '    'FRACLI
        End If
    End If
    IntegraFactura = B
    AlgunAsientoActualizado = B
    
    If Not DentroBeginTrans Then
        If B Then
            Conn.CommitTrans
        Else
            Conn.RollbackTrans
        End If
    End If
End Function

Private Function IntegraLaFactura(ByRef A_Donde As String) As Boolean
Dim cad As String
Dim Cad2 As String
Dim Cad3 As String
Dim Amplia2 As String
Dim DocConcAmp As String
Dim RF As Recordset
Dim ImporteNegativo As Boolean
Dim Importe0 As Boolean
Dim PrimeraContrapartida As String
    
    Dim SqlIva As String
    Dim RsIvas As ADODB.Recordset

    IntegraLaFactura = False
    'Sabemos que
    'numfac     --> CODIGO FACTURA
    'NumDiari       --> A�O FACTURA
    'NUmSerie       --> SERIE DE LA FACTURA
    'FechaAsiento   --> Fecha factura
    'FecFactuAnt    --> FecFactura Anterior
    
    'Obtenemos los datos de la factura
    A_Donde = "Leyendo datos factura"
    Set RF = New ADODB.Recordset
    Sql = "SELECT * FROM factcli"
    Sql = Sql & " WHERE numserie='" & NUmSerie
    Sql = Sql & "' AND numfactu= " & NumFac
    Sql = Sql & " AND anofactu=" & NumDiari
    RF.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RF.EOF Then
        MsgBox "No se encuentra la factura: " & vbCrLf & Sql, vbExclamation
        RF.Close
        Exit Function
    End If
    
 
    Sql = "select count(*) from hcabapu where numdiari = " & DBSet(DiarioFacturas, "N") & " and fechaent = " & DBSet(FechaAnterior, "F") & " and numasien = " & DBSet(NumAsiento, "N")
    If TotalRegistros(Sql) > 0 Then
        A_Donde = "Actualiza cabecera hco apuntes"
        
        Sql = "UPDATE hcabapu SET "
        Sql = Sql & " fechaent = " & DBSet(Fecha, "F")
        Sql = Sql & " where numdiari = " & DBSet(DiarioFacturas, "N")
        Sql = Sql & " and fechaent = " & DBSet(FechaAnterior, "F")
        Sql = Sql & " and numasien = " & DBSet(NumAsiento, "N")
    
        Conn.Execute Sql
    Else
        'Cabecera del hco de apuntes
        A_Donde = "Inserta cabecera hco apuntes"
        Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES ("
        Sql = Sql & DiarioFacturas & ",'" & Fecha & "'," & NumAsiento
        Sql = Sql & ","
        'Marzo 2010
        'Si tiene observaciones las llevo al apunte
        cad = DBLet(RF!observa, "T")
        If cad = "" Then
            cad = "NULL,"
        Else
            cad = "'" & DevNombreSQL(cad) & "',"
        End If
        cad = cad & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilizaci�n Factura de Cliente " & NUmSerie & Format(NumFac, "0000000") & " " & Fecha & "')"
        
        
        Sql = Sql & cad
        Conn.Execute Sql
    End If
    
    'Lineas fijas, es decir la linea de cliente, importes y tal y tal
    'Para el sql
    cad = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, "
    cad = cad & "codconce,ampconce, timporteD, timporteH,codccost, ctacontr, idcontab, punteada)"
    cad = cad & " VALUES (" & DiarioFacturas & ",'" & Fecha & "'," & NumAsiento & ","
    Mes = 1 'Contador de lineas
    
    
    A_Donde = "Linea cliente"
    '-------------------------------------------------------------------
    'LINEA Cliente
    Sql = Mes & ",'" & RF!codmacta & "',"
    
    DocConcAmp = "'" & NUmSerie & Format(NumFac, "0000000") & "'," & vParam.concefcl & ",'"
    
    
    'Ampliacion segun parametros
    Select Case vParam.nctafact
    Case 1
        If RF!totfaccl < 0 Then
            Cad2 = RecuperaValor(vParam.AmpliacionFacurasCli, 2)
        Else
            Cad2 = RecuperaValor(vParam.AmpliacionFacurasCli, 1)
        End If
        Cad2 = Cad2 & " " & NUmSerie & Format(NumFac, "0000000")
    Case 2
        Cad2 = DevNombreSQL(DBLet(RF!Nommacta))
    Case Else
        Cad2 = DBLet(RF!confaccl)
    End Select
    
    '   Modificacion para k aparezca en la ampliacio el CC en la ampliacion de codmacta
    '
    Amplia2 = Cad2
    If vParam.CCenFacturas Then
        A_Donde = "CC en Facturas."
        Cad3 = DevuelveCentroCosteFactura(True, PrimeraContrapartida)
        If Cad3 <> "" Then
            If Len(Amplia2) > 21 Then Amplia2 = Mid(Amplia2, 1, 21)
            Amplia2 = Amplia2 & " [" & Cad3 & "]"
        End If
    End If
    A_Donde = "Linea cliente"
    
    
    Sql = Sql & DocConcAmp & Amplia2 & "'"
    DocConcAmp = DocConcAmp & Cad2 & "'"   'DocConcAmp Sirve para el IVA
    
    'Esta variable sirve para las demas
    ImporteNegativo = (DBLet(RF!totfaccl, "N") < 0)
    
    'Importes, atencion importes negativos
    '  antes --> Cad2 = CadenaImporte(ImporteNegativo, True, RF!totfaccl)
    Cad2 = CadenaImporte(True, DBLet(RF!totfaccl, "N"), Importe0)
    Sql = Sql & "," & Cad2 & ",NULL,"
    
    'Contrpartida. 28 Marzo 2006
    If PrimeraContrapartida <> "" Then
        Sql = Sql & "'" & PrimeraContrapartida & "'"
    Else
        Sql = Sql & "NULL"
    End If
    Sql = Sql & ",'FRACLI',0)"
    
    
    Conn.Execute cad & Sql
    Mes = Mes + 1 'Es el contador de lineaapunteshco
    
    ' cuentas de iva ahora se sacan de las tablas de totales
    SqlIva = "select * from factcli_totales "
    SqlIva = SqlIva & " WHERE numserie='" & NUmSerie
    SqlIva = SqlIva & "' AND numfactu= " & NumFac
    SqlIva = SqlIva & " AND anofactu=" & NumDiari
    SqlIva = SqlIva & " order by numlinea "
    
    Set RsIvas = New ADODB.Recordset
    RsIvas.Open SqlIva, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RsIvas.EOF
        Cad3 = "cuentarr"
        Cad2 = DevuelveDesdeBD("cuentare", "tiposiva", "codigiva", RsIvas!codigiva, "N", Cad3)
        If Cad2 <> "" Then
        
            Sql = Mes & ",'" & Cad2 & "'," & DocConcAmp
            Cad2 = CadenaImporte(False, RsIvas!Impoiva, Importe0)
            Sql = Sql & "," & Cad2 & ","
            Sql = Sql & "NULL,'" & RF!codmacta & "','FRACLI',0)"
            'dependiendo de si ContabilizarAptIva0 = 1 se contabiliza o no el iva
            If Importe0 Then
                If vParam.ContabApteIva0 Then
                    Conn.Execute cad & Sql
                    Mes = Mes + 1
                End If
            Else
                Conn.Execute cad & Sql
                Mes = Mes + 1
            End If
            
            'La de recargo  1-----------------
            If Not IsNull(RsIvas!ImpoRec) Then
                     Sql = Mes & "," & Cad3 & "," & DocConcAmp
                    'Importes, atencion importes negativos
                    Cad2 = CadenaImporte(False, RsIvas!ImpoRec, Importe0)
                    Sql = Sql & "," & Cad2 & ","
                    Sql = Sql & "NULL,'" & RF!codmacta & "','FRACLI',0)"
                    If Not Importe0 Then
                        Conn.Execute cad & Sql
                        Mes = Mes + 1
                    End If
            End If
        Else
            MsgBox "Error leyendo TIPO de IVA: " & RsIvas!codigiva, vbExclamation
            RF.Close
            Exit Function
        End If
    
        RsIvas.MoveNext
    Wend
    Set RsIvas = Nothing
    
    '-------------------------------------
    ' RETENCION
    A_Donde = "Retencion"
    If Not IsNull(RF!cuereten) Then
        Sql = Mes & ",'" & RF!cuereten & "'," & DocConcAmp
        'Importes, atencion importes negativos
        Cad2 = CadenaImporte(True, RF!trefaccl, Importe0)
        Sql = Sql & "," & Cad2 & ","
        Sql = Sql & "NULL,NULL,'FRACLI',0)"
       
        Conn.Execute cad & Sql
        Mes = Mes + 1 'Es el contador de lineaapunteshco
    End If
    
    
    IncrementaProgres 2
    
    '------------------------------------------------------------
    'Las lineas de la factura. Para ello guardaremos algunos datos
    Cad2 = RF!codmacta
    ImporteD = DBLet(RF!totfaccl, "N")
    
    
    'Cerramos el RF
    Cuenta = RF!codmacta
    RF.Close
    
    
    
    A_Donde = "Leyendo lineas factura"
'    SQL = "Select factcli_lineas.* , cuentas.codmacta FROM factcli_lineas,Cuentas "
    Sql = "Select cuentas.codmacta, factcli_lineas.codccost, sum(factcli_lineas.baseimpo) baseimpo FROM factcli_lineas,Cuentas "
    Sql = Sql & " WHERE numserie='" & NUmSerie
    Sql = Sql & "' AND numfactu= " & NumFac
    Sql = Sql & " AND anofactu=" & NumDiari
    Sql = Sql & " AND factcli_lineas.codmacta = Cuentas.codmacta"
    Sql = Sql & " group by 1,2 "
    Sql = Sql & " order by 1,2 "
    RF.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Para cada linea insertamos
    Cad2 = ""
    A_Donde = "Procesando lineas"
    While Not RF.EOF
        'Importes, atencion importes negativos
        If Cad2 = "" Then PrimeraContrapartida = RF!codmacta
        Sql = Mes & ",'" & RF!codmacta & "'," & DocConcAmp
        Cad2 = CadenaImporte(False, RF!Baseimpo, Importe0)
        Sql = Sql & "," & Cad2 & ","
        If IsNull(RF!codccost) Then
            Cad2 = "NULL"
        Else
            Cad2 = "'" & RF!codccost & "'"
        End If
        
        Sql = Sql & Cad2 & ",'" & Cuenta & "','FRACLI',0)"
    
        Conn.Execute cad & Sql
        Mes = Mes + 1 'Es el contador de lineaapunteshco
        
        'Siguiente
        IncrementaProgres 1
        RF.MoveNext
        If Not RF.EOF Then PrimeraContrapartida = ""
    Wend
    RF.Close
    
    
    
    
    'AHora viene lo bueno.  MARZO 2006
    'Si el valor fuera true YA lo habria insertado en la cabcera
    If Not vParam.CCenFacturas Then
        If PrimeraContrapartida <> "" Then
            Sql = "UPDATE factcli_lineas SET codmacta ='" & PrimeraContrapartida & "'"
            Sql = Sql & " WHERE numdiari = " & DiarioFacturas & " AND fechaent ='" & Fecha & "' and numasien = " & NumAsiento
            Sql = Sql & " AND numlinea =1 " 'LA PRIMERA LINEA SIEMPRE ES LA DE LA CUENTA
            EjecutaSQL Sql  'Lo hacemos aqui para controlar el error y que no explote
        End If
    End If
        
    
    
    
    'Actualimos en factura, el n� de asiento
    Sql = "UPDATE factcli SET numdiari = " & DiarioFacturas & ", fechaent = '" & Fecha & "', numasien =" & NumAsiento
    Sql = Sql & " WHERE numserie='" & NUmSerie
    Sql = Sql & "' AND numfactu= " & NumFac
    Sql = Sql & " AND anofactu= " & NumDiari
    Conn.Execute Sql
    
    'Para los saldos ponemos el numero de asiento donde toca
    '
    A_Donde = "Saldos factura"
    NumDiari = vParam.numdiacl
    NumDiari = DiarioFacturas
    Mes = Month(FechaAsiento)
    Anyo = Year(FechaAsiento)
    'Actualizaremos los saldos
'    If Not CalcularLineasYSaldosFacturas Then Exit Function
    
    IntegraLaFactura = True
End Function



'////////////////////////////////////////////////////////////////////
'
'           Facturas proveedores
Private Function IntegraLaFacturaProv(ByRef A_Donde As String) As Boolean
Dim cad As String
Dim Cad2 As String
Dim Cad3 As String
Dim DocConcAmp As String
Dim Amplia2 As String
Dim RF As Recordset
Dim ImporteNegativo As Boolean
Dim Importe0 As Boolean 'Para saber si el importe es 0
Dim PrimeraContrapartida As String  'Si hay solo una linea entonces la pondremos como contrapartida de la primera base


'Modificacion de 31 Enero 2005
'-------------------------------------
'-------------------------------------
Dim ColumnaIVA As String
Dim TipoDIva As Byte
    
    Dim SqlIva As String
    Dim RsIvas As ADODB.Recordset

    IntegraLaFacturaProv = False
    
    'Sabemos que
    'numfac     --> CODIGO FACTURA
    'NumDiari       --> A�O FACTURA
    'FechaAsiento   --> Fecha factura
    
    
    'Obtenemos los datos de la factura
    A_Donde = "Leyendo datos factura"
    Set RF = New ADODB.Recordset
    Sql = "SELECT * FROM factpro"
    Sql = Sql & " WHERE numregis = " & NumFac
    Sql = Sql & " AND anofactu=" & NumDiari
    RF.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RF.EOF Then
        MsgBox "No se encuentra la factura: " & vbCrLf & Sql, vbExclamation
        RF.Close
        Exit Function
    End If
    
    Sql = "select count(*) from hcabapu where numdiari = " & DBSet(DiarioFacturas, "N") & " and fechaent = " & DBSet(FechaAnterior, "F") & " and numasien = " & DBSet(NumAsiento, "N")
    If TotalRegistros(Sql) > 0 Then
        A_Donde = "Actualiza cabecera hco apuntes"
        
        Sql = "UPDATE hcabapu SET "
        Sql = Sql & " fechaent = " & DBSet(Fecha, "F")
        Sql = Sql & " where numdiari = " & DBSet(DiarioFacturas, "N")
        Sql = Sql & " and fechaent = " & DBSet(FechaAnterior, "F")
        Sql = Sql & " and numasien = " & DBSet(NumAsiento, "N")
    
        Conn.Execute Sql
    Else
        'Cabecera del hco de apuntes
        A_Donde = "Inserta cabecera hco apuntes"
        Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari,feccreacion,usucreacion,desdeaplicacion) VALUES ("
        Sql = Sql & DiarioFacturas & ",'" & Fecha & "'," & NumAsiento
        
        'Marzo 2010
        'Si tiene observaciones las llevo al apunte
        cad = DBLet(RF!observa, "T")
        If cad = "" Then
            cad = "NULL,"
        Else
            cad = "'" & DevNombreSQL(cad) & "',"
        End If
        
        cad = cad & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilizaci�n Factura Proveedor Registro " & Format(NumFac, "0000000") & " " & Fecha & "')"
        
        Sql = Sql & "," & cad
        
        Conn.Execute Sql
        
    End If
    
    
    
    'Lineas fijas, es decir la linea de cliente, importes y tal y tal
    'Para el sql
    cad = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, "
    cad = cad & "codconce,ampconce, timporteD, timporteH,codccost, ctacontr, idcontab, punteada)"
    cad = cad & " VALUES (" & DiarioFacturas & ",'" & Fecha & "'," & NumAsiento & ","
    Mes = 1 'Contador de lineas
    PrimeraContrapartida = ""
    
    'Esta variable sirve para las demas
    ImporteNegativo = (RF!totfacpr < 0)
    A_Donde = "Linea proveedor"
    '-------------------------------------------------------------------
    'LINEA Proveedor
    Sql = Mes & ",'" & RF!codmacta & "',"
    
    'Documento "numdocum"
    If vParam.CodiNume = 1 Then
        Cad2 = Format(NumFac, "0000000000")
    Else
        Cad2 = DBLet(RF!NumFactu)
    End If
    

    DocConcAmp = "'" & Cad2 & "'," & vParam.concefpr & ",'"
    
    
    'Ampliacion segun parametros
    Select Case vParam.nctafact
    Case 1
        If RF!totfacpr < 0 Then
            Cad2 = RecuperaValor(vParam.AmpliacionFacurasPro, 2)
        Else
            Cad2 = RecuperaValor(vParam.AmpliacionFacurasPro, 1)
        End If
        Cad2 = Cad2 & " " & DevNombreSQL(RF!NumFactu)
        
        Cad2 = Cad2 & " (" & Format(RF!FecFactu, "ddmmyy") & ")"
    Case 2
        Cad2 = DevNombreSQL(DBLet(RF!Nommacta))
    Case Else
        Cad2 = DBLet(RF!confacpr)
    End Select
    
        
    
    'Modificacion para k aparezca en la ampliacio el CC en la ampliacion de codmacta
    '
    Amplia2 = Cad2
    If vParam.CCenFacturas Then
        A_Donde = "CC en Facturas."
        Cad3 = DevuelveCentroCosteFactura(False, PrimeraContrapartida)
        If Cad3 <> "" Then
            If Len(Amplia2) > 26 Then Amplia2 = Mid(Amplia2, 1, 26)
            Amplia2 = Amplia2 & "[" & Cad3 & "]"
        End If
    End If
    A_Donde = "Linea cliente"
    
    
    Sql = Sql & DocConcAmp & Amplia2 & "'"
    DocConcAmp = DocConcAmp & Cad2 & "'"   'DocConcAmp Sirve para el IVA
    
    
    'Importes, atencion importes negativos
    Cad2 = CadenaImporte(False, RF!totfacpr, Importe0)
    Sql = Sql & "," & Cad2 & ",NULL,"
    
    'Contrpartida. 28 Marzo 2006
    If PrimeraContrapartida <> "" Then
        Sql = Sql & "'" & PrimeraContrapartida & "'"
    Else
        Sql = Sql & "NULL"
    End If
    Sql = Sql & ",'FRAPRO',0)"
    
    Conn.Execute cad & Sql
    Mes = Mes + 1 'Es el contador de lineaapunteshco
    
    ' cuentas de iva ahora se sacan de las tablas de totales
    SqlIva = "select * from factpro_totales "
    SqlIva = SqlIva & " WHERE numserie='" & NUmSerie
    SqlIva = SqlIva & "' AND numregis= " & NumFac
    SqlIva = SqlIva & " AND anofactu=" & NumDiari
    SqlIva = SqlIva & " order by numlinea "
    
    
    Dim EsSujetoPasivo As Boolean
    Dim EsImportacion As Boolean
    
    EsImportacion = (DBLet(RF!codopera, "N") = 2)
    EsSujetoPasivo = ((DBLet(RF!codopera, "N") = 1) Or (DBLet(RF!codopera, "N") = 4))
    
    Set RsIvas = New ADODB.Recordset
    RsIvas.Open SqlIva, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RsIvas.EOF
        TipoDIva = DevuelveValor("select tipodiva from tiposiva where codigiva = " & DBSet(RsIvas!codigiva, "N"))
        If TipoDIva = 1 Then
            'Es iva NO deducible
            ColumnaIVA = "cuentasn"
        Else
            ColumnaIVA = "cuentaso"   'La normal
        End If
        
        Cad3 = "cuentasr"
        Cad2 = DevuelveDesdeBD(ColumnaIVA, "tiposiva", "codigiva", RsIvas!codigiva, "N", Cad3)
        If Cad2 <> "" Then
            Sql = Mes & ",'" & Cad2 & "'," & DocConcAmp
            Cad2 = CadenaImporte(True, RsIvas!Impoiva, Importe0)
            Sql = Sql & "," & Cad2 & ","
            Sql = Sql & "NULL,'" & RF!codmacta & "','FRAPRO',0)"
            
            If Importe0 Then
                If vParam.ContabApteIva0 Then
                    If Not EsImportacion Then
                        Conn.Execute cad & Sql
                        Mes = Mes + 1
                    End If
                End If
            Else
                If Not EsImportacion Then
                    Conn.Execute cad & Sql
                    Mes = Mes + 1
                End If
            End If
            
            'La de recargo  1-----------------
            If Not IsNull(RsIvas!ImpoRec) Then
                Sql = Mes & "," & Cad3 & "," & DocConcAmp
                'Importes, atencion importes negativos
                Cad2 = CadenaImporte(True, RsIvas!ImpoRec, Importe0)
                Sql = Sql & "," & Cad2 & ","
                Sql = Sql & "NULL,'" & RF!codmacta & "','FRAPRO',0)"
                If Not Importe0 Then
                    Conn.Execute cad & Sql
                    Mes = Mes + 1
                End If
            End If
            
            If EsSujetoPasivo Then
                Cad3 = "cuentarr"
                Cad2 = DevuelveDesdeBD("cuentare", "tiposiva", "codigiva", RsIvas!codigiva, "N", Cad3)
                
                Cad3 = Cad2 & "|" & Cad3 & "|"
                
                
                Sql = Mes & ",'" & RecuperaValor(Cad3, 1) & "'," & DocConcAmp
                Cad2 = CadenaImporte(False, RsIvas!Impoiva, Importe0)
                Sql = Sql & "," & Cad2 & ","
                Sql = Sql & "NULL,'" & RF!codmacta & "','FRAPRO',0)"
                'If Not Importe0 Then
                    Conn.Execute cad & Sql
                    Mes = Mes + 1
                'End If
               
                If Not IsNull(RsIvas!ImpoRec) Then
                     Sql = Mes & "," & RecuperaValor(Cad3, 2) & "," & DocConcAmp
                    'Importes, atencion importes negativos
                    Cad2 = CadenaImporte(False, RsIvas!ImpoRec, Importe0)
                    Sql = Sql & "," & Cad2 & ","
                    Sql = Sql & "NULL,'" & RF!codmacta & "','FRAPRO',0)"
                    If Not Importe0 Then
                        Conn.Execute cad & Sql
                        Mes = Mes + 1
                    End If
                End If
            End If
            
        Else
            MsgBox "Error leyendo TIPO de IVA: " & RsIvas!codigiva, vbExclamation
            RF.Close
            Exit Function
        End If
    
        RsIvas.MoveNext
    Wend
    Set RsIvas = Nothing
    
    '-------------------------------------
    
    '-------------------------------------
    ' RETENCION
    A_Donde = "Retencion"
    If Not IsNull(RF!cuereten) Then
        Sql = Mes & ",'" & RF!cuereten & "'," & DocConcAmp
        'Importes, atencion importes negativos
        Cad2 = CadenaImporte(False, RF!trefacpr, Importe0)
        Sql = Sql & "," & Cad2 & ","
        Sql = Sql & "NULL,NULL,'FRAPRO',0)"
       
        Conn.Execute cad & Sql
        Mes = Mes + 1 'Es el contador de lineaapunteshco
    End If
    
    
    IncrementaProgres 2
    
    '------------------------------------------------------------
    'Las lineas de la factura. Para ello guardaremos algunos datos
    Cad2 = RF!codmacta
    ImporteD = RF!totfacpr
    
    
    
    'Cerramos el RF
    Cuenta = RF!codmacta
    RF.Close
    
    
    
    A_Donde = "Leyendo lineas factura"
    Sql = "Select factpro_lineas.codmacta, factpro_lineas.codccost, sum(factpro_lineas.baseimpo) baseimpo  FROM factpro_lineas "
    Sql = Sql & " WHERE numregis= " & NumFac
    Sql = Sql & " AND anofactu=" & NumDiari
    Sql = Sql & " GROUP BY 1,2 "
    Sql = Sql & " ORDER BY 1,2 "
    RF.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Para cada linea insertamos
    A_Donde = "Procesando lineas"
    Cad2 = ""
    While Not RF.EOF
        'Importes, atencion importes negativos
        If Cad2 = "" Then PrimeraContrapartida = RF!codmacta
        Sql = Mes & ",'" & RF!codmacta & "'," & DocConcAmp
        Cad2 = CadenaImporte(True, RF!Baseimpo, Importe0)
        Sql = Sql & "," & Cad2 & ","
        If IsNull(RF!codccost) Then
            Cad2 = "NULL"
        Else
            Cad2 = "'" & RF!codccost & "'"
        End If
        
        Sql = Sql & Cad2 & ",'" & Cuenta & "','FRAPRO',0)"
    
        Conn.Execute cad & Sql
        Mes = Mes + 1 'Es el contador de lineaapunteshco
        
        'Siguiente
        IncrementaProgres 1
        RF.MoveNext
        If Not RF.EOF Then PrimeraContrapartida = ""
    Wend
    RF.Close
    
    
    'AHora viene lo bueno.  MARZO 2006
    'Si el valor fuera true YA lo habria insertado en la cabcera
    If Not vParam.CCenFacturas Then
        If PrimeraContrapartida <> "" Then
            Sql = "UPDATE hlinapu SET ctacontr ='" & PrimeraContrapartida & "'"
            Sql = Sql & " WHERE numdiari = " & DiarioFacturas & " AND fechaent ='" & Fecha & "' and numasien = " & NumAsiento
            Sql = Sql & " AND linliapu =1 " 'LA PRIMERA LINEA SIEMPRE ES LA DE LA CUENTA
            EjecutaSQL Sql  'Lo hacemos aqui para controlar el error y que no explote
        End If
    End If
    
    'Actualimos en factura, el n� de asiento
    Sql = "UPDATE factpro SET numdiari = " & DiarioFacturas & ", fechaent = '" & Fecha & "', numasien =" & NumAsiento
    Sql = Sql & " WHERE  numregis = " & NumFac
    Sql = Sql & " AND anofactu=" & NumDiari
    Conn.Execute Sql
    
    'Para los saldos ponemos el numero de asiento donde toca
    '
    A_Donde = "Saldos factura"
    NumDiari = DiarioFacturas
    Mes = Month(FechaAsiento)
    Anyo = Year(FechaAsiento)
    
    IntegraLaFacturaProv = True
End Function

Private Function ActualizaAsiento() As Boolean
    Dim bol As Boolean
    Dim Donde As String
    On Error GoTo EActualizaAsiento
    
    'Obtenemos el mes y el a�o
    Mes = Month(FechaAsiento)
    Anyo = Year(FechaAsiento)
    Fecha = Format(FechaAsiento, FormatoFecha)
    
    'Comprobamos que no existe en historico
    If AsientoExiste Then
        If OpcionActualizar = 1 Then
            MsgBox "El asiento ya existe. Fecha: " & Fecha & "     N�: " & NumAsiento, vbExclamation
            Exit Function
        Else
            Sql = "Comprobar  -> El asiento ya existe. Fecha: " & Fecha & "     N�: " & NumAsiento
            InsertaError Sql
        End If
    End If
    
    'Aqui bloquearemos
    
    Conn.BeginTrans
    bol = ActualizaElASiento(Donde)
    
EActualizaAsiento:
        If Err.Number <> 0 Then
            Sql = "Actualiza Asiento." & vbCrLf & "----------------------------" & vbCrLf
            Sql = Sql & Donde
            If OpcionActualizar = 1 Then
                MuestraError Err.Number, Sql, Err.Description
            Else
                Sql = Donde & " -> " & Err.Description
                Sql = Mid(Sql, 1, 200)
                InsertaError Sql
            End If
            bol = False
        End If
        If bol Then
            Conn.CommitTrans
            ActualizaAsiento = True
            AlgunAsientoActualizado = True
        Else
            If OpcionActualizar = 1 Then
                MsgBox "Error: " & Donde, vbExclamation
            Else
                'FALTA###
            End If
            Conn.RollbackTrans
        End If
End Function


Private Function ActualizaElASiento(ByRef A_Donde As String) As Boolean



    ActualizaElASiento = False
    
    'Insertamos en cabeceras
    A_Donde = "Insertando datos en historico cabeceras asiento"
    If Not InsertarCabecera Then Exit Function
    IncrementaProgres 1
    
    'Insertamos en lineas
    A_Donde = "Insertando datos en historico lineas asiento"
    If Not InsertarLineas Then Exit Function
    IncrementaProgres 2
    
    'Modificar saldos
    A_Donde = "Calculando Lineas y saldos "
    If Not CalcularLineasYSaldos(False) Then Exit Function
    
    
    'Borramos cabeceras y lineas del asiento
    A_Donde = "Borrar cabeceras y lineas en asientos"
    If Not BorrarASiento(True) Then Exit Function
    IncrementaProgres 2
    ActualizaElASiento = True
End Function


Private Function InsertarCabecera() As Boolean
On Error Resume Next

    Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari) SELECT numdiari,fechaent,numasien,obsdiari from cabapu where "
    Sql = Sql & " numdiari =" & NumDiari
    Sql = Sql & " AND fechaent='" & Fecha & "'"
    Sql = Sql & " AND numasien=" & NumAsiento

    Conn.Execute Sql

    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarCabecera = False
    Else
        InsertarCabecera = True
    End If
End Function




Private Function AsientoExiste() As Boolean
    AsientoExiste = True
    Sql = "SELECT numdiari from hcabapu"
    Sql = Sql & " WHERE numdiari =" & NumDiari
    Sql = Sql & " AND fechaent='" & Fecha & "'"
    Sql = Sql & " AND numasien=" & NumAsiento
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs.EOF Then AsientoExiste = False
    Rs.Close
    Set Rs = Nothing
End Function


Private Function CalcularLineasYSaldos(EsDesdeRecalcular As Boolean) As Boolean
Dim Reparto As Boolean
Dim T As String



    On Error GoTo ECalcularLineasYSaldos

    Dim RL As Recordset
    Set RL = New ADODB.Recordset
    
    Sql = "SELECT sum(timporteD) AS SD, sum(timporteH) AS SH, codmacta"
    Sql = Sql & "  FROM"
    If EsDesdeRecalcular Then
        Sql = Sql & " hlinapu"
    Else
        Sql = Sql & " linapu"
    End If
    Sql = Sql & " WHERE (((numdiari)= " & NumDiari
    Sql = Sql & ") AND ((fechaent)='" & Fecha & "'"
    Sql = Sql & ") AND ((numasien)=" & NumAsiento
    Sql = Sql & ")) group by codmacta"
    
        
    
   
    Set RL = New ADODB.Recordset
    RL.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        Cuenta = RL!codmacta
        If IsNull(RL!sD) Then
            ImporteD = 0
        Else
            ImporteD = RL!sD
        End If
        If IsNull(RL!sH) Then
            ImporteH = 0
        Else
            ImporteH = RL!sH
        End If
        
        
        'Sig
        RL.MoveNext
    Wend
    RL.Close
    IncrementaProgres 3
    If Not vParam.autocoste Then
        'NO tiene analitica
        CalcularLineasYSaldos = True
        Exit Function
    End If
    
    
    If EsDesdeRecalcular Then
        T = "h"
    Else
        T = ""
    End If
    
    Sql = "SELECT timporteD AS SD, timporteH AS SH, codmacta,idsubcos," & T & "linapu.codccost"
    Sql = Sql & " FROM " & T & "linapu,ccoste WHERE ccoste.codccost=" & T & "linapu.codccost"
    Sql = Sql & " AND numdiari=" & NumDiari
    Sql = Sql & " AND fechaent='" & Fecha & "'"
    Sql = Sql & " AND numasien=" & NumAsiento
    Sql = Sql & " AND " & T & "linapu.codccost Is Not Null;"
    
    RL.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        Cuenta = RL!codmacta
        CCost = RL!codccost
        ImporteD = DBLet(RL!sD, "N")
        ImporteH = DBLet(RL!sH, "N")
        Reparto = (RL!idsubcos = 1)
        If Not CalcularSaldosAnal Then
            RL.Close
            Exit Function
        End If
        If Reparto Then
            If Not HacerReparto(True) Then
                RL.Close
                Exit Function
            End If
        End If
        'Sig
        RL.MoveNext
    Wend
    RL.Close
    IncrementaProgres 2
    CalcularLineasYSaldos = True
    Exit Function
ECalcularLineasYSaldos:
    Err.Clear
End Function




Private Function HacerReparto(Actualizar As Boolean) As Boolean
Dim RR As ADODB.Recordset
Dim AD As Currency
Dim AH As Currency
Dim TD As Currency
Dim TH As Currency
Dim B As Boolean

    HacerReparto = False
    TD = ImporteD
    TH = ImporteH
    AD = 0
    AH = 0
    Set RR = New ADODB.Recordset
    Sql = "Select * from ccoste_lineas WHERE codccost = '" & CCost & "'"
    RR.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RR.EOF
        'Cargamos los porcentajes
        CCost = RR!subccost
        ImporteD = (RR!porccost) / 100
        ImporteH = ImporteD
        'Importe porcentajeado
        ImporteD = Round(ImporteD * TD, 2)
        ImporteH = Round(ImporteH * TH, 2)
        'Movemos al sguiente
        RR.MoveNext
        'Por si acaso los decimales quedan sueltos entonces
        'Los valores para el ultimo subcentro de reaparto se obtienen por diferencias
        'con el acumulado
        If RR.EOF Then
            ImporteD = TD - AD
            ImporteH = TH - AH
        Else
            'Acumulo
            AD = AD + ImporteD
            AH = AH + ImporteH
        End If
        If Actualizar Then
            B = CalcularSaldosAnal
        Else
            B = CalcularSaldosAnalDesactualizar
        End If
        If Not B Then
            RR.Close
            Exit Function
        End If
    Wend
    RR.Close
    HacerReparto = True
End Function




Private Function InsertarLineas() As Boolean
On Error Resume Next
    Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab,punteada,traspasado)"
    Sql = Sql & " SELECT numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab,punteada,traspasado From linapu"
    Sql = Sql & " WHERE numasien = " & NumAsiento
    Sql = Sql & " AND numdiari = " & NumDiari
    Sql = Sql & " AND fechaent='" & Fecha & "'"
    Conn.Execute Sql
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        
        InsertarLineas = False
    Else
        InsertarLineas = True
    End If
End Function




'-------------------------------------------------------
'-------------------------------------------------------
'ANALITICA
'-------------------------------------------------------
'-------------------------------------------------------

Private Function CalcularSaldosAnal() As Boolean
    
    CalcularSaldosAnal = CalcularSaldos1NivelAnal(vEmpresa.numnivel)

End Function

Private Function CalcularSaldosAnalDesactualizar() As Boolean
    CalcularSaldosAnalDesactualizar = CalcularSaldos1NivelAnalDesactualizar(vEmpresa.numnivel)

End Function

Private Function CalcularSaldos1NivelAnal(Nivel As Integer) As Boolean
    Dim ImpD As Currency
    Dim ImpH As Currency
    Dim TD As String
    Dim TH As String
    Dim Cta As String
    Dim I As Integer
    
    
    CalcularSaldos1NivelAnal = False
    I = DigitosNivel(Nivel)
    If I < 0 Then Exit Function
    
    Cta = Mid(Cuenta, 1, I)
    Sql = "Select debccost,habccost from hsaldosanal where "
    Sql = Sql & " codccost='" & CCost & "' AND"
    Sql = Sql & " Codmacta = '" & Cta & "' AND anoccost = " & Anyo & " AND mesccost = " & Mes
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        I = 0   'Nuevo
        ImpD = 0
        ImpH = 0
    Else
        I = 1
        ImpD = Rs.Fields(0)
        ImpH = Rs.Fields(1)
    End If
    Rs.Close
    'Acumulamos
    ImpD = ImpD + ImporteD
    ImpH = ImpH + ImporteH
    TD = TransformaComasPuntos(CStr(ImpD))
    TH = TransformaComasPuntos(CStr(ImpH))
    If I = 0 Then
        'Nueva insercion
        Sql = "INSERT INTO hsaldosanal(codccost,codmacta,anoccost,mesccost,debccost,habccost)"
        Sql = Sql & " VALUES('" & CCost & "','" & Cta & "'," & Anyo & "," & Mes & "," & TD & "," & TH & ")"
        Else
        Sql = "UPDATE hsaldosanal SET debccost=" & TD & ", habccost = " & TH
        Sql = Sql & " WHERE Codmacta = '" & Cta & "' AND Anoccost = " & Anyo & " AND mesccost = " & Mes
        Sql = Sql & " AND codccost = '" & CCost & "';"
    End If
    Conn.Execute Sql
    CalcularSaldos1NivelAnal = True
End Function



Private Function CalcularSaldos1NivelAnalDesactualizar(Nivel As Integer) As Boolean
    Dim ImpD As Currency
    Dim ImpH As Currency
    Dim TD As String
    Dim TH As String
    Dim Cta As String
    Dim I As Integer
    Dim NoHaySaldoContinuar As Boolean
    
    CalcularSaldos1NivelAnalDesactualizar = False
    I = DigitosNivel(Nivel)
    If I < 0 Then Exit Function
    
    Cta = Mid(Cuenta, 1, I)
    Sql = "Select debccost,habccost from hsaldosanal where "
    Sql = Sql & " codccost='" & CCost & "' AND"
    Sql = Sql & " Codmacta = '" & Cta & "' AND anoccost = " & Anyo & " AND mesccost = " & Mes
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        I = 0
            Sql = "Error grave. No habia saldos en analitica: " & vbCrLf
            Sql = Sql & "Cuenta:    " & Cta & "      " & CCost & vbCrLf
            Sql = Sql & "Mes-a�o:     " & Mes & " / " & Anyo & vbCrLf & vbCrLf & "�Continuar?"
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then
                NoHaySaldoContinuar = False
            Else
                NoHaySaldoContinuar = True
            End If
            ImpD = 0
            ImpH = 0
        
        If Not NoHaySaldoContinuar Then
            
                
            Rs.Close
            Exit Function
        End If
    Else
        I = 1
        ImpD = Rs.Fields(0)
        ImpH = Rs.Fields(1)
    End If
    Rs.Close
    'Acumulamos
    ImpD = ImpD - ImporteD 'Con respecto a ACTUALIZAR CAMBIA EL SIGNO
    ImpH = ImpH - ImporteH
    TD = TransformaComasPuntos(CStr(ImpD))
    TH = TransformaComasPuntos(CStr(ImpH))
    If I > 0 Then
        Sql = "UPDATE hsaldosanal SET debccost=" & TD & ", habccost = " & TH
        Sql = Sql & " WHERE Codmacta = '" & Cta & "' AND Anoccost = " & Anyo & " AND mesccost = " & Mes
        Sql = Sql & " AND codccost = '" & CCost & "';"
        Conn.Execute Sql
    Else
        Sql = "INSERT INTO hsaldosanal (codmacta, anoccost, mesccost, debccost, habccost,codccost) VALUES "
        Sql = Sql & "('" & Cta & "'," & Anyo & "," & Mes & ","
        Sql = Sql & TD & "," & TH & ",'" & CCost & "')"
        EjecutaSQL Sql   'Para que si da error nos deje tranquilos
    End If
    CalcularSaldos1NivelAnalDesactualizar = True
End Function




Private Function BorrarASiento(BorrarCabecera As Boolean) As Boolean

On Error GoTo EBorrarASiento
    BorrarASiento = False
    
    'Borramos las lineas
    Sql = "Delete from hlinapu"
    Sql = Sql & " WHERE numasien = " & NumAsiento
    Sql = Sql & " AND numdiari = " & NumDiari
    Sql = Sql & " AND fechaent=" & DBSet(FechaAnterior, "F")
    Conn.Execute Sql
    
    If BorrarCabecera Then
        'La cabecera
        Sql = "Delete from hcabapu"
        Sql = Sql & " WHERE numdiari =" & NumDiari
        Sql = Sql & " AND fechaent=" & DBSet(FechaAnterior, "F")
        Sql = Sql & " AND numasien=" & NumAsiento
        
        Conn.Execute Sql
    Else
        'Actualizamos la fecha de la cabecera
        Sql = "Update hcabapu"
        Sql = Sql & " set fechaent = " & DBSet(Fecha, "F")
        Sql = Sql & " WHERE numdiari =" & NumDiari
        Sql = Sql & " AND fechaent=" & DBSet(FechaAnterior, "F")
        Sql = Sql & " AND numasien=" & NumAsiento
    
        Conn.Execute Sql
    End If
    
    BorrarASiento = True
    Exit Function
EBorrarASiento:
    Err.Clear
    
End Function

Private Sub ObtenFoco(ByRef T As TextBox)
T.SelStart = 0
T.SelLength = Len(T.Text)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If NumErrores > 0 Then CerrarFichero
End Sub

Private Sub CerrarFichero()
On Error Resume Next
If NE = 0 Then Exit Sub
Close #NE
If Err.Number <> 0 Then Err.Clear
End Sub








Private Function InsertaError(ByRef CADENA As String)
Dim vS As String
    'Insertamos en errores
    'Esta lo tratamos con error especifico
    
    On Error Resume Next

    If OpcionActualizar < 10 Then
        'Insertamos error para ASIENTOS
        vS = NumDiari & "|"
        vS = vS & Fecha & "|"
        vS = vS & NumAsiento & "|"
        vS = vS & CADENA & "|"
    
    Else
        vS = NUmSerie & " " & NumFac & "|"
        vS = vS & FechaAsiento & "|"
        vS = vS & CADENA & "|"
    End If
    A�adeError vS
    
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error." & vbCrLf & Err.Description & vbCrLf & vS
        Err.Clear
    End If
End Function


Private Function DesActualizaAsiento() As Boolean
    Dim bol As Boolean
    Dim Donde As String
    On Error GoTo EDesActualizaAsiento
    
    
    '2.- Desactualiza pero NO insertes en apuntes
    '3.- Desactualizar asiento desde hco
    
    'Obtenemos el mes y el a�o
    Mes = Month(FechaAsiento)
    Anyo = Year(FechaAsiento)
    Fecha = Format(FechaAsiento, FormatoFecha)
    
    'Comprobamos que no existe en APUNTES
    'Obviamente solo comprobamos si vamos a insertar
    'en apuntes
    If Me.OpcionActualizar = 3 Then
        If AsientoExiste Then Exit Function
    End If
    'Aqui bloquearemos
    
    Conn.BeginTrans
    
    bol = DesActualizaElASiento(Donde)
    
EDesActualizaAsiento:
        If Err.Number <> 0 Then
            Sql = "Actualiza Asiento." & vbCrLf & "----------------------------" & vbCrLf
            Sql = Sql & Donde
            MuestraError Err.Number, Sql, Err.Description
            bol = False
        End If
        If bol Then
            Conn.CommitTrans
            espera 0.2
            DesActualizaAsiento = True
            AlgunAsientoActualizado = True
        Else
            Conn.RollbackTrans
        End If
End Function


Private Function DesActualizaElASiento(ByRef A_Donde As String) As Boolean

    '2  .- Desactualiza pero NO insertes en apuntes
    '      Si viene FRACLI o FRAPROV habr� que volver
    '3  .- Desactualizar asiento desde hco
        


    DesActualizaElASiento = False
    
    Select Case Me.OpcionActualizar
    
    Case 2
        If NUmSerie = "FRACLI" Or NUmSerie = "FRAPRO" Then
            A_Donde = "Desvinculando facturas"
            If Not DesvincularFactura(NUmSerie = "FRACLI") Then Exit Function
            IncrementaProgres 1
        End If
    End Select
    
    
    'Borramos cabeceras y lineas del asiento
    A_Donde = "Borrar cabeceras y lineas en historico"
    
    If OpcionActualizar = 2 Then
        If Not BorrarASiento(False) Then Exit Function
    Else
        If Not BorrarASiento(True) Then Exit Function
    End If
    
    IncrementaProgres 2
    DesActualizaElASiento = True
End Function

Private Function DesvincularFactura(Clientes As Boolean) As Boolean
On Error Resume Next
    Set Rs = New ADODB.Recordset
    If Clientes Then
        CCost = "factcli"
    Else
        CCost = "factpro"
    End If
    Sql = "Select * From " & CCost
    Sql = Sql & " WHERE numasien=" & NumAsiento
    Sql = Sql & " AND numdiari = " & NumDiari
    Sql = Sql & " AND fechaent = '" & Fecha & "'"
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        
        Sql = "UPDATE " & CCost & " SET numasien=NULL, fechaent=NULL, numdiari=NULL"
        If Clientes Then
            Sql = Sql & " WHERE numfactu = " & Rs!codfaccl
            Sql = Sql & " AND anofaccl =" & Rs!anofaccl
            Sql = Sql & " AND numserie = '" & Rs!NUmSerie & "'"
        Else
            'proveedores
            Sql = Sql & " WHERE numregis = " & Rs!NumRegis
            Sql = Sql & " AND anofactu =" & Rs!anofactu
        End If
        Conn.Execute Sql
    End If
    If Err.Number <> 0 Then
        DesvincularFactura = False
        MuestraError Err.Number, "Desvincular factura"
    Else
        DesvincularFactura = True
    End If
End Function


Private Function CalcularLineasYSaldosDesactualizar() As Boolean
    Dim RL As Recordset
    Dim Reparto As Boolean
    Set RL = New ADODB.Recordset
    
    
    '------------------------------------------
    'SALDOS
    'Calculamos sumas importes asiento en hco
    CalcularLineasYSaldosDesactualizar = False
    
    Sql = "SELECT sum(timporteD) AS SD, sum(timporteH) AS SH, codmacta"
    Sql = Sql & "  FROM  hlinapu"
    Sql = Sql & " WHERE (((numdiari)= " & NumDiari
    Sql = Sql & ") AND ((fechaent)='" & Fecha & "'"
    Sql = Sql & ") AND ((numasien)=" & NumAsiento
    Sql = Sql & ")) group by codmacta"
    
    
    
    Set RL = New ADODB.Recordset
    RL.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        Cuenta = RL!codmacta
        ImporteD = DBLet(RL!sD, "N")
        ImporteH = DBLet(RL!sH, "N")
        'Sig
        RL.MoveNext
    Wend
    RL.Close
    IncrementaProgres 3
    
    If Not vParam.autocoste Then
        'NO tiene analitica
        CalcularLineasYSaldosDesactualizar = True
        Exit Function
    End If
    
    
    
    '       ANALITICA
    Sql = "SELECT hlinapu.timporteD AS SD, hlinapu.timporteH AS SH, hlinapu.codmacta,"
    Sql = Sql & " hlinapu.fechaent, hlinapu.numdiari, hlinapu.numasien, hlinapu.codccost,ccoste.idsubcos"
    Sql = Sql & " From hlinapu,ccoste WHERE hlinapu.codccost=ccoste.codccost"
    Sql = Sql & " AND hlinapu.numdiari=" & NumDiari
    Sql = Sql & " AND hlinapu.fechaent='" & Fecha & "'"
    Sql = Sql & " AND hlinapu.numasien=" & NumAsiento
    Sql = Sql & " AND hlinapu.codccost Is Not Null;"
    RL.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        Cuenta = RL!codmacta
        CCost = RL!codccost
        ImporteD = DBLet(RL!sD, "N")
        ImporteH = DBLet(RL!sH, "N")
        Reparto = (RL!idsubcos = 1)
        If Not CalcularSaldosAnalDesactualizar Then
            RL.Close
            Exit Function
        End If
        If Reparto Then
            If Not HacerReparto(False) Then
                RL.Close
                Exit Function
            End If
        End If
        'Sig
        RL.MoveNext
    Wend
    RL.Close
    IncrementaProgres 2
    CalcularLineasYSaldosDesactualizar = True
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub BorrarArchivoTemporal()
On Error Resume Next
If Dir(App.Path & "\ErrActua.txt") <> "" Then Kill App.Path & "\ErrActua.txt"
If Err.Number <> 0 Then MuestraError Err.Number, "Borrar fichero temporal"
End Sub





Private Function DevuelveCentroCosteFactura(Cliente As Boolean, LaPrimeraContrapartida As String) As String
Dim R As ADODB.Recordset
Dim Sql As String
    DevuelveCentroCosteFactura = ""
    If Cliente Then
        
        Sql = "SELECT codccost,numlinea,codtbase FROM linfact"
        Sql = Sql & " WHERE numserie='" & NUmSerie
        Sql = Sql & "' AND codfaccl= " & NumFac
        Sql = Sql & " AND anofaccl=" & NumDiari
        Sql = Sql & " AND not (codccost is null)"   'El primero k devuelva
        Sql = Sql & " ORDER BY numlinea"
    Else
        Sql = "SELECT codccost,numlinea,codtbase FROM linfactprov"
        Sql = Sql & " WHERE numregis = " & NumFac
        Sql = Sql & " AND anofacpr=" & NumDiari
        Sql = Sql & " AND not (codccost is null)"   'El primero k devuelva
        Sql = Sql & " ORDER BY numlinea"
    End If
    
    
    Set R = New ADODB.Recordset
    R.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not R.EOF Then
        If Not IsNull(R.Fields(0)) Then DevuelveCentroCosteFactura = R.Fields(0)
        LaPrimeraContrapartida = R!codtbase
        R.MoveNext
        If Not R.EOF Then LaPrimeraContrapartida = ""
    End If
    R.Close
    Set R = Nothing
End Function








