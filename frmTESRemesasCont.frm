VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESRemesasCont 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "1"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "frmTESRemesasCont.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6240
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameContabilRem2 
      Height          =   4215
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5535
      Begin VB.CheckBox chkAgrupaCancelacion 
         Caption         =   "Agrupa cancelacion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   450
         TabIndex        =   6
         Top             =   3120
         Width           =   2535
      End
      Begin VB.CommandButton cmdContabRemesa 
         Caption         =   "Contabilizar"
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
         Left            =   2490
         TabIndex        =   4
         Top             =   3600
         Width           =   1425
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1950
         TabIndex        =   3
         Text            =   "Text4"
         Top             =   2520
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   1980
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1920
         Width           =   1365
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
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
         Index           =   8
         Left            =   4020
         TabIndex        =   5
         Top             =   3600
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "Gastos (€)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   450
         TabIndex        =   8
         Top             =   2490
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   18
         Left            =   450
         TabIndex        =   7
         Top             =   1950
         Width           =   750
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   10
         Left            =   1260
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "CONTABILIZAR REMESA"
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
         Index           =   2
         Left            =   180
         TabIndex        =   1
         Top             =   810
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmTESRemesasCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Opcion As Byte
    '8.- Contabilizar remesa
        
    
    
Public SubTipo As Byte

    'Para la opcion 22
    '   Remesas cancelacion cliente.
    '       1:  Efectos
    '       2: Talones pagares
    
'Febrero 2010
'Cuando pago proveedores con un talon, y le he indicado el numero
Public NumeroDocumento As String
    
    
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBa As frmBanco
Attribute frmBa.VB_VarHelpID = -1
Private WithEvents frmRe As frmColRemesas2
Attribute frmRe.VB_VarHelpID = -1
Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmP As frmFormaPago
Attribute frmP.VB_VarHelpID = -1


Dim RS As ADODB.Recordset
Dim SQL As String
Dim I As Integer
Dim IT As ListItem  'Comun
Dim PrimeraVez As Boolean
Dim Cancelado As Boolean
Dim CuentasCC As String




Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 21 Or Index = 25 Or Index = 31 Then CadenaDesdeOtroForm = "" 'ME garantizo =""
    If Index = 31 Then
        If MsgBox("¿Cancelar el proceso?", vbQuestion + vbYesNo) = vbYes Then SubTipo = 0
    End If
    Unload Me
End Sub




Private Sub cmdContabRemesa_Click()
Dim B As Boolean
Dim Importe As Currency
Dim CC As String
Dim Opt As Byte
Dim AgrupaCance As Boolean
Dim ContabilizacionEspecialNorma19 As Boolean


'Dim ImporteEnRecepcion As Currency
'Dim TalonPagareBeneficios As String
    SQL = ""
    
    If Text1(10).Text = "" Then SQL = "Ponga la fecha de abono"
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    'Fecha pertenece a ejercicios contbles
    If FechaCorrecta2(CDate(Text1(10).Text), True) > 1 Then Exit Sub
    
    
    'Ahora miramos la remesa. En que sitaucion , y de que tipo es
    SQL = "Select * from remesas where codigo =" & RecuperaValor(NumeroDocumento, 1)
    SQL = SQL & " AND anyo =" & RecuperaValor(NumeroDocumento, 2)
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If RS.EOF Then
        MsgBox "Ninguna remesa con esos valores", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Sub

    End If
    
    'Tiene valor
    SQL = ""
    B = AdelanteConLaRemesa()
    ContabilizacionEspecialNorma19 = False
    If B Then
        'Si es norma19 y tiene le parametro de contabilizacion por fecha comprobaremos la fecha de los vtos
        If Opcion = 8 Then
        
            'Se podrian agrupar los IFs, pero asi de momento me entero mas
        
            'Para RECIBOS BANCARIOS SOLO
            If DBLet(RS!Tiporem, "N") = 1 Then
                If vParam.Norma19xFechaVto Then
                    If Not IsNull(RS!Tipo) Then
                        If RS!Tipo = 0 Then
                            'NORMA 19
                            'Contbiliza por fecha VTO
                            'Comprobaremos que toooodos estan en fecha ejercicio
                            SQL = ComprobacionFechasRemesaN19PorVto
                            If SQL <> "" Then SQL = "-Comprobando fechas remesas N19" & vbCrLf & SQL
                            
                            
                            If txtImporte(0).Text <> "" Then SQL = SQL & vbCrLf & "N19 no permite gastos bancario"
                            
                            
                            If SQL <> "" Then
                                B = False
                            Else
                                ContabilizacionEspecialNorma19 = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
    End If

    If Not B Then
        If SQL = "" Then SQL = "Error y punto"
        RS.Close
        Set RS = Nothing
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    SQL = "Select cobros.codmacta,nomclien,fecbloq from scobro,cuentas where cobros.codmacta = cuentas.codmacta"
    SQL = SQL & " and  codrem =" & Text3(3).Text
    SQL = SQL & " AND anyorem =" & Text3(4).Text
    SQL = SQL & " AND fecbloq <='" & Format(Text1(10).Text, FormatoFecha) & "' GROUP BY 1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not miRsAux.EOF
        SQL = SQL & miRsAux!codmacta & Space(10) & miRsAux!FecBloq & Space(10) & miRsAux!Nomclien & vbCrLf
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If SQL <> "" Then
        CC = "Cuenta          Fec. bloqueo           Nombre" & vbCrLf & String(80, "-") & vbCrLf
        CC = "Cuentas bloqueadas" & vbCrLf & vbCrLf & CC & SQL
        MsgBox CC, vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Sub
    End If
       
       
       
    'Bloqueariamos la opcion de modificar esa remesa
        
        Importe = TextoAimporte(txtImporte(0).Text)
  
        'Tiene gastos. Falta ver si tiene la cuenta de gastos configurada. ASi como
        'si es analitica, el CC asociado
        CC = ""
        If vParam.autocoste Then CC = "codccost"
            
        SQL = DevuelveDesdeBD("ctagastos", "bancos", "codmacta", RS!codmacta, "T", CC)
        If SQL = "" Then
            MsgBox "Falta configurar la cuenta de gastos del banco:" & RS!codmacta, vbExclamation
            Set RS = Nothing
            Exit Sub
        End If
        
        If vParam.autocoste Then
            If CC = "" Then
                MsgBox "Necesita asignar centro de coste a la cuenta de gastos del banco: " & RS!codmacta, vbExclamation
                Set RS = Nothing
                Exit Sub
            End If
        End If
        
        SQL = SQL & "|" & CC & "|"
        
        
        'Añado, si tiene, la cuenta de ingresos
        CC = DevuelveDesdeBD("ctaingreso", "bancos", "codmacta", RS!codmacta, "T")
        If CC = "" Then
            If Importe > 0 Then
                MsgBox "Falta configurar la cuenta de ingresos del banco:" & RS!codmacta, vbExclamation
                Set RS = Nothing
                Exit Sub
            End If
        End If
        
        SQL = SQL & CC & "|"   'La
        

    SQL = RS!codmacta & "|" & SQL
    
    
    'Contab. remesa. Si es talon/pagare vamos a comprobar si hay diferencias entre el importe del documento
    'y el total de lineas
    B = False    'Si ya se ha hecho la pregunta no la volveremos a repetir
    'TalonPagareBeneficios = ""    'Solo para TAL/PAG y si hay importe beneficios etc

    
    'Pregunta conbilizacion
    If Not B Then   'Si no hemos hecho la pregunta en otro sitio la hacemos ahora
        Select Case Opcion
        Case 8
            CC = "Va a abonar"
        Case 22
            CC = "Procede a realizar la cancelacion del cliente de"
        Case 23
            CC = "Procede a realizar la confirmacion de"
        End Select
        CC = CC & " la remesa: " & RS!Codigo & " / " & RS!Anyo & vbCrLf & vbCrLf
        CC = CC & Space(30) & "¿Continuar?"
        If SubTipo = 2 Then
            If Val(RS!Tiporem) = 3 Then
                CC = "Talón" & vbCrLf & CC
            Else
                CC = "Pagaré" & vbCrLf & CC
            End If
            CC = "Tipo: " & CC
        End If
    
        If MsgBox(CC, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    'Para llevarlos a hco
    Conn.Execute "DELETE from tmpactualizar  where codusu =" & vUsu.Codigo
    
        
    
    If Opcion = 8 Then
        'CONTABILIZACION    ABONO REMESA
        
        'NORMA 19
        '------------------------------------
        
        'Contabilizaremos la remesa
        Conn.BeginTrans
        
        'mayo 2012
        If ContabilizacionEspecialNorma19 Then
            'Utiliza Morales
            'Es para contabilizar los recibos por fecha de vto
            
            B = ContabNorma19PorFechaVto(RS!Codigo, RS!Anyo, SQL)
        Else
            'Toooodas las demas opciones estan aqui
        
                                    'Efecto(1),pagare(2),talon(3)
            B = ContabilizarRecordsetRemesa(RS!Tiporem, DBLet(RS!Tipo, "N") = 0, RS!Codigo, RS!Anyo, SQL, CDate(Text1(10).Text), Importe)
        
        End If
        
        'si se contabiliza entonces updateo y la pongo en
        'situacion Q. Contabilizada a falta de devueltos ,
        If B Then
            Conn.CommitTrans
            'AQUI updateamos el registro pq es una tabla myisam
            'y no debemos meterla en la transaccion
            SQL = "UPDATE remesas SET"
            SQL = SQL & " situacion= 'Q'"
            SQL = SQL & " WHERE codigo=" & RS!Codigo
            SQL = SQL & " and anyo=" & RS!Anyo

            If Not Ejecuta(SQL) Then MsgBox "Error actualizando tabla remesa.", vbExclamation
            
            
            'Ahora actualizamos los registros que estan en tmpactualziar
            frmActualizar2.OpcionActualizar = 20
            frmActualizar2.Show vbModal
            Screen.MousePointer = vbDefault
            'Cerramos
            RS.Close
            Unload Me
            Exit Sub
        Else
            'ANtes
            'Conn.RollbackTrans
            'Ahora
            TirarAtrasTransaccion
        End If
    
    
    Else
        Conn.BeginTrans
      
        'Cancelacion /confirmacion cliente
        If SubTipo = 1 Then
            'EFECTOS
            If Opcion <= 23 Then
            
                'YA NO EXISTE CONFIRMACION REMESA
                Opt = Opcion - 22 '0.Cancelar   1.Confirmar
                AgrupaCance = False
                If Me.chkAgrupaCancelacion.Visible Then
                    If Me.chkAgrupaCancelacion.Value = 1 Then AgrupaCance = True
                End If
                
                'para la 23 NO deberiamos llegar. Ese proceso lo hemos eliminado
                If Opt = 0 Then
                    B = RemesasCancelacionEfectos(RS!Codigo, RS!Anyo, SQL, CDate(Text1(10).Text), Importe, AgrupaCance)
                Else
                    B = False
                    MsgBox " NO deberia haber entrado con confirmacion remesas", vbExclamation
                End If
            Else
                B = False
                MsgBox "Opcion incorrecta (>23)", vbExclamation
            End If
            
        Else
            MsgBox "AHora no deberia estar aqui!!!!!", vbExclamation
            
                                 '
            'B = RemesasCancelacionTALONPAGARE(Val(Rs!tiporem) = 3, Rs!Codigo, Rs!Anyo, SQL, CDate(Text1(10).Text), Importe)
        End If
        If B Then
            Conn.CommitTrans
            
            
            'Ahora actualizamos los registros que estan en tmpactualziar
            frmActualizar2.OpcionActualizar = 20
            frmActualizar2.Show vbModal
            Screen.MousePointer = vbDefault
            'Cerramos
            RS.Close
            Unload Me
            Exit Sub
            
        Else
            TirarAtrasTransaccion
        End If
        
    End If
    
    
    
    RS.Close
    Set RS = Nothing
    Screen.MousePointer = vbDefault
End Sub




Private Function AdelanteConLaRemesa() As Boolean
Dim C As String

    AdelanteConLaRemesa = False
    SQL = ""
    
    'Efectos eliminados
    If RS!Situacion = "Z" Or RS!Situacion = "Y" Then SQL = "Efectos eliminados"
    
    'abierta sin llevar a banco. Esto solo es valido para las de efectos
    If SubTipo = 1 Then
        If RS!Situacion = "A" Then SQL = "Remesa abierta. Sin llevar al banco."
    
    End If
    'Ya contabilizada
    If RS!Situacion = "Q" Then SQL = "Remesa abonada."
    
    If SQL <> "" Then Exit Function
    
    
    
    
    If Opcion = 8 Then
        'COntbilizar / abonar remesa
        '---------------------------------------------------------------------------
        If SubTipo = 1 Then
            'Febrero 2009
            'Ahora toooodas las remesas se hace lo mismmo
            ' De llevada a banco a cancelar cliente. De cancelar a abonar y de abonar a eliminar. NO
            'hay distinciones entre remesas. Para podrer abonar una remesa esta tiene que estar cancelada
            If vParam.EfectosCtaPuente Then
                If RS!Situacion <> "F" Then SQL = "La remesa NO puede abonarse. Faltan cancelacion "
            End If
            
        Else
            If RS!Tiporem = 2 And vParam.PagaresCtaPuente Then
                If RS!Situacion <> "F" Then SQL = "La remesa NO puede abonarse. Falta cancelación "
            End If
            
            If RS!Tiporem = 3 And vParam.TalonesCtaPuente Then
                If RS!Situacion <> "F" Then SQL = "La remesa NO puede abonarse. Falta cancelación "
            End If
        End If
        
            
    Else
       'Vamos a proceder al proceso de generacion cancelacion  /* CANCELACION */
       If SubTipo = 1 Then
            'Para los efectos la norma no tiene que ser 19
            'Febrero 2009.  Para tooodas las normas
            'If Rs!Tipo = 0 Then
            '    SQL = "Proceso no válido para NORMA 19"
            '    Exit Function
            'End If
        
       End If
       
       'Para elos tipos 1,2
       If Opcion = 22 Then
            'Cancelacion cliente
            'Para los efectos, tiene que estar generado soporte. Para talones/pagares no es obligado
            If SubTipo = 1 Then
                If RS!Situacion <> "B" Then SQL = "Para cancelar la remesa deberia esta en situación 'Soporte generado'"
            Else
                If RS!Situacion = "F" Then SQL = "Remesa YA cancelada"
            End If
        Else
            'Febrero 2009
            'No hay confirmacion
            SQL = "Opción de confirmacion NO es válida"
            'Confirmacion
            'If Rs!situacion <> "F" Then SQL = "Para confirmar la remesa esta deberia estar 'Cancelacion cliente'"
       End If
       
       
       'Si hasta aqui esta bien:
       'Compruebo que tiene configurado en parametros
       If SQL = "" Then
            'Comprobamos si esta bien configurada
            '
            If SubTipo = 1 Then
                    If Opcion = 22 Then
                        'SQL = "4310"
                        SQL = "RemesaCancelacion"
                    Else
                        SQL = "RemesaConfirmacion"
                    End If
                    SQL = DevuelveDesdeBD(SQL, "paramtesor", "codigo", "1")
                    If SQL = "" Then
                        SQL = "Falta configurar parámetros cuentas confirmación/cancelación remesa. "
                    Else
                        'OK. Esta configurado
                        SQL = ""
                    End If
                    
            Else
                'talones pagares
                'Veremos si esta configurado(y bien configurado) para el proceso
                If RS!Tiporem = 2 Then
                    'Pagare
                    C = "contapagarepte"
                ElseIf RS!Tiporem = 3 Then
                    'Talones
                    C = "contatalonpte"
                Else
                    'NO DEBIA HABERSE METIDO AQUI
                    C = ""
                    
                End If
                If C = "" Then
                    SQL = "Error validando tipo de remesa"
                    
                Else
                    C = DevuelveDesdeBD(C, "paramtesor", "codigo", 1)
                    If C = "" Then C = "0"
                    If Val(C) = 0 Then
                        SQL = "Falta configurar la aplicacion para las remesas de talones / pagares"
                    Else
                        SQL = ""
                    End If
                End If
            End If
       End If
    End If
    AdelanteConLaRemesa = SQL = ""
    
End Function











'
'
'
Private Sub NuevaRem()

Dim ForPa As String
Dim Cad As String
Dim Impor As Currency
Dim colCtas As Collection

'Algunas conideraciones

    If SubTipo <> vbTipoPagoRemesa Then
        'Para talones y pagares obligado la cuenta bancaria
        If txtCta(3).Text = "" Then
            MsgBox "Indique la cuenta bancaria", vbExclamation
            Exit Sub
        End If
    End If


    'Fecha remesa tiene k tener valor
    If Text1(8).Text = "" Then
        MsgBox "Fecha de remesa debe tener valor", vbExclamation
        Ponerfoco Text1(8)
        Exit Sub
    End If
    
    
    
    'VEMOS SI LA FECHA ESTA DENTRO DEL EJERCICIO
    If FechaCorrecta2(CDate(Text1(8).Text), True) > 1 Then
        Ponerfoco Text1(8)
        Exit Sub
    End If
    
    'Para talones pagares, vemos si esta configurado en parametros
    If SubTipo <> vbTipoPagoRemesa Then
        If Me.cmbRemesa.ListIndex = 0 Then
            SQL = "contapagarepte"
        Else
            SQL = "contatalonpte"
        End If
        SQL = DevuelveDesdeBD(SQL, "paramtesor", "codigo", "1")
        If SQL = "" Then SQL = "0"
        If SQL = "0" Then
            MsgBox "Falta configurar la opción en parametros", vbExclamation
            Exit Sub
        End If
    End If
    
    'mayo 2015
     If SubTipo = vbTipoPagoRemesa Then
        If vParam.RemesasPorEntidad Then
            If chkAgruparRemesaPorEntidad.Value = 1 Then
                'Si agrupa pro entidad, necesit el banco por defacto
                If txtCta(3).Text = "" Then
                    MsgBox "Si agrupa por entidad debe indicar el banco por defecto", vbExclamation
                    Exit Sub
                End If
            End If
        End If
    End If
    'A partir de la fecha generemos leemos k remesa corresponde
    SQL = "select max(codigo) from remesas where anyo=" & Year(CDate(Text1(8).Text))
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then
        NumRegElim = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    
    NumRegElim = NumRegElim + 1
    txtRemesa.Text = NumRegElim

    
    If SubTipo = vbTipoPagoRemesa Then
        SQL = " sforpa.tipforpa = " & vbTipoPagoRemesa
    Else
        If Me.cmbRemesa.ListIndex = 0 Then
            SQL = " talon = 0"
        Else
            SQL = " talon = 1"
        End If
    
    End If
    
    If SubTipo = vbTipoPagoRemesa Then
        'Del efecto
        If Text1(6).Text <> "" Then SQL = SQL & " AND scobro.fecvenci >= '" & Format(Text1(6).Text, FormatoFecha) & "'"
        If Text1(7).Text <> "" Then SQL = SQL & " AND scobro.fecvenci <= '" & Format(Text1(7).Text, FormatoFecha) & "'"
    Else
        'de la recepcion de factura
        If Text1(6).Text <> "" Then SQL = SQL & " AND fechavto >= '" & Format(Text1(6).Text, FormatoFecha) & "'"
        If Text1(7).Text <> "" Then SQL = SQL & " AND fechavto <= '" & Format(Text1(7).Text, FormatoFecha) & "'"
    End If
        
    
    
    'Si ha puesto importe desde Hasta
    If txtImporte(6).Text <> "" Then SQL = SQL & " AND impvenci >= " & TransformaComasPuntos(ImporteFormateado(txtImporte(6).Text))
    If txtImporte(7).Text <> "" Then SQL = SQL & " AND impvenci <= " & TransformaComasPuntos(ImporteFormateado(txtImporte(7).Text))

    
    
    'Desde hasta cuenta
    If SubTipo = vbTipoPagoRemesa Then
        If Me.txtCtaNormal(3).Text <> "" Then SQL = SQL & " AND scobro.codmacta >= '" & txtCtaNormal(3).Text & "'"
        If Me.txtCtaNormal(4).Text <> "" Then SQL = SQL & " AND scobro.codmacta <= '" & txtCtaNormal(4).Text & "'"
        'El importe
        SQL = SQL & " AND impvenci > 0"
        
        
        
        'MODIFICACION DE 2 DICIEMBRE del 05
        '------------------------------------
        'Hay un campo que indicara si el vto se remesa o NO
        SQL = SQL & " AND noremesar=0"


        'Si esta en situacion juridica TAMPOCO se remesa
        SQL = SQL & " AND situacionjuri=0"

        'JUNIO 2010
        'Si tiene algio  cobrado NO dejo remesar
        SQL = SQL & " AND impcobro is null"
    

    End If
    

    'Marzo 2015
    'Comprobar
    
    
 
    
    
    
    'Modificacion 28 Abril 06
    '------------------------
    ' Es para acotar mas el conjunto de recibos a remesar
    'Serie
    If SubTipo = vbTipoPagoRemesa Then
        If txtSerie(0).Text <> "" Then _
            SQL = SQL & " AND scobro.numserie >= '" & txtSerie(0).Text & "'"
        If txtSerie(1).Text <> "" Then _
            SQL = SQL & " AND scobro.numserie <= '" & txtSerie(1).Text & "'"
        
        'Fecha factura
        If Text1(22).Text <> "" Then _
            SQL = SQL & " AND scobro.fecfaccl >= '" & Format(Text1(22).Text, FormatoFecha) & "'"
        If Text1(21).Text <> "" Then _
            SQL = SQL & " AND scobro.fecfaccl <= '" & Format(Text1(21).Text, FormatoFecha) & "'"
        
        'Codigo factura
        If txtNumFac(0).Text <> "" Then _
            SQL = SQL & " AND scobro.codfaccl >= '" & txtNumFac(0).Text & "'"
        If txtNumFac(1).Text <> "" Then _
            SQL = SQL & " AND scobro.codfaccl <= '" & txtNumFac(1).Text & "'"
        
        
    Else
        'Fecha factura
        If Text1(22).Text <> "" Then SQL = SQL & " AND fecharec >= '" & Format(Text1(22).Text, FormatoFecha) & "'"
        If Text1(21).Text <> "" Then SQL = SQL & " AND fecharec <= '" & Format(Text1(21).Text, FormatoFecha) & "'"
    
    End If
     
    
    Screen.MousePointer = vbHourglass
    Set RS = New ADODB.Recordset
    
    'Marzo 2015
    'Ver si entre los desde hastas hay importes negativos... ABONOS
    
    If SubTipo = vbTipoPagoRemesa Then
    
        'Vemos las cuentas que vamos a girar . Sacaremos codmacta
        Cad = SQL
        Cad = "scobro.codmacta=cuentas.codmacta AND (siturem is null) AND " & Cad
        Cad = Cad & " AND scobro.codforpa = sforpa.codforpa ORDER BY codmacta,codfaccl "
        Cad = "Select distinct scobro.codmacta FROM scobro,cuentas,sforpa WHERE " & Cad
        RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Set colCtas = New Collection
        While Not RS.EOF
            colCtas.Add CStr(RS!codmacta)
            RS.MoveNext
        Wend
        RS.Close
        
        
        'Ahora veremos los negativos, de las cuentas que vamos a girar
        'Sol el select de los negativos , sin numserie ni na de na
        Cad = "impvenci < 0"
        Cad = "scobro.codmacta=cuentas.codmacta AND (siturem is null) AND " & Cad
        Cad = Cad & " AND scobro.codforpa = sforpa.codforpa  "
        Cad = "Select scobro.codmacta,nommacta,numserie,codfaccl,impvenci FROM scobro,cuentas,sforpa WHERE " & Cad
        
        
        If colCtas.Count > 0 Then
            Cad = Cad & " AND scobro.codmacta IN ("
            For I = 1 To colCtas.Count
                If I > 1 Then Cad = Cad & ","
                Cad = Cad & "'" & colCtas.Item(I) & "'"
            Next
            Cad = Cad & ") ORDER BY codmacta,codfaccl"
        
            'Seguimos
        
        
            Set colCtas = Nothing
            RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Cad = ""
            I = 0
            Set colCtas = New Collection
            While Not RS.EOF
                If I < 15 Then
                    Cad = Cad & vbCrLf & RS!codmacta & " " & RS!Nommacta & "  " & RS!NUmSerie & Format(RS!codfaccl, "000000") & "   -> " & Format(RS!ImpVenci, FormatoImporte)
                End If
                I = I + 1
                colCtas.Add CStr(RS!codmacta)
                RS.MoveNext
            Wend
            RS.Close
            
            If Cad <> "" Then
            
            
                If Me.chkComensaAbonos.Value = 0 Then
                
                    If I >= 15 Then Cad = Cad & vbCrLf & "....  y " & I & " vencimientos más"
                    Cad = "Clientes con abonos. " & vbCrLf & Cad & " ¿Continuar?"
                    If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then
                        Set RS = Nothing
                        Set colCtas = Nothing
                        Exit Sub
                    End If
                            
                Else
                    '-------------------------------------------------------------------------
                    For I = 1 To colCtas.Count
                        CadenaDesdeOtroForm = colCtas.Item(I)
                        frmListado.Opcion = 36
                        frmListado.Show vbModal
                        
                    Next
                    CadenaDesdeOtroForm = ""
                    
                    'Actualice BD
                    Screen.MousePointer = vbHourglass
                    espera 1
                    Screen.MousePointer = vbHourglass
                    Conn.Execute "commit"
                    espera 1
                    
                End If
            End If 'colcount
        End If
        Set colCtas = Nothing
    End If
        
    
    'Que la cuenta NO este bloqueada
    I = 0
    If SubTipo = vbTipoPagoRemesa Then
        Cad = " FROM scobro,sforpa,cuentas WHERE scobro.codforpa = sforpa.codforpa AND (siturem is null) AND "
        Cad = Cad & " scobro.codmacta=cuentas.codmacta AND (not (fecbloq is null) and fecbloq < '" & Format(CDate(Text1(8).Text), FormatoFecha) & "') AND "
        Cad = "Select scobro.codmacta,nommacta,fecbloq" & Cad & SQL & " GROUP BY 1 ORDER BY 1"
        
    Else
        Cad = "select cuentas.codmacta,nommacta from "
        Cad = Cad & "scarecepdoc,cuentas where scarecepdoc.codmacta=cuentas.codmacta"
        Cad = Cad & " AND (not (fecbloq is null) and fecbloq < '" & Format(CDate(Text1(8).Text), FormatoFecha) & "') "
        Cad = Cad & " AND " & SQL & " GROUP by 1"
    End If
    
    
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        Cad = ""
        I = 1
        While Not RS.EOF
            Cad = Cad & RS!codmacta & " - " & RS!Nommacta & " : " & RS!FecBloq & vbCrLf
            RS.MoveNext
        Wend
    End If

    RS.Close
    
    If I > 0 Then
        Cad = "Las siguientes cuentas estan bloquedas." & vbCrLf & String(60, "-") & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Screen.MousePointer = vbDefault
        
        Exit Sub
    End If
    
    
    
    
    
    If SubTipo = vbTipoPagoRemesa Then
        'Efectos bancario
    
        Cad = " FROM scobro,sforpa,cuentas WHERE scobro.codforpa = sforpa.codforpa AND (siturem is null) AND "
        Cad = Cad & " scobro.codmacta=cuentas.codmacta AND "
    
    Else
    
        'Talon / Pagare
        Cad = " FROM scarecepdoc,cuentas where scarecepdoc.codmacta=cuentas.codmacta AND"
    End If
    'Hacemos un conteo
    RS.Open "SELECT Count(*) " & Cad & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        I = DBLet(RS.Fields(0), "N")
    End If
    RS.Close
    Cad = Cad & SQL
    
    
    
    If I > 0 Then
        If SubTipo <> vbTipoPagoRemesa Then
            'Para talones y pagares comprobaremos que
            'si esta configurado para contabilizar contra cta puente
            'entonces tiene la marca
            'PAGARE. Ver si tiene cta puente pagare
            If Me.cmbRemesa.ListIndex = 0 Then
                If Not vParam.PagaresCtaPuente Then I = 0
            Else
                If Not vParam.TalonesCtaPuente Then I = 0
            End If
            If I = 0 Then
                'NO contabilizaq contra cuenta puente
                
            Else
                'Comrpobaremos que todos los vtos estan en contabilizados.
                'Por eso la marca
                
                SQL = "(select numserie,codfaccl,fecfaccl,numorden " & Cad & ")"
                SQL = "select distinct(id) from slirecepdoc where (numserie,numfaccl,fecfaccl,numvenci) in " & SQL
                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                SQL = ""
                While Not RS.EOF
                    SQL = SQL & ", " & RS!Id
                    RS.MoveNext
                Wend
                RS.Close
                'Ya tengo el numero de las recepciones
                If SQL = "" Then
                    'ummmmmmmm, n deberia haber pasado
                    
                Else
                    SQL = "(" & Mid(SQL, 3) & ")"
                    SQL = "SELECT * from scarecepdoc where Contabilizada=0 and codigo in " & SQL
                    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                    SQL = ""
                    While Not RS.EOF
                        SQL = SQL & vbCrLf & Format(RS!Codigo, "0000") & "         " & RS!codmacta & "    " & Format(RS!fecharec, "dd/mm/yyyy") & "   " & RS!numeroref
                        RS.MoveNext
                    Wend
                    RS.Close
                    If SQL <> "" Then
                        'Hay taloes / pagares que estan recepcionados y o estan contabilizados
                        SQL = String(70, "-") & SQL
                        SQL = vbCrLf & "Codigo      Cuenta            Fecha         Referencia " & vbCrLf & SQL
                        SQL = "Hay talones / pagares que estan recepcionados pero no estan contabilizados" & vbCrLf & vbCrLf & SQL
                        MsgBox SQL, vbExclamation
                        Set RS = Nothing
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
                
            End If
        End If
        I = 1  'Para que siga por abajo
        
    End If
    
    

    'La suma
    If I > 0 Then
        SQL = "select sum(impvenci),sum(impcobro),sum(gastos) " & Cad
        Impor = 0
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then Impor = DBLet(RS.Fields(0), "N") - DBLet(RS.Fields(1), "N") + DBLet(RS.Fields(2), "N")
        RS.Close
        If Impor = 0 Then I = 0
    End If
        

    Set RS = Nothing
    
    If I = 0 Then
        MsgBox "Ningun dato a remesar con esos valores", vbExclamation
    Else
         
         
        'Preparamos algunas cosillas
        'Aqui guardaremos cuanto llevamos a cada banco
        SQL = "Delete from tmpCierre1 where codusu =" & vUsu.Codigo
        Conn.Execute SQL
        
        'Si son talones o pagares NO hay reajuste en bancos
        'Con lo cual cargare la tabla con el banco
        
        If SubTipo <> vbTipoPagoRemesa Then
            ' Metermos cta banco, nºremesa . El resto no necesito
            SQL = "INSERT INTO tmpcierre1 (codusu, cta, nomcta, acumPerD) VALUES ("
            SQL = SQL & vUsu.Codigo & ",'" & txtCta(3).Text & "','"
            'ANTES
            'SQL = SQL & DevNombreSQL(Me.txtDescCta(3).Text) & "'," & TransformaComasPuntos(CStr(Impor)) & ")"
            'AHora.
            SQL = SQL & txtRemesa.Text & "',0)"
            Conn.Execute SQL
        Else
            If Not chkAgruparRemesaPorEntidad.Visible Then Me.chkAgruparRemesaPorEntidad.Value = 0
            SQL = Cad 'Le paso el SELECT
            If Me.chkAgruparRemesaPorEntidad.Value = 1 Then DividiVencimentosPorEntidadBancaria
                                
        End If
        
        
        'Lo qu vamos a hacer es , primero bloquear la opcioin de remesar
        If BloqueoManual(True, "Remesas", "Remesas") Then
            
            Me.Visible = False
            
            If SubTipo = vbTipoPagoRemesa Then
                'REMESA NORMAL Y CORRIENTE
                'La de efectos de toda la vida
                'Mostraremos el otro form, el de remesas
                
                frmRemesas.Opcion = 0
                frmRemesas.vSQL = CStr(Cad)
                
                If chkAgruparRemesaPorEntidad.Value = 1 Then
                    Cad = txtCta(3).Text
                Else
                    Cad = ""
                End If
                Cad = txtRemesa.Text & "|" & Year(CDate(Text1(8).Text)) & "|" & Text1(8).Text & "|" & Cad & "|"
                frmRemesas.vRemesa = Cad
                
                frmRemesas.ImporteRemesa = Impor
                frmRemesas.Show vbModal

                
               
            Else
                'Remesas de talones y pagares
                frmRemeTalPag.vRemesa = "" 'NUEVA
                frmRemeTalPag.SQL = Cad
                frmRemeTalPag.Talon = cmbRemesa.ListIndex = 1 '0 pagare   1 talon
                frmRemeTalPag.Text1(0).Text = Me.txtCta(3).Text & " - " & txtDescCta(3).Text
                frmRemeTalPag.Text1(1).Text = Text1(8).Text
                frmRemeTalPag.Show vbModal
            End If
            'Desbloqueamos
            BloqueoManual False, "Remesas", ""
            Unload Me
        Else
            MsgBox "Otro usuario esta generando remesas", vbExclamation
        End If

    End If
    
    Screen.MousePointer = vbDefault
End Sub




Private Sub NuevaRemTalPag()
Dim CtaPuente As Boolean
Dim ForPa As String
Dim Cad As String
Dim Impor As Currency

'Algunas conideraciones

        'Para talones y pagares obligado la cuenta bancaria
        If txtCta(3).Text = "" Then
            MsgBox "Indique la cuenta bancaria", vbExclamation
            Exit Sub
        End If



    'Fecha remesa tiene k tener valor
    If Text1(8).Text = "" Then
        MsgBox "Fecha de remesa debe tener valor", vbExclamation
        Ponerfoco Text1(8)
        Exit Sub
    End If
    
    
    
    'VEMOS SI LA FECHA ESTA DENTRO DEL EJERCICIO
    If FechaCorrecta2(CDate(Text1(8).Text), True) > 1 Then Exit Sub
    
        'NO hago la pregunta. Si no tiene la cuenta puente dejo continuar igual
'        If Me.cmbRemesa.ListIndex = 0 Then
'            SQL = Abs(vParam.PagaresCtaPuente)
'        Else
'            SQL = Abs(vParam.TalonesCtaPuente)
'        End If
'        If SQL = "0" Then
'
'            MsgBox "Falta configurar la opción en parametros", vbExclamation
'            Exit Sub
'        End If
    
    If Me.cmbRemesa.ListIndex = 0 Then
        CtaPuente = vParam.PagaresCtaPuente
    Else
        CtaPuente = vParam.TalonesCtaPuente
    End If



    'A partir de la fecha generemos leemos k remesa corresponde
    SQL = "select max(codigo) from remesas where anyo=" & Year(CDate(Text1(8).Text))
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then
        NumRegElim = DBLet(miRsAux.Fields(0), "N")
    End If
    miRsAux.Close
    
    NumRegElim = NumRegElim + 1
    txtRemesa.Text = NumRegElim

    
    
        If Me.cmbRemesa.ListIndex = 0 Then
            SQL = " talon = 0"
        Else
            SQL = " talon = 1"
        End If
    
        'Si no lleva cuenta puente, no hace falta que este contabilizada
        'Es decir. Solo mirare contabilizados si llevo ctapuente
        If CtaPuente Then SQL = SQL & " AND contabilizada = 1 "
        SQL = SQL & " AND LlevadoBanco = 0 "
        
        'de la recepcion de factura
        If Text1(6).Text <> "" Then SQL = SQL & " AND fechavto >= '" & Format(Text1(6).Text, FormatoFecha) & "'"
        If Text1(7).Text <> "" Then SQL = SQL & " AND fechavto <= '" & Format(Text1(7).Text, FormatoFecha) & "'"

        
    
    
    
    
    
    
        
        
    
    
    
    
    
    
    
 
    
    
        'Fecha recepcion
        If Text1(22).Text <> "" Then SQL = SQL & " AND fecharec >= '" & Format(Text1(22).Text, FormatoFecha) & "'"
        If Text1(21).Text <> "" Then SQL = SQL & " AND fecharec <= '" & Format(Text1(21).Text, FormatoFecha) & "'"
    
    
     
    
    Screen.MousePointer = vbHourglass
    Set RS = New ADODB.Recordset
    
    'Que la cuenta NO este bloqueada
    I = 0
    Cad = "select cuentas.codmacta,nommacta,FecBloq from "
    Cad = Cad & "scarecepdoc,cuentas where scarecepdoc.codmacta=cuentas.codmacta"
    Cad = Cad & " AND (not (fecbloq is null) and fecbloq < '" & Format(CDate(Text1(8).Text), FormatoFecha) & "') "
    Cad = Cad & " AND " & SQL & " GROUP by 1"

    
    
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        Cad = ""
        I = 1
        While Not RS.EOF
            Cad = Cad & RS!codmacta & " - " & RS!Nommacta & " : " & RS!FecBloq & vbCrLf
            RS.MoveNext
        Wend
    End If

    RS.Close
    
    If I > 0 Then
        Cad = "Las siguientes cuentas estan bloquedas." & vbCrLf & String(60, "-") & vbCrLf & Cad
        MsgBox Cad, vbExclamation
        Screen.MousePointer = vbDefault
        
        Exit Sub
    End If
    

    Cad = " FROM scarecepdoc,cuentas where scarecepdoc.codmacta=cuentas.codmacta AND"

    'Hacemos un conteo
    RS.Open "SELECT * " & Cad & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    Cad = ""
    While Not RS.EOF
        I = I + 1
        Cad = Cad & " OR ( id = " & RS!Codigo & ") "
        RS.MoveNext
    Wend
    RS.Close
    If I = 0 Then
        MsgBox "Ningun dato con esos valores", vbExclamation
        Exit Sub
    End If
    Cad = "(" & Mid(Cad, 4) & ")"
    SQL = " from scobro where (numserie,codfaccl,fecfaccl,numorden) in (select numserie ,numfaccl,fecfaccl,numvenci from slirecepdoc where " & Cad & ")"
    SQL = "select sum(impvenci),sum(impcobro),sum(gastos) " & SQL
    
    
    

    'La suma
    If I > 0 Then
        
        Impor = 0
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'If Not Rs.EOF Then Impor = DBLet(Rs.Fields(0), "N") - DBLet(Rs.Fields(1), "N") + DBLet(Rs.Fields(2), "N")
        
        'Solo el impcobro
        If Not RS.EOF Then Impor = DBLet(RS.Fields(1), "N")
        RS.Close
        If Impor = 0 Then I = 0
    End If
        

    Set RS = Nothing
    
    If I = 0 Then
        MsgBox "Ningun dato a remesar con esos valores(II)", vbExclamation
    Else
         
         
        'Preparamos algunas cosillas
        'Aqui guardaremos cuanto llevamos a cada banco
        SQL = "Delete from tmpCierre1 where codusu =" & vUsu.Codigo
        Conn.Execute SQL
        
        'Si son talones o pagares NO hay reajuste en bancos
        'Con lo cual cargare la tabla con el banco
        
        If SubTipo <> vbTipoPagoRemesa Then
            ' Metermos cta banco, nºremesa . El resto no necesito
            SQL = "INSERT INTO tmpcierre1 (codusu, cta, nomcta, acumPerD) VALUES ("
            SQL = SQL & vUsu.Codigo & ",'" & txtCta(3).Text & "','"
            'ANTES
            'SQL = SQL & DevNombreSQL(Me.txtDescCta(3).Text) & "'," & TransformaComasPuntos(CStr(Impor)) & ")"
            'AHora.
            SQL = SQL & txtRemesa.Text & "',0)"
            Conn.Execute SQL
        End If
        
        
        'Lo qu vamos a hacer es , primero bloquear la opcioin de remesar
        If BloqueoManual(True, "Remesas", "Remesas") Then
            
            Me.Visible = False
            

            'Remesas de talones y pagares
            frmRemeTalPag.vRemesa = "" 'NUEVA
            frmRemeTalPag.SQL = Cad
            frmRemeTalPag.Talon = cmbRemesa.ListIndex = 1 '0 pagare   1 talon
            frmRemeTalPag.Text1(0).Text = Me.txtCta(3).Text & " - " & txtDescCta(3).Text
            frmRemeTalPag.Text1(1).Text = Text1(8).Text
            frmRemeTalPag.Show vbModal

            'Desbloqueamos
            BloqueoManual False, "Remesas", ""
            Unload Me
        Else
            MsgBox "Otro usuario esta generando remesas", vbExclamation
        End If

    End If
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CrearDisco()
Dim B As Boolean
Dim FCobro As String
    
    
        
    
    
        If cboTipoRemesa.ListIndex < 0 Then
            MsgBox "Seleccione la norma para la remesa", vbExclamation
            Exit Sub
        End If
    
        'El identificador REFERENCIA solo es valido para la norma 19
        If Me.cmbReferencia.ListIndex = 3 Then
            B = cboTipoRemesa.ListIndex = 0 Or cboTipoRemesa.ListIndex = 3
            If Not B Then
                MsgBox "Campo 'REFERENCIA EN EL VTO.' solo es válido para la norma 19", vbExclamation
                Exit Sub
            End If
        End If
                
                
        If Text1(9).Text = "" Then
            MsgBox "Fecha cobro en blanco", vbExclamation
            Exit Sub
        End If
        
        If Text1(18).Text = "" Then
            MsgBox "Fecha presentacion en blanco", vbExclamation
            Exit Sub
        End If
        
        
        
        
        If Trim(Text7(0).Text) = "" Then Text7(0).Text = UCase(vEmpresa.nomempre)
        
        
        B = False
        If cboTipoRemesa.ListIndex = 0 Or cboTipoRemesa.ListIndex = 3 Then
            
            
            FCobro = Text1(9).Text
            If optSepaXML(1).Value Then FCobro = ""  'Ha selccionado por vencimiento
        
            SQL = Mid(Text7(1).Text & "   ", 1, 3) & "|" & Mid(Text7(0).Text & Space(40), 1, 40) & "|"
            If GrabarDisketteNorma19(App.Path & "\tmpRem.ari", Text3(0).Text & "|" & Text3(1).Text & "|", Text1(18).Text, SQL, Me.cmbReferencia.ListIndex, FCobro, True, chkSEPA_GraboNIF(0).Value = 1, chkSEPA_GraboNIF(1).Value = 1, cboTipoRemesa.ListIndex = 3) Then
                SQL = App.Path & "\tmpRem.ari"
                'Copio el disquete
                B = CopiarArchivo
            End If
        Else
        
            If cboTipoRemesa.ListIndex = 1 Then
                'NORMA 32
                If GrabarDisketteNorma32(App.Path & "\tmpRem32.ari", Text3(0).Text & "|" & Text3(1).Text & "|" & Text7(1).Text & "|", Text1(9).Text) Then
                    SQL = App.Path & "\tmpRem32.ari"
                    'Copio el disquete
                    B = CopiarArchivo
                End If
                
            Else
                'NORMA 58
                SQL = Mid(Text7(1).Text & "   ", 1, 3) & "|" & Mid(Text7(0).Text & Space(40), 1, 40) & "|"
                If GrabarDisketteNorma58(App.Path & "\tmpRem58.ari", Text3(0).Text & "|" & Text3(1).Text & "|", Text1(18).Text, SQL, Me.cmbReferencia.ListIndex, CDate(Text1(9).Text)) Then
                    SQL = App.Path & "\tmpRem58.ari"
                    'Copio el disquete
                    B = CopiarArchivo
                End If
                
                
                
            End If
        End If
        
        
        
        If B Then
            MsgBox "Fichero creado con exito", vbInformation
            If Text3(2).Text = "A" Or Text3(2).Text = "B" Then
                'Cambio la situacion de la remesa
                SQL = "UPDATE Remesas SET fecremesa = '" & Format(Text1(9).Text, FormatoFecha)
                SQL = SQL & "' , tipo = " & cboTipoRemesa.ListIndex & ", Situacion = 'B'"
                SQL = SQL & " WHERE codigo=" & Text3(0).Text
                SQL = SQL & " AND anyo =" & Text3(1).Text
                If Ejecuta(SQL) Then CadenaDesdeOtroForm = "OK"
                
                
                
                If CadenaDesdeOtroForm = "OK" Then
                
                    Set miRsAux = New ADODB.Recordset
                    If Not UpdatearCobrosRemesa Then MsgBox "Error updateando cobros remesa", vbExclamation
                    Set miRsAux = Nothing
                End If
                
            End If
            
        End If
        
        
        
        
        
End Sub


Private Function UpdatearCobrosRemesa() As Boolean
Dim Im As Currency
    On Error GoTo EUpdatearCobrosRemesa
    UpdatearCobrosRemesa = False
    
    SQL = "Select * from scobro WHERE codrem=" & Text3(0).Text
    SQL = SQL & " AND anyorem =" & Text3(1).Text
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
            While Not miRsAux.EOF
                SQL = "UPDATE scobro SET fecultco = '" & Format(Text1(9).Text, FormatoFecha) & "', impcobro = "
                Im = miRsAux!ImpVenci
                If Not IsNull(miRsAux!Gastos) Then Im = Im + miRsAux!Gastos
                SQL = SQL & TransformaComasPuntos(CStr(Im))
                
                SQL = SQL & " ,siturem = 'B'"
                
                
                'WHERE
                SQL = SQL & " WHERE numserie='" & miRsAux!NUmSerie
                SQL = SQL & "' AND  codfaccl =  " & miRsAux!codfaccl
                SQL = SQL & "  AND  fecfaccl =  '" & Format(miRsAux!fecfaccl, FormatoFecha)
                SQL = SQL & "' AND  numorden =  " & miRsAux!numorden
                'Muevo siguiente
                miRsAux.MoveNext
                
                'Ejecuto SQL
                If Not Ejecuta(SQL) Then MsgBox "Error: " & SQL, vbExclamation
            Wend
    End If
    miRsAux.Close
                    
                    
                    
    UpdatearCobrosRemesa = True
    Exit Function
EUpdatearCobrosRemesa:
    
End Function

Private Function SugerirCodigoSiguienteTransferencia() As String
    
    SQL = "Select Max(codigo) from stransfer"
    If SubTipo = 0 Then SQL = SQL & "cob"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, , , adCmdText
    SQL = "1"
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            SQL = CStr(RS.Fields(0) + 1)
        End If
    End If
    RS.Close
    Set RS = Nothing
    SugerirCodigoSiguienteTransferencia = SQL
End Function




Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer
Dim W As Integer
    Limpiar Me
    PrimeraVez = True
    Me.Icon = frmPpal.Icon
    
    
    'Cago los iconos
    CargaImagenesAyudas Me.imgCtaNorma, 1, "Seleccionar cuenta"
    CargaImagenesAyudas Me.Image1, 2
    CargaImagenesAyudas imgCuentas, 1, "Cuenta contable banco"
    CargaImagenesAyudas imgRem, 1, "Seleccionar remesa"
    CargaImagenesAyudas imgFP, 1, "Seleccionar Forma de pago"
    CargaImagenesAyudas imgConcepto, 1, "Concepto"
    CargaImagenesAyudas ImageAyuda, 3
    
    Carga1ImagenAyuda Me.Image4, 1
    Carga1ImagenAyuda Me.Image3, 1
    Carga1ImagenAyuda imgFra, 1
    
    
    FrameContabilRem2.Visible = False
    
    Select Case Opcion
    Case 8, 22, 23
        'Utilizare el mismo FRAM para
        '   8.- Contabilizar / Abono remesa
        '   22- Cancelacion cliente
        '   23- Confirmacion remesa
        '  TANTO DE EFECTOS como de talones pagares
        FrameContabilRem2.Visible = True
        
        Caption = "Remesas"
        If SubTipo = 1 Then
            Caption = Caption & " EFECTOS"
        Else
            Caption = Caption & " talones/pagarés"
        End If
        chkAgrupaCancelacion.Visible = False
        
        If Opcion = 8 Then
            SQL = "Abono remesa"
            CuentasCC = "Contabilizar"
        Else
        
            If Opcion = 22 Then
            
                SQL = DevuelveDesdeBD("RemesaCancelacion", "paramtesor", "codigo", "1", "N")
                chkAgrupaCancelacion.Visible = Len(SQL) = vEmpresa.DigitosUltimoNivel
                SQL = "Cancelacion cliente"
                CuentasCC = "Can. cliente"
            Else
                SQL = "Confirmacion remesa"
                CuentasCC = "Confirmar"
            End If
            
        End If
        Label5(2).Caption = SQL
        cmdContabRemesa.Caption = CuentasCC
        
        If Opcion = 8 Then
            Me.Caption = "Abono remesa"
            Label5(2).Caption = "Remesa : " & RecuperaValor(NumeroDocumento, 1) & "/" & RecuperaValor(NumeroDocumento, 2) & " Banco : " & RecuperaValor(NumeroDocumento, 4) & " de " & RecuperaValor(NumeroDocumento, 5)
        End If
        
        CuentasCC = ""
        'Los gastos solo van en la contabilizacion
        Label3(0).Visible = Opcion = 8
        txtImporte(0).Visible = Opcion = 8
        
        'noviembre 2009
        'Opcion 8. Contabilizar(ABONO)
        ' tipo  efectos
        ' si tiene cta efectos comerciales descontados y es de ultimo nivel
        ' mostrar el agrupar efectos comerciales descontad
        ' DEBERIA IR AQUI el check visible o no.
        'Veremos si hay que ponerlo o no
        
        
        W = FrameContabilRem2.Width
        h = FrameContabilRem2.Height
    End Select
    
    
    Me.Height = h + 360
    Me.Width = W + 90
    
    h = Opcion
    Me.cmdCancelar(h).Cancel = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NumeroDocumento = "" 'Para reestrablecerlo siempre
End Sub


Private Sub frmBa_DatoSeleccionado(CadenaSeleccion As String)
    I = CInt(imgCuentas(0).Tag)
    Me.txtCta(I).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescCta(I).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(Image1(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQL = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
    txtFP(I).Text = RecuperaValor(CadenaSeleccion, 1)
    txtFPDesc(I).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmRe_DatoSeleccionado(CadenaSeleccion As String)
    If I = 0 Then
        Text3(3).Text = RecuperaValor(CadenaSeleccion, 1)
        Text3(4).Text = RecuperaValor(CadenaSeleccion, 2)
        Text1(10).Text = RecuperaValor(CadenaSeleccion, 3)
    Else
        'DEVOLUCIOIN
        Text3(5).Text = RecuperaValor(CadenaSeleccion, 1)
        Text3(6).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
    
End Sub

Private Sub Image1_Click(Index As Integer)
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text1(Index).Text <> "" Then frmC.Fecha = CDate(Text1(Index).Text)
    Image1(0).Tag = Index
    frmC.Show vbModal
    Set frmC = Nothing
    If Text1(Index).Text <> "" Then Ponerfoco Text1(Index)
End Sub


Private Sub Ponerfoco(ByRef O As Object)
    On Error Resume Next
    O.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub ObtenerFoco(ByRef T As TextBox)
    T.SelStart = 0
    T.SelLength = Len(T.Text)
End Sub

Private Sub KEYpress(ByRef Tecla As Integer)
    If Tecla = 13 Then
        Tecla = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Image3_Click()
        Set frmCCtas = New frmColCtas
        SQL = ""
        frmCCtas.DatosADevolverBusqueda = "0"
        frmCCtas.Show vbModal
        Set frmCCtas = Nothing
        If SQL <> "" Then
            'TEngo cuenta contable
            Text5.Text = SQL
            SQL = "nommacta"
            Text4.Text = DevuelveDesdeBD("nifdatos", "cuentas", "codmacta", Text5.Text, "T", SQL)
            If Text4.Text = "" Then
                Text5.Text = ""
                MsgBox "La cuenta no tiene NIF.", vbExclamation
            Else
                Text5.Text = SQL
            End If
        End If

End Sub

Private Sub Image4_Click()
    SQL = ""
    cd1.ShowOpen
    If cd1.FileName <> "" Then SQL = cd1.FileName
    If SQL <> "" Then
        If Dir(SQL, vbArchive) = "" Then
            MsgBox "Fichero NO existe", vbExclamation
            SQL = ""
        End If
    End If
    If SQL <> "" Then Text8.Text = SQL
End Sub

Private Sub ImageAyuda_Click(Index As Integer)
    
    Select Case Index
    Case 0
        If vParam.NuevasNormasSEPA Then

            SQL = "Adeudos directos SEPA" & vbCrLf & vbCrLf & vbCrLf
            SQL = SQL & " Según la fecha seleccionada girará los vencimientos de la remesa:"
            
            SQL = SQL & vbCrLf & " Cobro.  Todos los vencimientos a esa fecha"
            SQL = SQL & vbCrLf & " Vencimiento.  Cada uno con su fecha"

        Else
            SQL = "Generacion antigua N19"
        End If
    End Select
    MsgBox ImageAyuda(Index).ToolTipText & vbCrLf & SQL, vbInformation
End Sub

Private Sub imgCC_Click(Index As Integer)
    LanzaBuscaGrid 2
    If SQL <> "" Then
        txtCC(Index).Text = SQL
        txtCC_LostFocus Index
    End If
End Sub

Private Sub imgCheck_Click(Index As Integer)

    If Index < 2 Then
        'Selecciona forma de pago
        For I = 1 To Me.lwtipopago.ListItems.Count
            Me.lwtipopago.ListItems(I).Checked = Index = 1
        Next

    ElseIf Index < 4 Then
        'Empresas
         For I = 1 To Me.ListView3.ListItems.Count
            Me.ListView3.ListItems(I).Checked = Index = 3
        Next
    Else
        'Reclamaciones
        If Me.optReclama(1).Value Then
            'Solo en correctos, los incorrectos se iran tooodos
            For I = 1 To Me.ListView6.ListItems.Count
                Me.ListView6.ListItems(I).Checked = Index = 5
            Next
        End If
    End If
End Sub

Private Sub imgcheckall_Click(Index As Integer)
    Cancelado = (Index = 0)
    For I = 1 To ListView4.ListItems.Count
        ListView4.ListItems(I).Checked = Cancelado
    Next I
    Cancelado = False
End Sub

Private Sub imgConcepto_Click(Index As Integer)
  
    LanzaBuscaGrid 1
    If SQL <> "" Then
        txtConcepto(Index).Text = SQL
        txtConcepto_LostFocus Index
    End If
End Sub

Private Sub imgCtaNorma_Click(Index As Integer)

        If Index <> 6 Then

               Set frmCCtas = New frmColCtas
               SQL = ""
               frmCCtas.DatosADevolverBusqueda = "0"
               frmCCtas.Show vbModal
               
               Set frmCCtas = Nothing
               If SQL <> "" Then
                   txtCtaNormal(Index).Text = SQL
                   txtCtaNormal_LostFocus Index
               End If
            
        Else
        
            'Para las cuentas agrupadas
            LanzaBuscaGrid 3
            If SQL <> "" Then
                If MsgBox("Va a insetar las cuentas del grupo de tesoreria: " & SQL & vbCrLf & "¿Continuar?", vbQuestion + vbYesNo) = vbYes Then
                    Screen.MousePointer = vbHourglass
                    Set miRsAux = New ADODB.Recordset
                    CargaGrupo
                    Set miRsAux = Nothing
                    Screen.MousePointer = vbDefault
                End If
            End If
        End If
            
            
End Sub

Private Sub imgCuentas_Click(Index As Integer)

    imgCuentas(0).Tag = Index
    Set frmBa = New frmBanco
    frmBa.DatosADevolverBusqueda = "OK"
    frmBa.Show vbModal
    Set frmBa = Nothing
End Sub


Private Sub imgDiario_Click(Index As Integer)
  
    LanzaBuscaGrid 0
    If SQL <> "" Then
        txtDiario(Index).Text = SQL
        txtDiario_LostFocus Index
    End If
End Sub

Private Sub imgEliminarCta_Click()
    If List1.SelCount = 0 Then Exit Sub
    
    SQL = "Desea quitar la(s) cuenta(s): " & vbCrLf
    For I = 0 To List1.ListCount - 1
        If List1.Selected(I) Then SQL = SQL & List1.List(I) & vbCrLf
    Next I
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        For I = List1.ListCount - 1 To 0 Step -1
            If List1.Selected(I) Then
                SQL = Trim(Mid(List1.List(I), 1, vEmpresa.DigitosUltimoNivel + 2))
                NumRegElim = InStr(1, CuentasCC, SQL)
                If NumRegElim > 0 Then CuentasCC = Mid(CuentasCC, 1, NumRegElim - 1) & Mid(CuentasCC, NumRegElim + vEmpresa.DigitosUltimoNivel + 1) 'para que quite el pipe final
                List1.RemoveItem I
            End If
        Next I
    
    End If
    NumRegElim = 0
End Sub

Private Sub imgFP_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmP = New frmFormaPago
    I = Index
    frmP.DatosADevolverBusqueda = "0|1"
    frmP.Show vbModal
    Set frmP = Nothing
End Sub

Private Sub imgFra_Click()
        CadenaDesdeOtroForm = ""
        SQL = ""
        If txtCtaNormal(11).Text <> "" Then SQL = "scobro.codmacta = '" & txtCtaNormal(11).Text & "'"
        frmVerCobrosPagos.vSQL = SQL
        frmVerCobrosPagos.OrdenarEfecto = False
        frmVerCobrosPagos.Regresar = True
        frmVerCobrosPagos.Cobros = True
        frmVerCobrosPagos.Show vbModal
        If CadenaDesdeOtroForm <> "" Then

            txtSerie(4).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            txtNumFac(4).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            Me.txtNumero.Text = RecuperaValor(CadenaDesdeOtroForm, 4)
            Ponerfoco Text1(11)
        End If
        CadenaDesdeOtroForm = ""
End Sub

Private Sub imgRem_Click(Index As Integer)
    I = Index
    Set frmRe = New frmColRemesas2
    frmRe.Tipo = SubTipo  'Para abrir efectos o talonesypagares
    frmRe.DatosADevolverBusqueda = "1|"
    frmRe.Show vbModal
    Set frmRe = Nothing
    'Por si ha puesto los datos
    CamposRemesaAbono
    
End Sub



Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If Shift And vbCtrlMask > 0 Then
            MsgBox "HOLITA VECINO. Has encontrado el huevo de pascua...., a curraaaaaarrr!!!!", vbExclamation
        End If
    End If
End Sub


Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).Text = "" Then Exit Sub
    
    If Not EsFechaOK(Text1(Index)) Then
        MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
        Text1(Index).Text = ""
        Ponerfoco Text1(Index)
    End If
    
End Sub



Private Sub CargaList()
    


        SQL = DevuelveDesdeBD("descformapago", "stipoformapago", "tipoformapago", CStr(SubTipo), "N")
        Text2(Opcion).Text = SQL
                
        
End Sub


Private Sub Text3_GotFocus(Index As Integer)
    With Text3(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Text3_LostFocus(Index As Integer)
    With Text3(Index)
        .Text = Trim(.Text)
        If .Text = "" Then Exit Sub
        
        If Not IsNumeric(.Text) Then
            MsgBox "Campo debe ser numérico: " & .Text, vbExclamation
            .Text = ""
            Ponerfoco Text3(Index)
        End If
        
        'Para que vaya a la tabal y traiga datos remesa
        If Index = 3 Or Index = 4 Then CamposRemesaAbono
    End With
End Sub


Private Sub PonerValoresDefectoRemesas()
Dim F As Date
    
    'Fecha remesa.. hoy
    Text1(8).Text = Format(Now, "dd/mm/yyyy")
    
    'Tipo. Por defecto siempre efecto
    Me.cmbRemesa.ListIndex = 0
    
    'Ahora vemos la fecha mas alta de remesas
    SQL = "select max(fecfin) from remesas "
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    F = CDate("01/01/2000")
    NumRegElim = 0
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            F = miRsAux.Fields(0)
            NumRegElim = 1
        End If
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    If NumRegElim = 0 Then
        Text1(6).Text = ""
    Else
        Text1(6).Text = Format(F, "dd/mm/yyyy")
    End If
    Text1(7).Text = Format(DateAdd("d", -1, Now), "dd/mm/yyyy")
End Sub

Private Function CopiarArchivo() As Boolean
On Error GoTo ECopiarArchivo

    CopiarArchivo = False
    'cd1.CancelError = True
    cd1.FileName = ""
    cd1.ShowSave
    If cd1.FileName <> "" Then
    
        If Dir(cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo " & cd1.FileName & " ya existe" & vbCrLf & vbCrLf & "¿Sobreescribir?", vbQuestion + vbYesNo) = vbNo Then Exit Function
            Kill cd1.FileName
        End If
        'Hacemos la copia
        FileCopy SQL, cd1.FileName
        CopiarArchivo = True
    End If
    
    
    Exit Function
ECopiarArchivo:
    MuestraError Err.Number, "Copiar Archivo"
End Function


Private Sub txtImporte_GotFocus(Index As Integer)
    With txtImporte(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
 Dim Valor
        txtImporte(Index).Text = Trim(txtImporte(Index))
        If txtImporte(Index).Text = "" Then Exit Sub
        

        If Not EsNumerico(txtImporte(Index).Text) Then
            txtImporte(Index).Text = ""
            Exit Sub
        End If
    
        
        If Index = 6 Or Index = 7 Then
           
            If InStr(1, txtImporte(Index).Text, ",") > 0 Then
                Valor = ImporteFormateado(txtImporte(Index).Text)
            Else
                Valor = CCur(TransformaPuntosComas(txtImporte(Index).Text))
            End If
            txtImporte(Index).Text = Format(Valor, FormatoImporte)
        End If
        
End Sub

Private Sub CargaImpagados()

    SQL = "Select fechadev,gastodev from sefecdev  WHERE numserie='" & RecuperaValor(CadenaDesdeOtroForm, 1)
    SQL = SQL & "' AND  codfaccl =  " & RecuperaValor(CadenaDesdeOtroForm, 2)
    SQL = SQL & "  AND  fecfaccl =  '" & Format(RecuperaValor(CadenaDesdeOtroForm, 3), FormatoFecha)
    SQL = SQL & "' AND  numorden =  " & RecuperaValor(CadenaDesdeOtroForm, 4)
    SQL = SQL & " order by fechadev"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set IT = ListView1.ListItems.Add
        IT.Text = Format(RS!fechadev, "dd/mm/yyyy")
        IT.SubItems(1) = Format(RS!gastodev, FormatoImporte)
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub


Private Sub CargaImagen()
On Error Resume Next
    Image2.Picture = LoadPicture(App.Path & "\minilogo.bmp")
    'Image1.Picture = LoadPicture(App.path & "\fondon.gif")
    Err.Clear
End Sub


Private Sub CargaRemesas()
    
    ListView2.ListItems.Clear
    
    If SubTipo > 2 Then
        CargaRemes 3  'Cargamos todo
        CargaRemes 2  'Cargamos todo
    Else
        CargaRemes SubTipo
    End If
    
    
End Sub


Private Sub CargaRemes(SubT As Byte)
Dim F As Date
Dim Dias As Integer

    On Error GoTo EC
    
    
    
    
 
    ' 3 es que esta cargando todo
    If SubT = 1 Or SubT = 3 Then
        'Efectos
        '
        SQL = "Select codigo,anyo, fecremesa,"
        If SubT = 3 Then
            SQL = SQL & " tiporemesa2.descripciont "
        Else
            SQL = SQL & " tiporemesa."
        End If
        SQL = SQL & "descripcion,descsituacion,remesas.codmacta,nommacta,remesadiasmenor, remesadiasmayor, "
        SQL = SQL & "Importe , remesas.descripcion as Desc1, remesas.Tipo,situacion,Tiporem from cuentas,tiposituacionrem,ctabancaria,"
        SQL = SQL & "remesas left join tiporemesa"
        If SubT = 3 Then SQL = SQL & "2" 'Para que carge, en lugar de norma19 norma52 etc que carge efectos, talon, pagare
        SQL = SQL & " on remesas.tipo"
        If SubT = 3 Then SQL = SQL & "rem"
        SQL = SQL & "=tiporemesa"
        If SubT = 3 Then SQL = SQL & "2" 'Para que carge, en lugar de norma19 norma52 etc que carge efectos, talon, pagare
        SQL = SQL & ".tipo where remesas.codmacta=cuentas.codmacta and situacio=situacion and ctabancaria.codmacta=remesas.codmacta"
        SQL = SQL & " AND tiporem = 1 "   'Efectos
        'Solo borrare las contabilizadas o pendientes de eliminar tooodos los efectos
        SQL = SQL & " AND (situacion ='Q' or situacion ='Y')"
                
        
    Else
        'Talones Remesesas
        SQL = "Select codigo,anyo, fecremesa,tiporemesa2.descripciont descripcion,descsituacion,remesas.codmacta,nommacta,talondias,pagaredias, "
        SQL = SQL & "Importe , remesas.descripcion as Desc1, remesas.Tipo,situacion,Tiporem from cuentas,tiposituacionrem,ctabancaria,"
        SQL = SQL & "remesas left join tiporemesa2 on remesas.tiporem=tiporemesa2.tipo "
        SQL = SQL & "where remesas.codmacta=cuentas.codmacta and situacio=situacion and ctabancaria.codmacta=remesas.codmacta"
        SQL = SQL & " AND tiporem > 1 "   'Pagares remesas
       'Solo borrare las contabilizadas o pendientes de eliminar tooodos los efectos
        SQL = SQL & " AND (situacion ='Q' or situacion ='Y')"
    
    End If
    
    SQL = SQL & " ORDER BY anyo,codigo"   'Solo borrare las contabilizadas
    Set RS = New ADODB.Recordset
    
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        'Ver los dias
        If SubT = 1 Or SubT = 3 Then
            'Efectos recibos
            Dias = DBLet(RS!remesadiasmenor, "N")
            I = DBLet(RS!remesadiasmayor, "N")
            If I < Dias And I > 0 Then Dias = I
        Else
            If RS!Tiporem = 2 Then
                'Pagare
                Dias = DBLet(RS!pagaredias, "N")
            Else
                'talon
                Dias = DBLet(RS!talondias, "N")
            End If
            
        End If
        F = RS!fecremesa
        
        If SubT = 2 Then
            'If RS!Codigo > 159 Then Stop
            SQL = "anyorem=" & RS!Anyo & " AND codrem "
            SQL = DevuelveDesdeBD("min(fecvenci)", "scobro", SQL, RS!Codigo, "N")
            If SQL <> "" Then
                If CDate(SQL) > F Then F = SQL
            End If
        End If
        
        F = DateAdd("d", Dias, F)
        If F < Now Then
            Set IT = ListView2.ListItems.Add
            IT.Text = RS!Anyo
            IT.SubItems(1) = RS!Codigo
            IT.SubItems(2) = RS!Descripcion
            IT.SubItems(3) = RS!fecremesa
            IT.SubItems(4) = RS!codmacta
            IT.SubItems(5) = RS!Nommacta
            IT.SubItems(6) = Format(RS!Importe, FormatoImporte)
            IT.SubItems(7) = RS!Desc1
            IT.Tag = RS!Tiporem
        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    Exit Sub
EC:
    MuestraError Err.Number, "Cargando vencimientos"
End Sub



'
Public Function GeneraCobrosPagosNIF() As Boolean
Dim Cad As String
Dim L As Long
Dim Empre As String
Dim Importe  As Currency

Dim QueTipoPago As String

    'Guardaremos en la variable QueTipoPago que tipos de pago ha seleccionado
    'Si selecciona todos los tipos de pago NO pondremos el IN en el select
    QueTipoPago = ""
    Cad = "" 'para saber si ha selccionado todos
    For L = 1 To Me.lwtipopago.ListItems.Count
        If lwtipopago.ListItems(L).Checked Then
            QueTipoPago = QueTipoPago & ", " & Me.lwtipopago.ListItems(L).Tag
        Else
            Cad = "NO" 'No estan todos seleccionados
        End If
    Next
    If Cad = "" Then
        'Estan todos. No tiene sentido hacer el Select in
        QueTipoPago = ""
    Else
        QueTipoPago = Mid(QueTipoPago, 2)
    End If
    
    
    
'En la tabla  INSERT INTO tmp347 (codusu, cliprov, cta, nif) VALUES ((
' Tendremos codccos: la empresa
'                  : cta, cada uno de los valores
'INSERT INTO ztesoreriacomun (codusu, codigo, texto1, texto2, texto3, texto4,
'texto5, texto6, importe1, importe2, fecha1, fecha2, fecha3, observa1,
'observa2, opcion) VALUES
    GeneraCobrosPagosNIF = False
    L = 1
    SQL = "Select * from tmp347 where codusu =" & vUsu.Codigo & " ORDER BY cliprov,cta"
    Set RS = New ADODB.Recordset
    Set miRsAux = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not RS.EOF
        If Cancelado Then
            RS.Close
            Exit Function
        End If
        'Los labels
        Label9.Caption = "Nif: " & RS!NIF & " - " & RS!Cta
        Label9.Refresh
        
        'SQL insert
        SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu,texto1, codigo,texto2,  texto3,texto4, texto5,fecha1,fecha2,"   'texto5, texto6,
        SQL = SQL & " importe1, importe2,opcion"
        SQL = SQL & ") VALUES ("
        'NIF      Nombre
        SQL = SQL & vUsu.Codigo & ",'" & RS!NIF & "',"
        
        
        '-------
        Empre = DameEmpresa(CStr(RS!cliprov))
        
        'COBROS
        Cad = "Select fecfaccl,numserie,codfaccl, numorden,impvenci,impcobro,gastos,fecvenci,nommacta from conta" & RS!cliprov & ".scobro as c1,"
        Cad = Cad & "conta" & RS!cliprov & ".cuentas as c2 "
        If QueTipoPago <> "" Then Cad = Cad & ", conta" & RS!cliprov & ".sforpa as sforpa"
        Cad = Cad & " where c1.codmacta = c2.codmacta AND c1.codmacta='" & RS!Cta & "'"
        If QueTipoPago <> "" Then Cad = Cad & " AND c1.codforpa=sforpa.codforpa AND sforpa.tipforpa in (" & QueTipoPago & ")"
        'Fechas
        If Text1(12).Text <> "" Then Cad = Cad & " AND fecvenci >='" & Format(Text1(12).Text, FormatoFecha) & "'"
        If Text1(13).Text <> "" Then Cad = Cad & " AND fecvenci <='" & Format(Text1(13).Text, FormatoFecha) & "'"
        
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            'Los label
            If Cancelado Then
                miRsAux.Close
                Exit Function
            End If
            
            'Insetamos codigo,  texto3
            '                    empresa
            Cad = L & ",'" & Empre & "','"
            Cad = Cad & miRsAux!NUmSerie & "/" & Format(miRsAux!codfaccl, "0000000000") & " : " & miRsAux!numorden & "','"
            Cad = Cad & RS!Cta & "','"
            Cad = Cad & DevNombreSQL(miRsAux!Nommacta) & "','"
            'texto4: fecha
            Cad = Cad & Format(miRsAux!fecfaccl, FormatoFecha) & "','"
            Cad = Cad & Format(miRsAux!FecVenci, FormatoFecha) & "',"
            
            
            'En importe1 estara el importe del cobro. En el 2 tb
'            Importe = DBLet(miRsAux!Gastos, "N") - DBLet(miRsAux!impcobro, "N")
'            Importe = Importe + miRsAux!impvenci
'             Cad = Cad & TransformaComasPuntos(CStr(Importe)) & "," & TransformaComasPuntos(CStr(Importe))


            Importe = DBLet(miRsAux!Gastos, "N")
            Cad = Cad & TransformaComasPuntos(CStr(Importe))
            Importe = miRsAux!ImpVenci - DBLet(miRsAux!impcobro, "N")
            Cad = Cad & "," & TransformaComasPuntos(CStr(Importe))
           
            
            
            'un cero para importe 2  y un cero para la opcion
            Cad = Cad & ",0)"
            
            'Ejecutamos
            Cad = SQL & Cad
            Ejecuta Cad
            
            L = L + 1
            miRsAux.MoveNext
            DoEvents
        Wend
        miRsAux.Close
        
        'PAGOS
        Cad = "Select numfactu,numorden,fecfactu,imppagad,fecefect,impefect,nommacta from conta" & RS!cliprov & ".spagop ,conta" & RS!cliprov & ".cuentas "
        If QueTipoPago <> "" Then Cad = Cad & ", conta" & RS!cliprov & ".sforpa as sforpa"
        Cad = Cad & " where ctaprove = codmacta AND ctaprove='" & RS!Cta & "'"
        If QueTipoPago <> "" Then Cad = Cad & " AND spagop.codforpa=sforpa.codforpa AND sforpa.tipforpa in (" & QueTipoPago & ")"
        
        
        'Fechas
        If Text1(12).Text <> "" Then Cad = Cad & " AND fecefect >='" & Format(Text1(12).Text, FormatoFecha) & "'"
        If Text1(13).Text <> "" Then Cad = Cad & " AND fecefect <='" & Format(Text1(13).Text, FormatoFecha) & "'"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            'Los label
            If Cancelado Then
                miRsAux.Close
                Exit Function
            End If
            
            'Insetamos codigo,  texto3,t5
            '                    empresa
            Cad = L & ",'" & Empre & "','"
            Cad = Cad & DevNombreSQL(miRsAux!NumFactu) & " : " & miRsAux!numorden & "','"
            Cad = Cad & RS!Cta & "','"
            Cad = Cad & DevNombreSQL(miRsAux!Nommacta) & "','"
            ' fecha1 y 2
            Cad = Cad & Format(miRsAux!FecFactu, FormatoFecha) & "','"
            Cad = Cad & Format(miRsAux!fecefect, FormatoFecha) & "',"
            
            
            'En importe1 estara el importe del cobro
            Importe = DBLet(miRsAux!imppagad, "N")

            Importe = miRsAux!ImpEfect - Importe
            Cad = Cad & TransformaComasPuntos(CStr(0)) & "," & TransformaComasPuntos(CStr(-1 * Importe))
            
            Cad = Cad & ",1)" '1: pago
            
            
            
            
            'Ejecutamos
            Cad = SQL & Cad
            Ejecuta Cad
            
            L = L + 1
            miRsAux.MoveNext
            
            DoEvents
        Wend
        miRsAux.Close
        
        
        'SIGUIENTE CUENTA
        RS.MoveNext
    Wend
    RS.Close
    
    Cad = "DELETE FROM usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo & " AND importe1+importe2=0"
    Conn.Execute Cad
    
    Cad = "select count(*) from usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    L = 0
    If Not RS.EOF Then
        L = DBLet(RS.Fields(0), "N")
    End If
    RS.Close
    
    Set RS = Nothing
    Set miRsAux = Nothing
    
    If L = 0 Then
        'ERROR. MO HAY DATOS
        MsgBox "Sin datos.", vbExclamation
    Else
        GeneraCobrosPagosNIF = True
    End If
        
End Function



Private Function DameEmpresa(ByVal S As String) As String
    DameEmpresa = "NO ENCONTRADA"
    For I = 1 To ListView3.ListItems.Count
        If ListView3.ListItems(I).Tag = S Then
            DameEmpresa = DevNombreSQL(ListView3.ListItems(I).Text)
            Exit For
        End If
    Next I
    
End Function






Private Sub cargaTipoPagos()
    'FALTARA VER LO DE QUITAR EMPRESAS NO PERMITIDAS
 
    lwtipopago.ListItems.Clear
    SQL = "select tipoformapago,descformapago,siglas from stipoformapago order by tipoformapago"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwtipopago.ListItems.Add
        IT.Key = "C" & miRsAux!tipoformapago
        IT.Text = miRsAux!descformapago
      '  IT.SubItems(1) = miRsAux!siglas
        IT.Tag = miRsAux!tipoformapago
        
        If miRsAux!tipoformapago > 0 Then IT.Checked = True  'menos el efectivo  todas
         
        miRsAux.MoveNext
        
    Wend
    miRsAux.Close
    Set miRsAux = New ADODB.Recordset
End Sub



Private Sub CargaCtasparaAgruparNIF()
    I = 0
    SQL = "select cuentas.codmacta,nifdatos from scobro,cuentas where scobro.codmacta=cuentas.codmacta"
    SQL = SQL & " and not (nifdatos is null)  "
    If txtCtaNormal(1).Text <> "" Then SQL = SQL & " and cuentas.codmacta >= '" & txtCtaNormal(1).Text & "'"
    If txtCtaNormal(2).Text <> "" Then SQL = SQL & " and cuentas.codmacta <= '" & txtCtaNormal(2).Text & "'"
    SQL = SQL & " group by  codmacta,nifdatos"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF
        If Cancelado Then
            miRsAux.Close
            Exit Sub
        End If
        SQL = "INSERT INTO tmpfaclin (codusu, codigo, NIF) VALUES (" & vUsu.Codigo & "," & I & ",'" & miRsAux!nifdatos & "')"
        Ejecuta SQL
        miRsAux.MoveNext
        DoEvents
        I = I + 1
    Wend
    miRsAux.Close
    If Cancelado Then Exit Sub
    'AHora los nifs en los pagos
    SQL = "select cuentas.codmacta,nifdatos from spagop,cuentas where ctaprove=cuentas.codmacta"
    SQL = SQL & " and not (nifdatos is null) "
    If txtCtaNormal(1).Text <> "" Then SQL = SQL & " and cuentas.codmacta >= '" & txtCtaNormal(1).Text & "'"
    If txtCtaNormal(2).Text <> "" Then SQL = SQL & " and cuentas.codmacta <= '" & txtCtaNormal(2).Text & "'"
    
    SQL = SQL & " group by  codmacta,nifdatos"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF
        If Cancelado Then
            miRsAux.Close
            Exit Sub
        End If
        SQL = "INSERT INTO tmpfaclin (codusu, codigo, NIF) VALUES (" & vUsu.Codigo & "," & I & ",'" & miRsAux!nifdatos & "')"
        Ejecuta SQL
        miRsAux.MoveNext
        I = I + 1
        DoEvents
    Wend
    
    miRsAux.Close
    If Cancelado Then Exit Sub
    
    'Ahora cargaremos la tabla tmp347 que tendra las cuentas
    'Para cada NIF generaremos sus datos, con las empresas
    SQL = "Select nif from tmpfaclin where codusu =" & vUsu.Codigo & " group by nif"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Label9.Caption = "Nif: " & miRsAux!NIF
        Label9.Refresh

        For I = 1 To ListView3.ListItems.Count
            If ListView3.ListItems(I).Checked Then
                If Cancelado Then
                    miRsAux.Close
                    Exit Sub
                End If
                SQL = "Select " & vUsu.Codigo & "," & Mid(ListView3.ListItems(I).Key, 2) & ",codmacta,'" & miRsAux!NIF & "'"
                SQL = "INSERT INTO tmp347 (codusu, cliprov, cta, nif) " & SQL
                SQL = SQL & " FROM Conta" & ListView3.ListItems(I).Tag & ".cuentas WHERE nifdatos = '" & miRsAux!NIF & "' ORDER BY codmacta"
                If Not Ejecuta(SQL) Then Exit Sub
            
                DoEvents
            
            End If
        Next I
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    Label9.Caption = "Cuentas obtenidas. Leyendo BD"
    Me.Refresh
    espera 0.5
    
End Sub


Private Function LeerComboReferencia(Leer As Boolean) As Integer
    LeerComboReferencia = 0
    On Error GoTo ELeerRef
    SQL = App.Path & "\CmbRefer.xdf"
    If Leer Then
        LeerComboReferencia = 2
        If Dir(SQL, vbArchive) <> "" Then
            I = FreeFile
            Open SQL For Input As #I
            Line Input #I, SQL
            Close #I
            If SQL <> "" Then
                If IsNumeric(SQL) Then LeerComboReferencia = Val(SQL)
            End If
        End If
        
    Else
        If Me.cmbReferencia.ListIndex = 2 Then
            If Dir(SQL, vbArchive) <> "" Then Kill SQL
        Else
            I = FreeFile
            Open SQL For Output As #I
            Print #I, cmbReferencia.ListIndex
            Close #I
        End If
    End If
    Exit Function
ELeerRef:
    Err.Clear
End Function


Private Sub CargaGastos()
Dim Importe As Currency
    Label11.Caption = "Cargando datos"
    Label11.Refresh


    'ESTO ES UN POCO MARCIANO
    '-------------------------------------------------
    '
    ' El recodset mirsaux  viene cargado desde la fase anterior
    ' De ese modo, con una unica .open . Si no es EOF lanzamos esta pantalla
    ' si es EOF ni nos molestamos en abrirla

    While Not miRsAux.EOF
        Set IT = ListView4.ListItems.Add()
        IT.Text = miRsAux!Descripcion
        IT.SubItems(1) = Format(miRsAux!Fecha, "dd/mm/yyyy")
        IT.SubItems(2) = Format(miRsAux!Importe, FormatoImporte)
        IT.Checked = True
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    Label11.Caption = ""
    
    
    
End Sub

Private Sub CargaDatosContabilizarGastos()
    txtCta(6).Text = RecuperaValor(CadenaDesdeOtroForm, 3)
    txtDescCta(6).Text = RecuperaValor(CadenaDesdeOtroForm, 4)
    txtCtaNormal(0).Text = RecuperaValor(CadenaDesdeOtroForm, 5)
    txtDCtaNormal(0).Text = RecuperaValor(CadenaDesdeOtroForm, 6)
    Text9.Text = RecuperaValor(CadenaDesdeOtroForm, 2)
    'Fecha e Importe
    SQL = RecuperaValor(CadenaDesdeOtroForm, 7)
    I = InStr(8, SQL, " ")
    Text1(19).Text = Trim(Mid(SQL, 1, I))
    txtImporte(3).Text = Trim(Mid(SQL, I))
    'ASignaremos cadenadesdeotroform el valor para hacer el UPDATE del registro SI se contabiliza
    SQL = RecuperaValor(CadenaDesdeOtroForm, 1) & "|"
    CadenaDesdeOtroForm = SQL & Text1(19).Text & "|" & Text9.Text & "|"
    
    VisibleCC
End Sub

Private Sub PonerCuentasCC()

    CuentasCC = ""
    If vParam.autocoste Then
        SQL = "Select * from parametros"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'NO PUEDE SER EOF
        CuentasCC = "|" & miRsAux!grupogto & "|" & miRsAux!grupovta & "|"
        miRsAux.Close
        Set miRsAux = Nothing
    End If
End Sub


Private Sub VisibleCC()
Dim B As Boolean

    B = False
    If vParam.autocoste Then
        If txtCtaNormal(0).Text <> "" Then
                SQL = "|" & Mid(txtCtaNormal(0).Text, 1, 1) & "|"
                If InStr(1, CuentasCC, SQL) > 0 Then B = True
        End If
    End If
    Label1(14).Visible = B
    txtCC(0).Visible = B
    txtDCC(0).Visible = B
    imgCC(0).Visible = B
End Sub



Private Sub LanzaBuscaGrid(Opcion As Integer)

'No tocar variable SQL
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String



    SQL = ""
    Screen.MousePointer = vbHourglass
    Set frmB = New frmBuscaGrid
    frmB.vSQL = ""
    
    '###A mano
    frmB.vDevuelve = "0|"   'Siempre el 0
    
    frmB.vSelElem = 0
    
    'Ejemplo
        'Cod Diag.|idDiag|N|10·
        Select Case Opcion
        Case 0
            'DIARIO
            Cad = "Codigo|numdiari|N|15·"
            Cad = Cad & "Descripcion|desdiari|T|60·"
            frmB.vTabla = "tiposdiario"
            frmB.vTitulo = "Diario"
        Case 1
            'CONCEPTO
            Cad = "Codigo|codconce|N|15·"
            Cad = Cad & "Descripcion|nomconce|T|60·"
            frmB.vTabla = "Conceptos"
            frmB.vTitulo = "Conceptos"
            
            frmB.vSQL = " codconce <900"
        
        Case 2
            'CC
            Cad = "Codigo|codccost|N|15·"
            Cad = Cad & "Descripcion|nomccost|T|60·"
            frmB.vTabla = "cabccost"
            frmB.vTitulo = "Centros de coste"
            
        Case 3
            'Cuentas agrupadas bajo el concepto: grupotesoreria
            Cad = "Grupo tesoreria|grupotesoreria|T|60·"
            frmB.vTabla = "cuentas"
            frmB.vSQL = " grupotesoreria <> '' GROUP BY 1"
            frmB.vTitulo = "Cuentas grupos tesoreria"
        End Select
           
  
        frmB.vCampos = Cad
        
        
        
        

'        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        

    Screen.MousePointer = vbDefault
End Sub




Private Function ContabilizarGastoFijo() As Boolean
Dim Mc As Contadores
Dim FechaAbono As Date
Dim Importe As Currency
    On Error GoTo EContabilizarGastoFijo
    ContabilizarGastoFijo = False
    Set Mc = New Contadores
    
    FechaAbono = CDate(Text1(19).Text)
    If Mc.ConseguirContador("0", FechaAbono <= vParam.fechafin, True) = 1 Then Exit Function
   
    
    
    'Insertamos la cabera
    SQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) VALUES ("
    SQL = SQL & txtDiario(0).Text & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ", 1, NULL, '"
    SQL = SQL & "Gasto fijo : " & RecuperaValor(CadenaDesdeOtroForm, 1) & " - " & DevNombreSQL(RecuperaValor(CadenaDesdeOtroForm, 3)) & vbCrLf
    SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy") & " por " & DevNombreSQL(vUsu.Nombre) & "');"
    If Not Ejecuta(SQL) Then Exit Function
    
    If InStr(1, txtImporte(3).Text, ",") > 0 Then
        'Texto formateado
        Importe = ImporteFormateado(txtImporte(3).Text)
    Else
        Importe = CCur(TransformaPuntosComas(txtImporte(3).Text))
    End If
    I = 1
    Do
        'Lineas de apuntes .
         SQL = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, "
         SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
         SQL = SQL & " timporteH, ctacontr, codccost,idcontab, punteada) "
         SQL = SQL & "VALUES (" & txtDiario(0).Text & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & "," & I & ",'"
         
         'Cuenta
         If I = 1 Then
            SQL = SQL & txtCtaNormal(0).Text
         Else
            SQL = SQL & txtCta(6).Text
        End If
        SQL = SQL & "','" & Format(Val(RecuperaValor(CadenaDesdeOtroForm, 1)), "000000000") & "'," & txtConcepto(0).Text & ",'"
        
        'Ampliacion
        SQL = SQL & DevNombreSQL(Mid(txtDConcpeto(0).Text & " " & Text9.Text, 1, 30)) & "',"
                        
        If I = 1 Then
            SQL = SQL & TransformaComasPuntos(CStr(Importe)) & ",NULL,'"
            'Contrapar
            SQL = SQL & txtCta(6).Text
        Else
            SQL = SQL & "NULL," & TransformaComasPuntos(CStr(Importe)) & ",'"
            'Contrpar
            SQL = SQL & txtCtaNormal(0).Text
        End If
        
        'Solo para la line NO banco
        If I = 1 And txtCC(0).Visible Then
            SQL = SQL & "','" & txtCC(0).Text & "'"
        Else
            SQL = SQL & "',NULL"
        End If
        SQL = SQL & ",'CONTAB',0)"
        
        If Not Ejecuta(SQL) Then Exit Function
        I = I + 1
    Loop Until I > 2  'Una para el banoc, otra para la cuenta
   
    
    'Insertamos para pasar a hco
    InsertaTmpActualizar Mc.Contador, txtDiario(0).Text, FechaAbono
    
    
    
    

    'AHora actualizamos el gasto
    FechaAbono = RecuperaValor(CadenaDesdeOtroForm, 2)
    SQL = "UPDATE sgastfijd SET"
    SQL = SQL & " contabilizado=1"
    SQL = SQL & " WHERE codigo=" & RecuperaValor(CadenaDesdeOtroForm, 1)
    SQL = SQL & " and fecha='" & Format(FechaAbono, FormatoFecha) & "'"
    Conn.Execute SQL


    
    
    ContabilizarGastoFijo = True
    Exit Function
EContabilizarGastoFijo:
    MuestraError Err.Number, "Contabilizar Gasto Fijo"
End Function



Private Function LeerGuardarOrdenacion(Leer As Boolean, Cobros As Boolean, Valor As Integer) As Integer
Dim C As String
Dim NF As Integer
Dim Fichero As String

On Error GoTo ELeerGuardarOrdenacion
    LeerGuardarOrdenacion = 0
    
    NF = FreeFile
    If Cobros Then
        Fichero = App.Path & "\OrdenCob.xdf"
    Else
        Fichero = App.Path & "\OrdenPag.xdf"
    End If
    If Leer Then
        
        If Dir(Fichero, vbArchive) <> "" Then
            
            Open Fichero For Input As #NF
            Line Input #NF, C
            Close #NF
            
            LeerGuardarOrdenacion = Val(C)
    
        End If
    Else
        
            Open Fichero For Output As #NF
            Print #NF, Valor
            Close #NF
    
    End If
    Exit Function
ELeerGuardarOrdenacion:
    Err.Clear
End Function



Private Sub PonerValoresPorDefectoDevilucionRemesa()
Dim FP As Ctipoformapago

    On Error GoTo EPonerValoresPorDefectoDevilucionRemesa
    
    
    Set FP = New Ctipoformapago
    FP.Leer vbTipoPagoRemesa
    Me.txtConcepto(1).Text = FP.condecli
    Me.txtConcepto(2).Text = FP.conhapro
    'Ampliaciones
    Combo2(0).ListIndex = FP.ampdecli
    Combo2(1).ListIndex = FP.amphapro
    
    'Que carge el concepto
    txtConcepto_LostFocus 1
    txtConcepto_LostFocus 2
    Set FP = Nothing
    Exit Sub
EPonerValoresPorDefectoDevilucionRemesa:
    MuestraError Err.Number, "PonerValoresPorDefectoDevilucionRemesa"
    Set FP = Nothing
End Sub


Private Sub CargalistaCuentas()
    List1.Clear
    If CadenaDesdeOtroForm <> "" Then
        Do
            I = InStr(1, CadenaDesdeOtroForm, "|")
            If I > 0 Then
                SQL = Mid(CadenaDesdeOtroForm, 1, I - 1)
                CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, I + 1)
                CuentaCorrectaUltimoNivel SQL, CuentasCC
                SQL = SQL & "      " & CuentasCC
                List1.AddItem SQL
            End If
        Loop Until I = 0
        CadenaDesdeOtroForm = ""
        
        'Genero Cuentas CC  (por no declarar mas variables vamos)
        CuentasCC = ""
        For I = 0 To List1.ListCount - 1
            SQL = Mid(List1.List(I), 1, vEmpresa.DigitosUltimoNivel)
            CuentasCC = CuentasCC & SQL & "|"
        Next I
    Else
        CuentasCC = ""
    End If
    
End Sub



Private Sub CargaGrupo()

    On Error GoTo ECargaGrupo
    
    SQL = "Select codmacta,nommacta FROM cuentas where grupotesoreria ='" & DevNombreSQL(SQL) & "'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not miRsAux.EOF
        SQL = miRsAux!codmacta
        If InStr(1, CuentasCC, SQL & "|") > 0 Then
            I = 1
        Else
            CuentasCC = CuentasCC & SQL & "|"
            SQL = SQL & "      " & miRsAux!Nommacta
            List1.AddItem SQL
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If I > 0 Then MsgBox "Algunas cuentas YA habian sido insertadas", vbExclamation
    Exit Sub
ECargaGrupo:
    MuestraError Err.Number, "CargaGrupo"
End Sub



Private Function ComprobarEfectosBorrar() As Boolean
Dim J As Integer
Dim Dias As Integer
Dim Tipopago As Byte
    ComprobarEfectosBorrar = False
    
    For J = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(J).Checked Then

                If ListView2.ListItems(J).Tag = 2 Then
                    'Tipopago = vbPagare
                    Tipopago = 2
                ElseIf ListView2.ListItems(J).Tag = 3 Then
                    'Tipopago = vbTalon
                    Tipopago = 3
                Else
                    'Tipopago = vbTipoPagoRemesa
                    Tipopago = 1
                End If
        
                    
                'Datos bancos. Importe maximo para dias 1, dias2 si no llega
                If Tipopago = 3 Then
                    'Sobre talones
                    'SQL = "100000000,talondias,talondias"
                    SQL = "talondias"
                ElseIf Tipopago = 2 Then
                    'SQL = "100000000,pagaredias,pagaredias"
                    SQL = "pagaredias"
                Else
                    'Efectos.
                    'SQL = "remesariesgo,remesadiasmenor,remesadiasmayor"
                    SQL = "remesadiasmenor"
                End If
   
                    
                'ANTES   Marzo 2011
                'Datos bancos. Importe maximo para dias 1, dias2 si no llega
''                If SubTipo = 3 Then
''                    'Sobre talones
''                    'SQL = "100000000,talondias,talondias"
''                    SQL = "talondias"
''                ElseIf SubTipo = 2 Then
''                    'SQL = "100000000,pagaredias,pagaredias"
''                    SQL = "pagaredias"
''                Else
''                    'Efectos.
''                    'SQL = "remesariesgo,remesadiasmenor,remesadiasmayor"
''                    SQL = "remesadiasmenor"
''                End If
                    
                SQL = "select " & SQL & " from remesas r,ctabancaria b where r.codmacta=b.codmacta and codigo=" & ListView2.ListItems(J).SubItems(1) & " AND anyo = " & ListView2.ListItems(J).Text
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If miRsAux.EOF Then
                    SQL = "Error grave datos banco" & vbCrLf & SQL
                Else
                    SQL = ""
                    Dias = DBLet(miRsAux.Fields(0), "N")
                End If
                
                miRsAux.Close
                
                If SQL <> "" Then
                    MsgBox SQL, vbExclamation
                    Exit Function
                End If
                
                SQL = "Select fecvenci from scobro WHERE codrem=" & ListView2.ListItems(J).SubItems(1)
                SQL = SQL & " AND anyorem = " & ListView2.ListItems(J).Text
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                SQL = ""
                If miRsAux.EOF Then
                    'NO hay ningun vencimiento menor.
                    SQL = "UPDATE remesas Set situacion=""Z"" where codigo =" & ListView2.ListItems(J).SubItems(1)
                    SQL = SQL & " AND anyo= " & ListView2.ListItems(J).Text
                    EjecutarSQL SQL
                    
                    
                    
                Else
                    While Not miRsAux.EOF
                        NumRegElim = DateDiff("d", miRsAux!FecVenci, Now)
                        
                        If NumRegElim > Dias Then SQL = "OK"
                        miRsAux.MoveNext
                    Wend
                    
                End If
                
                'Cierro el RS
                miRsAux.Close
                
                            
                
                
                
                
                If SQL = "OK" Then
                    ComprobarEfectosBorrar = True
                    Exit Function
                End If
                    
        End If 'De checked
    Next J


End Function


'Podria darse el caso que el importe del talon/pagare
'Se distinto a la suma de los vencimientos que lo comoponen
'con lo cual el apunte de abono debera contemplar esa diferencia
'y llevarlo a una cuenta 6-7
Private Function ComprobarImportesRemTalonPagare(ImporteRemesa As Currency, ByRef ImporteDocumentos As Currency) As Boolean
Dim DocumentosRecibido As Long

    On Error GoTo EComprobarImportesRemTalonPagare
    

    ComprobarImportesRemTalonPagare = False


    

    CuentasCC = "select l.id from   slirecepdoc l left join  scobro  on l.numserie=scobro.numserie and"
    CuentasCC = CuentasCC & " l.numfaccl=scobro.codfaccl and   l.fecfaccl=scobro.fecfaccl and l.numvenci=scobro.numorden"
    CuentasCC = CuentasCC & " WHERE codrem=" & Text3(3).Text & " AND anyorem=" & Text3(4).Text
    CuentasCC = CuentasCC & " group by id"
    
    
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open CuentasCC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ImporteDocumentos = 0
    DocumentosRecibido = 0
    CuentasCC = ""
    While Not miRsAux.EOF
        If IsNull(miRsAux!Id) Then
            CuentasCC = "Hay vencimientos asociados a la remesa sin estar en la recepcion de documentos."
        Else
        
            If DocumentosRecibido <> miRsAux!Id Then
                
                If DocumentosRecibido > 0 Then ImporteDocumentos = ImporteDocumentos + CCur(DBLet(DevuelveDesdeBD("importe", "scarecepdoc", "codigo", CStr(DocumentosRecibido))))
                DocumentosRecibido = miRsAux!Id
        
            End If
            
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If DocumentosRecibido > 0 Then ImporteDocumentos = ImporteDocumentos + CCur(DBLet(DevuelveDesdeBD("importe", "scarecepdoc", "codigo", CStr(DocumentosRecibido))))
    
    Set miRsAux = Nothing
    
    If CuentasCC <> "" Then MsgBox CuentasCC, vbExclamation
    
    
    
    
    ComprobarImportesRemTalonPagare = True
    
    
    
    Exit Function
EComprobarImportesRemTalonPagare:
    MuestraError Err.Number
End Function



Private Function DiferenciaEnImportes(Indice As Integer) As Boolean
Dim RB As ADODB.Recordset
Dim C As String
Dim Impor As Currency
Dim Codigo As Integer

    C = "select scobro.impvenci,l.importe,id from slirecepdoc l left join  scobro  on l.numserie=scobro.numserie and"
    C = C & " l.numfaccl=scobro.codfaccl and   l.fecfaccl=scobro.fecfaccl and l.numvenci=scobro.numorden"
    C = C & " WHERE anyorem = " & ListView2.ListItems(Indice).Text
    C = C & " AND codrem = " & ListView2.ListItems(Indice).SubItems(1) & " ORDER BY ID"
    
    Set RB = New ADODB.Recordset
    RB.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    DiferenciaEnImportes = False
    Codigo = 0
    While Not RB.EOF
        If RB!Id <> Codigo Then
            'Ha cambiado de documento
            If Codigo > 0 Then
                C = DevuelveDesdeBD("importe", "scarecepdoc", "codigo", CStr(Codigo))
                If CCur(C) <> Impor Then
                    'Ya esta clara la diferencia. Nos piramos
                    DiferenciaEnImportes = True
                    RB.Close
                    Exit Function
                End If
            End If
            'Reestablecemos
            Codigo = RB!Id
            Impor = 0
        End If
        'El importe
        Impor = Impor + RB!Importe
        'Siguiente
        RB.MoveNext
    Wend
    RB.Close
        
    If Codigo > 0 Then
        C = DevuelveDesdeBD("importe", "scarecepdoc", "codigo", CStr(Codigo))
        If CCur(C) <> Impor Then
            'Ya esta clara la diferencia. Nos piramos
            DiferenciaEnImportes = True
        End If
    End If
    Set RB = Nothing
End Function


'Cuando eliminamos un pagare/talon en los cuales el importe del talon
'no se corresponde con el de los vencimientos, entonces el program
'debe intentar que se eliminen todos a la vez
Private Function ComprobarTodosVencidos(Indice As Integer) As Boolean
Dim RV As ADODB.Recordset
Dim C As String
Dim Dias As Integer
        
        Set RV = New ADODB.Recordset
        If SubTipo = 3 Then
            C = "talondias"
        Else
            'SQL = "100000000,pagaredias,pagaredias"
            C = "pagaredias"
        End If
              
                    
        C = "select " & C & " from remesas r,ctabancaria b where r.codmacta=b.codmacta and codigo="
        C = C & ListView2.ListItems(Indice).SubItems(1) & " AND anyo = " & ListView2.ListItems(Indice).Text
        RV.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Dias = DBLet(RV.Fields(0), "N")
        RV.Close
    

        C = "select fecvenci from slirecepdoc l left join  scobro  on l.numserie=scobro.numserie and"
        C = C & " l.numfaccl=scobro.codfaccl and   l.fecfaccl=scobro.fecfaccl and l.numvenci=scobro.numorden"
        C = C & " WHERE anyorem= " & ListView2.ListItems(Indice).Text
        C = C & " AND codrem = " & ListView2.ListItems(Indice).SubItems(1)
        
        RV.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        C = ""
        While Not RV.EOF
            NumRegElim = DateDiff("d", RV!FecVenci, Now)
            If NumRegElim < Dias Then C = C & "#"
            RV.MoveNext
        Wend
        RV.Close
        Set RV = Nothing
        If C <> "" Then
            C = "Existen " & Len(C) & " vencimiento(s)  que no han vencido todavia."
            C = C & vbCrLf & "¿Continuar?"
            If MsgBox(C, vbQuestion + vbYesNo) = vbNo Then Exit Function
        End If
        
        ComprobarTodosVencidos = True
End Function


Private Sub CamposRemesaAbono()
       
   Me.txtTexto(0).Text = ""
   Me.txtTexto(1).Text = ""
   
   
   If Text3(3) <> "" And Text3(4).Text <> "" Then
        
        Set RS = New ADODB.Recordset
        SQL = "select importe,nommacta from remesas,cuentas where remesas.codmacta=cuentas.codmacta "
        SQL = SQL & " and anyo=" & Text3(4).Text & " and codigo=" & Text3(3).Text
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            Me.txtTexto(0).Text = RS!Nommacta
            Me.txtTexto(1).Text = Format(RS!Importe, FormatoImporte)
        End If
        RS.Close
        Set RS = Nothing
    End If
    
End Sub



Private Sub EliminarEnRecepcionDocumentos()
Dim CtaPte As Boolean
Dim J As Integer
Dim CualesEliminar As String
On Error GoTo EEliminarEnRecepcionDocumentos

    'Comprobaremos si hay datos
    
        'Si no lleva cuenta puente, no hace falta que este contabilizada
        'Es decir. Solo mirare contabilizados si llevo ctapuente
        CuentasCC = ""
        CualesEliminar = ""
        J = 0
        For I = 0 To 1
            ' contatalonpte
            SQL = "pagarecta"
            If I = 1 Then SQL = "contatalonpte"
            CtaPte = (DevuelveDesdeBD(SQL, "paramtesor", "codigo", "1") = "1")
            
            'Repetiremos el proceso dos veces
            SQL = "Select * from scarecepdoc where fechavto<='" & Format(Text1(17).Text, FormatoFecha) & "'"
            SQL = SQL & " AND   talon = " & I
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                    'Si lleva cta puente habra que ver si esta contbilizada
                    J = 0
                    If CtaPte Then
                        If Val(RS!Contabilizada) = 0 Then
                            'Veo si tiene lineas. S
                            SQL = DevuelveDesdeBD("count(*)", "slirecepdoc", "id", CStr(RS!Codigo))
                            If SQL = "" Then SQL = "0"
                            If Val(SQL) > 0 Then
                                CuentasCC = CuentasCC & RS!Codigo & " - No contabilizada" & vbCrLf
                                J = 1
                            End If
                        End If
                    End If
                    If J = 0 Then
                        'Si va benee
                        If Val(DBLet(RS!llevadobanco, "N")) = 0 Then
                            SQL = DevuelveDesdeBD("count(*)", "slirecepdoc", "id", CStr(RS!Codigo))
                            If SQL = "" Then SQL = "0"
                            If Val(SQL) > 0 Then
                                CuentasCC = CuentasCC & RS!Codigo & " - Sin llevar a banco" & vbCrLf
                                J = 1
                            End If
                    
                        End If
                    End If
                    'Esta la borraremos
                    If J = 0 Then CualesEliminar = CualesEliminar & ", " & RS!Codigo
                    
                    RS.MoveNext
            Wend
            RS.Close
            
            
            
        Next I
        
        

        
        If CualesEliminar = "" Then
            'No borraremos ninguna
            If CuentasCC <> "" Then
                CuentasCC = "No se puede eliminar de la recepcion de documentos los siguientes registros: " & vbCrLf & vbCrLf & CuentasCC
                MsgBox CuentasCC, vbExclamation
                
            End If
            Exit Sub
        End If
            
        
        
        'Si k hay para borrar
        CualesEliminar = Mid(CualesEliminar, 2)
        J = 1
        SQL = "X"
        Do
            I = InStr(J, CualesEliminar, ",")
            If I > 0 Then
                J = I + 1
                SQL = SQL & "X"
            End If
        Loop Until I = 0
        
        SQL = "Va a eliminar " & Len(SQL) & " registros de la recepcion de documentos." & vbCrLf & vbCrLf & vbCrLf
        If CuentasCC <> "" Then CuentasCC = "No se puede eliminar de la recepcion de documentos los siguientes registros: " & vbCrLf & vbCrLf & CuentasCC
        SQL = SQL & vbCrLf & CuentasCC
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
            SQL = "DELETE from slirecepdoc where id in (" & CualesEliminar & ")"
            Conn.Execute SQL
            
            SQL = "DELETE from scarecepdoc where codigo in (" & CualesEliminar & ")"
            Conn.Execute SQL
    
        End If

    Exit Sub
EEliminarEnRecepcionDocumentos:
    MuestraError Err.Number, Err.Description
End Sub



Private Sub txtTexto_GotFocus(Index As Integer)
    ObtenerFoco txtTexto(Index)
End Sub

Private Sub txtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
        KEYpress KeyAscii
End Sub

Private Sub GuardaDatosConceptoTalonPagare()
    CuentasCC = "DELETE FROM tmpimpbalance WHERE codusu = " & vUsu.Codigo
    Conn.Execute CuentasCC
  
    If txtTexto(3).Text <> "" Then
        CuentasCC = "Insert into `tmpimpbalance` (`codusu`,`Pasivo`,`codigo`,`QueCuentas`) VALUES (" & vUsu.Codigo
        CuentasCC = CuentasCC & ",'Z',1,'" & DevNombreSQL(txtTexto(3).Text) & "')"
        Ejecuta CuentasCC
        
    End If
    CuentasCC = ""
End Sub


Private Sub SituarComboReferencia(Leer As Boolean)
Dim NF As Integer
    
    On Error GoTo eSituarComboReferencia
    
    SQL = App.Path & "\cboremref.dat"
    If Leer Then
        I = 2
        If Dir(SQL, vbArchive) <> "" Then
            NF = FreeFile
            Open SQL For Input As #NF
            If Not EOF(NF) Then
                Line Input #NF, SQL
                If SQL <> "" Then
                    If IsNumeric(SQL) Then
                        If Val(SQL) > 0 And Val(SQL) < 3 Then I = Val(SQL)
                    End If
                End If
            End If
            Close #NF
            
        End If
        Me.cmbReferencia.ListIndex = I
    Else
        'GUARDAR
        If Me.cmbReferencia.ListIndex = 2 Then
            If Dir(SQL, vbArchive) <> "" Then Kill SQL
        Else
            Open SQL For Output As #NF
            Print #NF, Me.cmbReferencia.ListIndex
            Close #NF
        End If
    End If
    Exit Sub
eSituarComboReferencia:
    Err.Clear

End Sub



Private Function ComprobacionFechasRemesaN19PorVto() As String
Dim AUX As String

    ComprobacionFechasRemesaN19PorVto = ""
    AUX = "anyorem = " & RS!Anyo & " AND codrem "
    AUX = DevuelveDesdeBD("min(fecvenci)", "scobro", AUX, RS!Codigo)
    If AUX = "" Then
        ComprobacionFechasRemesaN19PorVto = "Error fechas vto"
    Else
        If CDate(AUX) < vParam.fechaini Then
            ComprobacionFechasRemesaN19PorVto = "Vtos con fecha menor que inicio de ejercicio"
        End If
    End If
    If ComprobacionFechasRemesaN19PorVto <> "" Then Exit Function
    
    ComprobacionFechasRemesaN19PorVto = ""
    AUX = "anyorem = " & RS!Anyo & " AND codrem "
    AUX = DevuelveDesdeBD("max(fecvenci)", "scobro", AUX, RS!Codigo)
    If AUX = "" Then
        ComprobacionFechasRemesaN19PorVto = "Error fechas vto"
        Exit Function
    End If
    If CDate(AUX) > DateAdd("yyyy", 1, vParam.fechafin) Then ComprobacionFechasRemesaN19PorVto = "Vtos con fecha mayor que fin de ejercicio"
    
    
    
End Function


Private Sub CargarVtosRecaudaEjecutiva()
Dim LineaOK As Boolean
Dim Importe As Currency


    On Error GoTo eCargarVtosRecaudaEjecutiva
    SQL = "Select numserie,codfaccl,fecfaccl,numorden,fecvenci,impvenci,gastos,impcobro,scobro.codmacta,nommacta,nifdatos"
    SQL = SQL & ",dirdatos,codposta,despobla,desprovi,codbanco ,codsucur,digcontr,scobro.cuentaba"
    SQL = SQL & NumeroDocumento
    SQL = SQL & " ORDER BY numserie,codfaccl,numorden"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Me.ListView5.ListItems.Clear
    
    While Not RS.EOF
        
        
        'If RS!codfaccl = 13188 Then Stop
        
        Set IT = ListView5.ListItems.Add
        IT.Text = RS!NUmSerie
        IT.SubItems(1) = Format(RS!codfaccl, "000000")
        IT.SubItems(2) = Format(RS!fecfaccl, "dd/mm/yyyy")
        IT.SubItems(3) = Format(RS!numorden, "00")
        IT.SubItems(4) = Format(RS!FecVenci, "dd/mm/yyyy")
        
        Importe = DBLet(RS!Gastos, "N")
        Importe = Importe - DBLet(RS!impcobro, "N")
         
        
        IT.SubItems(5) = Format(RS!ImpVenci - Importe, FormatoImporte)
        If Importe <> 0 Then IT.ListSubItems(5).ForeColor = vbBlue   'marcamos con Azul el lw wn importe que tienen gastos y/o parcial
     
    
        IT.SubItems(6) = RS!codmacta
        IT.SubItems(7) = Trim(RS!Nommacta)   'NOMBRE OBLIGADO
        
        'direc
        IT.SubItems(8) = Trim(DBLet(RS!nifdatos, "N"))
        IT.SubItems(10) = Trim(DBLet(RS!dirdatos, "N"))
        IT.SubItems(11) = Right("     " & DBLet(RS!codposta), 5) & " " & Trim(DBLet(RS!desPobla, "N"))
        
        
        
        'codbanco ,codsucur,digcontr,cuentaba
        If DBLet(RS!codbanco, "N") = 0 Then
            SQL = "----"
        Else
            SQL = Format(RS!codbanco, "0000")
        End If
        CuentasCC = SQL & " "
        If DBLet(RS!codsucur, "N") = 0 Then
            SQL = "----"
        Else
            SQL = Format(RS!codsucur, "0000")
        End If
        CuentasCC = CuentasCC & " " & SQL
        'CC,cuentaba
        If Trim(DBLet(RS!digcontr, "T")) = "" Then
            SQL = "--"
        Else
            If Not IsNumeric(RS!digcontr) Then
                SQL = "--"
            Else
                SQL = Right("--" & RS!digcontr, 2)
            End If
        End If
        CuentasCC = CuentasCC & " " & SQL
        If DBLet(RS!Cuentaba, "N") = 0 Then
            SQL = "----------"
        Else
            SQL = Format(RS!Cuentaba, "0000000000")
        End If
        CuentasCC = CuentasCC & " " & SQL
                
        IT.SubItems(9) = CuentasCC
        IT.ToolTipText = IT.SubItems(7)
        
        'Validaciones
        LineaOK = True
        
        
        'No pueden estar vacios ni NOMBRE, NIF,CTABANCO,direccion y boblacion
        'Ademas NIF y ctabanco tendras comprobaciones especiales
        For I = 7 To 11
            If IT.SubItems(I) = "" Then
                LineaOK = False
                IT.ListSubItems(I).ForeColor = vbRed
            End If
        Next
        'NIF
        If IT.SubItems(8) <> "" Then
            If Not Comprobar_NIF(IT.SubItems(8)) Then
                LineaOK = False
                IT.ListSubItems(8).ForeColor = vbRed
            End If
        End If
        
        'Cta banco
        If InStr(1, IT.SubItems(9), "-") > 0 Then
                'EROR tiene un -  que he puesto al formatearla
                LineaOK = False
                IT.ListSubItems(9).ForeColor = vbRed
        End If
        
        If Not LineaOK Then
            IT.Bold = True
            IT.ForeColor = vbRed
        End If
        RS.MoveNext
        
    Wend
    RS.Close
    
    
    
eCargarVtosRecaudaEjecutiva:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        cmdRecaudaEjec.Enabled = False
    End If
    Set RS = Nothing
End Sub






Private Sub ReclamacionGargarList()
    ListView6.ListItems.Clear
    
    SQL = "SELECT fechaadq,maidatos,razosoci,nommacta FROM  USUARIOS.zentrefechas,cuentas WHERE fechaadq=codmacta  "
    SQL = SQL & " AND codUsu = " & vUsu.Codigo & " AND "
    If Me.optReclama(0).Value Then
        'Sin email
        SQL = SQL & " coalesce(maidatos,'')='' "
        ListView6.Checkboxes = False
    Else
        SQL = SQL & " maidatos<>'' "
        ListView6.Checkboxes = True
    End If
    SQL = SQL & " GROUP BY fechaadq  "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set IT = ListView6.ListItems.Add
        IT.Text = RS!fechaadq
        IT.SubItems(1) = RS!Nommacta
        IT.SubItems(2) = DBLet(RS!maidatos, "T")
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing

End Sub




Private Sub DividiVencimentosPorEntidadBancaria()

    Set miRsAux = New ADODB.Recordset
    
    Conn.Execute "DELETE FROM tmp347 WHERE codusu = " & vUsu.Codigo
    '                                                               POR SI TUVIERAN MISMO BANCO, <> cta contable
    NumeroDocumento = "select oficina,entidad from ctabancaria where not sufijoem is null "
    NumeroDocumento = NumeroDocumento & " and entidad >0  and codmacta<>'" & Me.txtCta(3).Text & "' group by 1,2"
    miRsAux.Open NumeroDocumento, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumeroDocumento = ""
    While Not miRsAux.EOF
        NumeroDocumento = NumeroDocumento & ", (" & miRsAux!Entidad & ")"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If NumeroDocumento = "" Then
        NumeroDocumento = "(-1,-1)"
    Else
        NumeroDocumento = Mid(NumeroDocumento, 2) 'quitamos la primera coma
    End If
    
    NumeroDocumento = " (codbanco) in (" & NumeroDocumento & ")"
    
    'Agrupamos los vencimientos por entidad,oficina menos los del banco por defecto
    CuentasCC = "select codbanco,sum(impvenci + coalesce(gastos,0)) " & SQL
    CuentasCC = CuentasCC & " AND " & NumeroDocumento & " GROUP BY 1"
    miRsAux.Open CuentasCC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        CuentasCC = "insert into `tmpcierre1` (`codusu`,`cta`,`nomcta`,`acumPerD`) VALUES (" & vUsu.Codigo & ","
        CuentasCC = CuentasCC & miRsAux.Fields(0) & ",0," & TransformaComasPuntos(CStr(miRsAux.Fields(1))) & ")"
        Conn.Execute CuentasCC
        
         miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Los del banco por defecto, y lo que no tenemos banco, es decir, el resto
    '------------------------------------------------------------------------------
    CuentasCC = SQL & " AND NOT " & NumeroDocumento & " GROUP BY 1,2"
    'Vere la entidad y la oficina del PPAL
    NumeroDocumento = DevuelveDesdeBD("concat(entidad,',',oficina)", "ctabancaria", "codmacta", txtCta(3).Text, "T")
    NumeroDocumento = "Select " & NumeroDocumento & ",sum(impvenci + coalesce(gastos,0)) " & CuentasCC
    miRsAux.Open NumeroDocumento, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        CuentasCC = "insert into `tmpcierre1` (`codusu`,`cta`,`nomcta`,`acumPerD`) VALUES (" & vUsu.Codigo & ","
        CuentasCC = CuentasCC & miRsAux.Fields(0) & "," & miRsAux.Fields(1) & "," & TransformaComasPuntos(CStr(miRsAux.Fields(2))) & ")"
        Conn.Execute CuentasCC
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    espera 1
    
    
    'Pongo codmacta y nombanco como corresponde
    CuentasCC = "Select * from tmpcierre1 where codusu =" & vUsu.Codigo
    miRsAux.Open CuentasCC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        NumeroDocumento = "nommacta"
        CuentasCC = "ctabancaria.codmacta=cuentas.codmacta AND ctabancaria.entidad = " & miRsAux!Cta & " AND 1 "    'ctabancaria.oficina "
        CuentasCC = DevuelveDesdeBD("ctabancaria.codmacta", "ctabancaria,cuentas", CuentasCC, "1", "N", NumeroDocumento)  'miRsAux!nomcta
        If CuentasCC <> "" Then
            CuentasCC = "UPDATE tmpcierre1 SET cta = '" & CuentasCC & "',nomcta ='" & DevNombreSQL(NumeroDocumento)
            CuentasCC = CuentasCC & "' WHERE Cta = " & miRsAux!Cta & " AND nomcta =" & miRsAux!nomcta
            Conn.Execute CuentasCC
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Por si quiere borrar alguno de los repartios que hace
    'Por si casao luego BORRAN la remesa a generar para ese banco, es decir , no uqieren llevarlo ahora
    CuentasCC = "insert into tmp347(codusu,cta) select codusu,cta from tmpcierre1 WHERE codusu =" & vUsu.Codigo
    Conn.Execute CuentasCC
    
eDividir:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        
        
    End If
    NumeroDocumento = ""
    CuentasCC = ""
    Set miRsAux = Nothing
    Set RS = Nothing
End Sub



Private Sub LeerGuardarBancoDefectoEntidad(Leer As Boolean)
On Error GoTo eLeerGuardarBancoDefectoEntidad

    I = -1
    SQL = App.Path & "\BancRemEn.xdf"
    If Leer Then
        txtCta(3).Text = ""
        If Dir(SQL, vbArchive) <> "" Then
            I = FreeFile
            Open SQL For Input As #I
            If Not EOF(I) Then
                Line Input #I, SQL
                txtCta(3).Text = SQL
                txtCta(3).Tag = SQL
            End If
        End If
    
    Else
        'Guardar
        If Me.txtCta(3).Text = "" Then
            If Dir(SQL, vbArchive) <> "" Then Kill SQL
        Else
            I = FreeFile
            Open SQL For Output As #I
            Print #I, txtCta(3).Text
            
        End If
        
        
    End If
    
    If I >= 0 Then Close #I
    Exit Sub
eLeerGuardarBancoDefectoEntidad:
    Err.Clear
End Sub
