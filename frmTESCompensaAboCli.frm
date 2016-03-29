VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESCompensaAboCli 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listados"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   Icon            =   "frmTESCompensaAboCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCompensaAbonosCliente 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10935
      Begin VB.CommandButton cmdVtoDestino 
         Height          =   375
         Index           =   1
         Left            =   240
         Picture         =   "frmTESCompensaAboCli.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6120
         Width           =   375
      End
      Begin VB.TextBox txtimpNoEdit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   9120
         TabIndex        =   14
         Top             =   5640
         Width           =   1575
      End
      Begin VB.CommandButton cmdVtoDestino 
         Height          =   375
         Index           =   0
         Left            =   240
         Picture         =   "frmTESCompensaAboCli.frx":0A0E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5640
         Width           =   375
      End
      Begin VB.TextBox txtimpNoEdit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   8880
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtimpNoEdit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   7440
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin MSComctlLib.ListView lwCompenCli 
         Height          =   3975
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nº Fact"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fec. fact."
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Vto"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fecha Vto"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Forma pago"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Cobro"
            Object.Width           =   2884
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Abonos"
            Object.Width           =   2884
         EndProperty
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   1560
         TabIndex        =   5
         Text            =   "Text5"
         Top             =   1080
         Width           =   3675
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   17
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCompensar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   8520
         TabIndex        =   2
         Top             =   6120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   36
         Left            =   9720
         TabIndex        =   1
         Top             =   6120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Imprimir hco compensacion"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   17
         Top             =   6240
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   72
         Left            =   8160
         TabIndex        =   15
         Top             =   5685
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "  Establecer vencimiento destino"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   13
         Top             =   5760
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Rectifca./Abono"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   71
         Left            =   8880
         TabIndex        =   10
         Top             =   840
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cobro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   70
         Left            =   7440
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   69
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   17
         Left            =   1560
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Compensación abonos cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   20
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   4890
      End
   End
End
Attribute VB_Name = "frmTESCompensaAboCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SaltoLinea = """ + chr(13) + """

Public Opcion As Byte
    '1.- Cobros pendientes por cliente
    
    '3.- Reclamaciones por mail
    
    '4.- lISTADO agentes
    '5.- Departamentos
    
    '6.- Listado remesas
    
    '8.- Listado caja
    
    '9-  Devol remesas
    
    '10.- Listado formas de pago

    
    '11.- Transferencias PRovee   (o confirmings (domicilados o caixaconfirming)
    
    '12.- Listado previsional de gstos/pagos
    
    '13.- Transferencias ABONOS
    
    
    'Operaciones aseguradas
    '----------------------------
    '15.-  datos basicos
    '16.-  listado facturacion
    '17.-  Impagados asegurados
    
    
    '20.- Pregunta cuenta COBRO GENERICO
    '       La pongo aqui pq tengo implemntado todo
    
    
    '22.- Datos para la contabilizacion de las compensaciones
        
    '23.- Datos para la contbailiacion de la recpcion de documentos
    
    
    '24.-  Listado de documento(tal/pag) recibidos
    
    '25.-  Listado de pagos ordenados por banco  **** AHORA NO DEBERIA ENTRAR AQUI
    
    '26.-  Cancel remesa TAL/PAG.  Cando los importe no coinden. Solicitud cta y cc
    '27.-  Divide el vencimiento en dos vtos a partir del importe introducido en el text
        
        
    '30.-  Historico RECLAMACIONES
    '31.-   Gastos fijos
        
    '33.-  ASEGURADOS.  Listados avisos falta pago, avisos prorroga, aviso siniestro
        
    '34.-  Eliminar una recepcion de documentos, que ya ha sido contb con la puente
        
    '35.-  Gastos transferencias
        
    '36.-  Compensar ABONOS cobros
            
    '38.-  Recaudacion ejecutiva
        
    '39.-   Informe de comunicacion al seguro
    '40.-    Fras pendientes operaciones aseguradas
    
    '42.-   IMportar fichero norma 57 (recibos al cobro en ventanilla)
    
    '43.-   Confirmings
    '44.-   Caixaconfirming   igual que el de arriba
        
    '45.-   Listado cobros AGENTES -- >BACCHUS
    
Private WithEvents frmCta As frmColCtas
Attribute frmCta.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBa As frmBanco
Attribute frmBa.VB_VarHelpID = -1
Private WithEvents frmA As frmAgentes
Attribute frmA.VB_VarHelpID = -1
Private WithEvents frmP As frmFormaPago
Attribute frmP.VB_VarHelpID = -1

Dim SQL As String
Dim RC As String
Dim RS As Recordset
Dim PrimeraVez As Boolean

Dim Cad As String
Dim CONT As Long
Dim I As Integer
Dim TotalRegistros As Long

Dim Importe As Currency
Dim MostrarFrame As Boolean
Dim Fecha As Date

Dim DevfrmCCtas As String

Private Sub PonFoco(ByRef T1 As TextBox)
    T1.SelStart = 0
    T1.SelLength = Len(T1.Text)
End Sub









Private Function ComprobarObjeto(ByRef T As TextBox) As Boolean
    Set miTag = New CTag
    ComprobarObjeto = False
    If miTag.Cargar(T) Then
        If miTag.Cargado Then
            If miTag.Comprobar(T) Then ComprobarObjeto = True
        End If
    End If

    Set miTag = Nothing
End Function

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 20 Or Index = 23 Or Index >= 26 Then
        CadenaDesdeOtroForm = "" 'Por si acaso. Tiene que devolve "" para que no haga nada
    End If
    Unload Me
End Sub

Private Function PonerTipoPagoCobro_(ParaSelect As Boolean, EsReclamacion As Boolean) As String
Dim I As Integer
Dim Sele As Integer
Dim AUX As String
Dim Visibles As Byte

    AUX = ""
    Sele = 0
    Visibles = 0
    If Not EsReclamacion Then
        For I = 0 To Me.chkTipPago.Count - 1
            If Me.chkTipPago(I).Visible Then
                Visibles = Visibles + 1
                If Me.chkTipPago(I).Value = 1 Then
                    Sele = Sele + 1
                    If ParaSelect Then
                        AUX = AUX & ", " & I
                    Else
                        AUX = AUX & "-" & Me.chkTipPago(I).Caption
                    End If
                End If
            End If
        Next
        'No ha marcado ninguno o los ha nmarcado todos. NO pongo nada
        If Sele = 0 Or Sele = Visibles Then AUX = ""
        
    Else
        '************************
        'Reclamaciones
        
        For I = 0 To Me.chkTipPagoRec.Count - 1
            If Me.chkTipPagoRec(I).Visible Then
                Visibles = Visibles + 1
                If Me.chkTipPagoRec(I).Value = 1 Then
                    Sele = Sele + 1
                    If ParaSelect Then
                        AUX = AUX & ", " & I
                    Else
                        AUX = AUX & "-" & Me.chkTipPagoRec(I).Caption
                    End If
                End If
            End If
        Next
        'No ha marcado ninguno o los ha nmarcado todos. NO pongo nada
        If Sele = 0 Or Sele = Visibles Then AUX = ""
    End If
    If AUX <> "" Then
        AUX = Mid(AUX, 2)
        AUX = "(" & AUX & ")"
    End If
    PonerTipoPagoCobro_ = AUX
End Function



Private Sub cmdCompensar_Click()
    
    Cad = DevuelveDesdeBD("informe", "scryst", "codigo", 10) 'Orden de pago a bancos
    If Cad = "" Then
        MsgBox "No esta configurada la aplicación. Falta informe(10)", vbCritical
        Exit Sub
    End If
    Me.Tag = Cad
    
    Cad = ""
    RC = ""
    CONT = 0
    TotalRegistros = 0
    NumRegElim = 0
    For I = 1 To Me.lwCompenCli.ListItems.Count
        If Me.lwCompenCli.ListItems(I).Checked Then
            If Trim(lwCompenCli.ListItems(I).SubItems(6)) = "" Then
                'Es un abono
                TotalRegistros = TotalRegistros + 1
            Else
                NumRegElim = NumRegElim + 1
            End If
        End If
        If Me.lwCompenCli.ListItems(I).Bold Then
            Cad = Cad & "A"
            If CONT = 0 Then CONT = I
            
            
            
        End If
    Next
    
    I = 0
    SQL = ""
    If Len(Cad) <> 1 Then
        'Ha seleccionado o cero o mas de uno
        If txtimpNoEdit(0).Text <> txtimpNoEdit(1).Text Then
            'importes distintos. Solo puede seleccionar UNO
            SQL = "Debe selecionar uno(y solo uno) como vencimiento destino"
            
        End If
    Else
        'Comprobaremos si el selecionado esta tb checked
        If Not lwCompenCli.ListItems(CONT).Checked Then
            SQL = "El vencimiento seleccionado no esta marcado"
        
        Else
            'Si el importe Cobro es mayor que abono, deberia estar
            Importe = CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag)  'txtimpNoEdit(1).Tag es negativo
            If Importe <> 0 Then
                If Importe > 0 Then
                    'Es un abono
                    If Trim(lwCompenCli.ListItems(CONT).SubItems(6)) = "" Then SQL = "cobro"
                Else
                    If Trim(lwCompenCli.ListItems(CONT).SubItems(6)) <> "" Then SQL = "abono"
                End If
                If SQL <> "" Then SQL = "Debe marcar un " & SQL
            End If
            
        End If
    End If
    If TotalRegistros = 0 Or NumRegElim = 0 Then SQL = "Debe selecionar cobro(s) y abono(s)" & vbCrLf & SQL
        
    'Sep 2012
    'NO se pueden borrar las observaciones que ya estuvieran
    'RecuperaValor("text41csb|text42csb|text43csb|text51csb|text52csb|text53csb|text61csb|text62csb|text63csb|text71csb|text72csb|text73csb|text81csb|text82csb|text83csb|", J)
    If CONT > 0 Then
        'Hay seleccionado uno vto
        Set miRsAux = New ADODB.Recordset
        Cad = "text41csb,text42csb,text43csb,text51csb,text52csb,text53csb,text61csb,text62csb,text63csb,text71csb,text72csb,text73csb,text81csb,text82csb,text83csb"
        Cad = "Select " & Cad & " FROM scobro WHERE numserie ='" & lwCompenCli.ListItems(CONT).Text & "' AND codfaccl="
        Cad = Cad & lwCompenCli.ListItems(CONT).SubItems(1) & " AND fecfaccl='" & Format(lwCompenCli.ListItems(CONT).SubItems(2), FormatoFecha)
        Cad = Cad & "' AND numorden = " & lwCompenCli.ListItems(CONT).SubItems(3)
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        If miRsAux.EOF Then
            SQL = SQL & vbCrLf & " NO se ha encontrado el veto. destino"
        Else
            'Vamos a ver cuantos registros son
            CadenaDesdeOtroForm = ""
            RC = "0"
            For I = 0 To 14
                If DBLet(miRsAux.Fields(I), "T") = "" Then
                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & miRsAux.Fields(I).Name & "|"
                    RC = Val(RC) + 1
                End If
            Next I
            
                
                
            'If TotalRegistros + NumRegElim > 15 Then SQL = SQL & vbCrLf & "No caben los textos de los vencimientos"
            If TotalRegistros + NumRegElim > Val(RC) Then SQL = SQL & vbCrLf & "No caben los textos de los vencimientos"
        End If
        miRsAux.Close
        Set miRsAux = Nothing
    End If
    
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    Else
        If MsgBox("Seguro que desea realizar la compensación?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    
    
    
    
    Me.FrameCompensaAbonosCliente.Enabled = False
    Me.Refresh
    Screen.MousePointer = vbHourglass
    
    RealizarCompensacionAbonosClientes
    Me.FrameCompensaAbonosCliente.Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub






Private Function HacerPrevisionCuenta(Cta As String, Nommacta As String) As Boolean
Dim SaldoArrastrado As Currency
Dim ID As Currency
Dim IH As Currency


    HacerPrevisionCuenta = False
    
    lblPrevInd.Caption = Cta & " - " & Nommacta
    lblPrevInd.Refresh
    ' Las fechas son del periodo, luego me importa una mierda las fechas desde hasta
    '
    '
    CargaDatosConExt Cta, Now, Now, " 1 = 1", Nommacta
    
    Conn.Execute "insert into Usuarios.ztmpconextcab select * from tmpconextcab where codusu =" & vUsu.Codigo
    
    Conn.Execute "DELETE FROM tmpfaclin where codusu =" & vUsu.Codigo
    
    RC = "INSERT INTO tmpfaclin (codusu, IVA,codigo, Fecha, Cliente, cta,"
    RC = RC & " ImpIVA, Total) VALUES (" & vUsu.Codigo & ","
    
    'PARA CADA CUENTA
    'mETEREMOS TODOS LOS REGISTROS EN LA TABLA
    '
    '           TMPFACLIN
    '
    'TANTO COBROS COMO PAGOS I GASTOS
    '
    'Luego, en funcion del orden(TIPO o fecha) los iremos insertando en la tabla, para que
    'el saldo que va arrastrando sea el correcto
    
    
       
        
    CONT = 0
    
    
    '--------------------
    'DETALLAR COBROS
    lblPrevInd.Caption = Cta & " - Cobros"
    lblPrevInd.Refresh
    SQL = " WHERE fecvenci<='" & Format(Text3(18).Text, FormatoFecha) & "'"
    SQL = SQL & " AND ctabanc1 ='" & Cta & "'"
    If chkPrevision(0).Value = 0 Then
        SQL = "select sum(impvenci),sum(impcobro),fecvenci from scobro " & SQL
        SQL = SQL & " GROUP BY fecvenci"
        
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        While Not miRsAux.EOF
        
            ID = DBLet(miRsAux.Fields(0), "N")
            IH = DBLet(miRsAux.Fields(1), "N")
            Importe = ID - IH

            If Importe <> 0 Then
                CONT = CONT + 1
                Cad = "'COBRO'," & CONT & ",'" & Format(miRsAux!FecVenci, FormatoFecha) & "','COBROS PENDIENTES',NULL,"
                'HAY COBROS
                If Importe < 0 Then
                    Cad = Cad & "NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
                Else
                    Cad = Cad & TransformaComasPuntos(CStr(Importe)) & ",NULL"
                End If
                Cad = RC & Cad & ")"
                Conn.Execute Cad
                
            End If
                miRsAux.MoveNext
        Wend
        miRsAux.Close
                
    Else
         'DETALLAR PAGOS COBROS
            '(codusu, cta, ccost,Pos, fechaent, nomdocum, ampconce,"
            'timporteD,timporteH, saldo
            
        'SQL = "select scobro.*,nommacta from scobro,cuentas where scobro.codmacta=cuentas.codmacta"
        'SQL = SQL & " AND fecvenci<='2006-01-01'"
         
        SQL = "select scobro.* from scobro " & SQL
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            CONT = CONT + 1
            Cad = "'COBRO'," & CONT & ",'" & Format(miRsAux!FecVenci, FormatoFecha) & "','"
            'NUmero factura
            Cad = Cad & miRsAux!NUmSerie & miRsAux!codfaccl & "/" & miRsAux!numorden & "',"
            
            Cad = Cad & "'" & miRsAux!codmacta & "',"
            Importe = miRsAux!ImpVenci - DBLet(miRsAux!impcobro, "N")
            If Importe <> 0 Then
                If Importe < 0 Then
                    Cad = Cad & "NULL," & TransformaComasPuntos(CStr(Abs(Importe)))
                Else
                    Cad = Cad & TransformaComasPuntos(CStr(Importe)) & ",NULL"
                End If
                Cad = Cad & ")"
                Cad = RC & Cad
                Conn.Execute Cad
            End If
            miRsAux.MoveNext
            
        Wend
        miRsAux.Close
        
    End If
    
    
    '--------------------
    '--------------------
    '--------------------
    'DETALLAR PAGOS
    '--------------------
    '--------------------
    lblPrevInd.Caption = Cta & " - pagos"
    lblPrevInd.Refresh
    SQL = " WHERE fecefect<='" & Format(Text3(18).Text, FormatoFecha) & "'"
    SQL = SQL & " AND ctabanc1 ='" & Cta & "'"
    
    If chkPrevision(1).Value = 0 Then
        SQL = "select sum(impefect),sum(imppagad),fecefect from spagop " & SQL & " GROUP BY fecefect"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Importe = 0
        While Not miRsAux.EOF

                ID = DBLet(miRsAux.Fields(0), "N")
                IH = DBLet(miRsAux.Fields(1), "N")
                Importe = ID - IH
            
                If Importe <> 0 Then
                    CONT = CONT + 1
                    Cad = "'PAGO'," & CONT & ",'" & Format(miRsAux!fecefect, FormatoFecha) & "','PAGOS PENDIENTES',NULL,"
                    'HAY COBROS
                    If Importe > 0 Then
                        Cad = Cad & "NULL," & TransformaComasPuntos(CStr(Importe))
                    Else
                        Cad = Cad & TransformaComasPuntos(CStr(Abs(Importe))) & ",NULL"
                    End If
                    Cad = RC & Cad & ")"
                    Conn.Execute Cad
                End If
                miRsAux.MoveNext
        Wend
        miRsAux.Close
    Else
         'DETALLAR PAGOS COBROS
        'codusu, IVA,codigo, Fecha, Cliente, cta,"
        ' ImpIVA, Total) VALUES (" & vUsu.Codigo & ","
        
        'SQL = "select spagop.*,nommacta from spagop,cuentas where ctaprove=codmacta"
        'SQL = SQL & " AND fecefect<='2006-01-01'"
        
        SQL = "select spagop.* from spagop " & SQL
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            CONT = CONT + 1
            Cad = "'PAGO'," & CONT & ",'" & Format(miRsAux!fecefect, FormatoFecha) & "','"
            'NUmero factura
            Cad = Cad & DevNombreSQL(miRsAux!NumFactu) & "/" & miRsAux!numorden & "',"
            
            Cad = Cad & "'" & miRsAux!ctaprove & "',"
            Importe = miRsAux!ImpEfect - DBLet(miRsAux!imppagad, "N")
            If Importe <> 0 Then
                If Importe > 0 Then
                    Cad = Cad & "NULL," & TransformaComasPuntos(CStr(Importe))
                Else
                    Cad = Cad & TransformaComasPuntos(CStr(Abs(Importe))) & ",NULL"
                End If
                Cad = Cad & ")"
                Cad = RC & Cad
                Conn.Execute Cad
            End If
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        
    End If
    
    
    
    
    
    
    '--------------------
    '--------------------
    '--------------------
    'DETALLAR GASTOS GASTOS
    '--------------------
    '--------------------
    
    SQL = " from sgastfij,sgastfijd where sgastfij.codigo= sgastfijd.codigo"
    SQL = SQL & " and fecha >='" & Format(Now, FormatoFecha)
    SQL = SQL & "' AND fecha <='" & Format(Format(Text3(18).Text, FormatoFecha), FormatoFecha) & "'"
    SQL = SQL & " and ctaprevista='" & Cta & "'"
    
    'Desde 5 Abril 2006
    '------------------
    ' Si el gasto esta contbilizado desde la tesoreria, tiene la marca "contabilizado"
    SQL = SQL & " and contabilizado=0"
    
        ' ImpIVA, Total) VALUES (" & vUsu.Codigo & ","
        
        'SQL = "select spagop.*,nommacta from spagop,cuentas where ctaprove=codmacta"
        'SQL = SQL & " AND fecefect<='2006-01-01'"
     
     
    'ABro el recodset aqui.
    'Si es EOF entonces no necesito abrir la pantalla, puesto
    ' que no habran gastos para seleccionar
    'Si NO es EOF entonces abro el form y entonces alli(en frmvarios)
    'recorro el recodset
    SQL = " select sgastfij.codigo,descripcion,fecha,importe " & SQL
    
    
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If miRsAux.EOF Then
        miRsAux.Close
    Else
        NumRegElim = CONT
        CadenaDesdeOtroForm = "Gastos cuenta: " & Nommacta & "|" & Cta & "|" & Val(chkPrevision(2).Value) & "|"
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & RC & "|"
        frmVarios.Opcion = 18
        frmVarios.Show vbModal
        Set miRsAux = New ADODB.Recordset
        CONT = NumRegElim
        Me.Refresh
    End If
    
    
    If CONT = 0 Then Exit Function
    
    lblPrevInd.Caption = Cta & " - Informe"
    lblPrevInd.Refresh
    'Cargo INFORME
    '------------------------------------------------------------------------------------------
    'Leo el  saldo inicial
    RC = "Select * from tmpconextcab where codusu =" & vUsu.Codigo & " AND cta ='" & Cta & "'"
    miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SaldoArrastrado = 0
    If Not miRsAux.EOF Then SaldoArrastrado = DBLet(miRsAux!acumtotT, "N")
    miRsAux.Close
    
    'Si desgloso cobros, los detallo, si no hago el acumu
    RC = "INSERT INTO Usuarios.ztmpconext (codusu, cta, ccost,Pos, fechaent, nomdocum, ampconce,"
    RC = RC & "timporteD,timporteH, saldo) VALUES (" & vUsu.Codigo & ",'" & Cta & "','"
        
    
    
    'Ahora cogere todos los registros que estan cargados en tmpfaclin y los metere ya
    'en la tabla con los importes, ordenado como dice el option y
    'arrastrando saldo
    SQL = "select tmpfaclin.*,nommacta from tmpfaclin left join cuentas on cta=codmacta where codusu =" & vUsu.Codigo & " ORDER BY "
    'EL ORDEN
    If optPrevision(0).Value Then
        SQL = SQL & "fecha,cta"
    Else
        SQL = SQL & "cta,fecha"
    End If
    CONT = 1
    ID = 0
    IH = 0
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Cad = Mid(miRsAux!iva, 1, 4) & "'," & CONT & ",'" & Format(miRsAux!Fecha, FormatoFecha) & "','"
        
        
        
        If IsNull(miRsAux!Cta) Then
            'Stop
            Cad = Cad & "','" & DevNombreSQL(miRsAux!Cliente) & "'"
        Else
            Cad = Cad & Mid(DevNombreSQL(miRsAux!Cliente), 1, 10) & "',"
            If IsNull(miRsAux!Nommacta) Then
                Cad = Cad & "NULL"
            Else
                Cad = Cad & "'" & DevNombreSQL(miRsAux!Nommacta) & "'"
            End If
        End If
        If IsNull(miRsAux!Total) Then
            'VA AL DEBE
            Importe = miRsAux!ImpIva
            Cad = Cad & "," & TransformaComasPuntos(CStr(miRsAux!ImpIva)) & ",NULL,"
            ID = ID + Importe
        Else
            'HABER
            Importe = miRsAux!Total * -1
            Cad = Cad & ",NULL," & TransformaComasPuntos(CStr(miRsAux!Total)) & ","
            IH = IH + miRsAux!Total
        End If
        SaldoArrastrado = SaldoArrastrado + Importe
        Cad = Cad & TransformaComasPuntos(CStr(SaldoArrastrado)) & ")"
        Cad = RC & Cad
        Conn.Execute Cad
        
        
        CONT = CONT + 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    'Ajusto los importes de la tabla tmpconextcab
    SQL = "UPDATE Usuarios.ztmpconextcab SET acumantD=acumtotD,acumantH=acumtotH,acumantT=acumtotT"
    SQL = SQL & " where codusu =" & vUsu.Codigo & " AND cta ='" & Cta & "'"
    Conn.Execute SQL
    SQL = "UPDATE Usuarios.ztmpconextcab SET acumperD=" & TransformaComasPuntos(CStr(ID))
    SQL = SQL & ", acumperH=" & TransformaComasPuntos(CStr(IH))
    SQL = SQL & ", acumperT=" & TransformaComasPuntos(CStr(ID - IH))
    SQL = SQL & ", acumtott=" & TransformaComasPuntos(CStr(SaldoArrastrado))
    
    SQL = SQL & " where codusu =" & vUsu.Codigo & " AND cta ='" & Cta & "'"
    Conn.Execute SQL
    
    HacerPrevisionCuenta = True
    
End Function

Private Sub MontaSQLReclamacion()
    
    'Siempre hay que añadir el AND
    
    
    SQL = ""
    
    
    'Fecha factura
    RC = CampoABD(txtSerie(2), "T", "scobro.numserie", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtSerie(3), "T", "scobro.numserie", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    
    'Fecha factura
    RC = CampoABD(Text3(6), "F", "fecfaccl", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    
    RC = CampoABD(Text3(7), "F", "fecfaccl", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    'Fecha vto
    RC = CampoABD(Text3(9), "F", "fecvenci", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(Text3(10), "F", "fecvenci", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    'cuenta
    RC = CampoABD(txtCta(4), "T", "scobro.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(5), "T", "scobro.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC

    
    
    'Agente
    RC = CampoABD(txtAgente(3), "N", "scobro.agente", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtAgente(2), "N", "scobro.agente", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    
    'Forma de pago
    RC = CampoABD(txtFPago(3), "N", "scobro.codforpa", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtFPago(2), "N", "scobro.codforpa", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    'Solo devueltos
    If chkReclamaDevueltos.Value = 1 Then SQL = SQL & " AND devuelto = 1"
      
    
    'Marzo2015
    If chkExcluirConEmail.Value = 1 Then SQL = SQL & " AND coalesce(maidatos,'')=''"
    
    
    'LA de la fecha
    SQL = SQL & " AND ((ultimareclamacion  is null) OR (ultimareclamacion <= '" & Format(Fecha, FormatoFecha) & "'))"
    
    'QUE FALTE POR PAGAR
    SQL = SQL & " AND (impvenci>0)"
    
    
    RC = PonerTipoPagoCobro_(True, True)
    If RC <> "" Then SQL = SQL & " AND tipforpa IN " & RC
    
    
    
    'Select
    Cad = "Select scobro.*, cuentas.codmacta FROM scobro,cuentas,sforpa "
    Cad = Cad & " WHERE  sforpa.codforpa=scobro.codforpa AND scobro.codmacta = cuentas.codmacta"
    Cad = Cad & " AND sforpa.codforpa=scobro.codforpa "
    SQL = Cad & SQL
    
    
    
    
    
End Sub



Private Sub cmdVtoDestino_Click(Index As Integer)
    
    If Index = 0 Then
        TotalRegistros = 0
        If Not Me.lwCompenCli.SelectedItem Is Nothing Then TotalRegistros = Me.lwCompenCli.SelectedItem.Index
    
    
        For I = 1 To Me.lwCompenCli.ListItems.Count
            If Me.lwCompenCli.ListItems(I).Bold Then
                Me.lwCompenCli.ListItems(I).Bold = False
                Me.lwCompenCli.ListItems(I).ForeColor = vbBlack
                For CONT = 1 To Me.lwCompenCli.ColumnHeaders.Count - 1
                    Me.lwCompenCli.ListItems(I).ListSubItems(CONT).ForeColor = vbBlack
                    Me.lwCompenCli.ListItems(I).ListSubItems(CONT).Bold = False
                Next
            End If
        Next
        Me.Refresh
        
        If TotalRegistros > 0 Then
            I = TotalRegistros
            Me.lwCompenCli.ListItems(I).Bold = True
            Me.lwCompenCli.ListItems(I).ForeColor = vbRed
            For CONT = 1 To Me.lwCompenCli.ColumnHeaders.Count - 1
                Me.lwCompenCli.ListItems(I).ListSubItems(CONT).ForeColor = vbRed
                Me.lwCompenCli.ListItems(I).ListSubItems(CONT).Bold = True
            Next
        End If
        lwCompenCli.Refresh
        
        PonerfocoObj Me.lwCompenCli

    Else
    
        Cad = DevuelveDesdeBD("informe", "scryst", "codigo", 10) 'Orden de pago a bancos
        If Cad = "" Then
            MsgBox "No esta configurada la aplicación. Falta informe(10)", vbCritical
            Exit Sub
        End If
        CadenaDesdeOtroForm = Cad
    
        LanzaBuscaGrid 1, 4


    End If
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        Case 1
            Text3(1).SetFocus
        Case 3
            
            'Reclamaciones. Si no tiene configurado el envio web
            'no habilitaremos el check
            Cad = DevuelveDesdeBD("smtpHost", "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
            If Cad = "" Then
                Me.chkEmail.Value = 0
                chkEmail.Enabled = False
            End If
            'Text3(6).SetFocus
            txtSerie(2).SetFocus
        Case 10
            Me.cmdFormaPago.SetFocus
        Case 12
            txtCtaBanc(0).SetFocus
        Case 20
            PonerFoco txtCta(13)
            
        Case 22
            'Contabi efectos
            If CONT > 0 Then
                For I = 1 To Me.cboCompensaVto.ListCount
                    If Me.cboCompensaVto.ItemData(I) = CONT Then
                        CONT = I
                        Exit For
                    End If
                Next
            End If
            Me.cboCompensaVto.ListIndex = CONT
            PonerFoco Text3(23)
        Case 23
            CadenaDesdeOtroForm = ""  'Para que  no devuelva nada
        Case 30
            PonerFoco Text3(28)
            
        Case 31
            'gastos fijos
            Text3(30).Text = "01/01/" & Year(Now)
        Case 35
            PonerFoco txtImporte(2)
            
        Case 36
            If CadenaDesdeOtroForm <> "" Then
                txtCta(17).Text = CadenaDesdeOtroForm
                txtCta_LostFocus 17
            Else
                PonerFoco txtCta(17)
            End If
            CadenaDesdeOtroForm = ""
            
        Case 39
            PonerFoco Text3(34)
            
        Case 42
            
'            Me.Refresh
'            cmdNoram57Fich_Click


        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub



    
Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim Img As Image


    Limpiar Me
    Me.Icon = frmPpal.Icon
    CargaImagenesAyudas Me.imgCtaBanc, 1, "Cuenta contable bancaria"
    CargaImagenesAyudas Image2, 2
    CargaImagenesAyudas Me.imgFP, 1, "Forma de pago"
    CargaImagenesAyudas Me.Image3, 1, "Cuenta contable"
    CargaImagenesAyudas Me.Imagente, 1, "Seleccionar agente"
    CargaImagenesAyudas imgCCoste, 1, "Centro de coste"
    CargaImagenesAyudas Me.ImageAyudaImpcta, 3
    For Each Img In Me.imgConcepto
        Img.ToolTipText = "Concepto"
    Next
    For Each Img In Me.imgDiario
        Img.ToolTipText = "Diario"
    Next
    
    
    
    For Each Img In Me.imgDpto
        Img.ToolTipText = "Departamento"
    Next
    
    
    'Limpiamos el tag
    txtCta(6).Tag = ""
    PrimeraVez = True
    FrCobrosPendientesCli.Visible = False
    frpagosPendientes.Visible = False
    FramereclaMail.Visible = False
    FrameAgentes.Visible = False
    FrameDpto.Visible = False
    FrameListRem.Visible = False
    FrameListadoCaja.Visible = False
    FrameDevEfec.Visible = False
    Me.FrameFormaPago.Visible = False
    FrameTransferencias.Visible = False
    Me.FramePrevision.Visible = False
    FrameAseg_Bas.Visible = False
    FrameCobroGenerico.Visible = False
    FrameCompensaciones.Visible = False
    FrameRecepcionDocumentos.Visible = False
    FrameListaRecep.Visible = False
    frameListadoPagosBanco.Visible = False
    FrameDividVto.Visible = False
    FrameReclama.Visible = False
    FrameGastosFijos.Visible = False
    FrameGastosTranasferencia.Visible = False
    FrameCompensaAbonosCliente.Visible = False
    CommitConexion  'Porque son listados. No hay nada dentro transaccion
    
    Select Case Opcion
        
    Case 36
        
        
        H = FrameCompensaAbonosCliente.Height + 120
        W = FrameCompensaAbonosCliente.Width
        FrameCompensaAbonosCliente.Visible = True
        
        
        'cmdVtoDestino(1).Visible = (vUsu.Codigo Mod 100) = 0
        'Label1(1).Visible = (vUsu.Codigo Mod 100) = 0
        cmdVtoDestino(1).Visible = vUsu.Nivel = 0
        Label1(1).Visible = vUsu.Nivel = 0
        
        
    End Select
    
    Me.Width = W + 300
    Me.Height = H + 400
    
    I = Opcion
    If Opcion = 13 Or I = 43 Or I = 44 Then I = 11
    
    'Aseguradas
    If Opcion >= 15 And Opcion <= 18 Then I = 15  'aseguradoas
    If Opcion = 33 Then I = 15 'aseguradoas
    If Opcion = 34 Then I = 23 'Eliminar recepcion documento
    If Opcion = 40 Then I = 39
    Me.cmdCancelar(I).Cancel = True
    
    PonerFrameProgreso

End Sub


Private Sub PonerFrameProgreso()
Dim I As Integer

    'Ponemos el frame al pricnipio de todo
    FrameProgreso.Visible = False
    FrameProgreso.ZOrder 0
    
    'lo ubicamos
    'Posicion horizintal WIDTH
    I = Me.Width - FrameProgreso.Width
    If I > 100 Then
        I = I \ 2
    Else
        I = 0
    End If
    FrameProgreso.Left = I
    'Posicion  VERTICAL HEIGHT
    I = Me.Height - FrameProgreso.Height
    If I > 100 Then
        I = I \ 2
    Else
        I = 0
    End If
    FrameProgreso.Top = I
End Sub





Private Sub Form_Unload(Cancel As Integer)
    If Opcion = 1 Then
        CheckValueGuardar "Listcta", CByte(Abs(Me.optCuenta(0).Value))
        CheckValueGuardar "Infapa", chkApaisado(0)
    End If
    If Opcion = 23 Then CheckValueGuardar "Agrup0", Me.chkAgruparCtaPuente(0)
End Sub

Private Sub frmA_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtAgente(RC).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescAgente(RC).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    DevfrmCCtas = CadenaDevuelta
End Sub

Private Sub frmBa_DatoSeleccionado(CadenaSeleccion As String)
    SQL = CadenaSeleccion
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CInt(RC)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 1)
    DtxtCta(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtFPago(RC).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescFPago(RC).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
    txtSerie(RC).Text = RecuperaValor(CadenaSeleccion, 1)
    
End Sub

Private Sub Image3_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmCta = New frmColCtas
    RC = Index
    frmCta.DatosADevolverBusqueda = "0|1"
    frmCta.ConfigurarBalances = 3
    frmCta.Show vbModal
    Set frmCta = Nothing
    If Index = 17 Then PonerVtosCompensacionCliente
End Sub

Private Sub lwCompenCli_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim C As Currency
Dim Cobro As Boolean

    Cobro = True
    C = Item.Tag
    If Trim(Item.SubItems(6)) = "" Then
        'Es un abono
        Cobro = False
        C = -C
    
    End If
    
    'Si no es checkear cambiamos los signos
    If Not Item.Checked Then C = -C
    
    I = 0
    If Not Cobro Then I = 1
    
    Me.txtimpNoEdit(I).Tag = Me.txtimpNoEdit(I).Tag + C
    txtimpNoEdit(I).Text = Format(Abs(txtimpNoEdit(I).Tag))
    txtimpNoEdit(2).Text = Format(CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag), FormatoImporte)
            
End Sub

Private Sub txtCta_GotFocus(Index As Integer)
    PonFoco txtCta(Index)
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyPressGral KeyAscii
End Sub

Private Sub txtCta_LostFocus(Index As Integer)
Dim Cta As String
Dim B As Byte
    txtCta(Index).Text = Trim(txtCta(Index).Text)
    
    If Index = 6 Then
        'NO se ha cambiado nada de la cuenta
        If txtCta(6).Text = txtCta(6).Tag Then
        
            Exit Sub
        Else
            txtDpto(0).Text = ""
            txtDpto(1).Text = ""
            txtDescDpto(0).Text = ""
            txtDescDpto(0).Text = ""
        End If
    End If
     
     
    If txtCta(Index).Text = "" Then
        DtxtCta(Index).Text = ""
       ' txtCta(6).Tag = txtCta(6).Text
        Exit Sub
    End If
    
    If Index = 6 Then
        If txtCta(0).Text <> "" Or txtCta(1).Text <> "" Then
            MsgBox "Si selecciona desde / hasta cliente no podra seleccionar departamento", vbExclamation
            txtCta(6).Text = ""
            txtCta(6).Tag = txtCta(6).Text
            Exit Sub
        End If
        
    Else
        If Index = 0 Or Index = 1 Then
            If txtCta(6).Text <> "" Then
                MsgBox "Si seleciona departamento no puede seleccionar desde / hasta  cliente", vbExclamation
                txtCta(Index).Text = ""
                txtCta(6).Tag = txtCta(6).Text
                Exit Sub
            End If
        End If
    End If
    
    If Not IsNumeric(txtCta(Index).Text) Then
        MsgBox "La cuenta debe ser numérica: " & txtCta(Index).Text, vbExclamation
        txtCta(Index).Text = ""
        DtxtCta(Index).Text = ""
        txtCta(6).Tag = txtCta(6).Text
        PonerFoco txtCta(Index)
        
        If Index = 17 Then PonerVtosCompensacionCliente
        
        Exit Sub
    End If
    
    Select Case Index
    Case 0 To 7, 11, 12, 15, 16, 18, 19
        'NO hace falta que sean de ultimo nivel
        Cta = (txtCta(Index).Text)
                                '********
        B = CuentaCorrectaUltimoNivelSIN(Cta, SQL)
        If B = 0 Then
            MsgBox "NO existe la cuenta: " & txtCta(Index).Text, vbExclamation
            txtCta(Index).Text = ""
            DtxtCta(Index).Text = ""
        Else
            txtCta(Index).Text = Cta
            DtxtCta(Index).Text = SQL
            If B = 1 Then
                DtxtCta(Index).Tag = ""
            Else
                DtxtCta(Index).Tag = SQL
            End If
            
            
            'Index=1. Cliente en listado de cobros. Si pongo el desde pongo el hasta lo mismo
            If Index = 1 Then
                
                If Len(Cta) = vEmpresa.DigitosUltimoNivel Then
                    txtCta(0).Text = Cta
                    DtxtCta(0).Text = DtxtCta(1).Text
                End If
            End If
            
        End If
    Case Else
        'DE ULTIMO NIVEL
        Cta = (txtCta(Index).Text)
        If CuentaCorrectaUltimoNivel(Cta, SQL) Then
            txtCta(Index).Text = Cta
            DtxtCta(Index).Text = SQL
            
            
        Else
            MsgBox SQL, vbExclamation
            txtCta(Index).Text = ""
            DtxtCta(Index).Text = ""
            txtCta(Index).SetFocus
        End If
        If Index = 17 Then PonerVtosCompensacionCliente
        
    End Select
    txtCta(6).Tag = txtCta(6).Text
End Sub

Private Function ComprobarCuentas(Indice1 As Integer, Indice2 As Integer) As Boolean
Dim L1 As Integer
Dim L2 As Integer
    ComprobarCuentas = False
    If txtCta(Indice1).Text <> "" And txtCta(Indice2).Text <> "" Then
        L1 = Len(txtCta(Indice1).Text)
        L2 = Len(txtCta(Indice2).Text)
        If L1 > L2 Then
            L2 = L1
        Else
            L1 = L2
        End If
        If Val(Mid(txtCta(Indice1).Text & "000000000", 1, L1)) > Val(Mid(txtCta(Indice2).Text & "0000000000", 1, L1)) Then
            MsgBox "Cuenta desde mayor que cuenta hasta.", vbExclamation
            Exit Function
        End If
    End If
    ComprobarCuentas = True
End Function


Private Function ComprobarFechas(Indice1 As Integer, Indice2 As Integer) As Boolean
    ComprobarFechas = False
    If Text3(Indice1).Text <> "" And Text3(Indice2).Text <> "" Then
        If CDate(Text3(Indice1).Text) > CDate(Text3(Indice2).Text) Then
            MsgBox "Fecha inicio mayor que fecha fin", vbExclamation
            Exit Function
        End If
    End If
    ComprobarFechas = True
End Function



Private Sub SubSetFocus(Obje As Object)
    On Error Resume Next
    Obje.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


'Si tiene valor el campo fecha, entonces lo ponemos con el BD
Private Function CampoABD(ByRef T As TextBox, Tipo As String, CampoEnLaBD, Mayor_o_Igual As Boolean) As String

    CampoABD = ""
    If T.Text <> "" Then
        If Mayor_o_Igual Then
            CampoABD = " >= "
        Else
            CampoABD = " <= "
        End If
        Select Case Tipo
        Case "F"
            CampoABD = CampoEnLaBD & CampoABD & "'" & Format(T.Text, FormatoFecha) & "'"
        Case "T"
            CampoABD = CampoEnLaBD & CampoABD & "'" & T.Text & "'"
        Case "N"
            CampoABD = CampoEnLaBD & CampoABD & T.Text
        End Select
    End If
End Function



Private Function CampoBD_A_SQL(ByRef C As ADODB.Field, Tipo As String, Nulo As Boolean) As String

    If IsNull(C) Then
        If Nulo Then
            CampoBD_A_SQL = "NULL"
        Else
            If Tipo = "T" Then
                CampoBD_A_SQL = "''"
            Else
                CampoBD_A_SQL = "0"
            End If
        End If

    Else
    
        Select Case Tipo
        Case "F"
            CampoBD_A_SQL = "'" & Format(C.Value, FormatoFecha) & "'"
        Case "T"
            CampoBD_A_SQL = "'" & DevNombreSQL(C.Value) & "'"
        Case "N"
            CampoBD_A_SQL = TransformaComasPuntos(CStr(C.Value))
        End Select
    End If
End Function



Private Function DesdeHasta(Tipo As String, Desde As Integer, Hasta As Integer, Optional Des As String)
Dim C As String
    C = ""
    Select Case Tipo
    Case "F"
        'Son los text3(desde)....
        If Text3(Desde).Text <> "" Then C = C & "Desde " & Text3(Desde).Text
        
        If Text3(Hasta).Text <> "" Then
            If C <> "" Then C = C & "  "
            C = C & "Hasta " & Text3(Hasta).Text
        End If
        
    Case "C"
        'Cuenta
        If txtCta(Desde).Text <> "" Then C = "Desde " & txtCta(Desde).Text & "-" & DtxtCta(Desde).Text
        
        
        If txtCta(Hasta).Text <> "" Then
            If C <> "" Then C = C & "   "
            C = C & "Hasta " & txtCta(Hasta).Text & "-" & DtxtCta(Hasta).Text
        End If
        
        
    Case "FP"
        'FORMA DE PAGO
        If txtFPago(Desde).Text <> "" Then C = "Desde " & txtFPago(Desde).Text & "-" & txtDescFPago(Desde).Text
        
        
        If txtFPago(Hasta).Text <> "" Then
            If C <> "" Then C = C & "   "
            C = C & "Hasta " & txtFPago(Hasta).Text & "-" & txtDescFPago(Hasta).Text
        End If
    
    Case "BANCO"
        '    'txtCtaBanc  txtDescBanc
        If txtCtaBanc(Desde).Text <> "" Then C = "Desde " & txtCtaBanc(Desde).Text & "-" & txtDescBanc(Desde).Text
        
        If txtCtaBanc(Hasta).Text <> "" Then
            If C <> "" Then C = C & "   "
            C = C & "Hasta " & txtCtaBanc(Hasta).Text & "-" & txtDescBanc(Hasta).Text
        End If
    
    
    Case "S"
        'Serie
        If txtSerie(Desde).Text <> "" Then C = C & "Desde " & txtSerie(Desde).Text
        
        If txtSerie(Hasta).Text <> "" Then
            If C <> "" Then C = C & "  "
            C = C & "Hasta " & txtSerie(Hasta).Text
        End If
    
    Case "NF"
        'Numero factura
        If txtNumfac(Desde).Text <> "" Then C = C & "Desde " & txtNumfac(Desde).Text
        
        If txtNumfac(Hasta).Text <> "" Then
            If C <> "" Then C = C & "  "
            C = C & "Hasta " & txtNumfac(Hasta).Text
        End If
    
    Case "GF"
        'Gasto fijo
        
        If txtGastoFijo(Desde).Text <> "" Then C = C & "Desde " & txtGastoFijo(Desde).Text & " - " & Me.txtDescGastoFijo(Desde).Text
        
        If txtGastoFijo(Hasta).Text <> "" Then
            If C <> "" Then C = C & "  "
            C = C & "Hasta " & txtGastoFijo(Hasta).Text & " - " & Me.txtDescGastoFijo(Hasta).Text
        End If
    
    
    End Select
    If C <> "" Then C = "  " & Des & " " & C
    DesdeHasta = C
End Function


Private Sub PonerFrameProgressVisible(Optional TEXTO As String)
        If TEXTO = "" Then TEXTO = "Generando datos"
        Me.lblPPAL.Caption = TEXTO
        Me.lbl2.Caption = ""
        Me.ProgressBar1.Value = 0
        Me.FrameProgreso.Visible = True
        Me.Refresh
End Sub





'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------

Private Function CobrosPendientesCliente(ByVal ListadoCuentas As String) As Boolean
Dim TieneRemesa As Boolean
Dim RemesaTalones As Boolean
Dim RemesaPagares As Boolean
Dim RemesaEfectos As Boolean
Dim SePuedeRemesar As Boolean
Dim InsertarLinea As Boolean


Dim CadenaInsert As String

    On Error GoTo ECobrosPendientesCliente
    CobrosPendientesCliente = False

    
    'De parametros contapagarepte contatalonpte
    Cad = DevuelveDesdeBD("contatalonpte", "paramtesor", "codigo", "1")
    If Cad = "" Then Cad = "0"
    RemesaTalones = (Val(Cad) = 1)
    
    Cad = DevuelveDesdeBD("contapagarepte", "paramtesor", "codigo", "1")
    If Cad = "" Then Cad = "0"
    RemesaPagares = (Val(Cad) = 1)
    
    Cad = DevuelveDesdeBD("contaefecpte", "paramtesor", "codigo", "1")
    If Cad = "" Then Cad = "0"
    RemesaEfectos = (Val(Cad) = 1)
    

    
    
    
    
    'Trozo basico
    Cad = " FROM scobro ,cuentas,sforpa ,stipoformapago"
    Cad = Cad & " WHERE  scobro.codmacta = cuentas.codmacta"
    Cad = Cad & " AND scobro.codforpa = sforpa.codforpa"
    Cad = Cad & " AND sforpa.tipforpa = stipoformapago.tipoformapago"

    
    
    'Fecha factura
    RC = CampoABD(Text3(1), "F", "fecfaccl", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(Text3(2), "F", "fecfaccl", False)
    If RC <> "" Then Cad = Cad & " AND " & RC



    'Se me habia olvidado
    'A G E N T E
    RC = CampoABD(txtAgente(0), "N", "agente", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(txtAgente(1), "N", "agente", False)
    If RC <> "" Then Cad = Cad & " AND " & RC




    'Fecha vencimiento
    RC = CampoABD(Text3(19), "F", "fecvenci", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(Text3(20), "F", "fecvenci", False)
    If RC <> "" Then Cad = Cad & " AND " & RC

    'SERIE
    RC = CampoABD(txtSerie(0), "T", "numserie", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(txtSerie(1), "T", "numserie", False)
    If RC <> "" Then Cad = Cad & " AND " & RC
    
    'Numero factura
    RC = CampoABD(txtNumfac(0), "T", "codfaccl", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(txtNumfac(1), "T", "codfaccl", False)
    If RC <> "" Then Cad = Cad & " AND " & RC
    


    'Cliente
    RC = CampoABD(txtCta(1), "T", "scobro.codmacta", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(txtCta(0), "T", "scobro.codmacta", False)
    If RC <> "" Then Cad = Cad & " AND " & RC
    
    'Forma PAGO
    RC = CampoABD(txtFPago(0), "T", "scobro.codforpa", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(txtFPago(1), "T", "scobro.codforpa", False)
    If RC <> "" Then Cad = Cad & " AND " & RC
    
    
    'Cliente con departamento
    'If txtCta(0).Text <> "" Then
    '    If cad <> "" Then cad = cad & " AND "
    '    cad = cad & " scobro.codmacta = '" & txtCta(6).Text & "'"
    'End If
    
    'Departamento
    RC = CampoABD(Me.txtDpto(0), "N", "departamento", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(Me.txtDpto(1), "N", "departamento", False)
    If RC <> "" Then Cad = Cad & " AND " & RC
    
    
    'Solo los NO remesar
    If Me.chkNOremesar.Value = 1 Then
        Cad = Cad & " AND noremesar = 1 "
    End If
    
    'Docuemtno recibido y devuelto. Los combos  recedocu  Devuelto
    If Me.cboCobro(0).ListIndex > 0 Then Cad = Cad & " AND recedocu = " & cboCobro(0).ItemData(cboCobro(0).ListIndex)
    If Me.cboCobro(1).ListIndex > 0 Then Cad = Cad & " AND Devuelto = " & cboCobro(1).ItemData(cboCobro(1).ListIndex)
    
    
    'Y lista de cuentas

    If ListadoCuentas <> "" Then
        NumRegElim = 1
        SQL = ""
        Do
            TotalRegistros = InStr(NumRegElim, ListadoCuentas, "|")
            If TotalRegistros > 0 Then
                If NumRegElim > 1 Then SQL = SQL & ","
                NumRegElim = 2
                SQL = SQL & "'" & Mid(ListadoCuentas, 1, TotalRegistros - 1) & "'"
                ListadoCuentas = Mid(ListadoCuentas, TotalRegistros + 1)
            End If
           
            
        Loop Until TotalRegistros = 0
        NumRegElim = 0
        Cad = Cad & " AND scobro.codmacta IN (" & SQL & ")"
    End If
    
    
    
    'Si ha marcado alguna forma de pago
    RC = PonerTipoPagoCobro_(True, False)
    If RC <> "" Then Cad = Cad & " AND tipoformapago IN " & RC
    RC = ""
    
    'Contador
    SQL = "Select count(*) "
    SQL = SQL & Cad
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalRegistros = 0
    If Not RS.EOF Then
        'Total registros
        TotalRegistros = RS.Fields(0)
    End If
    RS.Close
    
    If TotalRegistros = 0 Then
        'NO hay registros
        MsgBox "Ningun dato con esos valores", vbExclamation
        Exit Function
    End If
    
    'Llegados aqui, si k hay registros.
    '100 POR EJEMPLO
    If TotalRegistros > 100 Then
        'Ponemos visible el frame
        MostrarFrame = True
        PonerFrameProgressVisible
    Else
        MostrarFrame = False 'NO lo mostramos
    End If
    
    
    
    'Borramos tmp
    Conn.Execute "DELETE FROM Usuarios.zpendientes where codusu = " & vUsu.Codigo
    
    'Marzo 2015
    'Si agrupamos por forma de pago, agruparemos antes por Tipo de pago
    If chkFormaPago.Value = 1 Then Conn.Execute "DELETE FROM Usuarios.zsimulainm where codusu = " & vUsu.Codigo
    
    
    
    
    
    SQL = "SELECT scobro.*, cuentas.nommacta, nifdatos,stipoformapago.descformapago ,stipoformapago.tipoformapago,nomforpa " & Cad
    ' ----------------
    '  20 Diciembre 2005
    '  Remesados o no remesados
    '
    CONT = 1
    If Me.ChkAgruparSituacion.Value = 1 Then
        '
        CONT = 0
    Else
        If Me.chkEfectosContabilizados.Value = 1 Then CONT = 0
    End If
    If CONT = 1 Then SQL = SQL & " AND codrem is null"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    TieneRemesa = False
    SQL = "INSERT INTO Usuarios.zpendientes (codusu, serie_cta, factura, fecha, numorden,"
    SQL = SQL & "codforpa, nomforpa, codmacta,nombre, fecVto, importe, pag_cob,gastos,vencido,Situacion"
    'Nuevo Enero 2009
    'Si esta apaisado ponemos los departamentos
    If Me.chkApaisado(0).Value = 1 Then
        SQL = SQL & ",coddirec,nomdirec"
    Else
        'Metemos el NIF para futors listados. Pej. El de cobors por cliente lo pondra
        SQL = SQL & ",nomdirec"
    End If
    SQL = SQL & ",devuelto,recibido"
    'SQL = SQL & ",observa) VALUES (" & vUsu.Codigo & ",'"
    'Dic 2013 . Acelerar proceso
    SQL = SQL & ",observa) VALUES "
    
    
    CadenaInsert = "" 'acelerar carga datos
    Fecha = CDate(Text3(0).Text)
    While Not RS.EOF
        CONT = CONT + 1
        If MostrarFrame Then
            lbl2.Caption = "Registro: " & CONT
            lbl2.Refresh
        End If
        
        
        'If Rs!codmacta = "4300019" Then Stop
        
        Cad = RS!NUmSerie & "','" & Format(RS!codfaccl, "0000") & "','" & Format(RS!fecfaccl, FormatoFecha) & "'," & RS!numorden
        
        'Modificacion. Enero 2010. Tiene k aparacer la forma de pago, no el tipo
        'Cad = Cad & "," & Rs!codforpa & ",'" & DevNombreSQL(Rs!descformapago) & "','"
        Cad = Cad & "," & RS!codforpa & ",'" & DevNombreSQL(RS!nomforpa) & "','"
        
        Cad = Cad & RS!codmacta & "','" & DevNombreSQL(RS!Nommacta) & "','"
        Cad = Cad & Format(RS!FecVenci, FormatoFecha) & "',"
        Cad = Cad & TransformaComasPuntos(CStr(RS!ImpVenci)) & ","
        'Cobrado, si no es nulo
        If Not IsNull(RS!impcobro) Then
            Cad = Cad & TransformaComasPuntos(CStr(RS!impcobro))
        Else
            Cad = Cad & "0"
        End If
        
        'Gastos
        Cad = Cad & "," & TransformaComasPuntos(DBLet(RS!Gastos, "N"))
        
        If Fecha > RS!FecVenci Then
            Cad = Cad & ",1"
        Else
            Cad = Cad & ",0"
        End If

        'Hay que añadir la situacion. Bien sea juridica....
        ' Si NO agrupa por situacion, en ese campo metere la referencia del cobro (rs!referencia)
         'vbTalon = 2 vbPagare = 3
        InsertarLinea = True
        
        If Me.ChkAgruparSituacion.Value = 0 Then
            Cad = Cad & ",'" & DevNombreSQL(DBLet(RS!referencia, "T")) & "'"
            
            'Si no agrupa por situacion de vto y no tiene el riesgo marcado
            'Talones pagares
            'Si se ha recepcionado documento, NO deben salir
            
            
            'Enero 2010
            'Comentamos esto ya que tiene la marca de recibidos si/no
'            If Me.chkEfectosContabilizados.Value = 0 Then
'                If Val(Rs!tipoformapago) = vbTalon Or Val(Rs!tipoformapago) = vbPagare Then
'                    If DBLet(Rs!recedocu, "N") = 1 Then InsertarLinea = False
'                End If
'            End If

            
        Else
            'La situacion.
            'Si esta en situacion juridica.
            ' Si no, si esta devuelto y no es una remesa
            ' y luego si es una remesa, sitaucion o devuelto
            If RS!situacionjuri = 1 Then
                Cad = Cad & ",'SITUACION JURIDICA'"
            Else
                'Cambio Marzo 2009
                ' Ahora tb se remesan los pagares y talones
                
                If Not IsNull(RS!siturem) Then
                    TieneRemesa = True
                    Cad = Cad & ",'R" & Format(RS!AnyoRem, "0000") & Format(RS!CodRem, "0000000000") & "'"
                    
                Else
                    
                    If RS!Devuelto = 1 Then
                        Cad = Cad & ",'DEVUELTO'"
                    Else
                            
                        SePuedeRemesar = False
                        If RemesaEfectos Then SePuedeRemesar = RS!tipoformapago = vbTipoPagoRemesa
                        If RemesaPagares Then SePuedeRemesar = RS!tipoformapago = vbPagare
                        If RemesaTalones Then SePuedeRemesar = RS!tipoformapago = vbTalon
                        
                    
                        If Not SePuedeRemesar Then
                            Cad = Cad & ",'PENDIENTE COBRO'"
                        Else
                            Cad = Cad & ",'PENDIENTE REMESAR'" '& Rs!anyorem
                        End If
                        
                        
                        
                        'Talones pagares
                        'Si se ha recepcionado documento, NO deben salir
                        'ENERO 2010
                        'Tiene la marca de recibidos
                        
                        'If Val(Rs!tipoformapago) = vbTalon Or Val(Rs!tipoformapago) = vbPagare Then
                        '    If Me.chkEfectosContabilizados.Value = 0 Then
                        '        If DBLet(Rs!recedocu, "N") = 1 Then InsertarLinea = False
                        '    End If
                        'End If
                        
                    
                    End If  'De devuelto
               End If  'de SITUREM null
            End If ' de situacion juridica
        End If  'de agrupa por sitacuib
        Cad = Cad & ","
        If Me.chkApaisado(0).Value = 1 Then
            'SI carga departamentos. Esto podriamos mejorar la velocidad si
            'pregarmos un rs o en la select linkamos con departamento...
            If IsNull(RS!departamento) Then
                Cad = Cad & "NULL,NULL,"
            Else
                Cad = Cad & "'" & RS!departamento & "','"
                Cad = Cad & DevNombreSQL(DevuelveDesdeBD("Descripcion", "departamentos", "codmacta = '" & RS!codmacta & "' AND dpto", RS!departamento, "N")) & "',"
            End If
            
        Else
            'Nif datos
            'Stop
             Cad = Cad & "'" & DevNombreSQL(DBLet(RS!nifdatos, "T")) & "',"
        End If
        
        If DBLet(RS!Devuelto, "N") = 0 Then
            Cad = Cad & "'',"
        Else
            Cad = Cad & "'S',"
        End If
        If DBLet(RS!recedocu, "N") = 0 Then
            Cad = Cad & "''"
        Else
            Cad = Cad & "'S'"
        End If
            
        Cad = Cad & ",'"
        If Me.ChkObserva.Value Then
            Cad = Cad & DevNombreSQL(DBLet(RS!obs, "T"))
'        Else
'            Cad = Cad & "''"
        End If
        Cad = Cad & "')"
        
        If InsertarLinea Then
        
            CadenaInsert = CadenaInsert & ", (" & vUsu.Codigo & ",'" & Cad
        
            If Len(CadenaInsert) > 20000 Then
                Cad = SQL & Mid(CadenaInsert, 2)
                Conn.Execute Cad
                CadenaInsert = ""
            End If
            'Cad = SQL & Cad
            'Conn.Execute Cad
        Else
            'Tiramos para atras el contador total
            CONT = CONT - 1
        End If
        'Siguiente
        RS.MoveNext
    Wend
    RS.Close
    
    If Len(CadenaInsert) > 0 Then
        Cad = SQL & Mid(CadenaInsert, 2)
        Conn.Execute Cad
        CadenaInsert = ""
    End If

    
    'Si esta seleccacona SITIACUIN VENCIMIENTO
    ' y tenia remesas , entonces updateo la tabla poniendo
    ' la situacion de la remesa
    If TieneRemesa Then
        Cad = "Select codigo,anyo,  descsituacion"
        Cad = Cad & " from remesas left join tiposituacionrem on situacion=situacio"
        RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            Debug.Print RS!Codigo
            If Not IsNull(RS!descsituacion) Then
                Cad = "R" & Format(RS!Anyo, "0000") & Format(RS!Codigo, "0000000000")
                Cad = " WHERE situacion='" & Cad & "'"
                Cad = "UPDATE Usuarios.zpendientes set Situacion='Remesados: " & RS!descsituacion & "' " & Cad
                Conn.Execute Cad
            End If
            RS.MoveNext
        Wend
        RS.Close
    End If
    
    'Marzo 2015.
    'Nivel de anidacion para los agrupados por forma de pago
    ' que es TIPO DE PAGO
    If chkFormaPago.Value = 1 Then
    
        Cad = "select codforpa from Usuarios.zpendientes where codusu =" & vUsu.Codigo & " group by 1"
        RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not RS.EOF
            Cad = Cad & ", " & RS!codforpa
            RS.MoveNext
        Wend
        RS.Close
        
        If Cad <> "" Then
            Cad = Mid(Cad, 2)
            Cad = " and codforpa IN (" & Cad & ")"
            Cad = " FROM sforpa , stipoformapago WHERE sforpa.tipforpa=stipoformapago.tipoformapago" & Cad
            Cad = "SELECT " & vUsu.Codigo & ",codforpa,tipoformapago,descformapago " & Cad
            Cad = "INSERT INTO Usuarios.zsimulainm(codusu,codigo,conconam,nomconam) " & Cad
        
            Conn.Execute Cad
        End If
    End If
    
    
    
    'Voy a comprobar si ha metido algun registo
    espera 0.3
    SQL = "Select count(*) FROM  Usuarios.zpendientes where codusu = " & vUsu.Codigo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    If Not RS.EOF Then CONT = DBLet(RS.Fields(0), "N")
    RS.Close
    If CONT = 0 Then
        MsgBox "No se ha generado ningun valor(2)", vbExclamation
    Else
        CobrosPendientesCliente = True 'Para imprimir
    End If
    Exit Function
ECobrosPendientesCliente:
    MuestraError Err.Number, Err.Description
End Function



Private Function PagosPendienteProv(ByVal ListadoCuentas As String) As Boolean
'Dim MismaClavePrimaria As String


    On Error GoTo EPagosPendienteProv
    PagosPendienteProv = False
    
    'Trozo basico de PAGOS
    Cad = "FROM spagop ,cuentas ,sforpa,stipoformapago"
    Cad = Cad & " WHERE spagop.ctaprove = cuentas.codmacta"
    Cad = Cad & " AND spagop.codforpa = sforpa.codforpa"
    Cad = Cad & " AND sforpa.tipforpa = stipoformapago.tipoformapago"

    
    'Fecha
    RC = CampoABD(Text3(3), "F", "fecefect", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(Text3(4), "F", "fecefect", False)
    If RC <> "" Then Cad = Cad & " AND " & RC

    'Cliente
    RC = CampoABD(txtCta(2), "T", "ctaprove", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(txtCta(3), "T", "ctaprove", False)
    If RC <> "" Then Cad = Cad & " AND " & RC
    
    
    'FORMA PAGO
    RC = CampoABD(txtFPago(6), "N", "spagop.codforpa", True)
    If RC <> "" Then Cad = Cad & " AND " & RC
    RC = CampoABD(txtFPago(7), "N", "spagop.codforpa", False)
    If RC <> "" Then Cad = Cad & " AND " & RC
    
    
    
    
    
    
    'Y lista de cuentas

    If ListadoCuentas <> "" Then
        NumRegElim = 1
        SQL = ""
        Do
            TotalRegistros = InStr(NumRegElim, ListadoCuentas, "|")
            If TotalRegistros > 0 Then
                If NumRegElim > 1 Then SQL = SQL & ","
                NumRegElim = 2
                SQL = SQL & "'" & Mid(ListadoCuentas, 1, TotalRegistros - 1) & "'"
                ListadoCuentas = Mid(ListadoCuentas, TotalRegistros + 1)
            End If
           
            
        Loop Until TotalRegistros = 0
        NumRegElim = 0
        Cad = Cad & " AND spagop.ctaprove IN (" & SQL & ")"
        
    End If
    
    
    'ORDEN
    Cad = Cad & " ORDER BY numfactu"
   
    
    
    
    'Contador
    SQL = "Select count(*) "
    SQL = SQL & Cad
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TotalRegistros = 0
    If Not RS.EOF Then
        'Total registros
        TotalRegistros = RS.Fields(0)
    End If
    RS.Close
    
    If TotalRegistros = 0 Then
        'NO hay registros
        MsgBox "Ningun dato con esos valores", vbExclamation
        Exit Function
    End If
    
    'Llegados aqui, si k hay registros.
    '100 POR EJEMPLO
    If TotalRegistros > 100 Then
        'Ponemos visible el frame
        MostrarFrame = True
        PonerFrameProgressVisible
    Else
        MostrarFrame = False 'NO lo mostramos
    End If
    
    
    
    'Borramos tmp
    Conn.Execute "DELETE FROM Usuarios.zpendientes where codusu = " & vUsu.Codigo
    
    SQL = "SELECT spagop.*, cuentas.nommacta, stipoformapago.descformapago, stipoformapago.siglas,nomforpa " & Cad
    
    'Cambiamos
''''''    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''''''    'Compruebo si hay repetidos en fecfactu|numfactu|siglas
''''''    SQL = ""
''''''    MismaClavePrimaria = "|"
''''''    While Not RS.EOF
''''''        RC = Format(RS!fecfactu, "yymmdd") & RS!siglas & RS!numfactu
''''''        If RC = SQL Then
''''''            RC = RC & "|"
''''''            If InStr(1, MismaClavePrimaria, "|" & RC) = 0 Then MismaClavePrimaria = MismaClavePrimaria & RC
''''''        Else
''''''            SQL = RC
''''''        End If
''''''        RS.MoveNext
''''''    Wend
''''''    SQL = RS.Source
''''''    RS.Close
    
    'Vuelvo a abrirlo
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Agosto 2013
    'Añadimos en campo SITUACION donde pondra si esta emitido o no (emitdocum)
    
    'Mayo 2014
    'La factura la metemos en nomdirec. Asi NO da error duplicados
    
    CONT = 0
    SQL = "INSERT INTO Usuarios.zpendientes (codusu, serie_cta, factura, fecha, numorden,nomdirec,"
    SQL = SQL & "codforpa, nomforpa, codmacta,nombre, fecVto, importe, pag_cob,vencido,situacion) VALUES (" & vUsu.Codigo & ",'"
    Fecha = CDate(Text3(5).Text)
    DevfrmCCtas = ""
    While Not RS.EOF
        CONT = CONT + 1
        If MostrarFrame Then
            lbl2.Caption = "Registro: " & CONT
            lbl2.Refresh
        End If
        
'        'Por si se repiten misma fecfactura, mismo numero factura, mismo tipo de pago
'        If MismaClavePrimaria <> "" Then
'            'Hay claves repetidas no tiene en cuenta el vto
'            RC = "|" & Format(RS!fecfactu, "yymmdd") & RS!siglas & RS!numfactu & "|"
'            'Enero 2011
'            RC = "|" & Format(RS!fecfactu, "yymmdd") & RS!siglas & RS!numfactu & "#" & RS!numorden & "|"
'
'
'            If InStr(1, MismaClavePrimaria, RC) > 0 Then
'                RC = DevNombreSQL(RS!numfactu)
'                RC = FijaNumeroFacturaRepetido(RC)
'                Cad = RS!siglas & "','" & RC & "','" & Format(RS!fecfactu, FormatoFecha) & "'," & RS!numorden
'            Else
'                'El mismo de abajo
'                Cad = RS!siglas & "','" & DevNombreSQL(RS!numfactu) & "','" & Format(RS!fecfactu, FormatoFecha) & "'," & RS!numorden
'            End If
'        Else
'            'El procedimiento de antes
'            Cad = RS!siglas & "','" & DevNombreSQL(RS!numfactu) & "','" & Format(RS!fecfactu, FormatoFecha) & "'," & RS!numorden
'        End If
'
        
        'mayo 2014
        Cad = RS!siglas & "','" & Format(CONT, "00000") & "','" & Format(RS!FecFactu, FormatoFecha) & "'," & RS!numorden & ",'" & DevNombreSQL(RS!NumFactu) & "'"
        
        
        'optMostraFP
        Cad = Cad & "," & RS!codforpa & ",'"
        If Me.optMostraFP(0).Value Then
            Cad = Cad & DevNombreSQL(RS!descformapago)
        Else
            Cad = Cad & DevNombreSQL(RS!nomforpa)
        End If
        Cad = Cad & "','" & RS!ctaprove & "','" & DevNombreSQL(RS!Nommacta) & "','"
        Cad = Cad & Format(RS!fecefect, FormatoFecha) & "',"
        Cad = Cad & TransformaComasPuntos(CStr(RS!ImpEfect)) & ","
        'Cobrado, si no es nulo
        If Not IsNull(RS!imppagad) Then
            Cad = Cad & TransformaComasPuntos(CStr(RS!imppagad))
        Else
            Cad = Cad & "0"
        End If
        If Fecha > RS!fecefect Then
            Cad = Cad & ",1"
        Else
            Cad = Cad & ",0"
        End If
        
        'Agosto 2013
        'Si esta en un tal-pag
        Cad = Cad & ",'"
        If DBLet(RS!emitdocum, "N") > 0 Then Cad = Cad & "*"
        
        Cad = Cad & "')"  'lleva el apostrofe
        Cad = SQL & Cad
        Conn.Execute Cad
        'Siguiente
        RS.MoveNext
    Wend
    RS.Close
     
    PagosPendienteProv = True 'Para imprimir
    Exit Function
EPagosPendienteProv:
    MuestraError Err.Number, Err.Description
End Function



Private Function FijaNumeroFacturaRepetido(Numerofactura) As String
Dim I As Integer
Dim AUX As String
        If Len(Numerofactura) >= 10 Then
            MsgBox "Clave duplicada. Imposible insertar. " & RS!NumFactu & ": " & RS!FecFactu, vbExclamation
            FijaNumeroFacturaRepetido = Numerofactura
            Exit Function
        End If
        
        'Añadiremos guienos por detras
        For I = Len(Numerofactura) To 10
            'Añadirenos espacios en blanco al final
            AUX = RS!NumFactu & String(I - Len(Numerofactura), "_")
            If InStr(1, DevfrmCCtas, "|" & AUX & "|") = 0 Then
                'Devolvemos este numero de factura
                FijaNumeroFacturaRepetido = AUX
                If DevfrmCCtas = "" Then DevfrmCCtas = "|"
                DevfrmCCtas = DevfrmCCtas & AUX & "|"
                Exit Function
            End If
            
        Next I
        
        'Si llega aqui probaremos con los -
        For I = Len(Numerofactura) + 1 To 10
            'Añadirenos espacios en blanco al final
            AUX = String(I - Len(Numerofactura), "_") & RS!NumFactu
            If InStr(1, DevfrmCCtas, "|" & AUX & "|") = 0 Then
                'Devolvemos este numero de factura
                FijaNumeroFacturaRepetido = AUX
                DevfrmCCtas = DevfrmCCtas & AUX & "|"
                Exit Function
            End If
            
        Next I
End Function




Private Function ListadoRemesas() As Boolean
Dim AUX As String
    On Error GoTo EListadoRemesas
    ListadoRemesas = False
    
    SQL = ""
    RC = CampoABD(txtRem(0), "N", "codigo", True)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(1), "N", "codigo", False)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(2), "N", "anyo", True)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(3), "N", "anyo", False)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    'Tipo remesa
    RC = RemesaSeleccionTipoRemesa(chkTipoRemesa(0).Value = 1, chkTipoRemesa(1).Value = 1, chkTipoRemesa(2).Value = 1)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    Set RS = New ADODB.Recordset
    
    'ANTES
    RC = "SELECT remesas.*,nommacta from remesas,cuentas "
    RC = RC & " WHERE remesas.codmacta = cuentas.codmacta"
    If SQL <> "" Then RC = RC & " AND " & SQL
    
    
    'AHORA
    RC = "Select codigo,anyo, fecremesa,tiporemesa.descripcion as d1,descsituacion,remesas.codmacta,nommacta,"
    RC = RC & " Importe , remesas.descripcion, remesas.Tipo,situacion,tiporem"
    RC = RC & " from cuentas,tiposituacionrem,remesas left join tiporemesa on remesas.tipo=tiporemesa.tipo where remesas.codmacta=cuentas.codmacta"
    RC = RC & " and situacio=situacion"
    If SQL <> "" Then RC = RC & " AND " & SQL
    
    
    RS.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        MsgBox "Ninguna remesa para listar", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Function
    End If
    
    SQL = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "Delete from Usuarios.ztmplibrodiario where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    Set miRsAux = New ADODB.Recordset
    
    
    'Para insertar en una                       codigo /año desc                tipo   situacion
    SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2,texto3,texto4, texto5,importe1,  fecha1) VALUES ("
    'Para insertar en la otra
    Cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien, fechaent,linliapu, codmacta, nommacta, numdocum, ampconce, debe,haber) VALUES ("
    
    
    
    TotalRegistros = 0
    
    While Not RS.EOF
        'Insertamos la cabecera de la remesas
        'Para ello el codigo sera: 200500001   es decir remesa 1 del año 2005 ...
        CONT = RS!Anyo * 100000 + RS!Codigo
        
        
        RC = vUsu.Codigo & "," & CONT & ",'" & DevNombreSQL(DBLet(RS!Descripcion, "T")) & "','" & DevNombreSQL(RS!Nommacta) & "',"
        'TIPO   situacion
        
        RC = RC & "'" & DBLet(RS!d1, "t") & "','" & DBLet(RS!descsituacion, "T") & "','"
        
        'Tipo remesa
        If RS!Tiporem = 2 Then
            RC = RC & "PAG"
        ElseIf RS!Tiporem = 3 Then
            RC = RC & "TAL"
        Else
            RC = RC & "EFE"
        End If
        RC = RC & "'," & TransformaComasPuntos(CStr(RS!Importe)) & ",'" & Format(RS!fecremesa, FormatoFecha) & "')"
    
        RC = SQL & RC
        Conn.Execute RC
       
        I = 1
        If Me.chkRem(0).Value = 1 Then
            'fecfaccl,scobro.codmacta,siturem,impcobro,impvenci,gastos,codfaccl,numserie,codbanco,codsucur,digcontr,cuentaba
            RC = "SELECT fecfaccl,scobro.codmacta,siturem,impcobro,impvenci,gastos,codfaccl,numserie,codbanco,codsucur,digcontr,scobro.cuentaba,nommacta"
            RC = RC & " ,fecvenci,scobro.iban from scobro,cuentas where codrem=" & RS!Codigo & " AND anyorem =" & RS!Anyo
            RC = RC & " AND cuentas.codmacta = scobro.codmacta  ORDER BY "
            If Me.optOrdenRem(1).Value Then
                'Codmacta
                RC = RC & "scobro.codmacta,numserie,codfaccl,fecfaccl"
            ElseIf Me.optOrdenRem(2).Value Then
                'Nommacta
                RC = RC & "nommacta,numserie,codfaccl,fecfaccl"
            ElseIf Me.optOrdenRem(0).Value Then
                'Numero factura
                RC = RC & "numserie,codfaccl,fecfaccl"
            Else
                'fcto
                RC = RC & "fecvenci,numserie,codfaccl,fecfaccl"
            
            End If
            miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                'If CONT = 200900004 Then Stop
                'cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien,
                'fechaent,linliapu, codmacta, nommacta, numdocum, ampconce, debe
                RC = vUsu.Codigo & "," & CONT & ",'" & Format(miRsAux!fecfaccl, FormatoFecha) & "',"
                RC = RC & I & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
                'Importe = miRsAux!impvenci - DBLet(miRsAux!impcobro, "N") + DBLet(miRsAux!Gastos, "N")
                If miRsAux!siturem > "B" Then
                    'No deberia ser NULL
                    Importe = DBLet(miRsAux!impcobro, "N")
                Else
                    Importe = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N")
                End If
                RC = RC & Format(miRsAux!codfaccl, "000000000") & "','"
                
                'Aqui pondre el CCC para los efectos
                '---------------------------------------------------
                'rc=rc & codbanco,codsucur,digcontr,scobro.cuentaba
                AUX = ""
                If RS!Tiporem = 1 Then
                        If IsNull(miRsAux!codbanco) Then
                            AUX = "0000"
                        Else
                            AUX = Format(miRsAux!codbanco, "0000")
                        End If
                        'AUX = AUX & " "
                        If IsNull(miRsAux!codsucur) Then
                            AUX = AUX & "0000"
                        Else
                            AUX = AUX & Format(miRsAux!codsucur, "0000")
                        End If
                        'AUX = AUX & " "
                        If IsNull(miRsAux!digcontr) Then
                            AUX = AUX & "**"
                        Else
                            AUX = AUX & Format(miRsAux!digcontr, "00")
                        End If
                        'AUX = AUX & " "
                        If IsNull(miRsAux!Cuentaba) Then
                            AUX = AUX & "0000"
                        Else
                            AUX = AUX & Format(miRsAux!Cuentaba, "0000000000")
                        End If
                Else
                    'Talon / Pagare. Si tiene numero puesto lo pondre
                 
                End If
                
                'Nuevo ENERO 2010
                'Fecha vto
                AUX = DBLet(miRsAux!IBAN, "T") & AUX
                If Len(AUX) > 24 Then AUX = Mid(AUX, 1, 24)
                AUX = AUX & "|" & Format(miRsAux!FecVenci, "dd/mm/yy")
                
                RC = RC & AUX & "'," & TransformaComasPuntos(CStr(Importe))
                
                'En el haber pongo el ascii de la serie
                '--------------------------------------
                RC = RC & "," & Asc(Left(DBLet(miRsAux!NUmSerie, "T") & " ", 1)) & ")"
                RC = Cad & RC
            
                Conn.Execute RC
            
                'Sig
                I = I + 1
                miRsAux.MoveNext
            Wend
            miRsAux.Close
       
        Else
            'Tenemos k insertar una unica linea a blancos
            RC = vUsu.Codigo & "," & CONT & ",'1999-12-31'," & I & ",'','','','',0,0)"
            RC = Cad & RC
            
            Conn.Execute RC
        End If
        TotalRegistros = TotalRegistros + 1
        RS.MoveNext
    Wend
    
    RS.Close
    
    
    
    
    Set RS = Nothing
    Set miRsAux = Nothing
    
    If Me.chkRem(0).Value = 1 Then
        If TotalRegistros = 0 Then
            MsgBox "No hay vencimientos asociados a las remesas", vbExclamation
            Exit Function
        End If
    End If
    ListadoRemesas = True
    Exit Function
EListadoRemesas:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
    Set miRsAux = Nothing

End Function









Private Function ListadoRemesasBanco() As Boolean
Dim AUX As String
Dim Cad2 As String
Dim J As Integer
    On Error GoTo EListadoRemesas
    ListadoRemesasBanco = False
    
    SQL = ""
    RC = CampoABD(txtRem(0), "N", "codigo", True)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(1), "N", "codigo", False)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(2), "N", "anyo", True)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtRem(3), "N", "anyo", False)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    'Tipo remesa
    RC = RemesaSeleccionTipoRemesa(chkTipoRemesa(0).Value = 1, chkTipoRemesa(1).Value = 1, chkTipoRemesa(2).Value = 1)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    Set RS = New ADODB.Recordset
    
    'ANTES
    RC = "SELECT remesas.*,nommacta from remesas,cuentas "
    RC = RC & " WHERE remesas.codmacta = cuentas.codmacta"
    If SQL <> "" Then RC = RC & " AND " & SQL
    
    
    'AHORA
    RC = "Select codigo,anyo, fecremesa,tiporemesa.descripcion as d1,descsituacion,remesas.codmacta,nommacta,"
    RC = RC & " Importe , remesas.descripcion, remesas.Tipo,situacion,tiporem"
    RC = RC & " from cuentas,tiposituacionrem,remesas left join tiporemesa on remesas.tipo=tiporemesa.tipo where remesas.codmacta=cuentas.codmacta"
    RC = RC & " and situacio=situacion"
    If SQL <> "" Then RC = RC & " AND " & SQL
    
    
    RS.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        MsgBox "Ninguna remesa para listar", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Function
    End If
    
    SQL = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "Delete from Usuarios.ztmplibrodiario where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    Set miRsAux = New ADODB.Recordset
    
    
    'Para insertar en una                       codigo /año desc                tipo   situacion
    SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2,texto3,texto4, texto5,importe1,  fecha1,observa1) VALUES ("
    
    
    
    TotalRegistros = 0
    
    While Not RS.EOF
        'Insertamos la cabecera de la remesas
        'Para ello el codigo sera: 200500001   es decir remesa 1 del año 2005 ...
        CONT = RS!Anyo * 100000 + RS!Codigo
        
        
        RC = vUsu.Codigo & "," & CONT & ",'" & DevNombreSQL(DBLet(RS!Descripcion, "T")) & "','" & DevNombreSQL(RS!Nommacta) & "',"
        'TIPO   situacion
        
        RC = RC & "'" & DBLet(RS!d1, "t") & "','" & DBLet(RS!descsituacion, "T") & "','"
        
        'Tipo remesa
        If RS!Tiporem = 2 Then
            RC = RC & "PAG"
        ElseIf RS!Tiporem = 3 Then
            RC = RC & "TAL"
        Else
            RC = RC & "EFE"
        End If
        RC = RC & "'," & TransformaComasPuntos(CStr(RS!Importe)) & ",'" & Format(RS!fecremesa, FormatoFecha) & "','"
        
        Cad2 = "Select * from ctabancaria where codmacta = '" & RS!codmacta & "'"
        miRsAux.Open Cad2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad2 = "NO ENCONTRADO"
        If Not miRsAux.EOF Then
            Cad2 = Trim(DBLet(miRsAux!IBAN, "T") & " ") & Format(DBLet(miRsAux!Entidad, "N"), "0000") & " " & Format(DBLet(miRsAux!Oficina, "N"), "0000") & " "
            If IsNull(miRsAux!Control) Then
                Cad2 = Cad2 & "*"
            Else
                Cad2 = Cad2 & miRsAux!Control
            End If
            Cad2 = Cad2 & " " & Format(DBLet(miRsAux!CtaBanco, "N"), "0000000000")
        End If
        miRsAux.Close
        RC = RC & Cad2 & "')"
        'ctabancaria(entidad,oficina,control,ctabanco)
        Cad2 = ""
        
        RC = SQL & RC
        Conn.Execute RC
       
        I = 1
        
            'Voy a comprobar que existen
            RC = "SELECT codmacta,reftalonpag FROM scobro "
            RC = RC & "  WHERE codrem=" & RS!Codigo & " AND anyorem =" & RS!Anyo
            RC = RC & " GROUP BY 1,2 "
            miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Cad2 = ""
            While Not miRsAux.EOF
                Cad2 = Cad2 & " scarecepdoc.codmacta = '" & miRsAux!codmacta & "' AND numeroref = '" & DevNombreSQL(miRsAux!reftalonpag) & "'|"
                miRsAux.MoveNext
            Wend
            miRsAux.Close
            
            If Cad2 = "" Then
                MsgBox "Error obteniendo cuenta/referenciatalon", vbExclamation
                RS.Close
                Exit Function
            End If
                
            'Comprobare que existen y de paso los inserto
            While Cad2 <> ""
                J = InStr(1, Cad2, "|")
                AUX = Mid(Cad2, 1, J - 1)
                Cad2 = Mid(Cad2, J + 1)
                
                'RC = "SELECT * FROM scarecepdoc ,slirecepdoc,cuentas WHERE codigo=id AND scarecepdoc.codmacta=cuentas.codmacta AND " & Aux
                RC = "SELECT * FROM scarecepdoc ,cuentas WHERE  scarecepdoc.codmacta=cuentas.codmacta AND " & AUX
               
                miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If miRsAux.EOF Then
                    MsgBox "No se encuentra la referencia; " & AUX, vbExclamation
                    miRsAux.Close
                    RS.Close
                    Exit Function
                End If
                
                While Not miRsAux.EOF
            
                
                
                
                
                    
                    'Para insertar en la otra
                    Cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien, fechaent,linliapu,  nommacta,codmacta, numdocum, ampconce, debe,haber) VALUES ("
                
                    'Trampas:  Entre codmacta numdocum llevare el numero de talon. Ya que suman 20 y reftal es len20
                    RC = vUsu.Codigo & "," & CONT & ",'" & Format(miRsAux!fechavto, FormatoFecha) & "',"
                    RC = RC & I & ",'" & DevNombreSQL(miRsAux!Nommacta) & "','"
                    Importe = DBLet(miRsAux!Importe, "N")
                    
                    'Referencia talon
                    AUX = DevNombreSQL(miRsAux!numeroref)
                    If Len(AUX) > 10 Then
                        RC = RC & Mid(AUX, 1, 10) & "','" & Mid(AUX, 11)
                    Else
                        RC = RC & AUX & "','"
                    End If
                    'Banco
                    RC = RC & "','" & DevNombreSQL(miRsAux!banco) & "',"
                    
                    'Talon / Pagare. Si tiene numero puesto lo pondre
                    RC = RC & TransformaComasPuntos(CStr(Importe))
                    
                    'En el haber pongo el ascii de la serie
                    '--------------------------------------
                    RC = RC & ",0)"
                    RC = Cad & RC
                
                    Conn.Execute RC
                
                    'Sig
                    I = I + 1
                    miRsAux.MoveNext
                Wend
                miRsAux.Close
            Wend

        TotalRegistros = TotalRegistros + 1
        RS.MoveNext
    Wend
    
    RS.Close
    
    
    
    
      Set RS = Nothing
    Set miRsAux = Nothing
    
    If Me.chkRem(0).Value = 1 Then
        If TotalRegistros = 0 Then
            MsgBox "No hay vencimientos asociados a las remesas", vbExclamation
            Exit Function
        End If
    End If
    ListadoRemesasBanco = True
    Exit Function
EListadoRemesas:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
    Set miRsAux = Nothing

End Function




'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'
'               CREDITO CAUCION
'
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

Private Function ListadoTransferencias() As Boolean
Dim Importe As Currency

    On Error GoTo EListadoTransferencias
    ListadoTransferencias = False
    
    SQL = ""
    RC = CampoABD(txtNumero(0), "N", "codigo", True)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    RC = CampoABD(txtNumero(1), "N", "codigo", False)
    If RC <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & RC
    End If
    
    
    Cad = RC
    
    Set RS = New ADODB.Recordset
    
    RC = "SELECT stransfer.*,nommacta from stransfer"
    If Opcion = 13 Then RC = RC & "cob"
    RC = RC & " as stransfer,cuentas "
    RC = RC & " WHERE stransfer.codmacta = cuentas.codmacta"
    If SQL <> "" Then RC = RC & " AND " & SQL
    
    RS.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        MsgBox "Ninguna valor para listar", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Function
    End If
    
    SQL = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = "Delete from Usuarios.ztmplibrodiario where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    If Opcion = 13 Then Conn.Execute "Delete from usuarios.zcuentas where codusu =" & vUsu.Codigo
        
    
    Set miRsAux = New ADODB.Recordset
    
    
    'Para insertar en una                       codigo /año desc
    SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2, importe1,  fecha1) VALUES ("
    'Para insertar en la otra
    Cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien, fechaent,linliapu, codmacta, nommacta, numdocum, ampconce, debe) VALUES ("
    
    
    

    
    While Not RS.EOF
        'Insertamos la cabecera de la remesas
        'Para ello el codigo sera: 200500001   es decir remesa 1 del año 2005 ...
        CONT = RS!Codigo
        
        
        RC = vUsu.Codigo & "," & CONT & ",'" & DevNombreSQL(DBLet(RS!Descripcion, "T")) & "','" & DevNombreSQL(RS!Nommacta) & "',"
        RC = RC & TransformaComasPuntos("0") & ",'" & Format(RS!Fecha, FormatoFecha) & "')"
    
        RC = SQL & RC
        Conn.Execute RC
       
        I = 1
     
            
            If Opcion = 13 Then
                RC = "scobro"
            Else
                RC = "spagop"
            End If
            RC = "SELECT " & RC & ".*,nommacta from cuentas," & RC
            RC = RC & " WHERE transfer = " & RS!Codigo
            RC = RC & " AND cuentas.codmacta = "
            If Opcion = 13 Then
                RC = RC & " scobro.codmacta "
                RC = RC & " ORDER BY scobro.codmacta,fecfaccl"
            Else
                RC = RC & " spagop.ctaprove "
                RC = RC & " ORDER BY ctaprove,fecfactu"
            End If
            miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not miRsAux.EOF
                'cad = "INSERT INTO Usuarios.ztmplibrodiario (codusu,  numasien,
                'fechaent,linliapu, codmacta, nommacta, numdocum, ampconce, debe "
                If Opcion = 13 Then
                    Fecha = miRsAux!fecfaccl
                Else
                    Fecha = miRsAux!FecFactu
                End If
                RC = vUsu.Codigo & "," & CONT & ",'" & Format(Fecha, FormatoFecha) & "',"
                RC = RC & I & ",'"
                If Opcion = 13 Then
                    RC = RC & miRsAux!codmacta
                Else
                    RC = RC & miRsAux!ctaprove
                End If
                
                RC = RC & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
                
                
                'Cuenta
                If Opcion <> 13 Then
                    RC = RC & DevNombreSQL(miRsAux!NumFactu) & "','"
                    
                    'Noviembre 2013
                    'Añadimos el IBAN
                    
                    RC = RC & Trim(DBLet(miRsAux!IBAN, "T") & " " & Format(DBLet(miRsAux!Entidad, "T"), "0000")) & " " & Format(DBLet(miRsAux!Oficina, "T"), "0000") & " "
                    RC = RC & Mid(DBLet(miRsAux!CC, "T") & "**", 1, 2) & " " & Right(String(10, "0") & DBLet(miRsAux!Cuentaba, "T"), 10)
                    Importe = miRsAux!ImpEfect - (DBLet(miRsAux!imppagad, "N"))
                    RC = RC & "'," & TransformaComasPuntos(CStr(Importe)) & ")"
                Else
                    RC = RC & DevNombreSQL(miRsAux!codfaccl) & "','"
                    
                    CadenaDesdeOtroForm = "NO"
                    If DBLet(miRsAux!codbanco, "N") > 0 Then
                        'Es especifico para ESCALONO, pero no molesta a nadie
                        If DBLet(miRsAux!Cuentaba, "T") = "8888888888" Then
                            'Seguira poniendo  cuenta no domiciliada
                        Else
                            CadenaDesdeOtroForm = ""
                        End If
                    End If
                    If CadenaDesdeOtroForm = "" Then
                        'OK, ponemos la cuenta
                        CadenaDesdeOtroForm = Trim(DBLet(miRsAux!IBAN, "T") & " " & Format(DBLet(miRsAux!codbanco, "N"), "0000")) & " " & Format(DBLet(miRsAux!codsucur, "N"), "0000") & " "
                        CadenaDesdeOtroForm = CadenaDesdeOtroForm & "**  ******" & Right(String(4, "0") & DBLet(miRsAux!Cuentaba, "T"), 4)
                        
                    Else
                        'CUENTANODOMICILIADA
                        CadenaDesdeOtroForm = "NODOM"  'en el rpt haremos un replace
                    End If
                    RC = RC & CadenaDesdeOtroForm
                    Importe = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N") - DBLet(miRsAux!impcobro, "N")
                    RC = RC & "'," & TransformaComasPuntos(CStr(Importe)) & ")"
                End If
                RC = Cad & RC
            
                Conn.Execute RC
            
                'Sig
                I = I + 1
                miRsAux.MoveNext
            Wend
            miRsAux.Close
       
'        Else
'            'Tenemos k insertar una unica linea a blancos
'            RC = vUsu.Codigo & "," & CONT & ",''," & I & ",'','','','',0)"
'            RC = Cad & RC
'
'            Conn.Execute RC
'        End If
        RS.MoveNext
    Wend
    RS.Close
    CadenaDesdeOtroForm = ""
    
    Set RS = Nothing
    Set miRsAux = Nothing
    
    If Opcion = 13 Then
        'Puede ser carta
        If chkCartaAbonos.Value Then
            'En nommacta pongo la provincia  (desprovi)
            Cad = "INSERT INTO usuarios.zcuentas(codusu,codmacta,nommacta,razosoci,dirdatos,codposta,despobla,nifdatos)"
            Cad = Cad & " Select " & vUsu.Codigo & ",codmacta,desprovi,razosoci,dirdatos,codposta,despobla,nifdatos FROM cuentas WHERE "
            Cad = Cad & " codmacta IN (select distinct(codmacta) from usuarios.ztmplibrodiario where codusu =" & vUsu.Codigo & ")"
            Ejecuta Cad
        
        
            Cad = "apoderado"
            RC = DevuelveDesdeBD("contacto", "empresa2", "1", "1", "N", Cad)
            If RC = "" Then RC = Cad
            If RC <> "" Then
                Cad = "UPDATE usuarios.ztesoreriacomun SET observa1='" & DevNombreSQL(RC) & "'"
                Cad = Cad & " WHERE codusu = " & vUsu.Codigo
                Conn.Execute Cad
            End If
        End If
    End If
    
    If Me.chkRem(0).Value = 1 Then
        If I = 1 Then
            MsgBox "No hay vencimientos asociados a las transferencias", vbExclamation
            Exit Function
        End If
    End If
    ListadoTransferencias = True
    Exit Function
EListadoTransferencias:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
    Set miRsAux = Nothing
End Function





Private Function ListAseguBasico() As Boolean
    On Error GoTo EListAseguBasico
    ListAseguBasico = False
    
    Cad = "DELETE FROM Usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    Cad = "Select * from cuentas where numpoliz<>"""""
    SQL = ""
    RC = CampoABD(Text3(21), "F", "fecsolic", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(22), "F", "fecconce", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    If SQL <> "" Then Cad = Cad & SQL
        
    
    'ORDENACION
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    Cad = Cad & " ORDER BY " & RC
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    Cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,fecha1,fecha2,importe1,"
    Cad = Cad & "importe2,observa1,observa2) VALUES (" & vUsu.Codigo & ","
        
    While Not miRsAux.EOF
        CONT = CONT + 1
        SQL = CONT & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','" & DBLet(miRsAux!nifdatos, "T") & "','"
        SQL = SQL & DevNombreSQL(miRsAux!numpoliz) & "',"
        'Fecha sol y concesion
        SQL = SQL & CampoBD_A_SQL(miRsAux!fecsolic, "F", True) & "," & CampoBD_A_SQL(miRsAux!fecconce, "F", True) & ","
        'Importes sol y concesion
        SQL = SQL & CampoBD_A_SQL(miRsAux!credisol, "N", True) & "," & CampoBD_A_SQL(miRsAux!credicon, "N", True) & ","
        'Observaciones
        RC = Memo_Leer(miRsAux!observa)
        If Len(RC) = 0 Then
            'Los dos campos NULL
            SQL = SQL & "NULL,NULL"
        Else
            If Len(RC) < 255 Then
                SQL = SQL & "'" & DevNombreSQL(RC) & "',NULL"
            Else
                SQL = SQL & "'" & DevNombreSQL(Mid(RC, 1, 255))
                RC = Mid(RC, 256)
                SQL = SQL & "','" & DevNombreSQL(Mid(RC, 1, 255)) & "'"
            End If
        End If
        
        SQL = SQL & ")"
        Conn.Execute Cad & SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If CONT > 0 Then
        ListAseguBasico = True
    Else
        MsgBox "Ningun datos con esos valores", vbExclamation
    End If
    Exit Function
EListAseguBasico:
    MuestraError Err.Number, "ListAseguBasico"
End Function





Private Function ListAsegFacturacion() As Boolean
Dim FP As Integer 'Forma de pago
Dim Cadpago As String
    On Error GoTo EListAsegFacturacion
    ListAsegFacturacion = False
    
    Cad = "DELETE FROM Usuarios.zpendientes  where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    
    If Me.optFecgaASig(0).Value Then
        Cad = "fecfaccl"
    Else
        Cad = "fecvenci"
    End If
        
    SQL = ""
    RC = CampoABD(Text3(21), "F", Cad, True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(22), "F", Cad, False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    
    
    
    Cad = "Select scobro.*,nommacta,numpoliz,nomforpa,forpa from scobro,cuentas,sforpa where scobro.codmacta= cuentas.codmacta AND numpoliz<>"""""
    Cad = Cad & " and scobro.codforpa=sforpa.codforpa AND sforpa.tipforpa <> " & vbEfectivo 'EL EFECTIVO NO se comprueba
    If SQL <> "" Then Cad = Cad & SQL
        
    
    'ORDENACION
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    Cad = Cad & " ORDER BY " & RC
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0

    Cad = "INSERT INTO Usuarios.zpendientes (codusu, serie_cta, factura, fecha, numorden,"
    Cad = Cad & "codforpa, nomforpa, codmacta, nombre, fecVto, importe,"
    Cad = Cad & "Situacion,pag_cob, vencido,  gastos) VALUES (" & vUsu.Codigo & ","
    Cadpago = ","
    While Not miRsAux.EOF
        CONT = CONT + 1
        SQL = "'" & miRsAux!NUmSerie & "','" & Format(miRsAux!codfaccl, "000000000") & "','" & Format(miRsAux!fecfaccl, FormatoFecha) & "',"
        FP = miRsAux!codforpa
        If optFP(1).Value Then
            If DBLet(miRsAux!Forpa, "N") > 0 Then
                FP = miRsAux!Forpa
                If InStr(1, Cadpago, "," & FP & ",") = 0 Then Cadpago = Cadpago & FP & ","
            End If
        End If
        SQL = SQL & miRsAux!numorden & "," & FP & ",'" & DevNombreSQL(miRsAux!nomforpa) & "','" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta)
        SQL = SQL & "','" & Format(miRsAux!FecVenci, FormatoFecha) & "',"
        'IMporte
        Importe = miRsAux!ImpVenci
        If Not IsNull(miRsAux!Gastos) Then Importe = Importe + miRsAux!Gastos
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        SQL = SQL & TransformaComasPuntos(CStr(Importe))
        'Situacion tengo numpoliza
        SQL = SQL & ",'" & DevNombreSQL(miRsAux!numpoliz) & "',"
        'Gastos e imvenci van a la columna pag_cob   Julio 2009
        Importe = miRsAux!ImpVenci + DBLet(miRsAux!Gastos, "N")
        SQL = SQL & TransformaComasPuntos(CStr(Importe))
        'El resto
        SQL = SQL & ",0,NULL)"
        
        Conn.Execute Cad & SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    If CONT = 0 Then
        MsgBox "Ningun datos con esos valores", vbExclamation
        Exit Function
    End If
    
    
    'Si ha cambiado valores en la forma de pago
    If Len(Cadpago) > 1 Then
        Cadpago = Mid(Cadpago, 2)
        Cadpago = Mid(Cadpago, 1, Len(Cadpago) - 1)
        Cad = "select codforpa,nomforpa from sforpa where codforpa in (" & Cadpago & ") GROUP by  codforpa"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = " WHERE codusu = " & vUsu.Codigo & " AND codforpa = "
        While Not miRsAux.EOF
            SQL = "UPDATE Usuarios.zpendientes SET nomforpa = '" & DevNombreSQL(miRsAux!nomforpa) & "'" & Cad & miRsAux!codforpa
            If Not Ejecuta(SQL) Then MsgBox "Error actualizando tmp.  Forpa: " & miRsAux!codforpa & " " & miRsAux!nomforpa, vbExclamation
            miRsAux.MoveNext
        Wend
        miRsAux.Close
    End If
    ListAsegFacturacion = True
    
    
    Exit Function
EListAsegFacturacion:
    MuestraError Err.Number, "ListAseguBasico"
End Function



Private Function ListAsegImpagos() As Boolean
    On Error GoTo EListAsegImpagos
    ListAsegImpagos = False
    
    Cad = "DELETE FROM Usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    Cad = "Select numserie,codfaccl,numorden,fecvenci,impvenci,impcobro,gastos,scobro.codmacta,nommacta,despobla,desprovi,numpoliz,nomforpa from "
    Cad = Cad & "scobro,cuentas,sforpa where scobro.codmacta= cuentas.codmacta AND numpoliz<>"""""
    Cad = Cad & " and scobro.codforpa=sforpa.codforpa AND sforpa.tipforpa <> " & vbEfectivo 'EL EFECTIVO NO se comprueba
    'Impagados
    Cad = Cad & " AND devuelto = 1"
    SQL = ""
    RC = CampoABD(Text3(21), "F", "fecvenci", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(22), "F", "fecvenci", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    If SQL <> "" Then Cad = Cad & SQL
        
    
    'ORDENACION
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    Cad = Cad & " ORDER BY " & RC
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    Cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,texto5,texto6,fecha1,importe1) VALUES (" & vUsu.Codigo & ","
        
    While Not miRsAux.EOF
        CONT = CONT + 1
        SQL = CONT & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','" & DevNombreSQL(DBLet(miRsAux!desPobla, "T")) & "','"
        SQL = SQL & DevNombreSQL(DBLet(miRsAux!desProvi, "T")) & "','" & DevNombreSQL(miRsAux!numpoliz) & "','"
        SQL = SQL & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000000") & "',"
        'Fecha vto
        SQL = SQL & CampoBD_A_SQL(miRsAux!FecVenci, "F", True) & ","
        'Importes sol y concesion
        Importe = miRsAux!ImpVenci
        If Not IsNull(miRsAux!Gastos) Then Importe = Importe + miRsAux!Gastos
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        SQL = SQL & TransformaComasPuntos(CStr(Importe))
        
    
        SQL = SQL & ")"
        Conn.Execute Cad & SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If CONT > 0 Then
        ListAsegImpagos = True
    Else
        MsgBox "Ningun datos con esos valores", vbExclamation
    End If
    Exit Function
EListAsegImpagos:
    MuestraError Err.Number, "ListAsegImpagos"
End Function


Private Function ListAsegEfectos() As Boolean
Dim TotalCred As Currency

    On Error GoTo EListAsegEfectos
    ListAsegEfectos = False
    
    Cad = "DELETE FROM Usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    Cad = "Select numserie,codfaccl,numorden,fecvenci,impvenci,impcobro,gastos,fecfaccl,devuelto,scobro.codmacta,nommacta,credicon from "
    Cad = Cad & "scobro,cuentas,sforpa where scobro.codmacta= cuentas.codmacta AND numpoliz<>"""""
    Cad = Cad & " and scobro.codforpa=sforpa.codforpa AND sforpa.tipforpa <> " & vbEfectivo 'EL EFECTIVO NO se comprueba

    SQL = ""
    RC = CampoABD(Text3(21), "F", "fecvenci", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(22), "F", "fecvenci", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    If SQL <> "" Then Cad = Cad & SQL
        
    
    'ORDENACION
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    Cad = Cad & " ORDER BY codmacta,fecvenci"
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    'Seran:                                                     codmac,nomma,credicon,numfac,fecfac,fvto,impvto,disponible,vencida
    Cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,fecha1,fecha2,importe1,importe2,opcion) VALUES (" & vUsu.Codigo & ","
    RC = ""
    
    While Not miRsAux.EOF
        If RC <> miRsAux!codmacta Then
            RC = miRsAux!codmacta
            TotalCred = DBLet(miRsAux!credicon, "N")
            CadenaDesdeOtroForm = ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
            If IsNull(miRsAux!credicon) Then
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & "0,00','"
            Else
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Format(miRsAux!credicon, FormatoImporte) & "','"
            End If
        End If
        CONT = CONT + 1
        SQL = CONT & CadenaDesdeOtroForm
        SQL = SQL & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000000") & "',"
        'Fecha fac
        SQL = SQL & CampoBD_A_SQL(miRsAux!fecfaccl, "F", True) & ","
        'Fecha vto
        SQL = SQL & CampoBD_A_SQL(miRsAux!FecVenci, "F", True) & ","
        'Importes sol y concesion
        Importe = miRsAux!ImpVenci
        If Not IsNull(miRsAux!Gastos) Then Importe = Importe + miRsAux!Gastos
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        SQL = SQL & TransformaComasPuntos(CStr(Importe))
        TotalCred = TotalCred - Importe
        SQL = SQL & "," & TransformaComasPuntos(CStr(TotalCred))
       
        'Devuelto
        SQL = SQL & "," & DBLet(miRsAux!Devuelto, "N") & ")"
    
        Conn.Execute Cad & SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If CONT > 0 Then
        ListAsegEfectos = True
    Else
        MsgBox "Ningun datos con esos valores", vbExclamation
    End If
    Exit Function
EListAsegEfectos:
    MuestraError Err.Number, "ListAsegEfec"
End Function



Private Sub GeneraComboCuentas()
'
'    If Opcion = 1 Then
'        'COBROS PENDIENTES
'    Else: Pagos
'
        cmbCuentas(Opcion - 1).Clear
        cmbCuentas(Opcion - 1).AddItem "Sin especificar"
        
        cmbCuentas(Opcion - 1).AddItem "Crear selección"
              
        'En el tag tendremos las cuentas seleccionadas
        If Me.cmbCuentas(Opcion - 1).Tag <> "" Then cmbCuentas(Opcion - 1).AddItem "Cuentas seleccionadas"


    'Cargo aqui tb los checks
    CargaTextosTipoPagos False
End Sub



Private Sub CargaTextosTipoPagos(Reclamaciones As Boolean)
    
    Set miRsAux = New ADODB.Recordset
    Cad = "Select tipoformapago, descformapago,siglas from stipoformapago order by tipoformapago "
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If Reclamaciones Then
            chkTipPagoRec(miRsAux!tipoformapago).Caption = miRsAux!siglas
            chkTipPagoRec(miRsAux!tipoformapago).Visible = True
            chkTipPagoRec(miRsAux!tipoformapago).Value = 1
        
        Else
            chkTipPago(miRsAux!tipoformapago).Caption = miRsAux!siglas
            chkTipPago(miRsAux!tipoformapago).Visible = True
            chkTipPago(miRsAux!tipoformapago).Value = 1
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



'Para conceptos y diarios
'Opcion: 0- Diario
'        1- Conceptos
'        2- Centros de coste
'        3- Gastos fijos
'        4. Hco compensaciones
Private Sub LanzaBuscaGrid(Indice As Integer, OpcionGrid As Byte)


    Select Case OpcionGrid
    Case 0
    'Diario
        DevfrmCCtas = "0"
        Cad = "Número|numdiari|N|30·"
        Cad = Cad & "Descripción|desdiari|T|60·"

        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = "Tiposdiario"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Diario"
        frmB.vSelElem = 0
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        If DevfrmCCtas <> "" Then
           Me.txtDiario(Indice) = RecuperaValor(DevfrmCCtas, 1)
           Me.txtDescDiario(Indice) = RecuperaValor(DevfrmCCtas, 2)
        End If
 Case 1
        'Conceptos
        DevfrmCCtas = "0"
        Cad = "Codigo|codconce|N|30·"
        Cad = Cad & "Descripción|nomconce|T|60·"

        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = "Conceptos"
        frmB.vSQL = ""
        
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "CONCEPTOS"
        frmB.vSelElem = 0
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        If DevfrmCCtas <> "" Then
           Me.txtConcpto(Indice) = RecuperaValor(DevfrmCCtas, 1)
           Me.txtDescConcepto(Indice) = RecuperaValor(DevfrmCCtas, 2)
        End If
        
    Case 2
        'Centros de coste
        DevfrmCCtas = "0"
        Cad = "Codigo|codccost|T|30·"
        Cad = Cad & "Descripción|nomccost|T|60·"

        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = "cabccost"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Centros de coste"
        frmB.vSelElem = 0
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        If DevfrmCCtas <> "" Then
            
           txtCCost(Indice) = RecuperaValor(DevfrmCCtas, 1)
           txtDescCCoste(Indice) = RecuperaValor(DevfrmCCtas, 2)
        End If
        
    Case 3
        'Gasto fijos  sgastfij codigo Descripcion
        DevfrmCCtas = "0"
        Cad = "Código|codigo|T|30·"
        Cad = Cad & "Descripción|Descripcion|T|60·"

        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = "sgastfij"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Gastos fijos"
        frmB.vSelElem = 0
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        If DevfrmCCtas <> "" Then
            
           txtGastoFijo(Indice) = RecuperaValor(DevfrmCCtas, 1)
           txtDescGastoFijo(Indice) = RecuperaValor(DevfrmCCtas, 2)
        End If
        
    Case 4
        'Gasto fijos  sgastfij codigo Descripcion
        '-------------------------------------------
        DevfrmCCtas = "0"
        Cad = "Código|codigo|T|10·"
        Cad = Cad & "Fecha|fecha|T|26·"
        Cad = Cad & "Cuenta|codmacta|T|14·"
        Cad = Cad & "Nombre|nommacta|T|50·"

        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = "scompenclicab"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|"
        frmB.vTitulo = "Hco. compensaciones cliente"
        frmB.vSelElem = 0
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        If DevfrmCCtas <> "" Then
            DevfrmCCtas = RecuperaValor(DevfrmCCtas, 1)
            If DevfrmCCtas = "" Then DevfrmCCtas = "0"
            If Val(DevfrmCCtas) Then
                CONT = Val(DevfrmCCtas)
                ImprimiCompensacion CONT
            End If
           
        End If
    End Select
End Sub

                                       '                Para saber el index del listview
Public Sub InsertaItemComboCompensaVto(TEXTO As String, Indice As Integer)
    Me.cboCompensaVto.AddItem TEXTO
    Me.cboCompensaVto.ItemData(Me.cboCompensaVto.NewIndex) = Indice
End Sub


Private Function GeneraDatosTalPag() As Boolean
Dim B As Boolean

    'Borramos los tmp
    SQL = "DELETE FROM usuarios.ztesoreriacomun WHERE codusu = " & vUsu.Codigo
    Conn.Execute SQL

    If chkLstTalPag(3).Value = 1 Then
        B = GeneraDatosTalPagDesglosado
    Else
        'Sin desglosar
        B = GeneraDatosTalPagSinDesglose
    End If
    GeneraDatosTalPag = B
End Function

Private Function GeneraDatosTalPagDesglosado() As Boolean
    On Error GoTo EGeneraDatosTalPag
    GeneraDatosTalPagDesglosado = False
    
    

       
       
    SQL = "select slirecepdoc.*,scarecepdoc.*,nommacta,nifdatos from slirecepdoc,scarecepdoc,cuentas "
    SQL = SQL & " where slirecepdoc.id =scarecepdoc.codigo and scarecepdoc.codmacta=cuentas.codmacta"
    If Text3(24).Text <> "" Then SQL = SQL & " AND fecharec >= '" & Format(Text3(24).Text, FormatoFecha) & "'"
    If Text3(25).Text <> "" Then SQL = SQL & " AND fecharec <= '" & Format(Text3(25).Text, FormatoFecha) & "'"
    'Contabilizado
    'SQL = SQL & " AND Contabilizada =  1"
    'Si esta llevada a banco o no
    If cboListPagare.ListIndex >= 1 Then SQL = SQL & " AND LlevadoBanco = " & Abs(cboListPagare.ListIndex = 1)
    '------------------------------------------------------------------------
    I = -1
    If (chkLstTalPag(0).Value = 1) Xor (chkLstTalPag(1).Value = 1) Then
        'Solo uno seleccionado
        I = 0
        If (chkLstTalPag(0).Value = 1) Then I = 1
            
    Else
        If (chkLstTalPag(0).Value = 0) Then
            MsgBox "Seleccione talon o pagare", vbExclamation
            Exit Function
        End If
    End If
    If I >= 0 Then SQL = SQL & " AND talon = " & I

    'Si ID
    If txtNumfac(2).Text <> "" Then SQL = SQL & " AND codigo >= " & txtNumfac(2).Text
    If txtNumfac(3).Text <> "" Then SQL = SQL & " AND codigo <= " & txtNumfac(3).Text

    Set RS = New ADODB.Recordset
    
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    RC = ""
    While Not RS.EOF
        I = I + 1
        'ztesoreriacomun (`codusu`,`codigo
        SQL = ", (" & vUsu.Codigo & "," & I & ","
        
        'texto1`,`texto2`,`texto3`,y el 4
        SQL = SQL & "'" & DevNombreSQL(RS!numeroref) & "','" & DevNombreSQL(RS!banco) & "','"
        SQL = SQL & DevNombreSQL(RS!codmacta) & "','" & DevNombreSQL(RS!Nommacta) & "','"
        
        
        '5 Serie y numero factura
        SQL = SQL & RS!NUmSerie & Format(RS!numfaccl, "000000") & "',"
        '`importe1`
        'SQL = SQL & TransformaComasPuntos(CStr(Rs!Importe)) & ","
        SQL = SQL & TransformaComasPuntos(CStr(RS.Fields(5))) & ",'"   'La columna 5 es sli.importe
        
        'texto6=nifdatos
        SQL = SQL & DevNombreSQL(DBLet(RS!nifdatos, "N"))
        
        '`fecha1`,`fecha2`,`fecha3`
        SQL = SQL & "','" & Format(RS!fecharec, FormatoFecha) & "',"
        SQL = SQL & "'" & Format(RS!fechavto, FormatoFecha) & "',"
        SQL = SQL & "'" & Format(RS!fecfaccl, FormatoFecha) & "')"
    
        RC = RC & SQL
        
        RS.MoveNext
    Wend
    RS.Close
    
    If I > 0 Then
        RC = Mid(RC, 3) 'QUITO LA PRIMERA COMA
        'OK hay datos. Insertamos
        SQL = "INSERT INTO usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,`texto2`,`texto3`,"
        SQL = SQL & "`texto4`,`texto5`,`importe1`,texto6,`fecha1`,`fecha2`,`fecha3`) VALUES "
        SQL = SQL & RC
        Conn.Execute SQL
    
        'Si estamos emitiendo el justicante de recepcion, guardare en z340 los campos
        'fiscales del cliente para su impresion
        If Me.chkLstTalPag(2).Value = 1 Then
            SQL = "DELETE FROM usuarios.z347 WHERE codusu = " & vUsu.Codigo
            Conn.Execute SQL
            
            SQL = "DELETE FROm Usuarios.z347carta WHERE codusu = " & vUsu.Codigo
            Conn.Execute SQL
            
            espera 0.3
            
            
            'En texto3 esta la codmacta
            SQL = "Select texto3 from usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo & " GROUP BY texto3"
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            RC = ""
            While Not RS.EOF
                RC = RC & ", '" & RS!texto3 & "'"
                RS.MoveNext
            Wend
            RS.Close
            
            
            
            
            
            'No puede ser EOF
            RC = Trim(Mid(RC, 2))
            'Monto un superselect
            'pongo el IGNORE por si acaso hay cuentas con el mismo NIF
            SQL = "insert ignore into usuarios.z347 (`codusu`,`cliprov`,`nif`,`razosoci`,`dirdatos`,`codposta`,`despobla`,`Provincia`)"
            SQL = SQL & " SELECT " & vUsu.Codigo & ",0,nifdatos,razosoci,dirdatos,codposta,despobla,desprovi FROM cuentas where codmacta in (" & RC & ")"
            Conn.Execute SQL
    
    
    
            'Ahora meto los datos de la empresa
            Cad = "INSERT INTO Usuarios.z347carta (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir,"
            Cad = Cad & "contacto) VALUES ("
            Cad = Cad & vUsu.Codigo
                
                
            'Monta Datos Empresa
            RS.Open "empresa2", Conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
            If RS.EOF Then
                MsgBox "Error en los datos de la empresa " & vEmpresa.nomempre
                RC = ",'','','','','',''"  '6 campos
            Else
                RC = DBLet(RS!siglasvia) & " " & DBLet(RS!Direccion) & "  " & DBLet(RS!numero) & ", " & DBLet(RS!puerta)
                RC = ",'" & DBLet(RS!nifempre) & "','" & vEmpresa.nomempre & "','" & RC & "','"
                RC = RC & DBLet(RS!codpos) & "','" & DBLet(RS!Poblacion) & "','" & DBLet(RS!provincia) & "','" & DBLet(RS!contacto) & "')"
            End If
            RS.Close
            Cad = Cad & RC
            Conn.Execute Cad
            
            
            
    
        End If
        GeneraDatosTalPagDesglosado = True
    Else
        'I>0
        MsgBox "No hay datos", vbExclamation
    End If

EGeneraDatosTalPag:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Function



Private Function GeneraDatosTalPagSinDesglose() As Boolean
    On Error GoTo EGeneraDatosTalPag
    GeneraDatosTalPagSinDesglose = False
    
    

       
       
    SQL = "select scarecepdoc.*,nommacta from scarecepdoc,cuentas "
    SQL = SQL & " where  scarecepdoc.codmacta=cuentas.codmacta"
    If Text3(24).Text <> "" Then SQL = SQL & " AND fecharec >= '" & Format(Text3(24).Text, FormatoFecha) & "'"
    If Text3(25).Text <> "" Then SQL = SQL & " AND fecharec <= '" & Format(Text3(25).Text, FormatoFecha) & "'"
    'Contabilizado
    'SQL = SQL & " AND Contabilizada =  1"
    'Si esta llevada a banco o no
    'SQL = SQL & " AND LlevadoBanco = " & Abs(chkLstTalPag(2).Value)
    If cboListPagare.ListIndex >= 1 Then SQL = SQL & " AND LlevadoBanco = " & Abs(cboListPagare.ListIndex = 1)
    
    I = -1
    If (chkLstTalPag(0).Value = 1) Xor (chkLstTalPag(1).Value = 1) Then
        'Solo uno seleccionado
        I = 0
        If (chkLstTalPag(0).Value = 1) Then I = 1
            
    Else
        If (chkLstTalPag(0).Value = 0) Then
            MsgBox "Seleccione talon o pagare", vbExclamation
            Exit Function
        End If
    End If
    If I >= 0 Then SQL = SQL & " AND talon = " & I
    'Si ID
    If txtNumfac(2).Text <> "" Then SQL = SQL & " AND codigo >= " & txtNumfac(2).Text
    If txtNumfac(3).Text <> "" Then SQL = SQL & " AND codigo <= " & txtNumfac(3).Text



    Set RS = New ADODB.Recordset
    
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    RC = ""
    While Not RS.EOF
        I = I + 1
        'ztesoreriacomun (`codusu`,`codigo
        SQL = ", (" & vUsu.Codigo & "," & I & ","
        
        'texto1`,`texto2`,`texto3`,y el 4
        SQL = SQL & "'" & DevNombreSQL(RS!numeroref) & "','" & DevNombreSQL(RS!banco) & "','"
        SQL = SQL & DevNombreSQL(RS!codmacta) & "','" & DevNombreSQL(RS!Nommacta) & "','"
        
        
        '5 Serie y numero factura
        SQL = SQL & "',"
        '`importe1`
        'SQL = SQL & TransformaComasPuntos(CStr(Rs.Fields(5))) & ","   'La columna 5 es sli.importe
        SQL = SQL & TransformaComasPuntos(CStr(RS!Importe)) & ","
        
        '
        '`fecha1`,`fecha2`,`fecha3`
        SQL = SQL & "'" & Format(RS!fecharec, FormatoFecha) & "',"
        SQL = SQL & "'" & Format(RS!fechavto, FormatoFecha) & "',"
        SQL = SQL & "'" & Format(Now, FormatoFecha) & "')"
    
        RC = RC & SQL
        
        RS.MoveNext
    Wend
    RS.Close
    
    If I > 0 Then
        RC = Mid(RC, 3) 'QUITO LA PRIMERA COMA
        'OK hay datos. Insertamos
        SQL = "INSERT INTO usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,`texto2`,`texto3`,"
        SQL = SQL & "`texto4`,`texto5`,`importe1`,`fecha1`,`fecha2`,`fecha3`) VALUES "
        SQL = SQL & RC
        Conn.Execute SQL
        GeneraDatosTalPagSinDesglose = True
    Else
        MsgBox "No hay datos", vbExclamation
    End If
    
    

EGeneraDatosTalPag:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Function





Private Function ListadoOrdenPago() As Boolean
    On Error GoTo EListadoOrdenPago
    ListadoOrdenPago = False

    'Borramos
    Cad = "DELETE from usuarios.zlistadopagos WHERE codusu = " & vUsu.Codigo
    Conn.Execute Cad
    Set miRsAux = New ADODB.Recordset
    'Inserttamos
    RC = ""
    If txtCtaBanc(3).Text <> "" Or txtCtaBanc(4).Text <> "" Then
        If txtCtaBanc(3).Text <> "" Then RC = " codmacta >= '" & txtCtaBanc(3).Text & "'"
        
        If txtCtaBanc(4).Text <> "" Then
            If RC <> "" Then RC = RC & " AND "
            RC = RC & " codmacta <= '" & txtCtaBanc(4).Text & "'"
        End If
        Cad = "Select codmacta from ctabancaria where " & RC
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        RC = ""
        While Not miRsAux.EOF
            RC = RC & ", '" & miRsAux!codmacta & "'"
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        If RC = "" Then
            MsgBox "Ningún banco con esos valores", vbExclamation
            Exit Function
        End If
           
        RC = Mid(RC, 2)
    End If
    
    
    SQL = ""
    If Text3(26).Text <> "" Then SQL = SQL & " AND fecefect >= '" & Format(Text3(26).Text, FormatoFecha) & "'"
    If Text3(27).Text <> "" Then SQL = SQL & " AND fecefect <= '" & Format(Text3(27).Text, FormatoFecha) & "'"
    If RC <> "" Then SQL = SQL & " AND ctabanc1 in (" & RC & ")"
    
    
    'JULIO2013
    'La variable contdocu grabaremos emitdocum, y en el listado sabremos si es de talon/pagere
    'para poder separalos
    'Antes: contdocu, ahora emitdocum
    
    'Agosto 2014
    'Tipo de pago
    Cad = "select " & vUsu.Codigo & ",`ctaprove`,`numfactu`,`fecfactu`,`numorden`,`fecefect`,`impefect`-coalesce(imppagad,0),"
    Cad = Cad & " `ctabanc1`,`ctabanc2`,`emitdocum`,spagop.entidad,spagop.oficina,spagop.CC,spagop.cuentaba,"
    Cad = Cad & " nommacta,'error','error',descformapago "
    
    Cad = Cad & " from spagop,cuentas,sforpa,stipoformapago "
    Cad = Cad & " WHERE spagop.ctaprove = cuentas.codmacta AND spagop.codforpa=sforpa.codforpa and tipoformapago=tipforpa"
    'Ponemos un check si salen negativos o no
    RC = " AND impefect >=0"
    If Me.chkPagBanco(0).Value = 1 And Me.chkPagBanco(1).Value = 1 Then RC = "" 'Salen todos
    Cad = Cad & RC 'todos o solo positivos
    Cad = Cad & SQL
    
    SQL = "INSERT INTO usuarios.zlistadopagos (`codusu`,`ctaprove`,`numfactu`,`fecfactu`,`numorden`,`fecefect`,"
    SQL = SQL & " `impefect`,`ctabanc1`,`ctabanc2`,`contdocu`,`entidad`,`oficina`,`CC`,`cuentaba`,"
    SQL = SQL & " `nomprove`,`nombanco`,`cuentabanco`,TipoForpa) " & Cad
    Conn.Execute SQL
    
    Cad = DevuelveDesdeBD("count(*)", "usuarios.zlistadopagos", "codusu", vUsu.Codigo)
    If Val(Cad) = 0 Then
        MsgBox "Ningun vencimiento con esos valores", vbExclamation
        Exit Function
    End If
    
    'Actualizo los datos de los bancos `nombanco`,`cuentabanco`
    Cad = "Select ctabanc1 from usuarios.zlistadopagos WHERE codusu = " & vUsu.Codigo
    Cad = Cad & " AND ctabanc1 <>'' GROUP BY ctabanc1"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not miRsAux.EOF
        Cad = Cad & miRsAux!ctabanc1 & "|"
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    While Cad <> ""
        I = InStr(1, Cad, "|")
        If I = 0 Then
            Cad = ""
        Else
            RC = Mid(Cad, 1, I - 1)
            Cad = Mid(Cad, I + 1)
            
            SQL = "Select ctabancaria.codmacta,ctabancaria.entidad, ctabancaria.oficina, ctabancaria.control, ctabancaria.ctabanco,"
            SQL = SQL & " ctabancaria.descripcion,nommacta from  ctabancaria,cuentas where ctabancaria.codmacta=cuentas.codmacta "
            SQL = SQL & " AND ctabancaria.codmacta ='" & RC & "'"
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                SQL = "Cuenta banco erronea: " & vbCrLf & "Hay vencimientos asociados a la cuenta " & RC & " que no esta en bancos"
                MsgBox SQL, vbExclamation
            Else
                SQL = DBLet(miRsAux!Descripcion, "T")
                If SQL = "" Then SQL = miRsAux!Nommacta
                SQL = DevNombreSQL(SQL) & "|"
                
                'Enti8dad...
                I = DBLet(miRsAux!Entidad, "0")
                SQL = SQL & Format(I, "0000")
                                'Oficina...
                I = DBLet(miRsAux!Oficina, "0")
                SQL = SQL & Format(I, "0000")
                                'CC...
                RC = DBLet(miRsAux!Control, "T")
                If RC = "" Then RC = "**"
                SQL = SQL & RC
                'cuenta
                RC = DBLet(miRsAux!CtaBanco, "T")
                If RC = "" Then RC = "    **"
                SQL = SQL & RC & "|"
                
                
                RC = "UPDATE usuarios.zlistadopagos set `nombanco`='" & RecuperaValor(SQL, 1)
                RC = RC & "',`cuentabanco`='" & RecuperaValor(SQL, 2) & "'"
                RC = RC & " WHERE ctabanc1 = '" & miRsAux!codmacta & "' AND codusu = " & vUsu.Codigo
                Conn.Execute RC
                
            End If
            miRsAux.Close
        End If
    Wend
    
    ListadoOrdenPago = True
    Set miRsAux = Nothing
    Exit Function
EListadoOrdenPago:
    MuestraError Err.Number, "ListadoOrdenPago"
End Function



Private Function ListadoReclamas() As Boolean

On Error GoTo EListadoReclamas

    ListadoReclamas = False
        

    SQL = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = ""
    Cad = ""
    
    If Text3(28).Text <> "" Or Text3(29).Text <> "" Then
        RC = DesdeHasta("F", 28, 29, "F.Reclama")
        If RC <> "" Then Cad = Cad & " " & RC
            
        RC = CampoABD(Text3(28), "F", "fecreclama", True)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
        RC = CampoABD(Text3(29), "F", "fecreclama", False)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
    End If
    
    
    If txtCta(15).Text <> "" Or txtCta(16).Text <> "" Then
        RC = DesdeHasta("C", 15, 16, "Cta")
        If RC <> "" Then Cad = Cad & " " & RC
            
        RC = CampoABD(txtCta(15), "T", "codmacta", True)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
        RC = CampoABD(txtCta(16), "T", "codmacta", False)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
    End If
    If SQL <> "" Then SQL = " WHERE " & SQL
    SQL = "Select * from shcocob" & SQL
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    RC = "insert into usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,`texto2`,`"
    RC = RC & "texto3`,`texto4`,`texto5`,`texto6`,`importe1`,`importe2`,`fecha1`,`fecha2`,"
    RC = RC & "`fecha3`,`texto`,`observa2`,`opcion`) VALUES "
    SQL = ""
    While Not RS.EOF
        NumRegElim = NumRegElim + 1
        SQL = SQL & ", (" & vUsu.Codigo & "," & NumRegElim & ",'" & RS!codmacta & "','"
        'text 2 y 3
        SQL = SQL & DevNombreSQL(RS!Nommacta) & "','" & RS!NUmSerie & Format(RS!codfaccl, "000000") & "','"
        '4 y 5
        SQL = SQL & RS!numorden & "','"
        If Val(RS!carta) = 1 Then
            SQL = SQL & "Email"
        ElseIf Val(RS!carta) = 2 Then
            SQL = SQL & "Teléfono"
        Else
            SQL = SQL & "Carta"
        End If
        'Text6, importe 1 y 2
        SQL = SQL & "',''," & TransformaComasPuntos(CStr(RS!ImpVenci)) & ",NULL,"
        'Fec1 reclama fec2 factra   fec3
        SQL = SQL & "'" & Format(RS!fecreclama, FormatoFecha) & "','" & Format(RS!fecfaccl, FormatoFecha) & "',NULL,"
        DevfrmCCtas = Memo_Leer(RS!observaciones)
        If DevfrmCCtas = "" Then
            DevfrmCCtas = "NULL"
        Else
            DevfrmCCtas = "'" & DevNombreSQL(DevfrmCCtas) & "'"
        End If
        SQL = SQL & DevfrmCCtas & ",NULL,0)"


        'Siguiente
        RS.MoveNext
        
        
        If Len(SQL) > 100000 Then
            SQL = Mid(SQL, 2) 'QUITO LA COMA
            SQL = RC & SQL
            Conn.Execute SQL
            SQL = ""
        End If
            
        
    Wend
    RS.Close
        If SQL <> "" Then
            SQL = Mid(SQL, 2) 'QUITO LA COMA
            SQL = RC & SQL
            Conn.Execute SQL
        End If
        
        
    If NumRegElim > 0 Then
        ListadoReclamas = True
    Else
        MsgBox "Ningun dato devuelto", vbExclamation
    End If
    Exit Function
EListadoReclamas:
    MuestraError Err.Number, Err.Description
End Function





'******************************************************************************************
'
'   Listado gastos fijos

Private Function ListadoGastosFijos() As Boolean

On Error GoTo EListadoGF

    ListadoGastosFijos = False
        

    SQL = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    SQL = ""
    Cad = ""
    
    
    DevfrmCCtas = "" ' ON del left join , NO al WHERE
    If Text3(30).Text <> "" Or Text3(31).Text <> "" Then
        RC = DesdeHasta("F", 30, 31, "Fecha")
        If RC <> "" Then Cad = Cad & " " & Trim(RC)
            
        RC = CampoABD(Text3(30), "F", "fecha", True)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
        RC = CampoABD(Text3(31), "F", "fecha", False)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
    End If
    DevfrmCCtas = SQL
    SQL = ""
    
    'Este si que va al where
    If txtGastoFijo(0).Text <> "" Or txtGastoFijo(1).Text <> "" Then
        RC = DesdeHasta("GF", 0, 1, "Codigo")
        If RC <> "" Then
            If Cad <> "" Then
                'Ya esta la fecha
                If Len(Cad & RC) > 100 Then Cad = Cad & """ + chr(13) + """
            End If
            Cad = Cad & " " & Trim(RC)
        End If
            
        RC = CampoABD(txtGastoFijo(0), "N", "sgastfij.codigo", True)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
        RC = CampoABD(txtGastoFijo(1), "N", "sgastfij.codigo", False)
        If RC <> "" Then
            If SQL <> "" Then SQL = SQL & " AND "
            SQL = SQL & RC
        End If
        
    End If
    
   
   
    RC = " FROM sgastfij left join sgastfijd ON sgastfij.Codigo = sgastfijd.Codigo"
    If DevfrmCCtas <> "" Then RC = RC & " AND " & DevfrmCCtas
    If SQL <> "" Then RC = RC & " WHERE " & SQL
    SQL = "SELECT sgastfij.codigo,descripcion,ctaprevista,fecha,importe" & RC
    
    

    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    RC = "insert into usuarios.ztesoreriacomun (`codusu`,`codigo`,`texto1`,`texto2`,`"
    RC = RC & "texto3`,`importe1`,`fecha1`) VALUES "
    SQL = ""
    While Not RS.EOF
        NumRegElim = NumRegElim + 1
        SQL = SQL & ", (" & vUsu.Codigo & "," & NumRegElim & ",'" & Format(RS!Codigo, "00000") & "','"
        'text 2 y 3
        SQL = SQL & DevNombreSQL(RS!Descripcion) & "','" & RS!Ctaprevista & "',"
       
  
        'Detalla
        If IsNull(RS!Fecha) Then
            SQL = SQL & "0,'" & Format(Now, FormatoFecha) & "'"
        Else
            SQL = SQL & TransformaComasPuntos(DBLet(RS!Importe, "N")) & ",'" & Format(RS!Fecha, FormatoFecha) & "'"
        End If
        SQL = SQL & ")"
        
        'Siguiente
        RS.MoveNext
            
        
    Wend
    RS.Close
    If SQL <> "" Then
        SQL = Mid(SQL, 2) 'QUITO LA COMA
        SQL = RC & SQL
        Conn.Execute SQL
    End If
        
        
    If NumRegElim = 0 Then
        MsgBox "Ningun dato devuelto", vbExclamation
        Exit Function
    End If
    
    
    'Updateo la cuenta bancaria
    RC = "Select texto3 from usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo & " GROUP BY 1"
    RS.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not RS.EOF
        SQL = SQL & RS!texto3 & "|"
        RS.MoveNext
    Wend
    RS.Close
    
    While SQL <> ""
        NumRegElim = InStr(1, SQL, "|")
        If NumRegElim = 0 Then
            SQL = ""
        Else
            RC = Mid(SQL, 1, NumRegElim - 1)
            SQL = Mid(SQL, NumRegElim + 1)
            
            RC = "Select codmacta,nommacta from cuentas where codmacta='" & RC & "'"
            RS.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                RC = "UPDATE usuarios.ztesoreriacomun SET texto4='" & DevNombreSQL(RS!Nommacta) & "' WHERE codusu =" & vUsu.Codigo & " AND texto3='" & RS!codmacta & "'"
                Conn.Execute RC
            End If
            RS.Close
        End If
    Wend
    ListadoGastosFijos = True
    Exit Function
EListadoGF:
    MuestraError Err.Number, Err.Description
End Function






'Listadoas aseguradoas
Private Function AvisosAseguradora() As Boolean



    On Error GoTo EListAsegEfectos
    AvisosAseguradora = False
    
    Cad = "DELETE FROM Usuarios.ztesoreriacomun  where codusu =" & vUsu.Codigo
    Conn.Execute Cad
    
    'feccomunica,fecprorroga,fecsiniestro
    SQL = ""
    If Me.optAsegAvisos(0).Value Then
        Cad = "feccomunica"
    ElseIf Me.optAsegAvisos(1).Value Then
        Cad = "fecprorroga"
    Else
        Cad = "fecsiniestro"
    End If
    RC = CampoABD(Text3(21), "F", Cad, True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(Text3(22), "F", Cad, False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    RC = CampoABD(txtCta(11), "T", "scobro.codmacta", True)
    If RC <> "" Then SQL = SQL & " AND " & RC
    RC = CampoABD(txtCta(12), "T", "scobro.codmacta", False)
    If RC <> "" Then SQL = SQL & " AND " & RC
    
    'Significa que no ha puesto fechas
    If InStr(1, SQL, Cad) = 0 Then SQL = SQL & " AND " & Cad & ">='1900-01-01'"
    
    'ORDENACION
    If Me.optAsegAvisos(0).Value Then
        RC = "feccomunica"
    ElseIf Me.optAsegAvisos(1).Value Then
        RC = "fecprorroga"
    Else
        RC = "fecsiniestro"
    End If
    
    Cad = "Select numserie,codfaccl,numorden,fecvenci,impvenci,impcobro,gastos,fecfaccl,devuelto,scobro.codmacta,nommacta,numpoliz"
    Cad = Cad & ",credicon," & RC & " LaFecha" 'alias
    Cad = Cad & "  FROM scobro,cuentas,sforpa where scobro.codmacta= cuentas.codmacta AND numpoliz<>"""""
    Cad = Cad & " and scobro.codforpa=sforpa.codforpa "
    If SQL <> "" Then Cad = Cad & SQL
    
    
    
    

    Cad = Cad & " ORDER BY " & RC & ","
    'ORDENACION 2º nivel
    If Me.optAsegBasic(1).Value Then
        RC = "nommacta"
    Else
        If Me.optAsegBasic(2).Value Then
            RC = "numpoliz"
        Else
            RC = "codmacta"
        End If
    End If
    Cad = Cad & RC
    
    
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CONT = 0
    'Seran:                                                     codmac,nomma,credicon,numfac,fecfac,faviso,fvto,impvto,disponible,vencida
    Cad = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo,texto1,texto2,texto3,texto4,fecha1,fecha2,fecha3,importe1,importe2,opcion) VALUES "
    RC = ""
    
    While Not miRsAux.EOF
        If Len(RC) > 500 Then
            RC = Mid(RC, 2)
            Conn.Execute Cad & RC
            RC = ""
        End If
        CONT = CONT + 1
        SQL = ", (" & vUsu.Codigo & "," & CONT & ",'" & miRsAux!codmacta & "','" & DevNombreSQL(miRsAux!Nommacta) & "','"
        SQL = SQL & DevNombreSQL(miRsAux!numpoliz) & "'"
        SQL = SQL & ",'" & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000000") & "',"  'texto4
        'Fecha fac
        SQL = SQL & CampoBD_A_SQL(miRsAux!fecfaccl, "F", True) & ","
        'Fecha aviso
        SQL = SQL & CampoBD_A_SQL(miRsAux!lafecha, "F", True) & ","
        'Fecha vto
        SQL = SQL & CampoBD_A_SQL(miRsAux!FecVenci, "F", True)
        
        SQL = SQL & "," & TransformaComasPuntos(CStr(miRsAux!ImpVenci))
        SQL = SQL & "," & TransformaComasPuntos(CStr(DBLet(miRsAux!Gastos, "N")))
        'Devuelto
        SQL = SQL & "," & DBLet(miRsAux!Devuelto, "N") & ")"
    
        RC = RC & SQL
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If RC <> "" Then
        RC = Mid(RC, 2)
        Conn.Execute Cad & RC
    End If
    
    
    If CONT > 0 Then
        AvisosAseguradora = True
    Else
        MsgBox "Ningun datos con esos valores", vbExclamation
    End If
    Exit Function
EListAsegEfectos:
    MuestraError Err.Number, "Avisos aseguradoras"
End Function



'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'
'       Compensaciones Cliente. Abonos vs Cobros
'
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

Private Sub PonerVtosCompensacionCliente()
Dim IT


    lwCompenCli.ListItems.Clear
    Me.txtimpNoEdit(0).Tag = 0
    Me.txtimpNoEdit(1).Tag = 0
    Me.txtimpNoEdit(0).Text = ""
    Me.txtimpNoEdit(1).Text = ""
    If Me.txtCta(17).Text = "" Then Exit Sub
    Set Me.lwCompenCli.SmallIcons = frmPpal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    Cad = "Select scobro.*,nomforpa from scobro,sforpa where scobro.codforpa=sforpa.codforpa "
    Cad = Cad & " AND codmacta = '" & Me.txtCta(17).Text & "'"
    Cad = Cad & " AND (transfer =0 or transfer is null) and codrem is null"
    Cad = Cad & " and estacaja=0 and recedocu=0"
    Cad = Cad & " ORDER BY fecvenci"
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwCompenCli.ListItems.Add()
        IT.Text = miRsAux!NUmSerie
        IT.SubItems(1) = Format(miRsAux!codfaccl, "000000")
        IT.SubItems(2) = Format(miRsAux!fecfaccl, "dd/mm/yyyy")
        IT.SubItems(3) = miRsAux!numorden
        IT.SubItems(4) = miRsAux!FecVenci
        IT.SubItems(5) = miRsAux!nomforpa
    
        Importe = DBLet(miRsAux!Gastos, "N")
        Importe = Importe + miRsAux!ImpVenci
        
        'Si ya he cobrado algo
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        
        If Importe > 0 Then
            IT.SubItems(6) = Format(Importe, FormatoImporte)
            IT.SubItems(7) = " "
        Else
            IT.SubItems(6) = " "
            IT.SubItems(7) = Format(-Importe, FormatoImporte)
        End If
        IT.Tag = Abs(Importe)  'siempre valor absoluto
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
End Sub



Private Sub RealizarCompensacionAbonosClientes()
Dim Borras As Boolean
    
    If BloqueoManual(True, "COMPEABONO", "1") Then

        Cad = DevuelveDesdeBD("max(codigo)", "scompenclicab", "1", "1")
        If Cad = "" Then Cad = "0"
        CONT = Val(Cad) + 1 'ID de la operacion
        
        Cad = "INSERT INTO scompenclicab(codigo,fecha,login,PC,codmacta,nommacta) VALUES (" & CONT
        Cad = Cad & ",now(),'" & DevNombreSQL(vUsu.Login) & "','" & DevNombreSQL(vUsu.PC)
        Cad = Cad & "','" & txtCta(17).Text & "','" & DevNombreSQL(DtxtCta(17).Text) & "')"
        
        Set miRsAux = New ADODB.Recordset
        Borras = True
        If Ejecuta(Cad) Then
            
            Borras = Not RealizarProcesoCompensacionAbonos
        
        End If


        Set miRsAux = Nothing
        If Borras Then
            Conn.Execute "DELETE FROM scompenclilin WHERE codigo = " & CONT
            Conn.Execute "DELETE FROM scompenclicab WHERE codigo = " & CONT
            
        End If

        'Desbloquamos proceso
        BloqueoManual False, "COMPEABONO", ""
        DevfrmCCtas = ""
        
        PonerVtosCompensacionCliente   'Volvemos a cargar los vencimientos
        
        'El nombre del report
        CadenaDesdeOtroForm = Me.Tag
        Me.Tag = ""
        If Not Borras Then
            ImprimiCompensacion CONT
            
        
        End If
        
        Set miRsAux = Nothing
    Else
        MsgBox "Proceso bloqueado", vbExclamation
    End If

End Sub



Private Sub ImprimiCompensacion(CodigoCompensacion As Long)

    On Error GoTo EImprimiCompensacion
        
        'CadenaDesdeOtroForm:  lleva el nombre del report
        
        
        'Ha ido bien. Imprimiremos la hoja por si quiere crear PDF
        Conn.Execute "DELETE FROM Usuarios.ztmpfaclin WHERE codusu =" & vUsu.Codigo
        Conn.Execute "Delete from Usuarios.zTesoreriaComun where codusu = " & vUsu.Codigo
        Conn.Execute "Delete from Usuarios.z347carta where codusu = " & vUsu.Codigo
        
        'insert into `ztmpfaclin` (`codusu`,`codigo`,`Numfac`,`Fecha`,`cta`,`Cliente`,`NIF`,
        '`Imponible`,`IVA`,`ImpIVA`,`Total`,`retencion`,`TipoIva`)
        SQL = "INSERT INTO usuarios.ztmpfaclin(`codusu`,`codigo`,`Numfac`,`Fecha`,`cta`,`Cliente`,`NIF`,`Imponible`,`ImpIVA`,`retencion`,`Total`,`IVA`,TipoIva)"
        SQL = SQL & "select " & vUsu.Codigo & ",@rownum:=@rownum+1 AS rownum,"
        SQL = SQL & "concat(numserie,right(concat(""000000"",codfaccl),8)) fecha,date_format(fecfaccl,'%d/%m/%Y') ffaccl,"
        SQL = SQL & "scompenclilin.codmacta,if (nommacta is null,nomclien,nommacta) nomcli,"
        SQL = SQL & "date_format(fecvenci,'%d/%m/%Y') venci,impvenci,gastos,impcobro,"
        SQL = SQL & "impvenci + coalesce(gastos,0) + coalesce(impcobro,0)  tot_al"
        SQL = SQL & ",if(fecultco is null,null,date_format(fecultco,'%d/&m')) fecco ,destino"
        SQL = SQL & " From (scompenclilin left join cuentas on scompenclilin.codmacta=cuentas.codmacta)"
        SQL = SQL & ",(SELECT @rownum:=0) r WHERE codigo=" & CONT & " order by destino desc,numserie,codfaccl"
        Conn.Execute SQL
            
        
            
        
   
    
    
        
    
    
    
    
        'Datos carta
        'Datos basicos de la empresa para la carta
        Cad = "INSERT INTO Usuarios.z347carta (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir, "
        Cad = Cad & "parrafo1, parrafo2, contacto, despedida,saludos,parrafo3, parrafo4, parrafo5, Asunto, Referencia)"
        Cad = Cad & " VALUES (" & vUsu.Codigo & ", "
        
        'Estos datos ya veremos com, y cuadno los relleno
        Set miRsAux = New ADODB.Recordset
        SQL = "select nifempre,siglasvia,direccion,numero,escalera,piso,puerta,codpos,poblacion,provincia,contacto from empresa2"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'Paarafo1 Parrafo2 contacto
        SQL = "'','',''"
        'sql= "'1234567890A','Ariadna Software ','Franco Tormo 3, Bajo Izda','46007','Valencia'"
        SQL = "'##########','" & vEmpresa.nomempre & "','#############','######','##########','##########'," & SQL
        If Not miRsAux.EOF Then
            SQL = ""
            For I = 1 To 6
                SQL = SQL & DBLet(miRsAux.Fields(I), "T") & " "
            Next I
            SQL = Trim(SQL)
            SQL = "'" & DBLet(miRsAux!nifempre, "T") & "','" & DevNombreSQL(vEmpresa.nomempre) & "','" & DevNombreSQL(SQL) & "'"
            SQL = SQL & ",'" & DBLet(miRsAux!codpos, "T") & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "'"

            'Contaccto
            SQL = SQL & ",NULL,NULL,'" & DevNombreSQL(DBLet(miRsAux!contacto)) & "' "
        End If
        miRsAux.Close
      
        Cad = Cad & SQL
        Cad = Cad & ",NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
        
        Conn.Execute Cad
        
        
        'Datos CLIENTE
         ', texto3, texto4, texto5,texto6
        Cad = DevuelveDesdeBD("codmacta", "scompenclicab", "codigo", CStr(CONT))
        Cad = "Select nommacta,razosoci,dirdatos,codposta,despobla,desprovi from cuentas where codmacta ='" & Cad & "'"
        miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        'NO PUEDE SER EOF
        Cad = miRsAux!Nommacta
        If Not IsNull(miRsAux!razosoci) Then Cad = miRsAux!razosoci
        Cad = "'" & DevNombreSQL(Cad) & "'"
        'Direccion
        Cad = Cad & ",'" & DevNombreSQL(CStr(DBLet(miRsAux!dirdatos))) & "'"
        'Poblacion
        SQL = DBLet(miRsAux!codposta)
        If SQL <> "" Then SQL = SQL & " - "
        SQL = SQL & DevNombreSQL(CStr(DBLet(miRsAux!desPobla)))
        Cad = Cad & ",'" & SQL & "'"
        'Provincia
        Cad = Cad & ",'" & DevNombreSQL(CStr(DBLet(miRsAux!desProvi))) & "'"
        miRsAux.Close
        

        
        SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2, texto3, texto4,texto5,texto6, observa1, "
        SQL = SQL & "importe1, importe2, fecha1, fecha2, fecha3, observa2, opcion)"
        SQL = SQL & " VALUES (" & vUsu.Codigo & ",1,'',''," & Cad
        
        'select Numfac,fecha from usuarios.ztmpfaclin where tipoiva=1 and codusu=2200
        Importe = 0
        RC = "NIF"   'RC = "fecha"   La fecha de VTo esta en el campo: NIF
        Cad = DevuelveDesdeBD("numfac", "Usuarios.ztmpfaclin", "codusu =" & vUsu.Codigo & " AND tipoiva", "1", "N", RC)
        If Cad = "" Then
            'Significa que la compesacion ha sido total, no quedaba resultante
            
        Else
            Cad = "Quedando el resultado en el vencimiento: " & Cad & " de " & Format(RC, "dd/mm/yyyy")
            Importe = 1
        End If
        
        If Importe > 0 Then
            RC = "select sum(impvenci + coalesce(gastos,0) + coalesce(impcobro,0)) from  scompenclilin  WHERE codigo=" & CONT
            miRsAux.Open RC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            RC = "0"
            If Not miRsAux.EOF Then Importe = DBLet(miRsAux.Fields(0), "N")
            miRsAux.Close
        Else
            RC = "0"
        End If
        
        'observa 1 texto 6 e importe1
        SQL = SQL & ",'" & Cad & "'," & TransformaComasPuntos(CStr(Importe))
        
        
        'importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion
        For I = 1 To 6
            SQL = SQL & ",NULL"
        Next
        SQL = SQL & ")"
        Conn.Execute SQL
        
        With frmImprimir
                .OtrosParametros = ""
                .NumeroParametros = 0
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                
                .Opcion = 91
                .Show vbModal
            End With



Exit Sub
EImprimiCompensacion:
    MuestraError Err.Number, Err.Description
End Sub

Private Function RealizarProcesoCompensacionAbonos() As Boolean
Dim Destino As Byte
Dim J As Integer

    'NO USAR CONT

    RealizarProcesoCompensacionAbonos = False











    'Vamos a seleccionar los vtos
    '(numserie,codfaccl,fecfaccl,numorden)
    'EN SQL
    SQLVtosSeleccionadosCompensacion NumRegElim, False    'todos  -> Numregelim tendr el destino
    
    'Metemos los campos en el la tabla de lineas
    ' Esto guarda el valor en CAD
    FijaCadenaSQLCobrosCompen
    
    
    'Texto compensacion
    DevfrmCCtas = ""
    
    RC = "Select " & Cad & " FROM scobro where (numserie,codfaccl,fecfaccl,numorden) IN (" & SQL & ")"
    miRsAux.Open RC, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "Error. EOF vencimientos devueltos ", vbExclamation
        Exit Function
    End If
    
    
    I = 0
    
    While Not miRsAux.EOF
        I = I + 1
        BACKUP_Tabla miRsAux, RC
        'Quito los parentesis
        RC = Mid(RC, 1, Len(RC) - 1)
        RC = Mid(RC, 2)
        
        Destino = 0
        If miRsAux!NUmSerie = Me.lwCompenCli.ListItems(NumRegElim).Text Then
            If miRsAux!codfaccl = Val(Me.lwCompenCli.ListItems(NumRegElim).SubItems(1)) Then
                If Format(miRsAux!fecfaccl, "dd/mm/yyyy") = Me.lwCompenCli.ListItems(NumRegElim).SubItems(2) Then
                    If miRsAux!numorden = Val(Me.lwCompenCli.ListItems(NumRegElim).SubItems(3)) Then Destino = 1
                End If
            End If
        End If
        
        RC = "INSERT INTO scompenclilin (codigo,linea,destino," & Cad & ") VALUES (" & CONT & "," & I & "," & Destino & "," & RC & ")"
        Conn.Execute RC
        
        'Para las observaciones de despues
        Importe = DBLet(miRsAux!Gastos, "N")
        Importe = Importe + miRsAux!ImpVenci
        'Si ya he cobrado algo
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        
        If Destino = 0 Then 'El destino
            DevfrmCCtas = DevfrmCCtas & miRsAux!NUmSerie & Format(miRsAux!codfaccl, "00000") & " " & Format(miRsAux!fecfaccl, "dd/mm/yy")
            DevfrmCCtas = DevfrmCCtas & " Vto:" & Format(miRsAux!FecVenci, "dd/mm/yy") & " " & Importe
            DevfrmCCtas = DevfrmCCtas & "|"
        Else
            'El DESTINO siempre ira en la primera observacion del texto
            RC = "Importe anterior vto: " & Importe
            DevfrmCCtas = RC & "|" & DevfrmCCtas
        End If
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Acutalizaremos el VTO destino
    
    Conn.BeginTrans
        'BORRAREMOS LOS VENCIMIENTOS QUE NO SEAN DESTINO a no ser que el importe restante sea 0
        Destino = 1
        If txtimpNoEdit(0).Text = txtimpNoEdit(1).Text Then Destino = 0
        SQLVtosSeleccionadosCompensacion 0, Destino = 1  'sin o con el destino
        RC = "DELETE FROM scobro where (numserie,codfaccl,fecfaccl,numorden) IN (" & SQL & ")"
        
        'Para saber si ha ido bien
        Destino = 0    '0 mal,1 bien
        If Ejecuta(RC) Then
            If txtimpNoEdit(0).Text = txtimpNoEdit(1).Text Then
                Destino = 1
            Else
                'Updatearemos los campos csb del vto restante. A partir del segundo
                'La variable CadenaDesdeOtroForm  tiene los que vamos a actualizar
                
                Cad = ""
                J = 0
                SQL = ""
                
                Do
                    I = InStr(1, DevfrmCCtas, "|")
                    If I = 0 Then
                        DevfrmCCtas = ""
                    Else
                        RC = Mid(DevfrmCCtas, 1, I - 1)
                        If Len(RC) > 40 Then RC = Mid(RC, 1, 40)
                        DevfrmCCtas = Mid(DevfrmCCtas, I + 1)
                        J = J + 1
                        'Antes desde aqui cogia el campo
                        'Ahora desde CadenaDesdeOtroForm que tiene los campos libres
                        'Cad = RecuperaValor("text41csb|text42csb|text43csb|text51csb|text52csb|text53csb|text61csb|text62csb|text63csb|text71csb|text72csb|text73csb|text81csb|text82csb|text83csb|", J)
                        Cad = RecuperaValor(CadenaDesdeOtroForm, J)
                        SQL = SQL & ", " & Cad & " = '" & DevNombreSQL(RC) & "'"
                
                    End If
                Loop Until DevfrmCCtas = ""
                Importe = CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag)  'txtimpNoEdit(1).Tag es negativo
                RC = "gastos=null, impcobro=null,fecultco=null,impvenci=" & TransformaComasPuntos(CStr(Importe))
                SQL = RC & SQL
                SQL = "UPDATE scobro SET " & SQL
                'WHERE
                RC = ""
                For J = 1 To Me.lwCompenCli.ListItems.Count
                    If Me.lwCompenCli.ListItems(J).Bold Then
                        'Este es el destino
                        RC = "NUmSerie = '" & Me.lwCompenCli.ListItems(J).Text
                        RC = RC & "' AND codfaccl = " & Val(Me.lwCompenCli.ListItems(J).SubItems(1))
                        RC = RC & " AND fecfaccl = '" & Format(Me.lwCompenCli.ListItems(J).SubItems(2), FormatoFecha)
                        RC = RC & "' AND numorden = " & Val(Me.lwCompenCli.ListItems(J).SubItems(3))
                        Exit For
                    End If
                Next
                If RC <> "" Then
                    Cad = SQL & " WHERE " & RC
                    If Ejecuta(Cad) Then Destino = 1
                Else
                    MsgBox "No encontrado destino", vbExclamation
                    
                End If
            End If
        End If
        If Destino = 1 Then
            Conn.CommitTrans
            RealizarProcesoCompensacionAbonos = True
        Else
            Conn.RollbackTrans
        End If
        
End Function

Private Sub SQLVtosSeleccionadosCompensacion(ByRef RegistroDestino As Long, SinDestino As Boolean)
Dim Insertar As Boolean
    SQL = ""
    For I = 1 To Me.lwCompenCli.ListItems.Count
        If Me.lwCompenCli.ListItems(I).Checked Then
        
            Insertar = True
            If Me.lwCompenCli.ListItems(I).Bold Then
                RegistroDestino = I
                If SinDestino Then Insertar = False
            End If
            If Insertar Then
                SQL = SQL & ", ('" & lwCompenCli.ListItems(I).Text & "'," & lwCompenCli.ListItems(I).SubItems(1)
                SQL = SQL & ",'" & Format(lwCompenCli.ListItems(I).SubItems(2), FormatoFecha) & "'," & lwCompenCli.ListItems(I).SubItems(3) & ")"
            End If
            
        End If
    Next
    SQL = Mid(SQL, 2)
            
End Sub


Private Sub FijaCadenaSQLCobrosCompen()

    Cad = "NUmSerie , codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, ctabanc1,"
    Cad = Cad & "codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, emitdocum, "
    Cad = Cad & "recedocu, contdocu, text33csb, text41csb, text42csb, text43csb, text51csb, text52csb,"
    Cad = Cad & "text53csb, text61csb, text62csb, text63csb, text71csb, text72csb, text73csb, text81csb,"
    Cad = Cad & "text82csb, text83csb, ultimareclamacion, agente, departamento, tiporem, CodRem, AnyoRem,"
    Cad = Cad & "siturem, Gastos, Devuelto, situacionjuri, noremesar, obs, transfer, estacaja, referencia,"
    Cad = Cad & "reftalonpag, nomclien, domclien, pobclien, cpclien, proclien, referencia1, referencia2,"
    Cad = Cad & "feccomunica, fecprorroga, fecsiniestro"
    
End Sub




Private Function DesdeHastaAgenteCobrosParciales() As String

    'SQL desde/hasta
    SQL = ""
    DevfrmCCtas = "" 'VISREPORT. Desde hasta
    RC = ""
    If Text3(36).Text <> "" Then
        SQL = SQL & " AND fecha >= '" & Format(Text3(36).Text, FormatoFecha) & "'"
        RC = RC & " desde " & Text3(36).Text
    End If
    If Text3(37).Text <> "" Then
        SQL = SQL & " AND fecha <= '" & Format(Text3(37).Text, FormatoFecha) & "'"
        RC = RC & " hasta " & Text3(37).Text
    End If
    If RC <> "" Then DevfrmCCtas = "Fecha cobro " & RC
    RC = ""
    If Me.txtAgente(6).Text <> "" Then
        SQL = SQL & " AND codagent >=" & txtAgente(6).Text
        RC = RC & " desde " & Me.txtAgente(6).Text & " " & Me.txtDescAgente(6).Text
    End If
    If Me.txtAgente(7).Text <> "" Then
        SQL = SQL & " AND codagent <=" & txtAgente(7).Text
        RC = RC & " hasta " & Me.txtAgente(7).Text & " " & Me.txtDescAgente(7).Text
    End If
    If RC <> "" Then DevfrmCCtas = Trim(DevfrmCCtas & "    Agentes " & RC)
    
    
    
    If SQL <> "" Then SQL = Mid(SQL, 5) 'quito el primer AND
    DesdeHastaAgenteCobrosParciales = SQL


End Function



Private Function RealizarProcesoUpdateCobrosAgente() As Boolean
Dim RCob As ADODB.Recordset
Dim vLog As cLOG

    Set vLog = New cLOG
    On Error GoTo eRealizarProcesoUpdateCobrosAgente
    RealizarProcesoUpdateCobrosAgente = False
    Set miRsAux = New ADODB.Recordset
    Set RCob = New ADODB.Recordset
    
    'tiene que sumar lo cobrado por factura y saldar el cobro
    RC = DesdeHastaAgenteCobrosParciales
    Cad = "select numserie,codfaccl,fecfaccl,numorden,sum(importe) as cobrado,count(*) as cuantos from scobrolin "
    If RC <> "" Then Cad = Cad & " WHERE " & RC
    Cad = Cad & " group by 1,2,3,4 "
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Label3(50).Caption = "Cobro " & miRsAux!NUmSerie & miRsAux!codfaccl
        Label3(50).Refresh
        RC = "numserie = '" & miRsAux!NUmSerie & "' AND codfaccl =" & miRsAux!codfaccl
        RC = RC & " AND fecfaccl = '" & Format(miRsAux!fecfaccl, FormatoFecha) & "' AND numorden=" & miRsAux!numorden
        Cad = "SELECT * from scobro where " & RC
        RCob.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'NO PUEDE SER EOF
        Cad = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", RCob!codmacta, "T")
        Cad = Cad & " " & miRsAux!NUmSerie & miRsAux!codfaccl & " de " & Format(miRsAux!fecfaccl, "dd/mm/yyyy") & " -" & miRsAux!numorden
        Cad = Cad & vbCrLf & "€Recibo: " & RCob!ImpVenci
        If Not IsNull(RCob!Gastos) Then Cad = Cad & " Gastos:" & RCob!Gastos
        If Not IsNull(RCob!impcobro) Then Cad = Cad & " Ult cobro:" & RCob!impcobro
        Cad = Cad & vbCrLf & "€ Cobrados agente: " & miRsAux!Cobrado
        Cad = Cad & vbCrLf & "Nº cobros: " & miRsAux!Cuantos
        
        
        Importe = DBLet(RCob!impcobro, "N") + miRsAux!Cobrado
        
        If Importe < RCob!ImpVenci + DBLet(RCob!Gastos, "N") Then
            SQL = "UPDATE scobro  SET impcobro = " & TransformaComasPuntos(CStr(Importe))
            SQL = SQL & ", fecultco = '" & Format(Now, FormatoFecha) & "'"
        Else
            Importe = Importe - (RCob!ImpVenci + DBLet(RCob!Gastos, "N"))
            If Importe > 0 Then Cad = Cad & vbCrLf & "Diferencia postiva: " & Importe
            SQL = "DELETE FROM scobro "
        End If
        SQL = SQL & " WHERE " & RC
        Conn.Execute SQL
        vLog.Insertar 101, vUsu, Cad
        
        RCob.Close
        
        espera 1
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    'Eliminamos cobros parciales
    Label3(50).Caption = "Eliminando parciales"
    Label3(50).Refresh
    RC = DesdeHastaAgenteCobrosParciales
    Cad = "DELETE from scobrolin "
    If RC <> "" Then Cad = Cad & " WHERE " & RC
    Conn.Execute Cad
    
        
    
    RealizarProcesoUpdateCobrosAgente = True
    
eRealizarProcesoUpdateCobrosAgente:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set miRsAux = Nothing
    Set vLog = Nothing
    Label3(50).Caption = ""
    
End Function
