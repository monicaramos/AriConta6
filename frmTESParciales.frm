VERSION 5.00
Begin VB.Form frmTESParciales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anticipo vto."
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRecibo 
      Caption         =   "Recibo"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   5
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Frame FrCobro 
      Height          =   4335
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   0
         Text            =   "Text2"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtDescCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   5
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   5160
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   3915
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   4
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2820
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   3
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   960
         Width           =   4815
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   3915
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cta banco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   22
         Top             =   1920
         Width           =   735
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   1200
         Picture         =   "frmTESParciales.frx":0000
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   4440
         TabIndex        =   20
         Top             =   3960
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pagado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   4440
         TabIndex        =   19
         Top             =   3240
         Width           =   540
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   6600
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gastos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   4440
         TabIndex        =   16
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe TOTAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   14
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vencimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   3960
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   1080
         Picture         =   "frmTESParciales.frx":6852
         Top             =   3930
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Cobro"
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
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   360
         Width           =   4890
      End
   End
End
Attribute VB_Name = "frmTESParciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Public Cobro As Boolean
Public Vto As String  'Llevara empipado las claves
Public Cta As String
Public Importes As String 'Empipado los importes
Public FormaPago As Integer

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBa As frmBanco
Attribute frmBa.VB_VarHelpID = -1

Dim impo As Currency
Dim Cad As String


Private Sub PonFoco(ByRef T1 As TextBox)
    T1.SelStart = 0
    T1.SelLength = Len(T1.Text)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub








Private Sub cmdRecibo_Click()

    If ImporteFormateado(Text2(0).Text) <= 0 Then
        MsgBox "No se pueden emitir recibos por importes menores o iguales a cero", vbExclamation
        Exit Sub
    End If
    If GenerarRecibos Then
        'DevuelveCadenaPorTipo True, Cad
        'If Cad = "" Then Cad = "He recibido de:"
        'Cad = "textoherecibido= """ & Cad & """|"
        'Imprimimos
        'Para que imprima el mismo que por el punto: generar cobros por...
        CadenaDesdeOtroForm = DevuelveNombreInformeSCRYST(6, "Recibo")
        
        frmImprimir.Opcion = 8
        frmImprimir.NumeroParametros = 1
        frmImprimir.OtrosParametros = Cad
        frmImprimir.FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
        frmImprimir.Show vbModal
            
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim B As Boolean
    CadenaDesdeOtroForm = ""
    If Index = 0 Then
        'Comprobamos importes. Y fecha de contabilizacioon
        If Not DatosOk Then Exit Sub
        
        If Cobro Then
            CadenaDesdeOtroForm = "cobro"
        Else
            CadenaDesdeOtroForm = "pago"
        End If
        CadenaDesdeOtroForm = "Desea generar el " & CadenaDesdeOtroForm & "?"
        B = True
        If MsgBox(CadenaDesdeOtroForm, vbQuestion + vbYesNo) = vbNo Then B = False
        CadenaDesdeOtroForm = ""
        If Not B Then Exit Sub
        
        'UPDATEAMOS EL Vencimiento y CONTABILIZAMOS EL COBRO/PAGO
        Screen.MousePointer = vbHourglass
        B = RealizarAnticipo
        Screen.MousePointer = vbDefault
        If Not B Then Exit Sub
        CadenaDesdeOtroForm = "OK" 'Para que refresque los datos en el form
    End If
    Unload Me
End Sub

Private Sub Form_Load()
        
        
    If Cobro Then
        Caption = "Cobro"
        Text1(0).Text = RecuperaValor(Vto, 1) & " / " & RecuperaValor(Vto, 2) & "      Fecha: " & RecuperaValor(Vto, 3) & "       Vto. num: " & RecuperaValor(Vto, 4)
        Text1(1).Text = RecuperaValor(Cta, 1)
        Text1(2).Text = RecuperaValor(Cta, 2)
        'Dos
        txtCta(1).Text = RecuperaValor(Cta, 3)
        Me.txtDescCta(1).Text = RecuperaValor(Cta, 4)
        
        'Importes
        Text1(3).Text = RecuperaValor(Importes, 1)
        Text1(4).Text = RecuperaValor(Importes, 2)
        Text1(5).Text = RecuperaValor(Importes, 3)
        Text3(0).Text = Format(Now, "dd/mm/yyyy")
        Label4(4).Caption = "Gastos"
        Label4(1).Caption = "Cliente"
                
        
        
    Else
        'PAGO
        Label4(1).Caption = "Proveedor"
        Caption = "Pago"
        Label4(4).Caption = "Pagado"
        'Cobro parcial de vencimientos
        Text1(0).Text = RecuperaValor(Vto, 2) & "      Fecha: " & RecuperaValor(Vto, 3) & "       Vto. num: " & RecuperaValor(Vto, 4)
        Text1(1).Text = RecuperaValor(Cta, 1)
        Text1(2).Text = RecuperaValor(Cta, 2)
        'Dos
        txtCta(1).Text = RecuperaValor(Cta, 3)
        Me.txtDescCta(1).Text = RecuperaValor(Cta, 4)
        
        'Importes
        Text1(3).Text = RecuperaValor(Importes, 1)
        Text1(4).Text = RecuperaValor(Importes, 2)  'Esto es lo pagado ya
        '''''Text1(5).Text = RecuperaValor(Importes, 3)
        Text3(0).Text = Format(Now, "dd/mm/yyyy")
        
    End If
    
    
    'IMPORTE Restante
    
    impo = ImporteFormateado(Text1(3).Text) 'Vto
    If Cobro Then
        'Gastos
        If Text1(4).Text <> "" Then impo = impo + ImporteFormateado(Text1(4).Text)
            
        'Ya cobrado
        If Text1(5).Text <> "" Then impo = impo - ImporteFormateado(Text1(5).Text)
        
    Else
        'Ya cobrado
        If Text1(4).Text <> "" Then impo = impo - ImporteFormateado(Text1(4).Text)
            
    End If
    Label1.Caption = "Pendiente: " & Format(impo, FormatoImporte)
    
    
    
    cmdRecibo.Visible = Cobro
    Label4(5).Visible = Cobro
    Text1(5).Visible = Cobro
    Me.Height = Me.FrCobro.Height + 1200 '240 + Me.Command1(0).Height + 240
    'Text2(0).Text = ""
    Text2(0).Text = Format(impo, FormatoImporte)
    Label2(1).Caption = Caption & " de vencimientos"
    Caption = Caption & " de vencimientos"
End Sub

Private Sub frmBa_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtCta(CInt(imgCuentas(1).Tag)).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescCta(CInt(imgCuentas(1).Tag)).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text3(CInt(Text3(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image2_Click(Index As Integer)
    
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text3(Index).Text <> "" Then frmC.Fecha = CDate(Text3(Index).Text)
    Text3(0).Tag = Index
    frmC.Show vbModal
    Set frmC = Nothing
End Sub



Private Sub imgCuentas_Click(Index As Integer)
    imgCuentas(1).Tag = Index
    Set frmBa = New frmBanco
    frmBa.DatosADevolverBusqueda = "OK"
    frmBa.Show vbModal
    Set frmBa = Nothing
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    PonFoco Text2(Index)
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text2_LostFocus(Index As Integer)
Dim Valor

        If Text2(Index).Text = "" Then Exit Sub
        If Not IsNumeric(Text2(Index).Text) Then
            MsgBox "importe debe ser numérico", vbExclamation
            Text2(Index).Text = ""
            PonerFoco Text2(Index)
        Else
            If InStr(1, Text2(Index).Text, ",") > 0 Then
                Valor = ImporteFormateado(Text2(Index).Text)
            Else
                Valor = CCur(TransformaPuntosComas(Text2(Index).Text))
            End If
            Text2(Index).Text = Format(Valor, FormatoImporte)
        End If
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    PonFoco Text3(Index)
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index))
    If Text3(Index) = "" Then Exit Sub
    If Not EsFechaOK(Text3(Index)) Then
        MsgBox "Fecha incorrecta: " & Text3(Index), vbExclamation
        Text3(Index).Text = ""
        Text3(Index).SetFocus
    End If
End Sub






Private Function GenerarRecibos() As Boolean
Dim SQL As String
Dim Poblacion As String
Dim J As Integer
Dim Aux As String

    On Error GoTo EGenerarRecibos
    GenerarRecibos = False
    
    
    If Text2(0).Text = "" Then
        MsgBox "Falta importe", vbExclamation
        Exit Function
    End If
    
        
    
    'Limpiamos
    Cad = "Delete from Usuarios.zTesoreriaComun where codusu = " & vUsu.Codigo
    Conn.Execute Cad


    'Guardamos datos empresa
    Cad = "Delete from Usuarios.z347carta where codusu = " & vUsu.Codigo
    Conn.Execute Cad
    
    Cad = "INSERT INTO Usuarios.z347carta (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir, saludos, "
    Cad = Cad & "parrafo1, parrafo2, parrafo3, parrafo4, parrafo5, despedida, contacto, Asunto, Referencia)"
    Cad = Cad & " VALUES (" & vUsu.Codigo & ", "
    
    'Estos datos ya veremos com, y cuadno los relleno
    Set miRsAux = New ADODB.Recordset
    SQL = "select nifempre,siglasvia,direccion,numero,escalera,piso,puerta,codpos,poblacion,provincia from empresa2"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    'sql= "'1234567890A','Ariadna Software ','Franco Tormo 3, Bajo Izda','46007','Valencia'"
    SQL = "'##########','" & vEmpresa.nomempre & "','#############','######','##########','##########'"
    If Not miRsAux.EOF Then
        SQL = ""
        For J = 1 To 6
            SQL = SQL & DBLet(miRsAux.Fields(J), "T") & " "
        Next J
        SQL = Trim(SQL)
        SQL = "'" & DBLet(miRsAux!nifempre, "T") & "','" & DevNombreSQL(vEmpresa.nomempre) & "','" & DevNombreSQL(SQL) & "'"
        SQL = SQL & ",'" & DBLet(miRsAux!codpos, "T") & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!Poblacion, "T")) & "'"
        Poblacion = DevNombreSQL(DBLet(miRsAux!Poblacion, "T"))
    End If
    miRsAux.Close
    
    Cad = Cad & SQL
    'otralinea,saludos
    Cad = Cad & ",NULL"
    'parrafo1
    Cad = Cad & ",''"
    
    
    '------------------------------------------------------------------------
    Cad = Cad & ",NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
    Conn.Execute Cad
    
    
    'NUEVO JUNIO 2011
    'Me cargo lo de aqui abajo
    'y traemos un sub desde el punto de menu imprimirrecibos
     
    RellenarCadenaSQLRecibo Poblacion
        
        
        
   
   


    GenerarRecibos = True
EGenerarRecibos:
    If Err.Number <> 0 Then
        MuestraError Err.Number
    End If
    Set miRsAux = Nothing
End Function


Private Function DatosOk() As Boolean
Dim Im As Currency

    On Error GoTo EDa
    DatosOk = False
    
    Cad = ""
    If Text2(0).Text = "" Then Cad = "importe"
    If Text3(0).Text = "" Then Cad = Cad & " fecha"
    If Cad <> "" Then
        MsgBox "Falta: " & Cad, vbExclamation
        Exit Function
    End If
    
    '----------------------------------
    'Junio 2011
    'YA dejamos cobros negativos
    Im = ImporteFormateado(Text2(0).Text)
    'If Im < 0 Then
    If Im = 0 Then
        'MsgBox "importes negativos", vbExclamation
        MsgBox "importes CERO", vbExclamation
        Exit Function
    End If
    
    
    If txtCta(1).Text = "" Then
        MsgBox "Falta cuenta banco", vbExclamation
        Exit Function
    End If
        
    
    'Fecha dentro ejercicios
    If FechaCorrecta2(CDate(Text3(0).Text), True) > 1 Then Exit Function
    
    
    
    If Cobro Then
        impo = ImporteFormateado(Text1(3).Text) 'Vto
        'Gastos
        If Text1(4).Text <> "" Then
            Im = ImporteFormateado(Text1(4).Text)
            impo = impo + Im
        End If
        
        'Ya cobrado
        If Text1(5).Text <> "" Then
            Im = ImporteFormateado(Text1(5).Text)
            impo = impo - Im
        End If
        
        
    Else
        impo = ImporteFormateado(Text1(3).Text) 'Vto
        'Gastos

        'Ya cobrado
        If Text1(4).Text <> "" Then
            Im = ImporteFormateado(Text1(4).Text)
            impo = impo - Im
        End If
        
        
    End If
    
    
    
    Im = ImporteFormateado(Text2(0).Text) 'Lo que voy a pagar
    Cad = ""
    If impo < 0 Then
        'Importes negativos
        If Im >= 0 Then
            Cad = "negativo"
        Else
            If Im < impo Then Cad = "X"
        End If
    Else
        If Im <= 0 Then
            Cad = "positivo"
        Else
            If Im > impo Then Cad = "X"
        End If
    End If
        
    If Cad <> "" Then
        
        If Cad = "X" Then
            Cad = "Importe a pagar mayor que el importe restante.(" & Format(Im, FormatoImporte) & " : " & Format(impo, FormatoImporte) & ")"
        Else
            Cad = "El importe debe ser " & Cad
        End If
        MsgBox Cad, vbExclamation
        Exit Function
    End If
    
        
        
    'Comprobaremos un par de cosillas
    If CuentaBloqeada(RecuperaValor(Cta, 1), CDate(Text3(0).Text), True) Then Exit Function
        
        
        
    DatosOk = True
    Exit Function
EDa:
    MuestraError Err.Number, "Datos Ok"
End Function


Private Function RealizarAnticipo() As Boolean

        Cad = "DELETE from tmpactualizar  where codusu =" & vUsu.Codigo
        Conn.Execute Cad
    
    
        Conn.BeginTrans
        If Contabilizar Then
            Conn.CommitTrans
            RealizarAnticipo = True
            
            
            '-----------------------------------------------------------
            'Ahora actualizamos los registros que estan en tmpactualziar
            frmActualizar2.OpcionActualizar = 20
            frmActualizar2.Show vbModal
            
        Else
            'Conn.RollbackTrans
            TirarAtrasTransaccion
            RealizarAnticipo = False
        End If
End Function


Private Function Contabilizar() As Boolean
Dim Mc As Contadores
Dim FP As Ctipoformapago
Dim SQL As String
Dim Ampliacion As String
Dim Numdocum As String
Dim Conce As Integer
Dim LlevaContr As Boolean
Dim Im As Currency
Dim Debe As Boolean
Dim ElConcepto As Integer
    On Error GoTo ECon
    Contabilizar = False
    Set Mc = New Contadores
    If Mc.ConseguirContador("0", CDate(Text3(0).Text) <= vParam.fechafin, True) = 1 Then Exit Function

    Set FP = New Ctipoformapago
    If FP.Leer(FormaPago) Then
        Set Mc = Nothing
        Set FP = Nothing
    End If
    
    'importe
    impo = ImporteFormateado(Text2(0).Text)
    
    'Inserto cabecera de apunte
    SQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) VALUES ("
    If Cobro Then
        SQL = SQL & FP.diaricli
    Else
        SQL = SQL & FP.diaripro
    End If
    SQL = SQL & ",'" & Format(Text3(0).Text, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ", 1, NULL, '"
    SQL = SQL & "Generado desde Tesorería el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & DevNombreSQL(vUsu.Nombre)
    If impo < 0 Then SQL = SQL & "  (ABONO)"
    SQL = SQL & "')"
    Conn.Execute SQL
        
        
    'Inserto en las lineas de apuntes
    SQL = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, "
    SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
    SQL = SQL & " timporteH, codccost, ctacontr, idcontab, punteada) VALUES ("
    If Cobro Then
        SQL = SQL & FP.diaricli
    Else
        SQL = SQL & FP.diaripro
    End If
    SQL = SQL & ",'" & Format(Text3(0).Text, FormatoFecha) & "'," & Mc.Contador & ","
    
    
    
    
    'numdocum
    Numdocum = DevNombreSQL(RecuperaValor(Vto, 2))
    If Cobro Then
        'Antes 21 Sept 2011
        'Numdocum = RecuperaValor(Vto, 1) & Format(Mid(Numdocum, 1, 9), "000000000")
        'Ahora
        Numdocum = SerieNumeroFactura(10, RecuperaValor(Vto, 1), Numdocum)
    End If
    
    
    
    'Concepto y ampliacion del apunte
    Ampliacion = ""
    If Cobro Then
        'CLIENTES
        Debe = False
        If impo < 0 Then
            If Not vParam.abononeg Then Debe = True
        End If
        If Debe Then
            Conce = FP.ampdecli
            LlevaContr = FP.ctrdecli = 1
            ElConcepto = FP.condecli
        Else
            ElConcepto = FP.conhacli
            Conce = FP.amphacli
            LlevaContr = FP.ctrhacli = 1
        End If
    Else
        'PAGOS
        Debe = True
        If impo < 0 Then
            If Not vParam.abononeg Then Debe = False
        End If
        If Debe Then
            Conce = FP.ampdepro
            LlevaContr = FP.ctrdepro = 1
            ElConcepto = FP.condepro
        Else
            ElConcepto = FP.conhapro
            Conce = FP.amphapro
            LlevaContr = FP.ctrhapro = 1
        End If

    End If
           
    'Si el importe es negativo y no permite abonos negativos
    'como ya lo ha cambiado de lado (dbe <-> haber)
    If impo < 0 Then
        If Not vParam.abononeg Then impo = Abs(impo)
    End If
       
           
           
    If Conce = 2 Then
       Ampliacion = Ampliacion & RecuperaValor(Vto, 3)  'Fecha vto
    ElseIf Conce = 4 Then
        'Contra partida
        Ampliacion = DevNombreSQL(txtDescCta(1).Text)
    Else
        
       If Conce = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
       If Cobro Then
            Ampliacion = Ampliacion & RecuperaValor(Vto, 1) & "/" & Mid(RecuperaValor(Vto, 2), 1, 9)
       Else
            Ampliacion = Ampliacion & Mid(RecuperaValor(Vto, 2), 1, 9)
       End If
    End If
    
    'Fijo en concepto el codconce
    Conce = ElConcepto
    Cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce), "N")
    Ampliacion = Cad & " " & Ampliacion
    Ampliacion = Mid(Ampliacion, 1, 35)
    
    
    
    
    'Ahora ponemos linliapu codmacta numdocum codconce ampconce timported timporte codccost ctacontr idcontab punteada
    'Cuenta Cliente/proveedor
    Cad = "1,'" & Text1(1).Text & "','" & Numdocum & "'," & Conce & ",'" & DevNombreSQL(Ampliacion) & "',"
    'Importe cobro-pago
    ' nos lo dire "debe"
    If Not Debe Then
        Cad = Cad & "NULL," & TransformaComasPuntos(CStr(impo))
    Else
        Cad = Cad & TransformaComasPuntos(CStr(impo)) & ",NULL"
    End If
    'Codccost
    Cad = Cad & ",NULL,"
    If LlevaContr Then
        Cad = Cad & "'" & txtCta(1).Text & "'"
    Else
        Cad = Cad & "NULL"
    End If
    Cad = Cad & ",'contab',0)"
    Cad = SQL & Cad
    Conn.Execute Cad
    
       
    'El banco    *******************************************************************************
    '---------------------------------------------------------------------------------------------
    
    'Vuelvo a fijar los valores
     'Concepto y ampliacion del apunte
    Ampliacion = ""
    If Cobro Then
       'CLIENTES
        'Si el apunte va al debe, el contrapunte va al haber
        If Not Debe Then
            Conce = FP.ampdecli
            LlevaContr = FP.ctrdecli = 1
            ElConcepto = FP.condecli
        Else
            ElConcepto = FP.conhacli
            Conce = FP.amphacli
            LlevaContr = FP.ctrhacli = 1
        End If
    Else
        'PAGOS
        'Si el apunte va al debe, el contrapunte va al haber
        If Not Debe Then
            Conce = FP.ampdepro
            LlevaContr = FP.ctrdepro = 1
            ElConcepto = FP.condepro
        Else
            ElConcepto = FP.conhapro
            Conce = FP.amphapro
            LlevaContr = FP.ctrhapro = 1
        End If
    End If
           
           
           
           
           
           
    If Conce = 2 Then
       Ampliacion = Ampliacion & RecuperaValor(Vto, 3)  'Fecha vto
    ElseIf Conce = 4 Then
        'Contra partida
        Ampliacion = DevNombreSQL(Text1(2).Text)
    Else
        
       If Conce = 1 Then Ampliacion = Ampliacion & FP.siglas & " "
       If Cobro Then
            Ampliacion = Ampliacion & RecuperaValor(Vto, 1) & "/" & Mid(RecuperaValor(Vto, 2), 1, 9)
       Else
            Ampliacion = Ampliacion & Mid(RecuperaValor(Vto, 2), 1, 9)
       End If
    End If
    
    
    Conce = ElConcepto
    Cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce), "N")
    Ampliacion = Cad & " " & Ampliacion
    Ampliacion = Mid(Ampliacion, 1, 35)
    
    
    
    
    
    
    
    
    
    Cad = "2,'" & txtCta(1).Text & "','" & Numdocum & "'," & Conce & ",'" & Ampliacion & "',"
    'Importe cliente
    'Si el cobro/pago va al debe el contrapunte ira al haber
    If Not Debe Then
        'al debe
        Cad = Cad & TransformaComasPuntos(CStr(impo)) & ",NULL"
    Else
        'al haber
        Cad = Cad & "NULL," & TransformaComasPuntos(CStr(impo))
    End If
    
    'Codccost
    Cad = Cad & ",NULL,"
    
    If LlevaContr Then
        Cad = Cad & "'" & Text1(1).Text & "'"
    Else
        Cad = Cad & "NULL"
    End If
    Cad = Cad & ",'idcontab',0)"
    Cad = SQL & Cad
    Conn.Execute Cad
    
    'Insertamos en la temporal para que lo ac
    If Cobro Then
        SQL = FP.diaricli
    Else
        SQL = FP.diaripro
    End If

    InsertaTmpActualizar Mc.Contador, SQL, Text3(0).Text
    
    'Actualizamos VTO
    ' o lo eliminamos. Segun sea el importe que falte
    'Tomomos prestada LlevaContr
    
    Im = ImporteFormateado(Text2(0).Text)  'lo que voy a anticipar
    
    impo = ImporteFormateado(Text1(3).Text)  'lo que me falta
    If Cobro Then
        If Text1(4).Text <> "" Then impo = impo + ImporteFormateado(Text1(4).Text)
        If Text1(5).Text <> "" Then impo = impo - ImporteFormateado(Text1(5).Text)
    Else
        If Text1(4).Text <> "" Then impo = impo - ImporteFormateado(Text1(4).Text)
    End If
    If impo - Im = 0 Then
        LlevaContr = True  'ELIMINAR VTO ya que esta totalmente pagado
    Else
        LlevaContr = False
    End If
    
    
    impo = ImporteFormateado(Text2(0).Text)
    If Cobro Then
        SQL = "scobro"
        Ampliacion = "fecultco"
        Numdocum = "impcobro"
        'El importe es el total. Lo que ya llevaba mas lo de ahora
        If Text1(5).Text <> "" Then impo = impo + ImporteFormateado(Text1(5).Text)
    Else
        
        SQL = "spagop"
        Ampliacion = "fecultpa"
        Numdocum = "imppagad"
        'El importe es el total. Lo que ya llevaba mas lo de ahora
        If Text1(4).Text <> "" Then impo = impo + ImporteFormateado(Text1(4).Text)
    End If
    
    
    If LlevaContr Then
        'ELIMINAMOS VTO
        Cad = "DELETE FROM " & SQL
    Else
        'UPDATEMAOS IMPORTES y fec ult pago/cobro
        Cad = "UPDATE " & SQL & " SET " & Ampliacion & " = '" & Format(Text3(0).Text, FormatoFecha) & "' , " & Numdocum & " = " & TransformaComasPuntos(CStr(impo))
    End If
    
    
    
    
    'EL WHERE
    If Cobro Then
        SQL = "numserie = '" & RecuperaValor(Vto, 1) & "' AND codfaccl=" & RecuperaValor(Vto, 2) & " and fecfaccl='" & Format(RecuperaValor(Vto, 3), FormatoFecha)
        SQL = SQL & "' AND numorden =" & RecuperaValor(Vto, 4)
    Else
        SQL = "ctaprove = '" & Text1(1).Text & "' AND numfactu='" & DevNombreSQL(RecuperaValor(Vto, 2)) & "' and fecfactu='" & Format(RecuperaValor(Vto, 3), FormatoFecha)
        SQL = SQL & "' AND numorden =" & RecuperaValor(Vto, 4)
    End If
    Cad = Cad & " WHERE " & SQL
    Conn.Execute Cad
    
    
    Contabilizar = True

    Set Mc = Nothing
    Set FP = Nothing

    Exit Function
ECon:
    MuestraError Err.Number, "Contabilizar anticipo"
    Set Mc = Nothing
    Set FP = Nothing
End Function
    
Private Sub txtCta_GotFocus(Index As Integer)
    PonFoco txtCta(1)
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCta_LostFocus(Index As Integer)

        txtCta(Index).Text = Trim(txtCta(Index).Text)
        Cad = txtCta(Index).Text
        impo = 0
        If Cad <> "" Then
            If CuentaCorrectaUltimoNivel(Cad, CadenaDesdeOtroForm) Then
                Cad = DevuelveDesdeBD("codmacta", "ctabancaria", "codmacta", Cad, "T")
                If Cad = "" Then
                    CadenaDesdeOtroForm = ""
                    MsgBox "La cuenta contable no esta asociada a ninguna cuenta bancaria", vbExclamation
                End If
            Else
                MsgBox CadenaDesdeOtroForm, vbExclamation
                Cad = ""
                CadenaDesdeOtroForm = ""
            End If
            impo = 1
        Else
            CadenaDesdeOtroForm = ""
        End If
        
        
        txtCta(Index).Text = Cad
        txtDescCta(Index).Text = CadenaDesdeOtroForm
        If Cad = "" And impo <> 0 Then
            PonerFoco txtCta(Index)
        End If
        CadenaDesdeOtroForm = ""
End Sub





'TROZO COPIADO DESDE frmcobrosimprimir
'Modificado para cuadrar datos
Private Sub RellenarCadenaSQLRecibo(Lugar As String)
Dim Aux As String
Dim QueDireccionMostrar As Byte
    '0. NO tiene
    '1. La del recibo
    '2. La de la cuenta

    

  
      ' IRan:   text5:  nomclien
      '         texto6: domclien
      '         observa2  cpclien  pobclien    + vbcrlf + proclien
  
      Cad = "select nomclien,domclien,pobclien,cpclien,proclien,razosoci,dirdatos,codposta,despobla,desprovi"
      'MAYO 2010
      Cad = Cad & ",codbanco,codsucur,digcontr,scobro.cuentaba,scobro.codmacta"
      Cad = Cad & " from scobro,cuentas where scobro.codmacta =cuentas.codmacta and"
      Cad = Cad & " numserie ='" & RecuperaValor(Me.Vto, 1) & "' and codfaccl=" & RecuperaValor(Me.Vto, 2)
      Cad = Cad & " and fecfaccl='" & Format(RecuperaValor(Me.Vto, 3), FormatoFecha) & "' and numorden=" & RecuperaValor(Me.Vto, 4)
      miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
      
      If Not miRsAux.EOF Then
          'El vto NO tiene datos de fiscales
          Aux = DBLet(miRsAux!nomclien, "T")
          If Aux = "" Then
              'La cuenta tampoco los tiene
              If IsNull(miRsAux!dirdatos) Then
                  QueDireccionMostrar = 0
              Else
                  QueDireccionMostrar = 2
              End If
          Else
              QueDireccionMostrar = 1
          End If
      Else
          QueDireccionMostrar = 0
      End If
      
      'texto1 , texto2, texto3, texto4, texto5,
      'texto6, importe1, importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion
      
          
      'Textos
      '---------
      '1.- Recibo nª
      Cad = "1,'" & RecuperaValor(Vto, 1) & "/" & RecuperaValor(Vto, 2) & "'"
      
     
      
      
      'Lugar Vencimiento
      Cad = Cad & ",'" & Lugar & "'"
      
      'text3 mostrare el codmacta
      'Cad = Cad & ",'" & DevNombreSQL(.SubItems(5)) & "',"
      Cad = Cad & ",'" & miRsAux!codmacta & "',"
      
      'MAYO 2010.    Ahora en este campo ira el CCC del cliente si es que lo tiene
      'Cad = Cad & "'" & .SubItems(6) & "'," ANTES
      Aux = DBLet(miRsAux!codbanco, "N")
      If Aux = "" Or Aux = "0" Then
          Aux = "NULL"
      Else
          'codbanco,codsucur,digcontr,cuentaba
          Aux = Format(DBLet(miRsAux!codbanco, "N"), "0000")
          Aux = Aux & " " & Format(DBLet(miRsAux!codsucur, "N"), "0000") & " "
          Aux = Aux & Mid(DBLet(miRsAux!digcontr, "T") & "  ", 1, 2) & " "
          Aux = Aux & Right(String(10, "0") & DBLet(miRsAux!Cuentaba, "N"), 10)
          Aux = "'" & Aux & "'"
      End If
      Cad = Cad & Aux & ","
  
      '5 y 6.
      'text5: nomclien
      'texto6:domclien
      If QueDireccionMostrar = 0 Then
          'Cad = Cad & "NULL,NULL"
          'Siempre el nomclien
          'cad = cad & "'" & DevNombreSQL(.SubItems(5)) & "',NULL"
      Else
          If QueDireccionMostrar = 1 Then
              Cad = Cad & "'" & DevNombreSQL(DBLet(miRsAux!nomclien, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!domclien, "T")) & "'"
          Else
              Cad = Cad & "'" & DevNombreSQL(DBLet(miRsAux!razosoci, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!dirdatos, "T")) & "'"
          End If
      End If
      
      
      
      'IMPORTES
      '--------------------
      Cad = Cad & "," & TransformaComasPuntos(CStr(ImporteFormateado(Text2(0).Text)))
      
      'El segundo importe NULL
      Cad = Cad & ",NULL"
      
      'FECFAS
      '--------------
      'Libramiento o pago     Auqi pone NOW
      Cad = Cad & ",'" & Format(Text3(0).Text, FormatoFecha) & "'"
      Cad = Cad & ",'" & Format(RecuperaValor(Vto, 5), FormatoFecha) & "'"
      
      '3era fecha  NULL
      Cad = Cad & ",NULL"
      
      'OBSERVACIONES
      '------------------
      Aux = EscribeImporteLetra(ImporteFormateado(Text2(0).Text))
      
      Aux = "       ** " & Aux
      Cad = Cad & ",'" & Aux & "**',"
      
      
      'Observa 2
      '         observa2:    cpclien  pobclien    + vbcrlf + proclien
      If QueDireccionMostrar = 0 Then
          Aux = "NULL"
      Else
          
          If QueDireccionMostrar = 1 Then
              Aux = DBLet(miRsAux!cpclien, "T") & "      " & DevNombreSQL(DBLet(miRsAux!pobclien, "T"))
              Aux = Trim(Aux)
              If Aux <> "" Then Aux = Aux & vbCrLf
              Aux = Aux & DevNombreSQL(DBLet(miRsAux!proclien, "T"))
          Else
              Aux = DBLet(miRsAux!codposta, "T") & "      " & DevNombreSQL(DBLet(miRsAux!desPobla, "T"))
              Aux = Trim(Aux)
              If Aux <> "" Then Aux = Aux & vbCrLf
              Aux = Aux & DevNombreSQL(DBLet(miRsAux!desProvi, "T"))
              
          End If
          Aux = "'" & Aux & "'"
      End If
      Cad = Cad & Aux
      
      
      
      'OPCION
      '--------------
      Cad = Cad & ",NULL)"
      
    'El sql completo
    Aux = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2, texto3, texto4, texto5, "
    Aux = Aux & "texto6, importe1, importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion)"
    Aux = Aux & " VALUES (" & vUsu.Codigo & "," & Cad
      
      
    Conn.Execute Aux

    miRsAux.Close
End Sub





