VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESInfSituacion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7290
   Icon            =   "frmTESInfSituacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConcepto 
      Caption         =   "Selección"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   16
      Top             =   30
      Width           =   6945
      Begin VB.TextBox txtFecha 
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
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "imgConcepto"
         Top             =   1140
         Width           =   1305
      End
      Begin VB.TextBox txtDias 
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
         Left            =   1350
         TabIndex        =   1
         Tag             =   "imgConcepto"
         Top             =   2010
         Width           =   1275
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   6270
         TabIndex        =   22
         Top             =   270
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ayuda"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   300
         TabIndex        =   24
         Top             =   1140
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Fechas"
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
         Left            =   270
         TabIndex        =   23
         Top             =   750
         Width           =   2280
      End
      Begin VB.Image ImgFec 
         Height          =   240
         Index           =   0
         Left            =   1050
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Días Anteriores"
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
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Top             =   1650
         Width           =   1890
      End
      Begin VB.Label lblFecha 
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
         Index           =   0
         Left            =   2580
         TabIndex        =   20
         Top             =   3630
         Width           =   4095
      End
      Begin VB.Label lblFecha1 
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
         Index           =   4
         Left            =   2580
         TabIndex        =   19
         Top             =   3990
         Width           =   4095
      End
      Begin VB.Label lblNumFactu 
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
         Index           =   0
         Left            =   2610
         TabIndex        =   18
         Top             =   2340
         Width           =   4035
      End
      Begin VB.Label lblNumFactu 
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
         Index           =   1
         Left            =   2580
         TabIndex        =   17
         Top             =   2700
         Width           =   4035
      End
   End
   Begin VB.Frame FrameTipoSalida 
      Caption         =   "Tipo de salida"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   150
      TabIndex        =   5
      Top             =   3090
      Width           =   6915
      Begin VB.OptionButton optTipoSal 
         Caption         =   "Impresora"
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
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "Archivo csv"
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
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1515
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "PDF"
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
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton optTipoSal 
         Caption         =   "eMail"
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
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtTipoSalida 
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
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   720
         Width           =   3345
      End
      Begin VB.TextBox txtTipoSalida 
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
         Index           =   1
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1200
         Width           =   4665
      End
      Begin VB.TextBox txtTipoSalida 
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
         Index           =   2
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1680
         Width           =   4665
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   0
         Left            =   6450
         TabIndex        =   8
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton PushButton2 
         Caption         =   ".."
         Height          =   315
         Index           =   1
         Left            =   6450
         TabIndex        =   7
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton PushButtonImpr 
         Caption         =   "Propiedades"
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
         Left            =   5190
         TabIndex        =   6
         Top             =   720
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
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
      Left            =   5850
      TabIndex        =   3
      Top             =   5940
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccion 
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
      Index           =   1
      Left            =   4260
      TabIndex        =   2
      Top             =   5940
      Width           =   1455
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "&Imprimir"
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
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   5910
      Width           =   1335
   End
End
Attribute VB_Name = "frmTESInfSituacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 903

' ***********************************************************************************************************
' ***********************************************************************************************************
' ***********************************************************************************************************
'
'  3 espacios
'       -Los desde hasta,
'       -las opciones / ordenacion
'       -el tipo salida
'
' ***********************************************************************************************************
' ***********************************************************************************************************
' ***********************************************************************************************************
Public numero As String

Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1

Private SQL As String
Dim cad As String
Dim RC As String
Dim I As Integer
Dim IndCodigo As Integer
Dim tabla As String

Dim F1 As Date
Dim F2 As Date

Dim PrimeraVez As Boolean
Dim Cancelado As Boolean

Public Sub InicializarVbles(AñadireElDeEmpresa As Boolean)
    cadFormula = ""
    cadselect = ""
    cadParam = "|"
    numParam = 0
    cadNomRPT = ""
    conSubRPT = False
    cadPDFrpt = ""
    ExportarPDF = False
    vMostrarTree = False
    
    If AñadireElDeEmpresa Then
        cadParam = cadParam & "pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    End If
    
End Sub

Private Function MontaSQL() As Boolean
Dim SQL As String
Dim SQL2 As String
Dim RC As String
Dim RC2 As String

    MontaSQL = False
    
    If Not PonerDesdeHasta("cobros.FecFactu", "F", Me.txtFecha(0), Me.txtFecha(0), Me.txtFecha(1), Me.txtFecha(1), "pDHFecha=""") Then Exit Function
    If Not PonerDesdeHasta("cobros.codmacta", "N", Me.txtCuentas(0), Me.txtNCuentas(0), Me.txtCuentas(1), Me.txtNCuentas(1), "pDHCuentas=""") Then Exit Function
            
    MontaSQL = True
End Function


Private Sub cmdAccion_Click(Index As Integer)

    If Not DatosOK Then Exit Sub
    
    
    'Exportacion a PDF
    If optTipoSal(3).Value + optTipoSal(2).Value + optTipoSal(1).Value Then
        If Not EliminarDocum(optTipoSal(2).Value) Then Exit Sub
    End If
    
    
    InicializarVbles True
    
    tabla = " "
    
    If Not MontaSQL Then Exit Sub
    
    If Not CargarTemporales Then Exit Sub
    
    If Not HayRegParaInforme("tmptesoreriacomun", "codusu = " & vUsu.Codigo) Then Exit Sub
    
    If optTipoSal(1).Value Then
        'EXPORTAR A CSV
        AccionesCSV
    
    Else
        'Tanto a pdf,imprimiir, preevisualizar como email van COntral Crystal
    
        If optTipoSal(2).Value Or optTipoSal(3).Value Then
            ExportarPDF = True 'generaremos el pdf
        Else
            ExportarPDF = False
        End If
        SoloImprimir = False
        If Index = 0 Then SoloImprimir = True 'ha pulsado impirmir
        
        AccionesCrystal
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        txtFecha(0).Text = Format(Now, "dd/mm/yyyy")
        txtDias(0).Text = 90
        
        PonFoco txtFecha(0)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub Form_Load()
    PrimeraVez = True
    Me.Icon = frmPpal.Icon
        
    'Otras opciones
    Me.Caption = "Informe de Situación"

    For I = 0 To 0
        Me.ImgFec(I).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next I
     
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
    
     
    PonerDatosPorDefectoImpresion Me, False, Me.Caption 'Siempre tiene que tener el frame con txtTipoSalida
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), 0
    
    
    
End Sub


Private Sub frmF_Selec(vFecha As Date)
    txtFecha(IndCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub imgFec_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    Case 0
        IndCodigo = Index
    
        'FECHA
        Set frmF = New frmCal
        frmF.Fecha = Now
        If txtFecha(Index).Text <> "" Then frmF.Fecha = CDate(txtFecha(Index).Text)
        frmF.Show vbModal
        Set frmF = Nothing
        PonFoco txtFecha(Index)
    End Select
    
    Screen.MousePointer = vbDefault

End Sub


Private Sub optTipoSal_Click(Index As Integer)
    ponerLabelBotonImpresion cmdAccion(1), cmdAccion(0), Index
End Sub

Private Sub optVarios_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub PushButton2_Click(Index As Integer)
    'FILTROS
    If Index = 0 Then
        frmPpal.cd1.Filter = "*.csv|*.csv"
         
    Else
        frmPpal.cd1.Filter = "*.pdf|*.pdf"
    End If
    frmPpal.cd1.InitDir = App.Path & "\Exportar" 'PathSalida
    frmPpal.cd1.FilterIndex = 1
    frmPpal.cd1.ShowSave
    If frmPpal.cd1.FileTitle <> "" Then
        If Dir(frmPpal.cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo ya existe. Reemplazar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        txtTipoSalida(Index + 1).Text = frmPpal.cd1.FileName
    End If
End Sub

Private Sub PushButtonImpr_Click()
    frmPpal.cd1.ShowPrinter
    PonerDatosPorDefectoImpresion Me, True
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub




Private Sub LanzaFormAyuda(Nombre As String, Indice As Integer)
    Select Case Nombre
    Case "imgFecha"
        imgFec_Click Indice
    End Select
End Sub


Private Sub AccionesCSV()
Dim SQL2 As String

    'Monto el SQL
    
    SQL = "SELECT `tmptesoreriacomun`.`texto1` cuenta , `tmptesoreriacomun`.`texto2` Conta, `tmptesoreriacomun`.`opcion` BD, `tmptesoreriacomun`.`texto5` Nombre, `tmptesoreriacomun`.`texto3` NroFra, `tmptesoreriacomun`.`fecha1` FecFra, `tmptesoreriacomun`.`fecha2` FecVto, `tmptesoreriacomun`.`importe1` Gasto, `tmptesoreriacomun`.`importe2` Recibo"
    SQL = SQL & " FROM   `tmptesoreriacomun` `tmptesoreriacomun`"
    SQL = SQL & " WHERE `tmptesoreriacomun`.codusu = " & vUsu.Codigo
    SQL = SQL & " ORDER BY `tmptesoreriacomun`.`texto1`, `tmptesoreriacomun`.`texto2`, `tmptesoreriacomun`.`opcion`, `tmptesoreriacomun`.`fecha1`"
    
    'LLamos a la funcion
    GeneraFicheroCSV SQL, txtTipoSalida(1).Text
    
End Sub


Private Sub AccionesCrystal()
Dim indRPT As String
Dim nomDocu As String
    
    vMostrarTree = False
    conSubRPT = False
        
    
    indRPT = "0903-00"
    
        
    
    If Not PonerParamRPT(indRPT, nomDocu) Then Exit Sub
    
    cadNomRPT = nomDocu ' "GastosFijos.rpt"
    
    cadFormula = "{tmptesoreriacomun.codusu} = " & vUsu.Codigo
    
    ImprimeGeneral
    
    If optTipoSal(1).Value Then CopiarFicheroASalida True, txtTipoSalida(1).Text
    If optTipoSal(2).Value Then CopiarFicheroASalida False, txtTipoSalida(2).Text, False
    If optTipoSal(3).Value Then LanzaProgramaAbrirOutlook 15
        
    If SoloImprimir Or ExportarPDF Then Unload Me
    Screen.MousePointer = vbDefault
    
    
    
End Sub


Private Function CargarTemporales() As Boolean
Dim SQL As String
Dim SQL2 As String
Dim I As Integer
Dim B As Boolean
Dim Rs As adodb.Recordset
Dim Rs2 As adodb.Recordset
Dim SqlInsert As String
Dim SqlValues As String

    CargarTemporales = False
    
    Label9.Caption = "Preparando tablas"
    Label9.Refresh
    SQL = "Delete from tmpbalancesumas where codusu =" & vUsu.Codigo
    Conn.Execute SQL
                
                
    F2 = CDate(txtFecha(0).Text)
    F1 = DateAdd("d", CLng(txtDias(0).Text) * (-1), F2)
                
                
    SQL = ""
    Screen.MousePointer = vbHourglass
    
            
    Label9.Caption = "Cálculo saldos de cuentas de banco"
    Label9.Refresh
            
    ' SITUACION DE LOS BANCOS
    SQL = "select * from bancos order by codbanco "
    
    Set Rs = New adodb.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SqlInsert = "insert into tmpbalancesumas (codusu,cta,nomcta,TotalD,TotalH) values "
    SqlValues = ""
    
    While Not Rs.EOF
        Label9.Caption = "Cuenta: " & Rs!codmacta & "-" & Rs!Descripcion
        Label9.Refresh
        
        SqlValues = SqlValues & "(" & vUsu.Codigo & "," & DBSet(Rs!codmacta, "T") & "," & DBSet(Rs!Descripcion, "T") & ","
        
        SQL2 = "select sum(coalesce(timporteD,0)), sum(coalesce(timporteh,0)) from hlinapu where codmacta = " & DBSet(Rs!codmacta, "T")
        SQL2 = SQL2 & " and fechaent <= " & DBSet(F2, "F")
        
        Set Rs2 = New adodb.Recordset
        Rs2.Open SQL2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs2.EOF Then
            SqlValues = SqlValues & DBSet(Rs2.Fields(0).Value, "N") & "," & DBSet(Rs2.Fields(1).Value, "N") & "),"
        Else
            SqlValues = SqlValues & "0,0),"
        End If
        
        Set Rs2 = Nothing
        
        Rs.MoveNext
    Wend
            
    Set Rs = Nothing
            
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        Conn.Execute SqlInsert, SqlValues
    End If
    
    'COBROS PENDIENTES
            
    Label9.Caption = "Cobros "
    Label9.Refresh
            
    SqlValues = ""
            
    SQL2 = "select impvenci + coalesce(gastos,0) - coalesce(impcobro,0) from cobros where impvenci + coalesce(gastos,0) - coalesce(impcobro,0) <> 0 "
    SQL2 = SQL2 & " and fechaent between " & DBSet(F1, "F") & " and " & DBSet(F2, "F")
    SQL2 = SQL2 & " and situjuri = 0"
    
    Set Rs2 = New adodb.Recordset
    Rs2.Open SQL2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs2.EOF Then
        SqlValues = "(" & vUsu.Codigo & ",'COBROS'," & ValorNulo & "," & DBSet(Rs2.Fields(0).Value, "N") & "," & ValorNulo & ")"
    Else
        SqlValues = "(" & vUsu.Codigo & ",'COBROS'," & ValorNulo & "," & DBSet(0, "N") & "," & ValorNulo & ")"
    End If
    Set Rs2 = Nothing
    
    Conn.Execute SqlInsert & SqlValues
    
    ' PAGOS PENDIENTES
    
    Label9.Caption = "Pagos "
    Label9.Refresh
            
    SqlValues = ""
            
    SQL2 = "select impefect  - coalesce(imppagad,0) from pagos where impefect - coalesce(imppagad,0) <> 0 "
    SQL2 = SQL2 & " and fecefect between " & DBSet(F1, "F") & " and " & DBSet(F2, "F")
    
    Set Rs2 = New adodb.Recordset
    Rs2.Open SQL2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs2.EOF Then
        SqlValues = "(" & vUsu.Codigo & ",'PAGOS'," & ValorNulo & "," & DBSet(Rs2.Fields(0).Value, "N") & "," & ValorNulo & ")"
    Else
        SqlValues = "(" & vUsu.Codigo & ",'PAGOS'," & ValorNulo & "," & DBSet(0, "N") & "," & ValorNulo & ")"
    End If
    Set Rs2 = Nothing
    
    Conn.Execute SqlInsert & SqlValues
    
    If vParam.HayAriges Then
        ' CALCULO DE LOS ALBARANES CON IVA Y VTOS SEGUN LA FORMA DE PAGO
       
    
    
    End If
            
            
    CargarTemporales = True
    
End Function



Private Function DameEmpresa(ByVal S As String) As String
    DameEmpresa = "NO ENCONTRADA"
    For I = 1 To ListView1(1).ListItems.Count
        If ListView1(1).ListItems(I).Tag = S Then
            DameEmpresa = DevNombreSQL(ListView1(1).ListItems(I).Text)
            Exit For
        End If
    Next I
  
End Function


Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    If txtFecha(0).Text = "" Then
        MsgBox "Debe dar un valor a la fecha.", vbExclamation
        PonFoco txtFecha(0)
        Exit Function
    End If
    
    If txtDias(0).Text = "" Then
        MsgBox "Debe dar un valor a los días.", vbExclamation
        PonFoco txtDias(0)
        Exit Function
    End If
    
    DatosOK = True


End Function


Private Sub txtFecha_LostFocus(Index As Integer)
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    PonerFormatoFecha txtFecha(Index)
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        KeyCode = 0
        
        LanzaFormAyuda txtFecha(Index).Tag, Index
    Else
        KEYdown KeyCode
    End If
End Sub


Private Sub txtTipoSalida_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtDias_GotFocus(Index As Integer)
    ConseguirFoco txtDias(Index), 3
End Sub

Private Sub txtDias_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtDias_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    txtDias(Index).Text = Trim(txtDias(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'dias
            txtDias(Index).Text = Format(txtDias(Index).Text, "####0")
            
    End Select

End Sub

