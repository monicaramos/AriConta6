VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPresu 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Presupuestos"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPresu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3600
      TabIndex        =   20
      Top             =   60
      Width           =   1155
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   1410
         TabIndex        =   22
         Top             =   300
         Visible         =   0   'False
         Width           =   795
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   120
         TabIndex        =   21
         Top             =   180
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generaci�n Masiva"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Agrupado"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   17
      Top             =   60
      Width           =   3405
      Begin VB.CheckBox Check2 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3750
         TabIndex        =   18
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   150
         TabIndex        =   19
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
               Object.Tag             =   "0"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Salir"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameFiltro 
      Height          =   705
      Left            =   5550
      TabIndex        =   15
      Top             =   60
      Width           =   2415
      Begin VB.ComboBox cboFiltro 
         Height          =   360
         ItemData        =   "frmPresu.frx":000C
         Left            =   90
         List            =   "frmPresu.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   210
         Width           =   2235
      End
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   1
      Left            =   1050
      TabIndex        =   14
      Text            =   "Dato2"
      Top             =   5790
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   -90
      TabIndex        =   0
      Tag             =   "Cuenta|T|N|||presupuestos|codmacta||S|"
      Text            =   "Dat"
      Top             =   5790
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   2670
      TabIndex        =   1
      Tag             =   "A�o|N|N|1900||presupuestos|anopresu|0000|S|"
      Text            =   "Dato2"
      Top             =   5790
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   4230
      TabIndex        =   2
      Tag             =   "Denominaci�n|N|N|1|12|presupuestos|mespresu|00|S|"
      Text            =   "Dato2"
      Top             =   5790
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   5670
      TabIndex        =   3
      Tag             =   "Importe|N|N|||presupuestos|imppresu|###,###,##0.00|N|"
      Text            =   "Dato2"
      Top             =   5790
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   90
      TabIndex        =   12
      Top             =   6930
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   90
         TabIndex        =   13
         Top             =   120
         Width           =   2550
      End
   End
   Begin VB.CommandButton cmdAux 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   0
      Left            =   810
      TabIndex        =   10
      Top             =   5760
      Width           =   195
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FEF7E4&
      Height          =   350
      Left            =   7290
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   6450
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6900
      TabIndex        =   4
      Top             =   6990
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8100
      TabIndex        =   6
      Top             =   6990
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8100
      TabIndex        =   5
      Top             =   6990
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmPresu.frx":0050
      Height          =   5295
      Left            =   90
      TabIndex        =   7
      Top             =   960
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   5970
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   390
      Left            =   8730
      TabIndex        =   11
      Top             =   150
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total �"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   9
      Top             =   6480
      Width           =   1065
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmPresu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 1101


'Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)
Private WithEvents frmCtas As frmColCtas
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private frmGen As frmPresuGenerar
Attribute frmGen.VB_VarHelpID = -1

Private CadenaConsulta As String
Private TextoBusqueda As String
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte
Dim jj As Integer
Dim SQL As String
Dim PrimeraVez As Boolean
Dim cadFiltro As String

Dim Agrupado As Boolean

Dim EjerciciosPartidos As Boolean

Private CadB As String

'----------------------------------------------
'----------------------------------------------
'   Deshabilitamos todos los botones menos
'   el de salir
'   Ademas mostramos aceptar y cancelar
'   Modo 0->  Normal
'   Modo 1 -> Lineas BUSCAR
'   Modo 2 -> Recorrer registros
'   Modo 3 -> Lineas  INSERTAR
'   Modo 4 -> Lineas MODIFICAR
'----------------------------------------------
'----------------------------------------------

Private Sub PonerModo(vModo)
Dim B As Boolean
    Modo = vModo
    
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    B = (Modo = 0 Or Modo = 2)
    
    For jj = 0 To 4
        txtaux(jj).Visible = Not B
    Next jj
    
    cmdAux(0).Visible = Not B
    
    For i = 0 To txtaux.Count - 1
        If i <> 1 Then txtaux(i).BackColor = vbWhite
    Next i
    
    Toolbar1.Buttons(1).Enabled = B
    Toolbar1.Buttons(2).Enabled = B
    cmdAceptar.Visible = Not B
    cmdCancelar.Visible = Not B
    'DataGrid1.Enabled = b
    
    txtaux(0).Enabled = (Modo <> 2)
    txtaux(2).Enabled = txtaux(0).Enabled
    txtaux(2).BackColor = txtaux(0).BackColor
    cmdAux(0).Enabled = txtaux(0).Enabled
    
    PonerModoUsuarioGnral Modo, "ariconta"
    
    
End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.Adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub

Private Sub BotonAnyadir()
    Dim anc As Single

    If Not Agrupado Then

         'Situamos el grid al final
         DataGrid1.AllowAddNew = True
         If Not Adodc1.Recordset.EOF Then
             DataGrid1.HoldFields
             Adodc1.Recordset.MoveLast
             DataGrid1.Row = DataGrid1.Row + 1
         End If
        
         If DataGrid1.Row < 0 Then
             anc = DataGrid1.Top + 210
             Else
             anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top
         End If
         txtaux(0).Text = ""
         For jj = 1 To 4
             txtaux(jj).Text = ""
         Next jj
         LLamaLineas anc, 3
         
         'Ponemos el foco
         PonFoco txtaux(0)
         
         
     Else
        Dim Quitar As Boolean
        Quitar = False
        If TieneEjercicio(Adodc1.Recordset.Fields(0), False) And TieneEjercicio(Adodc1.Recordset.Fields(0), True) Then
            If MsgBox("Esta cuenta ya tiene presupuestos para ejercicio actual y el siguiente. �Desea continuar?.", vbQuestion + vbYesNo + vbDefault) = vbNo Then
                Exit Sub
            End If
            Quitar = True
        End If
     
     
        Set frmGen = New frmPresuGenerar
        
        frmGen.opcion = 0
        frmGen.Caption = "Inserci�n de Presupuestos"
        If Not Quitar Then
            frmGen.txtCta(0) = Adodc1.Recordset(0).Value
            frmGen.txtDesCta(0) = Adodc1.Recordset(1).Value
        End If
        frmGen.Modo = 0 'insertar
        frmGen.Show vbModal
        
        Set frmGen = Nothing
     
     End If
     
     CargaGrid CadB
    
End Sub

Private Function TieneEjercicio(Cta As String, Actual As Boolean) As Boolean
Dim SQL As String

    SQL = "select count(*) from presupuestos where codmacta = " & DBSet(Cta, "T") & " and date(concat(anopresu,'-',right(concat('00',mespresu),2),'-1')) "
    If Actual Then
        SQL = SQL & " between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
    Else
        SQL = SQL & " between " & DBSet(DateAdd("yyyy", 1, vParam.fechaini), "F") & " and " & DBSet(DateAdd("yyyy", 1, vParam.fechafin), "F")
    End If

    TieneEjercicio = (TotalRegistros(SQL) <> 0)



End Function

Private Sub BotonVerTodos()
    DataGrid1.Enabled = False
    espera 0.1
    TextoBusqueda = ""
    CadB = ""
    CargaGrid CadB
    DataGrid1.Enabled = True
End Sub

Private Sub BotonBuscar()
    DataGrid1.Enabled = False
    If Agrupado Then
        CadB = " tmppresu1.cta is null"
    Else
        CadB = " presupuestos.codmacta is null"
    End If
    CargaGrid CadB
    DataGrid1.Enabled = True
    'Buscar
    For jj = 0 To 4
        txtaux(jj).Text = ""
    Next jj
    LLamaLineas DataGrid1.Top + 250, 1
    PonFoco txtaux(0)
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim Cad As String
    Dim anc As Single
    Dim i As Integer
    If Adodc1.Recordset.EOF Then Exit Sub
    'If Adodc1.Recordset.RecordCount < 1 Then Exit Sub

    If Agrupado Then
        If Adodc1.Recordset.Fields(3) < Year(vParam.fechaini) Then
            MsgBox "No se permite modificar de ejercicios cerrados.", vbExclamation
            Exit Sub
        End If
    End If



    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If Not Agrupado Then
    
        If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
            i = DataGrid1.Bookmark - DataGrid1.FirstRow
            DataGrid1.Scroll 0, i
            DataGrid1.Refresh
        End If
        
        If DataGrid1.Row < 0 Then
            anc = 320
            Else
            anc = DataGrid1.RowTop(DataGrid1.Row) + 600
        End If
    
         anc = FijarVariableAnc(DataGrid1)
         
         
         Cad = ""
         For i = 0 To 1
             Cad = Cad & DataGrid1.Columns(i).Text & "|"
         Next i
         'Llamamos al form
         For i = 0 To txtaux.Count - 1
             txtaux(i).Text = DataGrid1.Columns(i).Text
         Next i
         LLamaLineas anc, 4
        
        'Como es modificar
        PonFoco txtaux(4)
   
    Else
        Screen.MousePointer = vbDefault
        
        Set frmGen = New frmPresuGenerar
        frmGen.opcion = 0
        frmGen.Modo = 1 'modificar
        frmGen.Caption = "Modificacion de Presupuestos"
        frmGen.txtCta(0) = Adodc1.Recordset(0).Value
        frmGen.txtDesCta(0) = Adodc1.Recordset(1).Value
        frmGen.Ejercicio = Adodc1.Recordset.Fields(3).Value
        frmGen.Show vbModal
    
    End If
    
    CargaGrid CadB
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid
    PonerModo xModo
    'Fijamos el ancho
    For jj = 0 To 4
        txtaux(jj).Top = alto
    Next jj
    cmdAux(0).Top = alto
End Sub




Private Sub BotonEliminar()
Dim SQL As String
Dim vFecIni As String
Dim vFecFin As String

    On Error GoTo Error2
    'Ciertas comprobaciones
    If Adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub
    
    '### a mano
    SQL = "Seguro que desea eliminar la linea de presupuesto:" & vbCrLf
    SQL = SQL & vbCrLf & "Cuenta: " & Adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Denominaci�n: " & Adodc1.Recordset.Fields(1)
    
    If Not Agrupado Then
        SQL = SQL & vbCrLf & "A�o       : " & Adodc1.Recordset.Fields(2)
        SQL = SQL & vbCrLf & "Mes  : " & Adodc1.Recordset.Fields(3)
    Else
        SQL = SQL & vbCrLf & "Per�odo   : " & Adodc1.Recordset.Fields(2)
    End If
    
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from presupuestos where codmacta='" & Trim(Adodc1.Recordset.Fields(0)) & "'"
        If Not Agrupado Then
            SQL = SQL & " and anopresu =" & Adodc1.Recordset!anopresu
            SQL = SQL & " and mespresu = " & Adodc1.Recordset!mespresu & " ;"
        Else
            vFecIni = Adodc1.Recordset.Fields(2) & "-" & Format(Month(vParam.fechaini), "00") & "-01"
            vFecFin = DateAdd("d", -1, DateAdd("yyyy", 1, vFecIni))
        
            SQL = SQL & " and date(concat(anopresu,'-',right(concat('00',mespresu),2),'-01')) between " & DBSet(vFecIni, "F") & " and " & DBSet(vFecFin, "F")
        End If
        Conn.Execute SQL
        CargarTemporal
        CargaGrid CadB
        Adodc1.Recordset.Cancel
    End If
    Exit Sub
Error2:
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub OpcionesCambiadas()
    If txtaux(0).Visible Then Exit Sub

    Screen.MousePointer = vbHourglass
    CargarSqlFiltro
    CargaGrid CadB
    Screen.MousePointer = vbDefault
End Sub

Private Sub cboFiltro_Click()
    If PrimeraVez Then Exit Sub
    OpcionesCambiadas
End Sub


Private Sub cmdAceptar_Click()
Dim i As Integer
Select Case Modo
    Case 1
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        
'        If Agrupado Then
'            If txtAux(2).Text <> "" Then
'                CadB = "ejercicio = " & Mid(txtAux(2).Text, 1, 4)
'            End If
'        End If
        
        
        'Para el texto
        TextoBusqueda = ""
        If txtaux(0).Text <> "" Then TextoBusqueda = TextoBusqueda & "Cod. Inmov " & txtaux(0).Text
        If txtaux(2).Text <> "" Then TextoBusqueda = TextoBusqueda & "Fecha " & txtaux(2).Text
        If txtaux(3).Text <> "" Then TextoBusqueda = TextoBusqueda & "Porcentaje " & txtaux(3).Text
        If txtaux(4).Text <> "" Then TextoBusqueda = TextoBusqueda & "Importe " & txtaux(4).Text
        
        If CadB <> "" Then
            PonerModo 0
            DataGrid1.Enabled = False
            CargaGrid CadB
            DataGrid1.Enabled = True
        End If

    Case 3
        If DatosOK Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                Conn.Execute "commit"
                'MsgBox "Registro insertado.", vbInformation
                CargaGrid CadB
                BotonAnyadir
            End If
        End If
    
    Case 4
        'Modificar
        If DatosOK Then
            '-----------------------------------------
            'Hacemos insertar
            If ModificaDesdeFormulario(Me) Then
                Conn.Execute "commit"
'                i = Adodc1.Recordset.Fields(0)
'                PonerModo 0
'                CargaGrid cadB
'                Adodc1.Recordset.Find (Adodc1.Recordset.Fields(0).Name & " =" & i)
                PosicionarData
                PonerFocoGrid Me.DataGrid1
            End If
        End If
    End Select
End Sub

Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la cap�alera, no llevar els () ***
    Cad = "codmacta = '" & Adodc1.Recordset.Fields(0) & "' and anopresu = " & Adodc1.Recordset.Fields(2) & " and mespresu = " & Adodc1.Recordset.Fields(3)
    ' ***************************************
    CargaGrid CadB
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarDataMULTI(Adodc1, Cad, Indicador, True) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       PonerModo 0
    End If
   
    ' ***********************************************************************************
End Sub



Private Sub cmdAux_Click(Index As Integer)
Dim F As Date

    Select Case Index
        Case 0 ' Cuenta contable
            Screen.MousePointer = vbHourglass
            Set frmCtas = New frmColCtas
            frmCtas.DatosADevolverBusqueda = "0|1|2|"
            frmCtas.ConfigurarBalances = 3  'NUEVO
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            
            PonFoco txtaux(0)
        

    End Select
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1
            CargaGrid CadB
        Case 3
            If Not Agrupado Then DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
    End Select
    
    PonerModo 0
    lblIndicador.Caption = ""
    TextoBusqueda = ""
    DataGrid1.SetFocus

End Sub


Private Sub DataGrid1_DblClick()
'If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

'++
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then Unload Me
End Sub
'++

Private Sub Form_Activate()

    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        CadB = ""
        cboFiltro.ListIndex = vUsu.FiltroPresup
        CargarElGrid
        PrimeraVez = False
    End If
End Sub

Private Sub Form_Load()
    
    Me.Icon = frmPpal.Icon

    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(8).Image = 16
    End With

    ' Botonera Principal 2
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 47
        .Buttons(2).Image = 36
    End With

    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With

    PrimeraVez = True
    
    CargaFiltros
    
    Set miTag = New CTag
    '## A mano
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    PonerModo 0
    CadAncho = False
    PonerOpcionesMenu  'En funcion del usuario
    'Cadena consulta
   
    Agrupado = True
    
    EjerciciosPartidos = (Year(vParam.fechaini) <> Year(vParam.fechafin))
    
End Sub


Private Sub CargarTemporal()
Dim SQL As String
Dim SQL2 As String
Dim SqlInsert As String
Dim CadValues As String
Dim AnyoMin As Integer
Dim AnyoMax As Integer
Dim MesI As Integer
Dim MesF As Integer
Dim vImp As Currency

Dim Rs As ADODB.Recordset

    On Error GoTo eCargarTemporal



    SQL = "delete from tmppresu1 where codusu = " & DBSet(vUsu.Codigo, "N")
    Conn.Execute SQL

    SqlInsert = "insert into tmppresu1 (codusu,codigo,cta,ejercicio,ano,Importe) values "

    SQL = "select min(anopresu) from presupuestos"
    AnyoMin = DevuelveValor(SQL)
    
    
    SQL = "select max(anopresu) from presupuestos"
    AnyoMax = DevuelveValor(SQL)
    
    MesI = Month(vParam.fechaini)
    MesF = Month(vParam.fechafin)
    
    SQL = "select distinct codmacta from presupuestos order by 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    J = 0
    While Not Rs.EOF
        For i = AnyoMin To AnyoMax
            If EjerciciosPartidos Then
                SQL2 = "select count(*) "
                SQL2 = SQL2 & " from presupuestos "
                SQL2 = SQL2 & " Where codmacta = " & DBSet(Rs!codmacta, "T")
                SQL2 = SQL2 & " and ((anopresu = " & DBSet(i, "N") & " and mespresu >=" & DBSet(MesI, "N") & ") or "
                SQL2 = SQL2 & "  (anopresu = " & DBSet(i + 1, "N") & " and mespresu <= " & DBSet(MesF, "N") & "))"
                
                
                SQL = "select sum(coalesce(imppresu,0)) "
                SQL = SQL & " from presupuestos "
                SQL = SQL & " Where codmacta = " & DBSet(Rs!codmacta, "T")
                SQL = SQL & " and ((anopresu = " & DBSet(i, "N") & " and mespresu >=" & DBSet(MesI, "N") & ") or "
                SQL = SQL & " (anopresu = " & DBSet(i + 1, "N") & " and mespresu <= " & DBSet(MesF, "N") & "))"
            Else
                SQL2 = "select count(*) "
                SQL2 = SQL2 & " from presupuestos "
                SQL2 = SQL2 & " Where codmacta = " & DBSet(Rs!codmacta, "T")
                SQL2 = SQL2 & " and anopresu = " & DBSet(i, "N")
                
                SQL = "select sum(coalesce(imppresu,0)) "
                SQL = SQL & " from presupuestos "
                SQL = SQL & " Where codmacta = " & DBSet(Rs!codmacta, "T")
                SQL = SQL & " and anopresu = " & DBSet(i, "N")
            End If
            
            If TotalRegistros(SQL2) <> 0 Then
                J = J + 1
                vImp = DevuelveValor(SQL)
                
                Dim Ejer As String
                If EjerciciosPartidos Then
                    Ejer = Format(i, "0000") & "-" & Mid(Format(i + 1, "0000"), 3, 2)
                Else
                    Ejer = Format(i, "0000")
                End If
                
                CadValues = CadValues & "(" & DBSet(vUsu.Codigo, "N") & "," & DBSet(J, "N") & "," & DBSet(Rs!codmacta, "T") & "," & DBSet(Ejer, "T") & "," & DBSet(i, "N") & "," & DBSet(vImp, "N") & "),"
            End If
        Next i
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        Conn.Execute SqlInsert & CadValues
    End If
    
    Exit Sub
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal Totales", Err.Description
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set miTag = Nothing
    vUsu.ActualizarFiltro "ariconta", IdPrograma, Me.cboFiltro.ListIndex
End Sub


Private Sub frmEI_DatoSeleccionado(CadenaSeleccion As String)
    txtaux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtaux(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    txtaux(0).Text = RecuperaValor(CadenaSeleccion, 1)
    txtaux(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
    txtaux(2).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub





Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
                BotonAnyadir
        Case 2
                BotonModificar
        Case 3
                BotonEliminar
        Case 5
                BotonBuscar
        Case 6
                BotonVerTodos
        Case 8
                frmPresuList.Show vbModal
        Case Else
    End Select


End Sub



Private Sub CargaGrid(Optional vSQL As String)
Dim J As Integer
Dim TotalAncho As Integer
Dim i As Integer
Dim tots As String
    
    Text1.Text = ""
    Adodc1.ConnectionString = Conn
    
    
    If Not Agrupado Then
        vSQL = Replace(vSQL, "tmppresu1.cta", "presupuestos.codmacta")
        vSQL = Replace(vSQL, "tmppresu1.ejercicio", "presupuestos.anopresu")
        vSQL = Replace(vSQL, "tmppresu1.mes", "presupuestos.mespresu")
        vSQL = Replace(vSQL, "tmppresu1.importe", "presupuestos.imppresu")
    Else
        vSQL = Replace(vSQL, "presupuestos.codmacta", "tmppresu1.cta")
        vSQL = Replace(vSQL, "presupuestos.anopresu", "tmppresu1.ejercicio")
        vSQL = Replace(vSQL, "presupuestos.mespresu", "tmppresu1.mes")
        vSQL = Replace(vSQL, "presupuestos.imppresu", "tmppresu1.importe")
    End If
    
    SQL = CadenaConsulta & " and " & cadFiltro
    If vSQL <> "" Then SQL = SQL & " and " & vSQL

    If Agrupado Then
        CargarTemporal
        
        txtaux(0).Tag = "Cuenta|T|N|||tmppresu1|cta||S|"
        txtaux(2).Tag = "A�o|T|N|1900||tmppresu1|ejercicio||S|"
        txtaux(3).Tag = "Mes|N|N|||tmppresu1|mes|00|S|"
        txtaux(4).Tag = "Importe|N|N|||tmppresu1|importe|###,###,##0.00|N|"
       
    
'        Sql = Sql & " GROUP BY 1,2,3    "
        SQL = SQL & " ORDER BY 1,2,3 "
    Else
        
        txtaux(0).Tag = "Cuenta|T|N|||presupuestos|codmacta||S|"
        txtaux(2).Tag = "A�o|N|N|1900||presupuestos|anopresu|0000|S|"
        txtaux(3).Tag = "Mes|N|N|||presupuestos|mespresu|00|S|"
        txtaux(4).Tag = "Importe|N|N|||presupuestos|imppresu|###,###,##0.00|N|"
        
        SQL = SQL & " ORDER BY codmacta,anopresu,mespresu"
    End If

    
    
    CargaGridGnral Me.DataGrid1, Me.Adodc1, SQL, True 'PrimeraVez
    
    If Agrupado Then
        ' *******************canviar els noms i si fa falta la cantitat********************
        tots = "S|txtAux(0)|T|Cuenta|1555|;S|cmdAux(0)|B|||;S|txtAux(1)|T|T�tulo|4450|;S|txtAux(2)|T|Ejercicio|1000|;N||||0|;N||||0|;"
        tots = tots & "S|txtAux(4)|T|Importe|1450|;"
    Else
        ' *******************canviar els noms i si fa falta la cantitat********************
        tots = "S|txtAux(0)|T|Cuenta|1555|;S|cmdAux(0)|B|||;S|txtAux(1)|T|T�tulo|3800|;S|txtAux(2)|T|A�o|900|;"
        tots = tots & "S|txtAux(3)|T|Mes|800|;S|txtAux(4)|T|Importe|1450|;"
    End If
    
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft

    'Habilitamos modificar y eliminar
    CargarSumas SQL
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim C As String
    WheelUnHook
    Select Case Button.Index
        Case 1 ' generar de manera masiva
            Set frmGen = New frmPresuGenerar
            frmGen.opcion = 1
            frmGen.Show vbModal
            Set frmGen = Nothing
        
            CargaGrid CadB
        

        Case 2 ' AGRUPADO O NO
            ' ahora es al reves si no estaba agrupado hemos de agruparlo
            '                   y sino desagruparlo
            Agrupado = Not Agrupado
            
            CargarElGrid

    End Select
End Sub

Private Sub CargarElGrid()
    If Agrupado Then
        
        CargarTemporal

        CadenaConsulta = "SELECT tmppresu1.cta , nommacta, tmppresu1.ejercicio, ano, tmppresu1.mes, tmppresu1.importe imppresu "
        CadenaConsulta = CadenaConsulta & " FROM  tmppresu1,cuentas  WHERE tmppresu1.codusu= " & DBSet(vUsu.Codigo, "N") & " and "
        CadenaConsulta = CadenaConsulta & " tmppresu1.cta=cuentas.codmacta"
        
    Else
        CadenaConsulta = "SELECT presupuestos.codmacta, nommacta, anopresu,mespresu,imppresu "
        CadenaConsulta = CadenaConsulta & " FROM  presupuestos,cuentas  WHERE "
        CadenaConsulta = CadenaConsulta & " presupuestos.codmacta=cuentas.codmacta"
    End If
    
    If Agrupado Then
        Me.Toolbar2.Buttons(2).Image = 36
    Else
        Me.Toolbar2.Buttons(2).Image = 43
    End If
    
    CargarSqlFiltro
    CargaGrid CadB
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtaux(Index), Modo
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'elemento
            Case 2: KEYBusqueda KeyAscii, 1 'fecha
            
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    cmdAux_Click (indice)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim RC As String
Dim Valor As Currency

    If Not PerderFocoGnral(txtaux(Index), Modo) Then Exit Sub

    Select Case Index
    Case 0
        RC = txtaux(0).Text
        If RC = "" And Modo = 1 Then Exit Sub
        If CuentaCorrectaUltimoNivel(RC, SQL) Then
            txtaux(0).Text = RC
            txtaux(1).Text = SQL
        Else
            MsgBox SQL, vbExclamation
            txtaux(0).Text = ""
            txtaux(1).Text = ""
            PonFoco txtaux(0)
        End If
    
    Case 2, 3
        If Modo = 1 Then Exit Sub
            If Index = 2 Then
                RC = "A�o"
            Else
                If Index = 3 Then
                    RC = "mes"
                Else
                    RC = "importe"
                End If
            End If
            
            If Not IsNumeric(txtaux(Index).Text) Then
                MsgBox "El valor del " & RC & " debe de ser num�rico.", vbExclamation
                PonFoco txtaux(Index)
                Exit Sub
            End If
            
            'Particularidades
            If Index = 2 Then
                If Val(txtaux(2).Text) < 1000 Then
                    MsgBox "A�o incorrecto", vbExclamation
                    PonFoco txtaux(2)
                    Exit Sub
                End If
            Else
                If (Val(txtaux(3).Text) < 1) Or (Val(txtaux(3).Text) > 12) Then
                    MsgBox "Mes incorrecto", vbExclamation
                    PonFoco txtaux(3)
                    Exit Sub
                End If
            End If
    Case 4 ' importe
        PonerFormatoDecimal txtaux(Index), 1
    End Select
End Sub


Private Function DatosOK() As Boolean
Dim B As Boolean
B = CompForm(Me)
If Not B Then Exit Function

If Modo = 1 Then
    'Estamos insertando
    
End If
DatosOK = B
End Function

Private Sub DeseleccionaGrid()
    On Error GoTo EDeseleccionaGrid
        
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Function SepuedeBorrar() As Boolean
    SepuedeBorrar = True
End Function



Private Sub CargarSumas(ByRef vS As String)
On Error GoTo ECargarSumas

    Set miRsAux = New ADODB.Recordset
    SQL = "Select sum(imppresu) from (" & vS & ") aaaa"
'    If vS <> "" Then SQL = SQL & " WHERE  " & vS
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then _
            Text1.Text = Format(miRsAux.Fields(0), FormatoImporte)
    End If
    miRsAux.Close
ECargarSumas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar sumas"
    Set miRsAux = Nothing
End Sub


Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim Rs As ADODB.Recordset
Dim Cad As String
    
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    Cad = Cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") And (Modo = 0 Or Modo = 2)
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub


Private Sub CargaFiltros()
Dim AUx As String

    cboFiltro.Clear
    
    cboFiltro.AddItem "Sin Filtro "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 0
    cboFiltro.AddItem "Ejercicios Abiertos "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 1
    cboFiltro.AddItem "Ejercicio Actual "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 2
    cboFiltro.AddItem "Ejercicio Siguiente "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 3

End Sub


Private Sub CargarSqlFiltro()

    Screen.MousePointer = vbHourglass
    
    cadFiltro = ""
    
    If Not Agrupado Then
        Select Case Me.cboFiltro.ListIndex
            Case 0 ' sin filtro
                cadFiltro = "(1=1)"
            
            Case 1 ' ejercicios abiertos
                cadFiltro = "date(concat(anopresu,'-',mespresu,'-01')) >= " & DBSet(vParam.fechaini, "F")
            
            Case 2 ' ejercicio actual
                cadFiltro = "date(concat(anopresu,'-',mespresu,'-01')) between " & DBSet(vParam.fechaini, "F") & " and " & DBSet(vParam.fechafin, "F")
            
            Case 3 ' ejercicio siguiente
                cadFiltro = "date(concat(anopresu,'-',mespresu,'-01')) > " & DBSet(vParam.fechafin, "F")
        
        End Select
    Else
        Select Case Me.cboFiltro.ListIndex
            Case 0 ' sin filtro
                cadFiltro = "(1=1)"
            
            Case 1 ' ejercicios abiertos
                cadFiltro = "ano in (" & Year(vParam.fechaini) & "," & Year(vParam.fechaini) + 1 & ")"
            
            Case 2 ' ejercicio actual
                cadFiltro = "ano in (" & Year(vParam.fechaini) & ")"
            
            Case 3 ' ejercicio siguiente
                cadFiltro = "ano in (" & Year(vParam.fechaini) + 1 & ")"
        
        End Select
    
    End If
    
'    If Agrupado Then
'        cadFiltro = Replace(Replace(cadFiltro, "anopresu", "ano"), "mespresu", "mes")
'    Else
'        cadFiltro = Replace(Replace(cadFiltro, "ano", "anopresu"), "mes", "mespresu")
'    End If
    Screen.MousePointer = vbDefault


End Sub


