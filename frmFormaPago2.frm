VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFormaPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formas de pago"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   14280
   Icon            =   "frmFormaPago2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   14280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   5
      Left            =   5730
      MaxLength       =   40
      TabIndex        =   6
      Tag             =   "IBAN Transf.Clientes|T|S|||formapago|iban|||"
      Text            =   "Dat"
      Top             =   5640
      Width           =   4575
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   4
      Left            =   4830
      MaxLength       =   5
      TabIndex        =   5
      Tag             =   "Resto Vtos|N|N|0||formapago|restoven|####0||"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   3
      Left            =   3930
      MaxLength       =   5
      TabIndex        =   4
      Tag             =   "Primer Vto|N|N|0||formapago|primerve|####0||"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   2
      Left            =   3060
      MaxLength       =   5
      TabIndex        =   3
      Tag             =   "Nro Vto.|N|N|1||formapago|numerove|####0||"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3930
      TabIndex        =   16
      Top             =   60
      Width           =   2415
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   180
         TabIndex        =   17
         Top             =   150
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "�ltimo"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   13
      Top             =   60
      Width           =   3585
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3750
         TabIndex        =   14
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   15
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
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
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
      ItemData        =   "frmFormaPago2.frx":000C
      Left            =   2370
      List            =   "frmFormaPago2.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Tipo de pago|N|N|||formapago|tipforpa|||"
      Top             =   5670
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   11880
      TabIndex        =   7
      Top             =   6600
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      Left            =   13080
      TabIndex        =   8
      Top             =   6600
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   1
      Left            =   900
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Denominaci�n|T|N|||formapago|nomforpa|||"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   0
      Left            =   60
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "C�digo |N|N|0||formapago|codforpa|000|S|"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFormaPago2.frx":0010
      Height          =   5295
      Left            =   120
      TabIndex        =   12
      Top             =   1020
      Width           =   14030
      _ExtentX        =   24739
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
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
      Left            =   13080
      TabIndex        =   11
      Top             =   6600
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   9
      Top             =   6555
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
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
         Left            =   150
         TabIndex        =   10
         Top             =   180
         Width           =   2550
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   6030
      Top             =   30
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
      Left            =   13620
      TabIndex        =   18
      Top             =   180
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
Attribute VB_Name = "frmFormaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 206


Private CadenaConsulta As String
Private CadB As String

Dim PasamosPorCTA As Boolean

Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte

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
    
    For i = 0 To txtaux.Count - 1
        txtaux(i).Visible = Not B
        txtaux(i).BackColor = vbWhite
    Next i
    
    Combo1.Visible = Not B
    Combo1.BackColor = vbWhite
    'Prueba
    
    
    cmdAceptar.Visible = Not B
    cmdCancelar.Visible = Not B
    DataGrid1.Enabled = B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.Visible = B
    End If
    txtaux(0).Enabled = (Modo <> 4)
    
    txtaux(5).Enabled = ((Modo = 1) Or (ValorCombo(Combo1) = 1))

    PonerModoUsuarioGnral Modo, "ariconta"


End Sub


Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub



Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    Dim i As Integer
    
    PasamosPorCTA = False
    
    'Obtenemos la siguiente numero de factura
    NumF = SugerirCodigoSiguiente
    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    If DataGrid1.Row < 0 Then
        anc = DataGrid1.Top + 270
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top
    End If
    
    txtaux(0).Text = NumF
    For i = 0 To txtaux.Count - 1
        txtaux(i).Text = ""
    Next i
    
    Combo1.ListIndex = -1
    LLamaLineas anc, 3
    
    ' Por defecto el numero de vencimientos es 1
    txtaux(2).Text = 1
    txtaux(3).Text = 0
    txtaux(4).Text = 0
    
    'Ponemos el foco
    PonFoco txtaux(0)
End Sub



Private Sub BotonVerTodos()
    CargaGrid ""
End Sub

Private Sub BotonBuscar()
    CargaGrid "codforpa = -1"
    'Buscar
    For i = 0 To txtaux.Count - 1
        txtaux(i).Text = ""
    Next i
    Combo1.ListIndex = -1
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
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub

    PasamosPorCTA = False
    
    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    DeseleccionaGrid
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top 'DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    'Llamamos al form
    txtaux(0).Text = DataGrid1.Columns(0).Text
    txtaux(1).Text = DataGrid1.Columns(1).Text
    NumRegElim = adodc1.Recordset!TipForpa
    
    PosicionarCombo Combo1, CInt(NumRegElim)
    
    txtaux(2).Text = DataGrid1.Columns(4).Text
    txtaux(3).Text = DataGrid1.Columns(5).Text
    txtaux(4).Text = DataGrid1.Columns(6).Text
    txtaux(5).Text = DataGrid1.Columns(7).Text
    LLamaLineas anc, 4
   
    'Como es modificar
    PonFoco txtaux(1)
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    PonerModo xModo
    'Fijamos el ancho
    For i = 0 To txtaux.Count - 1
        txtaux(i).Top = alto
    Next i
    Combo1.Top = alto - 15
End Sub


Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub

    
    If Not SepuedeBorrar Then Exit Sub
    '### a mano
    SQL = "Seguro que desea eliminar la forma de pago:"
    SQL = SQL & vbCrLf & "C�digo: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Denominaci�n: " & adodc1.Recordset.Fields(1)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from formapago where codforpa=" & adodc1.Recordset!codforpa
        Conn.Execute SQL
        CargaGrid ""
        adodc1.Recordset.Cancel
    End If
    Exit Sub
Error2:
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub


Private Sub adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If adReason = adRsnMove And adStatus = adStatusOK Then PonLblIndicador Me.lblIndicador, adodc1
End Sub


Private Sub cmdAceptar_Click()
Dim i As Integer
Dim CadB As String
    Select Case Modo
    Case 1
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        If CadB <> "" Then
            PonerModo 0
            CargaGrid CadB
        End If
    Case 3
        If DatosOK Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                espera 0.5
                'MsgBox "Registro insertado.", vbInformation
                CargaGrid
                BotonAnyadir
            End If
        End If
    Case 4
            'Modificar
            If DatosOK Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    i = adodc1.Recordset.Fields(0)
                    PonerModo 0
                    CargaGrid
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & i)
                End If
            End If
    End Select


End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1
            CargaGrid
        Case 3
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
    End Select
    PonerModo 0
    lblIndicador.Caption = ""
    DataGrid1.SetFocus
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String

If adodc1.Recordset.EOF Then
    MsgBox "Ning�n registro a devolver.", vbExclamation
    Exit Sub
End If


Cad = adodc1.Recordset.Fields(0) & "|"
Cad = Cad & adodc1.Recordset.Fields(1) & "|"
RaiseEvent DatoSeleccionado(Cad)
Unload Me
End Sub

Private Sub cmdRegresar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus()
Dim SQL As String
    If Modo = 3 Then
        txtaux(5).Enabled = (ValorCombo(Combo1) = 1)
        If txtaux(5).Enabled Then
'            SQL = "select concat(iban,'-',right(concat('0000',entidad),4),'-', right(concat('0000',oficina),4),'-', right(concat('00',control),2), right(concat('00',mid(ctabanco,1,2)),2),'-', right(concat('0000',mid(ctabanco,3,4)),10) )  from bancos where ctatransfercli = 1 "
            SQL = "select iban from bancos where ctatransfercli = 1"
            txtaux(5).Text = DevuelveValor(SQL)
            If txtaux(5).Text = "0" Then txtaux(5).Text = ""
        Else
            txtaux(5).Text = ""
        End If
    End If
End Sub

Private Sub DataGrid1_DblClick()
If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()

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

    ' desplazamiento
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CargaCombo
    
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False
    PonerOpcionesMenu  'En funcion del usuario
    'Cadena consulta
    CadenaConsulta = "Select formapago.codforpa, formapago.nomforpa, formapago.tipforpa, tipofpago.descformapago, formapago.numerove, formapago.primerve, formapago.restoven, "
    CadenaConsulta = CadenaConsulta & " formapago.iban "
    'CadenaConsulta = CadenaConsulta & "if(if(formapago.ibantransfcli is null or formapago.ibantransfcli = '','',formapago.ibantransfcli) = '','',concat(mid(formapago.ibantransfcli,1,4),'-',mid(formapago.ibantransfcli,5,4),'-',mid(formapago.ibantransfcli,9,4),'-',mid(formapago.ibantransfcli,13,4),'-',mid(formapago.ibantransfcli,17,4),'-',mid(formapago.ibantransfcli,21,4)))  "
    CadenaConsulta = CadenaConsulta & " FROM formapago ,tipofpago"
    CadenaConsulta = CadenaConsulta & " WHERE formapago.tipforpa = tipofpago.tipoformapago"
    CargaGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
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



'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Private Function SugerirCodigoSiguiente() As String
    Dim SQL As String
    Dim Rs As ADODB.Recordset
    
    SQL = "Select Max(codforpa) from formapago"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, , , adCmdText
    SQL = "1"
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then
            SQL = CStr(Rs.Fields(0) + 1)
        End If
    End If
    Rs.Close
    SugerirCodigoSiguiente = SQL
End Function

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
                frmFormaPagoList.Show vbModal
        Case Else
        
    End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    FrameDesplazamiento.Visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub CargaGrid(Optional SQL As String)
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim i As Integer
    
    adodc1.ConnectionString = Conn
    If SQL <> "" Then
        SQL = CadenaConsulta & " AND " & SQL
        Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY codforpa "
    adodc1.RecordSource = SQL
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 350
    
    
    'Nombre producto
    i = 0
        DataGrid1.Columns(i).Caption = "C�digo"
        DataGrid1.Columns(i).Width = 800
        DataGrid1.Columns(i).NumberFormat = "000"
        
    
    'Leemos del vector en 2
    i = 1
        DataGrid1.Columns(i).Caption = "Denominaci�n"
        DataGrid1.Columns(i).Width = 3680
        TotalAncho = TotalAncho + DataGrid1.Columns(i).Width
    
    'El importe es campo calculado
    i = 2
        DataGrid1.Columns(i).Visible = False
        
    i = 3
        DataGrid1.Columns(i).Caption = "Tipo pago"
        DataGrid1.Columns(i).Width = 2500
            
    i = 4
        DataGrid1.Columns(i).Caption = "No.Vtos"
        DataGrid1.Columns(i).Width = 900
    i = 5
        DataGrid1.Columns(i).Caption = "1er.Vto"
        DataGrid1.Columns(i).Width = 900
    i = 6
        DataGrid1.Columns(i).Caption = "Resto Vtos"
        DataGrid1.Columns(i).Width = 1200
    i = 7
        DataGrid1.Columns(i).Caption = "IBAN Transferencia Clientes"
        DataGrid1.Columns(i).Width = 3450
    
        
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtaux(0).Width = DataGrid1.Columns(0).Width - 60
        txtaux(1).Width = DataGrid1.Columns(1).Width - 60
        txtaux(2).Width = DataGrid1.Columns(4).Width - 60
        txtaux(3).Width = DataGrid1.Columns(5).Width - 60
        txtaux(4).Width = DataGrid1.Columns(6).Width - 60
        txtaux(5).Width = DataGrid1.Columns(7).Width - 60
        
        Combo1.Width = DataGrid1.Columns(3).Width
        txtaux(0).Left = DataGrid1.Left + 340
        txtaux(1).Left = txtaux(0).Left + txtaux(0).Width + 45
        Combo1.Left = txtaux(1).Left + txtaux(1).Width + 45
        txtaux(2).Left = Combo1.Left + Combo1.Width + 45
        txtaux(3).Left = txtaux(2).Left + txtaux(2).Width + 45
        txtaux(4).Left = txtaux(3).Left + txtaux(3).Width + 45
        txtaux(5).Left = txtaux(4).Left + txtaux(4).Width + 75
        
        CadAncho = True
    End If
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
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim B As Boolean
Dim CADENA
    If Not PerderFocoGnral(txtaux(Index), Modo) Then Exit Sub

    txtaux(Index).Text = Trim(txtaux(Index).Text)
    If txtaux(Index).Text = "" Then Exit Sub
    If Modo = 1 Then Exit Sub 'Busquedas
    
    Select Case Index
        Case 0
            If Not IsNumeric(txtaux(0).Text) Then
                MsgBox "C�digo concepto tiene que ser num�rico", vbExclamation
                Exit Sub
            End If
            txtaux(0).Text = Format(txtaux(0).Text, "000")
        
        Case 5
            PasamosPorCTA = True
            If txtaux(Index).Text = "" Then
'                txtAux1.Text = ""
                Exit Sub
            End If
'            txtAux1.Text = Replace(txtAux(Index).Text, "-", "")
'            Cadena = txtAux1.Text
'            Cadena = Mid(Cadena, 1, 4) & "-" & Mid(Cadena, 5, 4) & "-" & Mid(Cadena, 9, 4) & "-" & Mid(Cadena, 13, 2) & "-" & Mid(Cadena, 15, 10)
'            txtAux(Index).Text = Cadena
            B = ComprobarIBANCuenta(txtaux(Index).Text)
            If Not B Then
                PonFoco txtaux(Index)
            Else
'                txtAux1.Text = Replace(txtAux(5).Text, "-", "")
                cmdAceptar.SetFocus
            End If
    
    End Select
End Sub
            
Private Function ComprobarIBANCuenta(CuentaCCC As String) As Boolean
Dim cadMen As String
Dim Cta As String
Dim BuscaChekc As String
Dim B As Boolean

    If Len(CuentaCCC) <> 24 Then
        cadMen = "La cuenta bancaria no tiene longitud correcta. � Desea continuar ?."
        If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            B = True
        Else
            B = False
        End If
    Else
        Cta = Mid(CuentaCCC, 5, 20)
        If Not Comprueba_CC(Cta) Then
            cadMen = "La cuenta bancaria no es correcta. � Desea continuar ?."
            If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                B = True
            Else
                B = False
            End If
        Else
            BuscaChekc = ""
            BuscaChekc = Mid(CuentaCCC, 1, 2)
                
            If DevuelveIBAN2(BuscaChekc, Cta, Cta) Then
                If Mid(CuentaCCC, 1, 4) <> BuscaChekc & Cta Then
                    cadMen = "El c�digo de IBAN no es correcto, deber�a ser " & BuscaChekc & Cta & vbCrLf & vbCrLf & "� Continuar ?"
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        B = True
                    Else
                        B = False
                    End If
                Else
                    B = True
                End If
            End If
        End If
    End If
    ComprobarIBANCuenta = B
End Function

Private Function DatosOK() As Boolean
Dim Datos As String
Dim B As Boolean
    B = CompForm(Me)
    If Not B Then Exit Function
    
    If Modo = 3 Then
        'Estamos insertando
         Datos = DevuelveDesdeBD("nomforpa", "formapago", "codforpa", txtaux(0).Text, "N")
         If Datos <> "" Then
            MsgBox "Ya existe la forma de pago : " & txtaux(0).Text & "-" & Datos, vbExclamation
            B = False
        End If
    End If
    
    If B And (Modo = 3 Or Modo = 4) Then
'        txtAux1.Text = Replace(txtAux(5).Text, "-", "")
        If txtaux(5).Text <> "" And Not PasamosPorCTA Then
            B = ComprobarIBANCuenta(txtaux(5).Text)
            If Not B Then PonFoco txtaux(5)
        End If
    End If
    
    DatosOK = B
End Function

Private Sub CargaCombo()
    Combo1.Clear
    'Conceptos
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select * from tipofpago order by descformapago", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF
        Combo1.AddItem miRsAux!descformapago
        Combo1.ItemData(Combo1.NewIndex) = miRsAux!tipoformapago
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Function SepuedeBorrar() As Boolean
Dim SQL As String
    SepuedeBorrar = False

    
    SepuedeBorrar = True
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


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub




' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del rat�n.
Private Sub DataGrid1_GotFocus()
  WheelHook DataGrid1
End Sub
Private Sub DataGrid1_LostFocus()
  WheelUnHook
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
        Toolbar1.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2) And Not vParam.HayAriges
        Toolbar1.Buttons(2).Enabled = DBLet(Rs!Modificar, "N") And (Modo = 0 Or Modo = 2) And Not vParam.HayAriges
        Toolbar1.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 0 Or Modo = 2) And Not vParam.HayAriges
        
        Toolbar1.Buttons(5).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        Toolbar1.Buttons(6).Enabled = DBLet(Rs!Ver, "N") And (Modo = 0 Or Modo = 2)
        
        Toolbar1.Buttons(8).Enabled = DBLet(Rs!Imprimir, "N") And (Modo = 0 Or Modo = 2)
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub

