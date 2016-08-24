VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTESGastosFijos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gastos Fijos"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10725
   Icon            =   "frmGastosFijo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   3960
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Contrapartida|T|S|||sgastfij|contrapar|||"
      Text            =   "Dato2"
      Top             =   5520
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   2400
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "Cta prevista|T|N|||sgastfij|ctaprevista|||"
      Text            =   "Dato2"
      Top             =   5520
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3600
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   6360
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   6360
      Width           =   3375
   End
   Begin VB.CommandButton cmdCab 
      Cancel          =   -1  'True
      Caption         =   "Cabeceras"
      Height          =   375
      Left            =   9600
      TabIndex        =   22
      Top             =   6840
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   3120
      TabIndex        =   16
      Top             =   2160
      Width           =   5775
      Begin VB.CommandButton cmdPregunta 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4440
         TabIndex        =   21
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdPregunta 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   20
         Top             =   1800
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Generar serie gastos periodicos"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Copiar gastos periodo anterior"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   4695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Un unico gasto puntual"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   2535
         Left            =   120
         Top             =   120
         Width           =   5535
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Siguiente"
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   15
      Top             =   7080
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Actual"
      Height          =   255
      Index           =   0
      Left            =   6840
      TabIndex        =   14
      Top             =   6840
      Value           =   -1  'True
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGastosFijo.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGastosFijo.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGastosFijo.frx":6B88
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5895
      Left            =   6840
      TabIndex        =   12
      Top             =   840
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   10398
      _Version        =   393217
      Indentation     =   1411
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   6900
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9540
      TabIndex        =   5
      Top             =   6900
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   900
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Nombre|T|N|||sgastfij|descripcion|||"
      Text            =   "Dato2"
      Top             =   5520
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   60
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "Código|N|N|0||sgastfij|codigo||S|"
      Text            =   "Dat"
      Top             =   5520
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmGastosFijo.frx":D3EA
      Height          =   5445
      Left            =   60
      TabIndex        =   9
      Top             =   600
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   9604
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Height          =   375
      Left            =   9540
      TabIndex        =   8
      Top             =   6900
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   6
      Top             =   6735
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar importes"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar gastos"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8760
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Image imgCta 
      Height          =   240
      Left            =   4680
      Picture         =   "frmGastosFijo.frx":D3FF
      Top             =   6120
      Width           =   240
   End
   Begin VB.Image ImBanco 
      Height          =   240
      Left            =   1440
      Picture         =   "frmGastosFijo.frx":13C51
      Top             =   6120
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Contrapartida"
      Height          =   195
      Index           =   1
      Left            =   3600
      TabIndex        =   26
      Top             =   6120
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "Cuenta prevista"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   6120
      Width           =   1110
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6840
      TabIndex        =   13
      Top             =   600
      Width           =   3375
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
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
   Begin VB.Menu mnPop 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu mnCOntabilizar 
         Caption         =   "Contabilizar"
      End
   End
End
Attribute VB_Name = "frmTESGastosFijos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmC As frmColCtas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBanco
Attribute frmB.VB_VarHelpID = -1

Private CadenaConsulta As String
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte
Dim PrimeraVez As Boolean
'----------------------------------------------
'----------------------------------------------
'
'   Deshabilitamos todos los botones menos
'   el de salir
'   Ademas mostramos aceptar y cancelar
'   Modo 0->  Normal
'   Modo 1 -> Lineas INSERTAR
'   Modo 2 -> Lineas MODIFICAR
'   Modo 3 -> Lineas BUSCAR
'   Modo 4 -> HABILITAR Modificar LISTVIEW
'   Modo 5 -> Modi Lineas
'
'----------------------------------------------
'----------------------------------------------


Private Sub PonerModo(vModo)
Dim B As Boolean
Modo = vModo



If Modo > 3 Then
    'MODIFICAR LINEAS
    B = Modo = 4
    cmdAceptar.Visible = Not B
    cmdCancelar.Visible = Not B
    cmdCab.Visible = B
'    Option1(0).Visible = False
'    Option1(1).Visible = False
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(8).Enabled = B
    Toolbar1.Buttons(7).Enabled = B
    Toolbar1.Buttons(6).Enabled = B
    Toolbar1.Buttons(10).Enabled = False
    Toolbar1.Buttons(13).Enabled = False
    
    Me.DataGrid1.Enabled = False
    Exit Sub
End If
Me.DataGrid1.Enabled = True
Me.cmdCab.Visible = False

B = (Modo = 0)




txtAux(0).Visible = Not B
txtAux(1).Visible = Not B
txtAux(2).Visible = Not B
txtAux(3).Visible = Not B
Me.ImBanco.Visible = Not B
Me.imgCta.Visible = Not B

mnOpciones.Enabled = B
Toolbar1.Buttons(1).Enabled = B
Toolbar1.Buttons(2).Enabled = B
Toolbar1.Buttons(8).Enabled = B
Toolbar1.Buttons(7).Enabled = B
Toolbar1.Buttons(6).Enabled = B
Toolbar1.Buttons(10).Enabled = B
Toolbar1.Buttons(13).Enabled = B
Me.Option1(0).Visible = B
Me.Option1(1).Visible = B

'Prueba
cmdAceptar.Visible = Not B
cmdCancelar.Visible = Not B
DataGrid1.Enabled = B

'Si es regresar
If DatosADevolverBusqueda <> "" Then
    cmdRegresar.Visible = B
End If
'Si estamo mod or insert
If Modo = 2 Then
   txtAux(0).BackColor = &H80000018
   Else
    txtAux(0).BackColor = &H80000005
End If
txtAux(0).Enabled = (Modo <> 2)
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    PonerModo 1
    Limpiar Me
    'Obtenemos la siguiente numero de factura
    NumF = SugerirCodigoSiguiente
    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DeseleccionaGrid
    DataGrid1.AllowAddNew = True
    If adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    
   
    If DataGrid1.Row < 0 Then
        On Error Resume Next
        anc = DataGrid1.RowTop(0) + DataGrid1.Top
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top
    End If
    txtAux(0).Text = NumF
    
    LLamaLineas anc, 0
    
    
    'Ponemos el foco
    txtAux(0).SetFocus
    
'    If FormularioHijoModificado Then
'        CargaGrid
'        BotonAnyadir
'        Else
'            'cmdCancelar.SetFocus
'            If Not Adodc1.Recordset.EOF Then _
'                Adodc1.Recordset.MoveFirst
'    End If
End Sub



Private Sub BotonVerTodos()
    CargaGrid ""
End Sub

Private Sub BotonBuscar()
    CargaGrid "codigo = -1"
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    txtAux(2).Text = ""
    txtAux(3).Text = ""
    LLamaLineas DataGrid1.Top + 206, 2
    txtAux(0).SetFocus
        
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim cad As String
    Dim anc As Single
    Dim I As Integer
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub


    
    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    DeseleccionaGrid
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top
    End If

    'Llamamos al form
    For I = 0 To 3
        txtAux(I).Text = DataGrid1.Columns(I).Text
    Next I
    I = adodc1.Recordset!Codigo
    LLamaLineas anc, 1
   
   'Como es modificar
   txtAux(1).SetFocus
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim I As Integer
    PonerModo xModo + 1
    'Fijamos el ancho
    For I = 0 To 3
        txtAux(I).Top = alto
    Next I
    'txtAux(1).Top = alto
End Sub




Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    If Not SepuedeBorrar Then Exit Sub
    '### a mano
    SQL = "Seguro que desea eliminar el gasto fijo:"
    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Descripcion: " & adodc1.Recordset.Fields(1)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
    
        'Hay que eliminar
        SQL = "Delete from sgastfij where codigo=" & adodc1.Recordset!Codigo
        Conn.Execute SQL
        
        'Hay que eliminar
        SQL = "Delete from sgastfijd where codigo=" & adodc1.Recordset!Codigo
        Conn.Execute SQL
        
        CargaGrid ""
        adodc1.Recordset.Cancel
    End If
    Exit Sub
Error2:
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub





Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    If adReason <> adRsnMove Then Exit Sub
    Label1.Caption = ""
    If Modo = 0 Then
        If Not adodc1.Recordset.EOF Then
            'Label1.Caption = adodc1.Recordset!Nombre
            If Not IsNull(adodc1.Recordset!Codigo) Then
                Label1.Caption = adodc1.Recordset!descripcion
                Label1.Tag = adodc1.Recordset!Codigo
            End If
        End If
    End If
    CargarTreeview
End Sub



Private Sub cmdAceptar_Click()
    If Modo = 4 Then
        'lineas
        
    Else
        AceptarCab
    End If
End Sub


Private Sub AceptarCab()
Dim I As Integer
Dim CadB As String
Select Case Modo
    Case 1
    If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                
                'MsgBox "Registro insertado.", vbInformation
                'SituarData1
                CargaGrid
                DataGrid1.AllowAddNew = False
                SituarData
                PonerModo 0
            End If
        End If
    Case 2
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    I = adodc1.Recordset.Fields(0)
                    PonerModo 0
                    CargaGrid
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
                End If
            End If
    Case 3
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        If CadB <> "" Then
            PonerModo 0
            CargaGrid CadB
        End If
    End Select


End Sub

Private Sub cmdCab_Click()
    PonerModo 0
    CargarTreeview
End Sub

Private Sub cmdCancelar_Click()
Select Case Modo
Case 1
    DataGrid1.AllowAddNew = False
    Me.Refresh
    CargaGrid
    If Not adodc1.Recordset.EOF Then
        adodc1.Recordset.MoveFirst
        Label1.Caption = adodc1.Recordset!descripcion
        Label1.Tag = adodc1.Recordset!Codigo
        
        'espera 1
    End If
    CargarTreeview
Case 3
    CargaGrid
End Select
PonerModo 0
lblIndicador.Caption = ""

End Sub

Private Sub cmdPregunta_Click(Index As Integer)
Dim B As Byte
    CadenaDesdeOtroForm = ""
    If Index = 0 Then
            B = 0
            If Option2(1).Value Then
                B = 1
            Else
                If Option2(2).Value Then B = 2
            End If
            frmGeneraGastos.Elemento = adodc1.Recordset!Codigo
            frmGeneraGastos.Opcion = B
            frmGeneraGastos.Show vbModal
            If CadenaDesdeOtroForm = "" Then Exit Sub
        
    
    End If
    PreparaGenerar False
    If CadenaDesdeOtroForm <> "" Then CargarTreeview
End Sub

Private Sub PreparaGenerar(Mostrar As Boolean)
    Toolbar1.Enabled = Not Mostrar
    Me.mnOpciones.Enabled = Not Mostrar
    Frame2.Visible = Mostrar
    If Mostrar Then
        cmdPregunta(1).Cancel = True
        cmdPregunta(0).SetFocus
    Else
        cmdCancelar.Cancel = True
    End If
End Sub


Private Sub cmdRegresar_Click()
    Dim cad As String
    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro a devolver.", vbExclamation
        Exit Sub
    End If
    cad = adodc1.Recordset.Fields(0) & "|"
    cad = cad & adodc1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub cmdRegresar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    PrimeraVez = True
          ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 10
        .Buttons(11).Image = 16
        .Buttons(13).Image = 21
        .Buttons(15).Image = 15
        
        
        .Buttons(17).Image = 6
        .Buttons(18).Image = 7
        .Buttons(19).Image = 8
        .Buttons(20).Image = 9
    End With
    
    Me.Icon = frmPpal.Icon
    CargaNodosIncial
    Frame2.Visible = False
    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    'Bloqueo de tabla, cursor type
 
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False
    PonerOpcionesMenu  'En funcion del usuario
    'Cadena consulta
    CadenaConsulta = "Select sgastfij.codigo,sgastfij.descripcion,sgastfij.ctaprevista,sgastfij.contrapar,nommacta from sgastfij,cuentas WHERE sgastfij.ctaprevista=cuentas.codmacta"
    CargaGrid
    lblIndicador.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(3).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub ImBanco_Click()
    Set frmB = New frmBanco
    frmB.DatosADevolverBusqueda = "0|1|"
    frmB.Show vbModal
    Set frmB = Nothing
End Sub

Private Sub imgcta_Click()
    Set frmC = New frmColCtas
    frmC.DatosADevolverBusqueda = "0|1|"
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnCOntabilizar_Click()


    If Modo <> 0 Then Exit Sub
        
    If TreeView1.Nodes.Count < 13 Then Exit Sub
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    If TreeView1.SelectedItem.Index < 13 Then Exit Sub
    
    If TreeView1.SelectedItem.BackColor = vbRed Then Exit Sub
    
    
    'AHORA CAONTABILZIAMOS
    CadenaDesdeOtroForm = adodc1.Recordset!Codigo & "|" & adodc1.Recordset!descripcion & "|"
    'Cuenta BANCO
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & adodc1.Recordset!Ctaprevista & "|" & Text1(0).Text & "|"
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & DBLet(adodc1.Recordset!contrapar, "T") & "|" & Text1(1).Text & "|"
    'Fecha e importe
    CadenaDesdeOtroForm = CadenaDesdeOtroForm & TreeView1.SelectedItem.Text & "|"
    
    frmVarios.Opcion = 19
    frmVarios.Show vbModal
    Me.Refresh
    Screen.MousePointer = vbHourglass
    CargarTreeview
    Screen.MousePointer = vbDefault
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
    Dim RS As ADODB.Recordset
    
    SQL = "Select Max(codigo) from sgastfij"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, , , adCmdText
    SQL = "1"
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            SQL = CStr(RS.Fields(0) + 1)
        End If
    End If
    RS.Close
    SugerirCodigoSiguiente = SQL
End Function

Private Sub Option1_Click(Index As Integer)
    CargarTreeview
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub

Private Sub HacerToolBar(ButtonIndex As Integer)

Select Case ButtonIndex
Case 1
        BotonBuscar
Case 2
        BotonVerTodos
Case 6
        If Modo < 4 Then
            BotonAnyadir
        Else
            'ñinea
            CadenaDesdeOtroForm = ""
            BotonLinea False
        End If
Case 7
        If Modo < 4 Then
            BotonModificar
        Else
            If TreeView1.SelectedItem.Index > 12 Then
                CadenaDesdeOtroForm = Trim(Mid(TreeView1.SelectedItem.Text, 1, 12)) & "|" & Trim(Mid(TreeView1.SelectedItem.Text, 12)) & "|"
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Abs(TreeView1.SelectedItem.BackColor = vbRed) & "|"
                    
            Else
                MsgBox "Seleccione el importe a modificar", vbExclamation
                Exit Sub
            End If
            BotonLinea True
        End If
Case 8
        If Modo < 4 Then
            BotonEliminar
        Else
            If TreeView1.SelectedItem.Index > 12 Then
                EliminaLinea
            Else
                MsgBox "Seleccione el importe a modificar", vbExclamation
            End If
            
        End If
        
        
Case 10
        'Si esta vacio no ponemosne lineas
        If adodc1.Recordset.EOF Then Exit Sub
        PonerModo 4
        
Case 11
    frmListado.Opcion = 31
    frmListado.Show vbModal
        
Case 13
        'Generar
        '--------
        If adodc1.Recordset.EOF Then Exit Sub
        If Modo = 0 Then PreparaGenerar True
        
Case 15
        Unload Me
Case Else

End Select
End Sub


Private Sub BotonLinea(Generar As Boolean)
        
    If adodc1.Recordset.EOF Then Exit Sub
        'En datosdesdeotroform sabremos si es nuevo, o modificar
            frmGeneraGastos.Opcion = 0

            frmGeneraGastos.Elemento = adodc1.Recordset!Codigo
            frmGeneraGastos.Show vbModal
            If CadenaDesdeOtroForm = "" Then
                'Generar
                PonerModo 4
                
            Else
                'Ha modificado
                'PonerModo 0
                CargarTreeview
            End If
            
End Sub


Private Sub EliminaLinea()
Dim cad As String
    If MsgBox("Eliminar el gasto: " & vbCrLf & TreeView1.SelectedItem.Text & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    cad = Format(Trim(Mid(TreeView1.SelectedItem.Text, 1, InStr(5, TreeView1.SelectedItem.Text, " "))), FormatoFecha)
    cad = "DELETE from sgastfijd where fecha='" & cad
    cad = cad & "' AND codigo =" & adodc1.Recordset!Codigo
    If Ejecuta(cad) Then CargarTreeview
    
End Sub

Private Sub DespalzamientoVisible(Bol As Boolean)
    Dim I
    For I = 17 To 20
        Toolbar1.Buttons(I).Visible = Bol
    Next I
End Sub

Private Sub CargaGrid(Optional SQL As String)
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim I As Integer
    
    adodc1.ConnectionString = Conn
    If SQL <> "" Then
        SQL = CadenaConsulta & " AND " & SQL
        Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY codigo"
    adodc1.RecordSource = SQL
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    
    'Nombre producto
    I = 0
        DataGrid1.Columns(I).Caption = "Cod."
        DataGrid1.Columns(I).Width = 700
        DataGrid1.Columns(I).NumberFormat = "000"
        
    
    'Leemos del vector en 2
    I = 1
        DataGrid1.Columns(I).Caption = "Nombre"
        DataGrid1.Columns(I).Width = 2800
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
        'Leemos del vector en 2
    I = 2
        DataGrid1.Columns(I).Caption = "Prevista"
        DataGrid1.Columns(I).Width = 1200
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
       
       
        'Leemos del vector en 2
    I = 3
        DataGrid1.Columns(I).Caption = "Contrapar."
        DataGrid1.Columns(I).Width = 1200
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
       
        
    DataGrid1.Columns(4).Visible = False
        
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(0).Left = DataGrid1.Columns(0).Left + 45
        txtAux(0).Width = DataGrid1.Columns(0).Width - 45
        For I = 1 To 3
            txtAux(I).Left = DataGrid1.Columns(I).Left + 75
            txtAux(I).Width = DataGrid1.Columns(I).Width - 30
        Next I
    '
        CadAncho = True
    End If
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
        Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
    End If
End Sub

Private Sub TreeView1_DblClick()
    
        If TreeView1.Nodes.Count > 12 Then
            If TreeView1.SelectedItem.Index > 12 Then
                If Modo = 4 Then
                    HacerToolBar 7
                Else
                    If Modo = 0 Then mnCOntabilizar_Click
                End If
            End If
        End If
 
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 2 And Modo = 0 Then
        PopupMenu mnPop
    End If
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    With txtAux(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim SQL As String
    txtAux(Index).Text = Trim(txtAux(Index).Text)
    
    If txtAux(Index).Text = "" Then
        If Index > 1 Then Text1(Index - 2).Text = ""
        Exit Sub
    End If
    If Modo = 3 Then Exit Sub 'Busquedas
    Select Case Index
    Case 0
        If Not IsNumeric(txtAux(0).Text) Then
            MsgBox "Código concepto tiene que ser numérico", vbExclamation
            Exit Sub
        End If
        txtAux(0).Text = Format(txtAux(0).Text, "000")
    Case 2, 3
        
        
        CadenaDesdeOtroForm = txtAux(Index).Text

        
        If CuentaCorrectaUltimoNivel(CadenaDesdeOtroForm, SQL) Then
            If Index = 2 Then
                CadenaDesdeOtroForm = DevuelveDesdeBD("codmacta", "ctabancaria", "codmacta", CadenaDesdeOtroForm, "T")
                If CadenaDesdeOtroForm = "" Then
                    SQL = ""
                    MsgBox "La cuenta contable no esta asociada a ninguna cuenta bancaria", vbExclamation
                End If
            End If
        Else
            MsgBox SQL, vbExclamation
            CadenaDesdeOtroForm = ""
            SQL = ""
        End If

        
        txtAux(Index).Text = CadenaDesdeOtroForm
        Text1(Index - 2).Text = SQL
        If CadenaDesdeOtroForm = "" Then Ponerfoco txtAux(Index)
        

        
        
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim Datos As String
Dim B As Boolean
B = CompForm(Me)
If Not B Then Exit Function

If Modo = 1 Then
    'Estamos insertando
     Datos = DevuelveDesdeBD("CODIGO", "sgastfij", "codigo", txtAux(0).Text, "N")
     If Datos <> "" Then
        MsgBox "Ya existe el codigo : " & txtAux(0).Text, vbExclamation
        B = False
    End If
End If
DatosOk = B
End Function



Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub



Private Function SepuedeBorrar() As Boolean
Dim SQL As String
    SepuedeBorrar = False
'    SQL = DevuelveDesdeBD("agente", "scobro", "agente", adodc1.Recordset!Codigo, "N")
'    If SQL <> "" Then
'        MsgBox "Existen cobros pendientes para este agente", vbExclamation
'        Exit Function
'    End If
'
    
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
    'Caption = KeyAscii
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            If Modo = 0 Then Unload Me
        End If
    End If
End Sub


Private Sub CargaNodosIncial()
Dim F As Date
Dim N As Node
    F = vParam.fechaini
    While F <= vParam.fechafin
        Set N = TreeView1.Nodes.Add
        N.Key = "F" & Format(F, "mm")
        N.Text = UCase(Format(F, "mmmm"))
        N.Image = 1
        F = DateAdd("m", 1, F)
    Wend
    
        
End Sub


Private Sub VaciaNodos()
Dim I As Integer
    'Los doce primero son los meses
    For I = 12 To 1 Step -1
        TreeView1.Nodes(I).Bold = False
        TreeView1.Nodes(I).Image = 1
    Next I
    
    If TreeView1.Nodes.Count > 12 Then
        For I = TreeView1.Nodes.Count To 13 Step -1
                TreeView1.Nodes.Remove I
        Next I
    End If
End Sub




Private Sub CargarTreeview()
Dim cad As String
Dim F As Date
Dim N As Node
Dim I As Integer
Dim Carga As Boolean

    VaciaNodos
    
    cad = "Select * from sgastfijd where codigo ="
    If Label1.Caption = "" Then
        Carga = False
        cad = cad & "-1"
        Text1(0).Text = "": Text1(1).Text = ""
    Else
        Carga = True
        Text1(0).Text = adodc1.Recordset!Nommacta
        
        If Not IsNull(adodc1.Recordset!contrapar) Then
            Text1(1).Text = DevuelveDesdeBD("nommacta", "cuentas", "codmacta", adodc1.Recordset!contrapar, "T")
        Else
            Text1(1).Text = ""
        End If
        cad = cad & Label1.Tag
        If Option1(0).Value Then
            F = vParam.fechaini
        Else
            F = DateAdd("d", 1, vParam.fechafin)
        End If
        
        cad = cad & " and fecha>= '" & Format(F, FormatoFecha)
        F = DateAdd("yyyy", 1, F)
        F = DateAdd("d", -1, F)
        cad = cad & "' and fecha<= '" & Format(F, FormatoFecha) & "'"
        cad = cad & " ORDER BY fecha"
    End If
    If Not Carga Then Exit Sub
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    TreeView1.Tag = Space(30)
    While Not miRsAux.EOF
        
        cad = "F" & Format(miRsAux!Fecha, "mm")
        Set N = TreeView1.Nodes.Add(cad, tvwChild, "E" & Format(miRsAux!Fecha, "dd/mm/yyyy"), , 2)
        N.Text = Format(miRsAux!Fecha, "dd/mm/yyyy") & Right(TreeView1.Tag & Format(miRsAux!Importe, FormatoImporte), 30)
        If miRsAux!contabilizado = 1 Then
            N.BackColor = vbRed
            N.ForeColor = vbWhite
        End If
        N.EnsureVisible
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If cad <> "" Then
        For I = 1 To 12
            If TreeView1.Nodes(I).Children > 0 Then
                TreeView1.Nodes(I).Bold = True
                TreeView1.Nodes(I).Image = 3
            End If
        Next I
    End If
End Sub

Private Sub SituarData()
    Me.adodc1.Recordset.Find " codigo = " & Val(Me.txtAux(0).Text), , , 1
End Sub
