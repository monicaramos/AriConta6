VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESCompensaAboCli 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compensaci�n abonos cliente"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13050
   Icon            =   "frmTESCompensaAboCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   13050
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6885
      Left            =   90
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   12735
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   1
         Left            =   11550
         TabIndex        =   20
         Top             =   6300
         Width           =   975
      End
      Begin VB.CommandButton Command1 
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
         Left            =   10350
         TabIndex        =   19
         Top             =   6300
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   240
         TabIndex        =   17
         Top             =   180
         Width           =   3585
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   180
            TabIndex        =   22
            Top             =   240
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
      Begin MSComctlLib.ListView lw1 
         Height          =   5085
         Left            =   240
         TabIndex        =   18
         Top             =   990
         Width           =   12315
         _ExtentX        =   21722
         _ExtentY        =   8969
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "C�digo"
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cuenta"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nombre"
            Object.Width           =   10054
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   390
         Left            =   12060
         TabIndex        =   21
         Top             =   300
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
   End
   Begin VB.Frame FrameCompensaAbonosCliente 
      BorderStyle     =   0  'None
      Height          =   6885
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   12735
      Begin VB.Frame FrameBotonGnral 
         Height          =   705
         Left            =   180
         TabIndex        =   13
         Top             =   150
         Width           =   1035
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   330
            Left            =   240
            TabIndex        =   14
            Top             =   180
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Establecer Vencimiento"
                  Object.Tag             =   "2"
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txtimpNoEdit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Left            =   10380
         TabIndex        =   11
         Top             =   5790
         Width           =   1845
      End
      Begin VB.TextBox txtimpNoEdit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Left            =   10260
         TabIndex        =   10
         Top             =   1230
         Width           =   2025
      End
      Begin VB.TextBox txtimpNoEdit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Left            =   8250
         TabIndex        =   7
         Top             =   1230
         Width           =   1965
      End
      Begin MSComctlLib.ListView lwCompenCli 
         Height          =   3975
         Left            =   240
         TabIndex        =   6
         Top             =   1650
         Width           =   12315
         _ExtentX        =   21722
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Factura"
            Object.Width           =   2116
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Vto"
            Object.Width           =   1234
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fecha Vto"
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Forma pago"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Cobro"
            Object.Width           =   3590
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Abonos"
            Object.Width           =   3590
         EndProperty
      End
      Begin VB.TextBox DtxtCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
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
         Index           =   17
         Left            =   1560
         TabIndex        =   4
         Text            =   "Text5"
         Top             =   1200
         Width           =   4905
      End
      Begin VB.TextBox txtCta 
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
         Index           =   17
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1305
      End
      Begin VB.CommandButton cmdCompensar 
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
         Left            =   10350
         TabIndex        =   2
         Top             =   6300
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
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
         Index           =   0
         Left            =   11550
         TabIndex        =   1
         Top             =   6300
         Width           =   975
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   12060
         TabIndex        =   15
         Top             =   300
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   72
         Left            =   9300
         TabIndex        =   12
         Top             =   5835
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Rectifca./Abono"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   71
         Left            =   10740
         TabIndex        =   9
         Top             =   960
         Width           =   1590
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cobro"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   70
         Left            =   9300
         TabIndex        =   8
         Top             =   960
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   69
         Left            =   240
         TabIndex        =   5
         Top             =   930
         Width           =   1440
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   17
         Left            =   1770
         Top             =   930
         Width           =   240
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

Private Const IdPrograma = 607


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

Dim SQL As String
Dim RC As String
Dim RS As Recordset
Dim PrimeraVez As Boolean

Dim cad As String
Dim CONT As Long
Dim I As Integer
Dim TotalRegistros As Long

Dim Importe As Currency
Dim MostrarFrame As Boolean
Dim Fecha As Date

Dim DevfrmCCtas As String

Dim CampoOrden As String
Dim Orden As Boolean
Dim Modo As Byte

Dim Txt33Csb As String
Dim Txt41Csb As String

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
    If Index = 0 Then
        Frame1.Visible = True
        Frame1.Enabled = True
    Else
        Unload Me
    End If
End Sub


Private Sub cmdCompensar_Click()
    
    cad = DevuelveDesdeBD("informe", "scryst", "codigo", IdPrograma) 'Orden de pago a bancos
    If cad = "" Then
        MsgBox "No esta configurada la aplicaci�n. Falta el informe", vbCritical
        Exit Sub
    End If
    Me.Tag = cad
    
    cad = ""
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
            cad = cad & "A"
            If CONT = 0 Then CONT = I
        End If
    Next
    
    I = 0
    SQL = ""
    If Len(cad) <> 1 Then
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
'        'Hay seleccionado uno vto
'        Set miRsAux = New ADODB.Recordset
'        cad = "text41csb"
'        cad = "Select " & cad & " FROM cobros WHERE numserie ='" & lwCompenCli.ListItems(CONT).Text & "' AND numfactu="
'        cad = cad & lwCompenCli.ListItems(CONT).SubItems(1) & " AND fecfactu='" & Format(lwCompenCli.ListItems(CONT).SubItems(2), FormatoFecha)
'        cad = cad & "' AND numorden = " & lwCompenCli.ListItems(CONT).SubItems(3)
'        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'        If miRsAux.EOF Then
'            SQL = SQL & vbCrLf & " NO se ha encontrado el vto. destino"
'        Else
'            'Vamos a ver cuantos registros son
'            CadenaDesdeOtroForm = ""
'            RC = "0"
'            For I = 0 To 0
'                If DBLet(miRsAux.Fields(I), "T") = "" Then
'                    CadenaDesdeOtroForm = CadenaDesdeOtroForm & miRsAux.Fields(I).Name & "|"
'                    RC = Val(RC) + 1
'                End If
'            Next I
'
'
'            'If TotalRegistros + NumRegElim > 15 Then SQL = SQL & vbCrLf & "No caben los textos de los vencimientos"
'            If TotalRegistros + NumRegElim > Val(RC) Then SQL = SQL & vbCrLf & "No caben los textos de los vencimientos"
'        End If
'        miRsAux.Close
'        Set miRsAux = Nothing

        Dim CadAux As String
        
        Txt33Csb = "Compensa: "
        Txt41Csb = ""
        For I = 1 To Me.lwCompenCli.ListItems.Count - 1
            If Me.lwCompenCli.ListItems(I).Checked Then
                CadAux = Trim(lwCompenCli.ListItems(I).Text & lwCompenCli.ListItems(I).SubItems(1)) & " " & Trim(lwCompenCli.ListItems(I).SubItems(2))
                If Len(Txt33Csb & " " & CadAux) < 80 Then
                    Txt33Csb = Txt33Csb & " " & CadAux
                Else
                    If Len(Txt41Csb & " " & CadAux) < 60 Then
                        Txt41Csb = Txt41Csb & CadAux
                    Else
                        Txt41Csb = Txt41Csb & ".."
                        Exit For
                    End If
                End If
            End If
        Next I
        

    End If
    
    
'    If SQL <> "" Then
'        MsgBox SQL, vbExclamation
'        Exit Sub
'    Else
        If MsgBox("Seguro que desea realizar la compensaci�n?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
'    End If
    
    
    Me.FrameCompensaAbonosCliente.Enabled = False
    Me.Refresh
    Screen.MousePointer = vbHourglass
    
    RealizarCompensacionAbonosClientes
    Me.FrameCompensaAbonosCliente.Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdVtoDestino(Index As Integer)
    
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
        
        PonerFocoLw Me.lwCompenCli

    Else
    
        frmTESCompensaAboCliImp.pCodigo = Me.lw1.SelectedItem
        frmTESCompensaAboCliImp.Show vbModal
    
    
    
'        cad = DevuelveDesdeBD("informe", "scryst", "codigo", 10) 'Orden de pago a bancos
'        If cad = "" Then
'            MsgBox "No esta configurada la aplicaci�n. Falta informe(10)", vbCritical
'            Exit Sub
'        End If
'        CadenaDesdeOtroForm = cad
    
        'LanzaBuscaGrid 1, 4


    End If
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        Case 36
            If CadenaDesdeOtroForm <> "" Then
                txtCta(17).Text = CadenaDesdeOtroForm
                txtCta_LostFocus 17
            Else
                PonFoco txtCta(17)
            End If
            CadenaDesdeOtroForm = ""
            
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub



    
Private Sub Form_Load()
Dim h As Integer
Dim W As Integer
Dim Img As Image


    Limpiar Me
    Me.Icon = frmPpal.Icon
    CargaImagenesAyudas Me.Image3, 1, "Cuenta contable"
    
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
    
    
    'La toolbar
    With Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 4
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
    
    
    'Limpiamos el tag
    PrimeraVez = True
    FrameCompensaAbonosCliente.Visible = False
    CommitConexion  'Porque son listados. No hay nada dentro transaccion
    
    Select Case Opcion
        
    Case 36
        
        h = FrameCompensaAbonosCliente.Height + 120
        W = FrameCompensaAbonosCliente.Width
        
        FrameCompensaAbonosCliente.Visible = False
        Me.Frame1.Visible = True
        
    End Select
    
    Me.Width = W + 300
    Me.Height = h + 400
    
    Me.cmdCancelar(0).Cancel = True
    
    PonerModoUsuarioGnral 0, "ariconta"
    
'    PonerFrameProgreso

End Sub


'Private Sub PonerFrameProgreso()
'Dim I As Integer
'
'    'Ponemos el frame al pricnipio de todo
'    FrameProgreso.Visible = False
'    FrameProgreso.ZOrder 0
'
'    'lo ubicamos
'    'Posicion horizintal WIDTH
'    I = Me.Width - FrameProgreso.Width
'    If I > 100 Then
'        I = I \ 2
'    Else
'        I = 0
'    End If
'    FrameProgreso.Left = I
'    'Posicion  VERTICAL HEIGHT
'    I = Me.Height - FrameProgreso.Height
'    If I > 100 Then
'        I = I \ 2
'    Else
'        I = 0
'    End If
'    FrameProgreso.Top = I
'End Sub



Private Sub frmCta_DatoSeleccionado(CadenaSeleccion As String)
    txtCta(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 1)
    DtxtCta(CInt(RC)).Text = RecuperaValor(CadenaSeleccion, 2)
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

Private Sub lw1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim Campo2 As Integer

    Orden = Not Orden
    
    Select Case ColumnHeader
        Case "C�digo"
            CampoOrden = "compensa.codigo"
        Case "Fecha"
            CampoOrden = "compensa.fecha"
        Case "Cuenta"
            CampoOrden = "compensa.codmacta"
        Case "Nombre"
            CampoOrden = "compensa.nommacta"
    End Select
    CargaList


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

Private Sub HacerToolBar(Boton As Integer)

    Select Case Boton
        Case 1
            BotonAnyadir
        Case 2
'            BotonModificar
        Case 3
'            BotonEliminar False
        Case 5
'            BotonBuscar
        Case 6 ' ver todos
            CargaList
        Case 8
            'Imprimir factura
            
            
'            frmFacturasCliList.NUmSerie = Text1(2).Text
'            frmFacturasCliList.NumFactu = Text1(0).Text
'            frmFacturasCliList.FecFactu = Text1(1).Text
'
'            frmFacturasCliList.Show vbModal

             cmdVtoDestino (1)


    End Select
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar Button.Index
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    HacerToolBar2 Button.Index
End Sub

Private Sub HacerToolBar2(Boton As Integer)

    Select Case Boton
        Case 1
            cmdVtoDestino (0)
    End Select
End Sub


Private Sub BotonAnyadir()

    Frame1.Visible = False
    Frame1.Enabled = False

    Me.FrameCompensaAbonosCliente.Visible = True
    Me.FrameCompensaAbonosCliente.Enabled = True
    

End Sub


Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtCta_GotFocus(Index As Integer)
    PonFoco txtCta(Index)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCta_LostFocus(Index As Integer)
Dim Cta As String
Dim B As Byte
    txtCta(Index).Text = Trim(txtCta(Index).Text)
    
     
     
    If txtCta(Index).Text = "" Then
        DtxtCta(Index).Text = ""
       ' txtCta(6).Tag = txtCta(6).Text
        Exit Sub
    End If
    
    If Not IsNumeric(txtCta(Index).Text) Then
        MsgBox "La cuenta debe ser num�rica: " & txtCta(Index).Text, vbExclamation
        txtCta(Index).Text = ""
        DtxtCta(Index).Text = ""
        txtCta(6).Tag = txtCta(6).Text
        PonFoco txtCta(Index)
        
        If Index = 17 Then PonerVtosCompensacionCliente
        
        Exit Sub
    End If
    
    Select Case Index
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

Private Sub PonerFrameProgressVisible(Optional TEXTO As String)
        If TEXTO = "" Then TEXTO = "Generando datos"
'        Me.lblPPAL.Caption = TEXTO
'        Me.lbl2.Caption = ""
'        Me.ProgressBar1.Value = 0
'        Me.FrameProgreso.Visible = True
        Me.Refresh
End Sub



'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'
'               CREDITO CAUCION
'
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------




'Para conceptos y diarios
'Opcion: 0- Diario
'        1- Conceptos
'        2- Centros de coste
'        3- Gastos fijos
'        4. Hco compensaciones
Private Sub LanzaBuscaGrid(Indice As Integer, OpcionGrid As Byte)

    Select Case OpcionGrid
    Case 4
'        'Gasto fijos  sgastfij codigo Descripcion
'        '-------------------------------------------
'        DevfrmCCtas = "0"
'        cad = "C�digo|codigo|T|10�"
'        cad = cad & "Fecha|fecha|T|26�"
'        cad = cad & "Cuenta|codmacta|T|14�"
'        cad = cad & "Nombre|nommacta|T|50�"
'
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vTabla = "scompenclicab"
'        frmB.vSQL = ""
'
'        '###A mano
'        frmB.vDevuelve = "0|"
'        frmB.vTitulo = "Hco. compensaciones cliente"
'        frmB.vSelElem = 0
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        If DevfrmCCtas <> "" Then
'            DevfrmCCtas = RecuperaValor(DevfrmCCtas, 1)
'            If DevfrmCCtas = "" Then DevfrmCCtas = "0"
'            If Val(DevfrmCCtas) Then
'                CONT = Val(DevfrmCCtas)
'                ImprimiCompensacion CONT
'            End If
'
'        End If
    End Select
End Sub


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
    cad = "Select cobros.*,nomforpa from cobros,formapago where cobros.codforpa=formapago.codforpa "
    cad = cad & " AND codmacta = '" & Me.txtCta(17).Text & "'"
    cad = cad & " AND (transfer =0 or transfer is null) and codrem is null"
    cad = cad & " and recedocu=0 and situacion = 0" ' pendientes de cobro
    cad = cad & " ORDER BY fecvenci"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwCompenCli.ListItems.Add()
        IT.Text = miRsAux!NUmSerie
        IT.SubItems(1) = Format(miRsAux!NumFactu, "000000")
        IT.SubItems(2) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
        IT.SubItems(3) = miRsAux!NUmorden
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

        cad = DevuelveDesdeBD("max(codigo)", "compensa", "1", "1")
        If cad = "" Then cad = "0"
        CONT = Val(cad) + 1 'ID de la operacion
        
        cad = "INSERT INTO compensa(codigo,fecha,login,PC,codmacta,nommacta) VALUES (" & CONT
        cad = cad & ",now(),'" & DevNombreSQL(vUsu.Login) & "','" & DevNombreSQL(vUsu.PC)
        cad = cad & "','" & txtCta(17).Text & "','" & DevNombreSQL(DtxtCta(17).Text) & "')"
        
        Set miRsAux = New ADODB.Recordset
        Borras = True
        If Ejecuta(cad) Then
            
            Borras = Not RealizarProcesoCompensacionAbonos
        
        End If


        Set miRsAux = Nothing
        If Borras Then
            Conn.Execute "DELETE FROM compensa WHERE codigo = " & CONT
            Conn.Execute "DELETE FROM compensa_facturas WHERE codigo = " & CONT
            
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
        cad = "INSERT INTO Usuarios.z347carta (codusu, nif, razosoci, dirdatos, codposta, despobla, otralineadir, "
        cad = cad & "parrafo1, parrafo2, contacto, despedida,saludos,parrafo3, parrafo4, parrafo5, Asunto, Referencia)"
        cad = cad & " VALUES (" & vUsu.Codigo & ", "
        
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
      
        cad = cad & SQL
        cad = cad & ",NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
        
        Conn.Execute cad
        
        
        'Datos CLIENTE
         ', texto3, texto4, texto5,texto6
        cad = DevuelveDesdeBD("codmacta", "scompenclicab", "codigo", CStr(CONT))
        cad = "Select nommacta,razosoci,dirdatos,codposta,despobla,desprovi from cuentas where codmacta ='" & cad & "'"
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        'NO PUEDE SER EOF
        cad = miRsAux!Nommacta
        If Not IsNull(miRsAux!razosoci) Then cad = miRsAux!razosoci
        cad = "'" & DevNombreSQL(cad) & "'"
        'Direccion
        cad = cad & ",'" & DevNombreSQL(CStr(DBLet(miRsAux!dirdatos))) & "'"
        'Poblacion
        SQL = DBLet(miRsAux!codposta)
        If SQL <> "" Then SQL = SQL & " - "
        SQL = SQL & DevNombreSQL(CStr(DBLet(miRsAux!desPobla)))
        cad = cad & ",'" & SQL & "'"
        'Provincia
        cad = cad & ",'" & DevNombreSQL(CStr(DBLet(miRsAux!desProvi))) & "'"
        miRsAux.Close
        

        
        SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2, texto3, texto4,texto5,texto6, observa1, "
        SQL = SQL & "importe1, importe2, fecha1, fecha2, fecha3, observa2, opcion)"
        SQL = SQL & " VALUES (" & vUsu.Codigo & ",1,'',''," & cad
        
        'select Numfac,fecha from usuarios.ztmpfaclin where tipoiva=1 and codusu=2200
        Importe = 0
        RC = "NIF"   'RC = "fecha"   La fecha de VTo esta en el campo: NIF
        cad = DevuelveDesdeBD("numfac", "Usuarios.ztmpfaclin", "codusu =" & vUsu.Codigo & " AND tipoiva", "1", "N", RC)
        If cad = "" Then
            'Significa que la compesacion ha sido total, no quedaba resultante
            
        Else
            cad = "Quedando el resultado en el vencimiento: " & cad & " de " & Format(RC, "dd/mm/yyyy")
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
        SQL = SQL & ",'" & cad & "'," & TransformaComasPuntos(CStr(Importe))
        
        
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
    
    RC = "Select " & cad & ", gastos, impvenci, impcobro, fecvenci FROM cobros where (numserie,numfactu,fecfactu,numorden) IN (" & SQL & ")"
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
            If miRsAux!NumFactu = Val(Me.lwCompenCli.ListItems(NumRegElim).SubItems(1)) Then
                If Format(miRsAux!FecFactu, "dd/mm/yyyy") = Me.lwCompenCli.ListItems(NumRegElim).SubItems(2) Then
                    If miRsAux!NUmorden = Val(Me.lwCompenCli.ListItems(NumRegElim).SubItems(3)) Then Destino = 1
                End If
            End If
        End If
        
        RC = "INSERT INTO compensa_facturas (codigo,linea,destino," & cad & ") VALUES (" & CONT & "," & I & "," & Destino & "," & DBSet(miRsAux!NUmSerie, "T")
        RC = RC & "," & DBSet(miRsAux!NumFactu, "N") & "," & DBSet(miRsAux!FecFactu, "F") & "," & DBSet(miRsAux!NUmorden, "N") & ")"
        Conn.Execute RC
        
        'Para las observaciones de despues
        Importe = DBLet(miRsAux!Gastos, "N")
        Importe = Importe + miRsAux!ImpVenci
        'Si ya he cobrado algo
        If Not IsNull(miRsAux!impcobro) Then Importe = Importe - miRsAux!impcobro
        
        If Destino = 0 Then 'El destino
            DevfrmCCtas = DevfrmCCtas & miRsAux!NUmSerie & Format(miRsAux!NumFactu, "0000000") & " " & Format(miRsAux!FecFactu, "dd/mm/yyyy")
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
        'Insertaremos registros en cobros_realizados BORRAREMOS LOS VENCIMIENTOS QUE NO SEAN DESTINO a no ser que el importe restante sea 0
        Destino = 1
        If txtimpNoEdit(0).Text = txtimpNoEdit(1).Text Then Destino = 0
        
        SQLVtosSeleccionadosCompensacion 0, Destino = 1  'sin o con el destino
        
        'Para saber si ha ido bien
        Destino = 0    '0 mal,1 bien
        If InsertarCobrosRealizados(SQL) Then
            If txtimpNoEdit(0).Text = txtimpNoEdit(1).Text Then
                Destino = 1
            Else
                'Updatearemos los campos csb del vto restante. A partir del segundo
                'La variable CadenaDesdeOtroForm  tiene los que vamos a actualizar
                
                cad = ""
                J = 0
                SQL = ""
                
                
                Importe = CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag)  'txtimpNoEdit(1).Tag es negativo
                
                RC = "gastos=null, impcobro=null,fecultco=null,impvenci=" & TransformaComasPuntos(CStr(Importe))
                RC = RC & ",text33csb=" & DBSet(Txt33Csb, "T")
                RC = RC & ",text41csb=" & DBSet(Txt41Csb, "T")
                
                SQL = RC & SQL
                SQL = "UPDATE cobros SET " & SQL
                'WHERE
                RC = ""
                For J = 1 To Me.lwCompenCli.ListItems.Count
                    If Me.lwCompenCli.ListItems(J).Bold Then
                        'Este es el destino
                        RC = "NUmSerie = '" & Me.lwCompenCli.ListItems(J).Text
                        RC = RC & "' AND numfactu = " & Val(Me.lwCompenCli.ListItems(J).SubItems(1))
                        RC = RC & " AND fecfactu = '" & Format(Me.lwCompenCli.ListItems(J).SubItems(2), FormatoFecha)
                        RC = RC & "' AND numorden = " & Val(Me.lwCompenCli.ListItems(J).SubItems(3))
                        Exit For
                    End If
                Next
                If RC <> "" Then
                    cad = SQL & " WHERE " & RC
                    If Ejecuta(cad) Then Destino = 1
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

Private Function InsertarCobrosRealizados(facturas As String) As Boolean
Dim SQL As String
Dim SQL2 As String
Dim CadInsert As String
Dim CadValues As String
Dim NumLin As Long

    On Error GoTo eInsertarCobrosRealizados


    InsertarCobrosRealizados = True

    CadInsert = "insert into cobros_realizados (numserie, numfactu, fecfactu, numorden, numlinea, usuariocobro,fecrealizado,impcobro) values  "
    
    SQL = "select * from cobros where (numserie, numfactu, fecfactu, numorden) in (" & facturas & ")"
    
    CadValues = ""
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
    
        Importe = DBLet(RS!ImpVenci) + DBLet(RS!Gastos, "N") - DBLet(RS!impcobro, "N")
        
        SQL2 = "select max(numlinea) from cobros_realizados where numserie = " & DBSet(RS!NUmSerie, "T") & " and numfactu = " & DBSet(RS!NumFactu, "N")
        SQL2 = SQL2 & " and fecfactu = " & DBSet(RS!FecFactu, "F") & " and numorden = " & DBSet(RS!NUmorden, "N")
        NumLin = DevuelveValor(SQL2)
        NumLin = NumLin + 1
    
        CadValues = CadValues & "(" & DBSet(RS!NUmSerie, "T") & "," & DBSet(RS!NumFactu, "N") & "," & DBSet(RS!FecFactu, "F") & "," & DBSet(RS!NUmorden, "N")
        CadValues = CadValues & "," & DBSet(NumLin, "N") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(Now, "FH") & "," & DBSet(Importe, "N") & "),"
        
        
        ' actualizamos la cabecera del cobro pq ya no lo eliminamos
        SQL = "update cobros set situacion = 2, impcobro = impvenci + coalesce(gastos,0) where numserie = " & DBSet(RS!NUmSerie, "T")
        SQL = SQL & " and numfactu = " & DBSet(RS!NumFactu, "N") & " and fecfactu = " & DBSet(RS!FecFactu, "F") & " and numorden = " & DBSet(RS!NUmorden, "N")
        
        Conn.Execute SQL
        
        RS.MoveNext
    Wend
    
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        Conn.Execute CadInsert & CadValues
    End If
    
    
    Set RS = Nothing
    Exit Function
    
eInsertarCobrosRealizados:
    InsertarCobrosRealizados = False
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

    cad = "numserie, numfactu, fecfactu, numorden "
    
'    cad = "numserie , numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, ctabanc1,"
'    cad = cad & "entidad, oficina, control, cuentaba, iban, fecultco, impcobro, emitdocum, "
'    cad = cad & "recedocu, contdocu, text33csb, text41csb, "
'    cad = cad & "ultimareclamacion, agente, departamento, tiporem, CodRem, AnyoRem,"
'    cad = cad & "siturem, Gastos, Devuelto, situacionjuri, noremesar, observa, transfer, referencia,"
'    cad = cad & "nomclien, domclien, pobclien, cpclien, proclien, referencia1, referencia2,"
'    cad = cad & "feccomunica, fecprorroga, fecsiniestro, fecejecutiva, nifclien, codpais, situacion  "
    
End Sub



Private Sub PonerModoUsuarioGnral(Modo As Byte, aplicacion As String)
Dim RS As ADODB.Recordset
Dim cad As String
    
    On Error Resume Next

    cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(aplicacion, "T")
    cad = cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        Toolbar1.Buttons(1).Enabled = DBLet(RS!creareliminar, "N")
        Toolbar1.Buttons(2).Enabled = False 'DBLet(RS!Modificar, "N") And (Modo = 2) And DesdeNorma43 = 0
        Toolbar1.Buttons(3).Enabled = False 'DBLet(RS!creareliminar, "N") And (Modo = 2) And DesdeNorma43 = 0
        
        Toolbar1.Buttons(5).Enabled = False 'DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2) And DesdeNorma43 = 0
        Toolbar1.Buttons(6).Enabled = DBLet(RS!Ver, "N")
        
        Toolbar1.Buttons(8).Enabled = DBLet(RS!Imprimir, "N") And Modo = 2
    
        Toolbar2.Buttons(1).Enabled = True
'        Toolbar2.Buttons(2).Enabled = DBLet(RS!especial, "N")
    End If
    
    RS.Close
    Set RS = Nothing
    
End Sub



Private Sub CargaList()
Dim IT

    lw1.ListItems.Clear
    Set Me.lw1.SmallIcons = frmPpal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    
    cad = "Select codigo,fecha,codmacta,nommacta from compensa "
    
    
    If CampoOrden = "" Then CampoOrden = "compensa.codigo"
    cad = cad & " ORDER BY " & CampoOrden
    If Orden Then cad = cad & " DESC"
    
    
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw1.ListItems.Add()
        IT.Text = miRsAux!Codigo
        IT.SubItems(1) = Format(miRsAux!Fecha, "dd/mm/yyyy hh:mm:ss")
        IT.SubItems(2) = miRsAux!codmacta
        IT.SubItems(3) = miRsAux!Nommacta
        
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing

    If lw1.ListItems.Count > 0 Then
        Modo = 2
    Else
        Modo = 0
    End If
    
    PonerModoUsuarioGnral Modo, "ariconta"
    
End Sub

