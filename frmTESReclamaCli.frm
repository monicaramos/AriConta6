VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESReclamaCli 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reclamaciones"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13050
   Icon            =   "frmTESReclamaCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   13050
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6930
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6885
      Left            =   90
      TabIndex        =   8
      Top             =   -30
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
         TabIndex        =   11
         Top             =   6300
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   240
         TabIndex        =   9
         Top             =   180
         Width           =   3585
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   180
            TabIndex        =   12
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
         TabIndex        =   10
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2647
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cuenta"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   5468
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Envio"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Importe"
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Observac"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Codigo"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   12060
         TabIndex        =   13
         Top             =   210
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
   Begin VB.Frame FrameReclamacionesCliente 
      BorderStyle     =   0  'None
      Height          =   6885
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   12735
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   270
         TabIndex        =   18
         Tag             =   "Cuenta|T|N|||reclama|codmacta|||"
         Text            =   "99/99/9999"
         Top             =   1230
         Width           =   1245
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
         Height          =   840
         Index           =   0
         Left            =   270
         TabIndex        =   17
         Tag             =   "Observaciones|T|N|||reclama|observaciones|||"
         Top             =   2070
         Width           =   12045
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   1560
         TabIndex        =   14
         Tag             =   "Fecha vencimiento|F|N|||cobros|fecvenci|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   510
         Width           =   1245
      End
      Begin VB.TextBox Text1 
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
         Index           =   4
         Left            =   10470
         TabIndex        =   6
         Top             =   1200
         Width           =   1845
      End
      Begin MSComctlLib.ListView lwReclamCli 
         Height          =   2565
         Left            =   240
         TabIndex        =   5
         Top             =   3660
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   4524
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
      Begin VB.TextBox Text1 
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
         Index           =   3
         Left            =   1590
         TabIndex        =   3
         Text            =   "Text5"
         Top             =   1230
         Width           =   6645
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
         Left            =   10140
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
         Left            =   11340
         TabIndex        =   1
         Top             =   6300
         Width           =   975
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   7950
         Top             =   150
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
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
         Index           =   0
         Left            =   270
         TabIndex        =   16
         Top             =   1740
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   270
         TabIndex        =   15
         Top             =   540
         Width           =   795
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1230
         Picture         =   "frmTESReclamaCli.frx":000C
         Top             =   540
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
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
         Left            =   9390
         TabIndex        =   7
         Top             =   1215
         Width           =   765
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
         TabIndex        =   4
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
Attribute VB_Name = "frmTESReclamaCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SaltoLinea = """ + chr(13) + """

Private Const IdPrograma = 608


    
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

Dim VerTodos As Boolean


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
    
    If Index = 0 Then BotonVerTodos True
End Sub


Private Sub cmdCompensar_Click()
    
    cad = DevuelveDesdeBD("informe", "scryst", "codigo", IdPrograma) 'Orden de pago a bancos
    If cad = "" Then
        MsgBox "No esta configurada la aplicación. Falta el informe", vbCritical
        Exit Sub
    End If
    Me.Tag = cad
    
    cad = ""
    RC = ""
    CONT = 0
    TotalRegistros = 0
    NumRegElim = 0
    For I = 1 To Me.lwReclamCli.ListItems.Count
        If Me.lwReclamCli.ListItems(I).Checked Then
            If Trim(lwReclamCli.ListItems(I).SubItems(6)) = "" Then
                'Es un abono
                TotalRegistros = TotalRegistros + 1
            Else
                NumRegElim = NumRegElim + 1
            End If
        End If
        If Me.lwReclamCli.ListItems(I).Bold Then
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
        If Not lwReclamCli.ListItems(CONT).Checked Then
            SQL = "El vencimiento seleccionado no esta marcado"
        
        Else
            'Si el importe Cobro es mayor que abono, deberia estar
            Importe = CCur(txtimpNoEdit(0).Tag) + CCur(txtimpNoEdit(1).Tag)  'txtimpNoEdit(1).Tag es negativo
            If Importe <> 0 Then
                If Importe > 0 Then
                    'Es un abono
                    If Trim(lwReclamCli.ListItems(CONT).SubItems(6)) = "" Then SQL = "cobro"
                Else
                    If Trim(lwReclamCli.ListItems(CONT).SubItems(6)) <> "" Then SQL = "abono"
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
        For I = 1 To Me.lwReclamCli.ListItems.Count - 1
            If Me.lwReclamCli.ListItems(I).Checked Then
                CadAux = Trim(lwReclamCli.ListItems(I).Text & lwReclamCli.ListItems(I).SubItems(1)) & " " & Trim(lwReclamCli.ListItems(I).SubItems(2))
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
        If MsgBox("Seguro que desea realizar la compensación?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
'    End If
    
    
    Me.FrameReclamacionesCliente.Enabled = False
    Me.Refresh
    Screen.MousePointer = vbHourglass
    
    RealizarCompensacionAbonosClientes
    Me.FrameReclamacionesCliente.Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdVtoDestino(Index As Integer)
    
    If Index = 0 Then
        TotalRegistros = 0
        If Not Me.lwReclamCli.SelectedItem Is Nothing Then TotalRegistros = Me.lwReclamCli.SelectedItem.Index
    
    
        For I = 1 To Me.lwReclamCli.ListItems.Count
            If Me.lwReclamCli.ListItems(I).Bold Then
                Me.lwReclamCli.ListItems(I).Bold = False
                Me.lwReclamCli.ListItems(I).ForeColor = vbBlack
                For CONT = 1 To Me.lwReclamCli.ColumnHeaders.Count - 1
                    Me.lwReclamCli.ListItems(I).ListSubItems(CONT).ForeColor = vbBlack
                    Me.lwReclamCli.ListItems(I).ListSubItems(CONT).Bold = False
                Next
            End If
        Next
        Me.Refresh
        
        If TotalRegistros > 0 Then
            I = TotalRegistros
            Me.lwReclamCli.ListItems(I).Bold = True
            Me.lwReclamCli.ListItems(I).ForeColor = vbRed
            For CONT = 1 To Me.lwReclamCli.ColumnHeaders.Count - 1
                Me.lwReclamCli.ListItems(I).ListSubItems(CONT).ForeColor = vbRed
                Me.lwReclamCli.ListItems(I).ListSubItems(CONT).Bold = True
            Next
        End If
        lwReclamCli.Refresh
        
        PonerFocoLw Me.lwReclamCli

    Else
    
'        frmTESReclamaCliImp.pCodigo = Me.lw1.SelectedItem
'        frmTESReclamaCliImp.Show vbModal

    End If
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Not Frame1.Visible Then
            If CadenaDesdeOtroForm <> "" Then
                txtCta(17).Text = CadenaDesdeOtroForm
                txtCta_LostFocus 17
            Else
                PonFoco txtCta(17)
            End If
            CadenaDesdeOtroForm = ""
        End If
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
        .Buttons(2).Image = 1
        .Buttons(4).Image = 7
        .Buttons(5).Image = 8
        
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26
    End With
    
    
    'Limpiamos el tag
    PrimeraVez = True
    CommitConexion  'Porque son listados. No hay nada dentro transaccion
    
        
    h = FrameReclamacionesCliente.Height + 120
    W = FrameReclamacionesCliente.Width
    
    FrameReclamacionesCliente.Visible = False
    Me.Frame1.Visible = True
    
    VerTodos = False
    
    Me.Width = W + 300
    Me.Height = h + 400
    
    Me.cmdCancelar(0).Cancel = True
    
    PonerModoUsuarioGnral 0, "ariconta"
    
    Orden = True
    
'    PonerFrameProgreso

End Sub


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
        Case "Código"
            CampoOrden = "reclama.codigo"
        Case "Fecha"
            CampoOrden = "reclama.fecreclama"
        Case "Cuenta"
            CampoOrden = "reclama.codmacta"
        Case "Nombre"
            CampoOrden = "reclama.nommacta"
        Case "Carta"
            CampoOrden = "reclama.carta"
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

'            frmTESReclamaCliImp.pCodigo = Me.lw1.SelectedItem
'            frmTESReclamaCliImp.Show vbModal


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
        Case 2 ' ver todos
            BotonVerTodos False
        Case 4 'cuenta anterior
            Desplazamiento 0
        Case 5 'cuenta siguiente
            Desplazamiento 1
    End Select
End Sub

Private Sub Desplazamiento(Index As Integer)
    If Data1.Recordset.EOF Then Exit Sub
    
    Select Case Index
        Case 0
            Data1.Recordset.MovePrevious
            If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
            
        Case 1
            Data1.Recordset.MoveNext
            If Data1.Recordset.EOF Then Data1.Recordset.MoveLast
    End Select
    txtCta(17).Text = Data1.Recordset.Fields(0)
    DtxtCta(17).Text = Data1.Recordset.Fields(1)
    PonerModoUsuarioGnral 0, "ariconta"
    PonerVtosCompensacionCliente
End Sub


Private Sub BotonVerTodos(Limpiar As Boolean)
Dim SQL As String
    'Ver todos
    
    VerTodos = True
    
    SQL = "select distinct cobros.codmacta, cuentas.nommacta from cobros inner join cuentas on cobros.codmacta = cuentas.codmacta where (1=1) "
    If Me.Check1(0).Value Then SQL = SQL & " and impvenci + coalesce(gastos,0) - coalesce(impcobro,0) < 0"
    If Limpiar Then SQL = SQL & " and cobros.codmacta is null"
    
    
    
    If TotalRegistrosConsulta(SQL) = 0 Then
        If Not Limpiar Then MsgBox "No hay cuentas con abonos.", vbExclamation
        
        VerTodos = False
    End If
    
    
    Data1.ConnectionString = Conn
    '***** canviar el nom de la PK de la capçalera; repasar codEmpre *************
    Data1.RecordSource = SQL
    Data1.Refresh
    
    If VerTodos Then
        txtCta(17).Text = Data1.Recordset.Fields(0)
        DtxtCta(17).Text = Data1.Recordset.Fields(1)
    Else
        txtCta(17).Text = ""
        DtxtCta(17).Text = ""
    End If
    PonerModoUsuarioGnral 0, "ariconta"
    PonerVtosCompensacionCliente
    
'    CadB = ""
'    CadB1 = ""
'
'    LimpiarCampos
'    CargaGrid 0, False
'    CargaGrid 1, False
''    If chkVistaPrevia.Value = 1 Then
''        MandaBusquedaPrevia ""
''    Else
''        CadenaConsulta = "Select * from " & NombreTabla
''        PonerCadenaBusqueda
''    End If
'
'    HacerBusqueda2
'
    
End Sub



Private Sub BotonAnyadir()

    Frame1.Visible = False
    Frame1.Enabled = False

    Me.FrameReclamacionesCliente.Visible = True
    Me.FrameReclamacionesCliente.Enabled = True
    
    VerTodos = False
    
    PonleFoco txtCta(17)

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
        MsgBox "La cuenta debe ser numérica: " & txtCta(Index).Text, vbExclamation
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


    lwReclamCli.ListItems.Clear
    Me.txtimpNoEdit(0).Tag = 0
    Me.txtimpNoEdit(1).Tag = 0
    Me.txtimpNoEdit(0).Text = ""
    Me.txtimpNoEdit(1).Text = ""
    If Me.txtCta(17).Text = "" Then Exit Sub
    Set Me.lwReclamCli.SmallIcons = frmPpal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    cad = "Select cobros.*,nomforpa from cobros,formapago where cobros.codforpa=formapago.codforpa "
    cad = cad & " AND codmacta = '" & Me.txtCta(17).Text & "'"
    cad = cad & " AND (transfer =0 or transfer is null) and codrem is null"
    cad = cad & " and recedocu=0 and situacion = 0" ' pendientes de cobro
    cad = cad & " ORDER BY fecvenci"
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwReclamCli.ListItems.Add()
        IT.Text = miRsAux!NUmSerie
        IT.SubItems(1) = Format(miRsAux!NumFactu, "0000000")
        IT.SubItems(2) = Format(miRsAux!FecFactu, "dd/mm/yyyy")
        IT.SubItems(3) = miRsAux!NUmORDEN
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
            Screen.MousePointer = vbDefault
'            frmTESReclamaCliImp.pCodigo = CONT
'            frmTESReclamaCliImp.Show vbModal
        End If
        
        Set miRsAux = Nothing
    Else
        MsgBox "Proceso bloqueado", vbExclamation
    End If

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
        If miRsAux!NUmSerie = Me.lwReclamCli.ListItems(NumRegElim).Text Then
            If miRsAux!NumFactu = Val(Me.lwReclamCli.ListItems(NumRegElim).SubItems(1)) Then
                If Format(miRsAux!FecFactu, "dd/mm/yyyy") = Me.lwReclamCli.ListItems(NumRegElim).SubItems(2) Then
                    If miRsAux!NUmORDEN = Val(Me.lwReclamCli.ListItems(NumRegElim).SubItems(3)) Then Destino = 1
                End If
            End If
        End If
        
        RC = "INSERT INTO reclama_facturas (codigo,linea,destino," & cad & ",impvenci,gastos,impcobro,fecvenci) VALUES (" & CONT & "," & I & "," & Destino & "," & DBSet(miRsAux!NUmSerie, "T")
        RC = RC & "," & DBSet(miRsAux!NumFactu, "N") & "," & DBSet(miRsAux!FecFactu, "F") & "," & DBSet(miRsAux!NUmORDEN, "N") & "," & DBSet(miRsAux!ImpVenci, "N")
        RC = RC & "," & DBSet(miRsAux!Gastos, "N") & "," & DBSet(miRsAux!impcobro, "N") & "," & DBSet(miRsAux!FecVenci, "F") & ")"
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
                For J = 1 To Me.lwReclamCli.ListItems.Count
                    If Me.lwReclamCli.ListItems(J).Bold Then
                        'Este es el destino
                        RC = "NUmSerie = '" & Me.lwReclamCli.ListItems(J).Text
                        RC = RC & "' AND numfactu = " & Val(Me.lwReclamCli.ListItems(J).SubItems(1))
                        RC = RC & " AND fecfactu = '" & Format(Me.lwReclamCli.ListItems(J).SubItems(2), FormatoFecha)
                        RC = RC & "' AND numorden = " & Val(Me.lwReclamCli.ListItems(J).SubItems(3))
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

    CadInsert = "insert into cobros_realizados (numserie, numfactu, fecfactu, numorden, numlinea, usuariocobro,fecrealizado,impcobro,numasien) values  "
    
    SQL = "select * from cobros where (numserie, numfactu, fecfactu, numorden) in (" & facturas & ")"
    
    CadValues = ""
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
    
        Importe = DBLet(RS!ImpVenci) + DBLet(RS!Gastos, "N") - DBLet(RS!impcobro, "N")
        
        SQL2 = "select max(numlinea) from cobros_realizados where numserie = " & DBSet(RS!NUmSerie, "T") & " and numfactu = " & DBSet(RS!NumFactu, "N")
        SQL2 = SQL2 & " and fecfactu = " & DBSet(RS!FecFactu, "F") & " and numorden = " & DBSet(RS!NUmORDEN, "N")
        NumLin = DevuelveValor(SQL2)
        NumLin = NumLin + 1
    
        CadValues = CadValues & "(" & DBSet(RS!NUmSerie, "T") & "," & DBSet(RS!NumFactu, "N") & "," & DBSet(RS!FecFactu, "F") & "," & DBSet(RS!NUmORDEN, "N")
        CadValues = CadValues & "," & DBSet(NumLin, "N") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(Now, "FH") & "," & DBSet(Importe, "N") & ",0),"
        
        
        ' actualizamos la cabecera del cobro pq ya no lo eliminamos
        SQL = "update cobros set situacion = 2, impcobro = impvenci + coalesce(gastos,0) where numserie = " & DBSet(RS!NUmSerie, "T")
        SQL = SQL & " and numfactu = " & DBSet(RS!NumFactu, "N") & " and fecfactu = " & DBSet(RS!FecFactu, "F") & " and numorden = " & DBSet(RS!NUmORDEN, "N")
        
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
    For I = 1 To Me.lwReclamCli.ListItems.Count
        If Me.lwReclamCli.ListItems(I).Checked Then
        
            Insertar = True
            If Me.lwReclamCli.ListItems(I).Bold Then
                RegistroDestino = I
                If SinDestino Then Insertar = False
            End If
            If Insertar Then
                SQL = SQL & ", ('" & lwReclamCli.ListItems(I).Text & "'," & lwReclamCli.ListItems(I).SubItems(1)
                SQL = SQL & ",'" & Format(lwReclamCli.ListItems(I).SubItems(2), FormatoFecha) & "'," & lwReclamCli.ListItems(I).SubItems(3) & ")"
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
        Toolbar1.Buttons(2).Enabled = DBLet(RS!Modificar, "N") And (Modo = 2)
        Toolbar1.Buttons(3).Enabled = DBLet(RS!creareliminar, "N") And (Modo = 2)
        
        Toolbar1.Buttons(5).Enabled = False 'DBLet(RS!Ver, "N") And (Modo = 0 Or Modo = 2) And DesdeNorma43 = 0
        Toolbar1.Buttons(6).Enabled = DBLet(RS!Ver, "N")
        
        Toolbar1.Buttons(8).Enabled = DBLet(RS!Imprimir, "N") And Modo = 2
    
        Toolbar2.Buttons(1).Enabled = True 'establecer cta
        Toolbar2.Buttons(2).Enabled = True 'ver todos
        Toolbar2.Buttons(4).Enabled = VerTodos
        Toolbar2.Buttons(5).Enabled = VerTodos
        
    End If
    
    RS.Close
    Set RS = Nothing
    
End Sub



Private Sub CargaList()
Dim IT

    lw1.ListItems.Clear
    Set Me.lw1.SmallIcons = frmPpal.ImgListviews
    Set miRsAux = New ADODB.Recordset
    
    cad = "Select codigo,fecreclama,codmacta,nommacta,carta,importes,observaciones from reclama "
    
    
    If CampoOrden = "" Then CampoOrden = "reclama.fecreclama"
    cad = cad & " ORDER BY " & CampoOrden
    If Orden Then cad = cad & " DESC"
    
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lw1.ListItems.Add()
        IT.Text = Format(miRsAux!Fecreclama, "dd/mm/yyyy")
        IT.SubItems(1) = miRsAux!codmacta
        IT.SubItems(2) = miRsAux!Nommacta
        IT.SubItems(3) = miRsAux!carta
        IT.SubItems(4) = Format(miRsAux!Importes, "###,###,##0.00")
        IT.SubItems(5) = DBLet(miRsAux!observaciones, "T")
        IT.SubItems(6) = miRsAux!Codigo

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

