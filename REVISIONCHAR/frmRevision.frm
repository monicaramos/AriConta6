VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRevision 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRevision.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameErrorRestore 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   180
      TabIndex        =   10
      Top             =   90
      Visible         =   0   'False
      Width           =   5775
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4215
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   7435
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label Label29 
         Caption         =   "Cambio caracteres"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   210
         Width           =   4515
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   5400
         Picture         =   "frmRevision.frx":6852
         ToolTipText     =   "Quitar seleccion"
         Top             =   720
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   5400
         Picture         =   "frmRevision.frx":699C
         ToolTipText     =   "Todos"
         Top             =   1080
         Width           =   240
      End
   End
   Begin VB.Timer tCuadre 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   6420
      Top             =   5400
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameMultibase 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   60
      TabIndex        =   1
      Top             =   -30
      Width           =   6135
      Begin VB.ComboBox Combo1 
         Height          =   360
         Left            =   1740
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   3540
         Width           =   2595
      End
      Begin VB.CommandButton cmdMultiBase 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   4
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton cmdMultiBase 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   3
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label34 
         Caption         =   "Label34"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   5040
         Width           =   5535
      End
      Begin VB.Label Label33 
         Caption         =   "Base de Datos a revisar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   300
         TabIndex        =   8
         Top             =   3060
         Width           =   4815
      End
      Begin VB.Label Label32 
         Caption         =   "A este proceso le puede costar mucho tiempo."
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
         Left            =   270
         TabIndex        =   7
         Top             =   2430
         Width           =   4815
      End
      Begin VB.Label Label31 
         Caption         =   "No debe trabajar nadie en la base de datos."
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
         Left            =   270
         TabIndex        =   6
         Top             =   1980
         Width           =   4815
      End
      Begin VB.Label Label30 
         Caption         =   "Utilidad para revisar los caracteres especiales que puedan quedar al realizar integraciones o recuperando un backup"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   270
         TabIndex        =   5
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label Label29 
         Caption         =   "Revisión caracteres"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   270
         TabIndex        =   2
         Top             =   120
         Width           =   4935
      End
   End
   Begin VB.Label Label11 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   120
   End
End
Attribute VB_Name = "frmRevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    
Public Parametros As String
    '1.- Vendran empipados: Cuenta, PunteadoD, punteadoH, pdteD,PdteH

Private PrimeraVez As Boolean

Dim I As Integer
Dim SQL As String
Dim RS As Recordset
Dim ItmX As ListItem
Dim Errores As String
Dim NE As Integer
Dim Ok As Integer

Dim Iniciado As Boolean



Private Sub cmdMultiBase_Click(Index As Integer)
Dim I As Integer
    If Index = 1 Then
        If Iniciado Then
            Me.FrameErrorRestore.Visible = False
            Iniciado = False
        Else
            Unload Me
        End If
        Exit Sub
    End If
    
    If Not Iniciado Then
        Me.TreeView1.Nodes.Clear
        Me.FrameErrorRestore.Visible = True
        Iniciado = True
        If Not CargaArbolTablas Then
            Me.FrameErrorRestore.Visible = False
            Iniciado = False
        End If
    Else
        For I = 1 To TreeView1.Nodes.Count
            If TreeView1.Nodes(I).Children = 0 Then
                'Debe seleccionar nodos hijos
                If TreeView1.Nodes(I).Checked Then
                    NE = NE + 1
                    Exit For
                End If
            End If
        Next
    
        If NE = 0 Then
            MsgBox "Seleccione donde se van a realizar los cambios", vbExclamation
            Exit Sub
        End If
    
        'Comprobacion si hay alguien trabajando
        If UsuariosConectados("") Then Exit Sub
        
        SQL = "Seguro que desea continuar con el proceso"
        
        If MsgBox(SQL, vbCritical + vbYesNoCancel) = vbNo Then Exit Sub
    
    
        UpdatearRestoreBakcup_
    
        Screen.MousePointer = vbDefault
        Label34.Caption = ""
        SQL = "Proceso finalizado" & vbCrLf
        SQL = SQL & "Se han actualizado  " & NumRegElim & " columna(s)."
        MsgBox SQL, vbInformation
        
        Me.FrameErrorRestore.Visible = False
        Iniciado = False
    End If
    
    
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
    Screen.MousePointer = vbDefault
End Sub




Private Sub Form_Load()
Dim w, h


    Me.tCuadre.Enabled = False
    PrimeraVez = True
    Iniciado = False
    
    'MULTIBASE
    Me.Caption = "Revisión de caracteres "
    w = Me.frameMultibase.Width
    h = Me.frameMultibase.Height + 300
    Me.frameMultibase.Visible = True
    Label34.Caption = ""
'        txtFecha(0).Text = Format(vParam.fechaini, "dd/mm/yyyy")
'        txtFecha(1).Text = Format(vParam.fechafin, "dd/mm/yyyy")
    cmdMultiBase(1).Cancel = True
        
    CargarCombo
        
        
    Me.Width = w + 120
    Me.Height = h + 120
End Sub


Private Sub CargarCombo()
Dim RS As ADODB.Recordset
Dim SQL As String

    Combo1.Clear

    'Tipo de factura
    Set RS = New ADODB.Recordset
    SQL = "show databases"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    I = 0
    While Not RS.EOF
        If DBLet(RS.Fields(0).Value, "T") <> "information_schema" Then
            Combo1.AddItem RS.Fields(0)
            Combo1.ItemData(Combo1.NewIndex) = I
            I = I + 1
        End If
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing

End Sub


Private Sub imgCheck_Click(Index As Integer)
    For NE = 1 To TreeView1.Nodes.Count
        TreeView1.Nodes(NE).Checked = Index = 1
    Next
    
    Select Case Index
        ' ICONOS VISIBLES EN EL LISTVIEW DEL FRMPPAL
        Case 1 ' marcar todos
            For I = 1 To TreeView1.Nodes.Count
                TreeView1.Nodes(I).Checked = True
            Next I
        Case 0 ' desmarcar todos
            For I = 1 To TreeView1.Nodes.Count
                TreeView1.Nodes(I).Checked = False
            Next I
    End Select
    
End Sub


'Private Sub optMultibas_Click(Index As Integer)
'    FrameErrorRestore.Visible = Me.optMultibas(1).Value
'    If Me.optMultibas(1).Value Then
'        If Me.TreeView1.Nodes.Count = 0 Then CargaArbolTablas
'    End If
'End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim N As Node
    'Si es padre
    If Node.Parent Is Nothing Then
        If Node.Children > 0 Then
            Set N = Node.Child
            Do
                N.Checked = Node.Checked
                Set N = N.Next
            Loop Until N Is Nothing
        End If
    End If
End Sub



'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'Restore desde backup
'
'
Private Function CargaArbolTablas() As Boolean
Dim N As Node
Dim I As Integer
Dim miRsAux As ADODB.Recordset
Dim vSql As String

    On Error GoTo eCargaArbolTablas

    CargaArbolTablas = False
    
    conn.Execute "USE " & Combo1.Text

    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "show tables", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        SQL = miRsAux.Fields(0)
        If LCase(Mid(SQL, 1, 3)) = "tmp" Then SQL = ""
        
        If SQL <> "" Then
            Set N = TreeView1.Nodes.Add(, , miRsAux.Fields(0), miRsAux.Fields(0))
            N.Checked = True
            N.Expanded = True
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    For I = 1 To TreeView1.Nodes.Count
        Label34.Caption = Space(20) & TreeView1.Nodes(I).Text
        Label34.Refresh
        
        vSql = "show columns from " & Combo1.Text & "." & TreeView1.Nodes(I) & " where type like 'varch%' or type like 'text%' "
        
        miRsAux.Open vSql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not miRsAux.EOF
            
            SQL = miRsAux!Field
            If DBLet(miRsAux!Key, "T") <> "" Then
                If DBLet(miRsAux!Key, "T") = "PRI" Then SQL = ""
 
            End If
            
            miRsAux.MoveNext
            
            If SQL <> "" Then
                Set N = TreeView1.Nodes.Add(TreeView1.Nodes(I).Key, tvwChild, , SQL)
                N.Checked = True
                
            End If
                
        Wend
        miRsAux.Close
   Next

    'Quito los que no voy a procesar
   Set N = TreeView1.Nodes(1)
   Set N = N.LastSibling
   While Not (N Is Nothing)
        I = 0
        If N.Children = 0 Then I = N.Index
        If N.Previous Is Nothing Then
            Set N = Nothing
        Else
            Set N = N.Previous
        End If
        If I > 0 Then TreeView1.Nodes.Remove I
    Wend
    
    
    Label34.Caption = ""
    
    CargaArbolTablas = True
    Exit Function
    
eCargaArbolTablas:
    MsgBox "Cargar Tablas en Arbol" & vbCrLf & vbCrLf & Err.Description, vbExclamation
End Function

Private Sub UpdatearRestoreBakcup_()
Dim J As Byte
Dim devuelve As String
Dim T1 As Single

    T1 = Timer
    For NE = 1 To TreeView1.Nodes.Count
        If Not TreeView1.Nodes(NE).Parent Is Nothing Then
            If Timer - T1 > 4 Then
                DoEvents
                Me.Refresh
                T1 = Timer
            End If
            Me.Label34.Caption = TreeView1.Nodes(NE).Parent.Text
            Me.Label34.Refresh
            If TreeView1.Nodes(NE).Checked Then
                NumRegElim = NumRegElim + 1
                For J = 1 To 8
                    CarcateresRestores J, Errores, devuelve
                    SQL = "UPDATE " & Combo1.Text & "." & TreeView1.Nodes(NE).Parent.Text & " SET "
                    SQL = SQL & TreeView1.Nodes(NE) & " = REPLACE(" & TreeView1.Nodes(NE) & ",'" & Errores & "','" & devuelve & "') "
                    If Not EjecutaSQL(SQL) Then Exit Sub
                Next J
            End If
        End If
    Next NE
End Sub

Private Sub CarcateresRestores(Cual As Byte, C1 As String, C2 As String)
    Select Case Cual
    Case 1
        C1 = "Ã‘": C2 = "Ñ"

    Case 2
        C1 = "Ã±": C2 = "ñ"
    Case 3
        C1 = "Ã©": C2 = "é"
    
    Case 4
        C1 = "Ã­": C2 = "í"
    Case 5
        C1 = "Âº": C2 = "º"

    Case 6
        C1 = "Ã³": C2 = "ó"
    Case 7
        C1 = "Â±": C2 = "±"
    Case Else
        C1 = "Ã¡": C2 = "á"
    End Select





    
'
'select domclien,REPLACE(domclien,'Ã‘','Ñ') from sclien
'select domclien,REPLACE(domclien,'Ã±','ñ') from sclien
'select domclien,REPLACE(domclien,'Ã©','é') from sclien
'select domclien,REPLACE(domclien,'Ã­','í') from sclien
'select domclien,REPLACE(domclien,'Âº','º') from sclien
'select domclien,REPLACE(domclien,'Ã³','ó') from sclien
'select domclien,REPLACE(domclien,'Ã¡','á') from sclien
    
End Sub




Public Function UsuariosConectados(vMens As String, Optional DejarContinuar As Boolean) As Boolean
Dim I As Integer
Dim cad As String
Dim metag As String
Dim SQL As String
cad = OtrosPCsContraContabiliad(False)
UsuariosConectados = False
If cad <> "" Then
    UsuariosConectados = True
    I = 1
    metag = vMens
    If vMens <> "" Then metag = metag & vbCrLf
    metag = metag & vbCrLf & "Los siguientes PC's están conectados a: " & vEmpresa.nomEmpre & " (" & ")" & vbCrLf & vbCrLf
    
    Do
        SQL = RecuperaValor(cad, I)
        If SQL <> "" Then
            metag = metag & "    - " & SQL & vbCrLf
            I = I + 1
        End If
    Loop Until SQL = ""
    If DejarContinuar Then
        'Hare la pregunta
        metag = metag & vbCrLf & "¿Continuar?"
        If MsgBox(metag, vbQuestion + vbYesNoCancel) = vbYes Then UsuariosConectados = False
    Else
        'Informa UNICAMENTE
        MsgBox metag, vbExclamation
    End If
End If
End Function


Public Function EjecutaSQL(ByRef SQL As String) As Boolean
    EjecutaSQL = False
    On Error Resume Next
    conn.Execute SQL
    If Err.Number <> 0 Then
        Err.Clear
    Else
        EjecutaSQL = True
    End If
End Function


Public Function OtrosPCsContraContabiliad(EsAlIniciar As Boolean) As String
Dim MiRS As Recordset
Dim cad As String
Dim Equipo As String
Dim EquipoConBD As Boolean

Dim SERVER As String

    On Error GoTo EOtrosPCsContraContabiliad
    
    Set MiRS = New ADODB.Recordset
    
'    SERVER = Servidor
'
'    EquipoConBD = (UCase(vUsu.pc) = UCase(SERVER)) Or (LCase(SERVER) = "localhost")
'
'    Cad = "show processlist"
'    MiRS.Open Cad, conn, adOpenKeyset, adLockOptimistic, adCmdText
'    Cad = ""
'    While Not MiRS.EOF
'        If UCase(MiRS.Fields(3)) = UCase(vUsu.CadenaConexion) Then
'            Equipo = MiRS.Fields(2)
'            'Primero quitamos los dos puntos del puerot
'            NumRegElim = InStr(1, Equipo, ":")
'            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
'
'            'El punto del dominio
'            NumRegElim = InStr(1, Equipo, ".")
'            If NumRegElim > 0 Then Equipo = Mid(Equipo, 1, NumRegElim - 1)
'
'            Equipo = UCase(Equipo)
'
'            If Equipo <> vUsu.pc Then
'
'                    NumRegElim = 0
'                    If Equipo <> "LOCALHOST" Then
'                        'Si no es localhost
'                        NumRegElim = 1
'                    Else
'                        'HAy un proceso de loclahost. Luego, si el equipo no tiene la BD
'                        If Not EquipoConBD Then NumRegElim = 1
'                    End If
'
'                    'Si hay que insertar
'                    If NumRegElim = 1 Then
'                        If InStr(1, Cad, Equipo & "|") = 0 Then Cad = Cad & Equipo & "|"
'                    End If
'            End If
'        End If
'        'Siguiente
'        MiRS.MoveNext
'    Wend
'    NumRegElim = 0
'    MiRS.Close
'    Set MiRS = Nothing
'    OtrosPCsContraContabiliad = Cad
    Exit Function
EOtrosPCsContraContabiliad:
    MuestraError Err.Number, Err.Description, "Leyendo PROCESSLIST"
    Set MiRS = Nothing
    If EsAlIniciar Then
        OtrosPCsContraContabiliad = "LEYENDOPC|"
    Else
        cad = "¿El sistema no puede determinar si hay PCs conectados. ¿Desea continuar igualmente?"
        If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
            OtrosPCsContraContabiliad = ""
        Else
            OtrosPCsContraContabiliad = "USUARIO ACTUAL|"
        End If
    End If
    
    
    
End Function




