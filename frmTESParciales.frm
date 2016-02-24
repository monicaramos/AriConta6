VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTESParciales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anticipo vto."
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkImprimir 
      Caption         =   "Imprimir Recibo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   17
      Top             =   7020
      Width           =   2685
   End
   Begin VB.CommandButton Command1 
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
      Index           =   1
      Left            =   5850
      TabIndex        =   4
      Top             =   7140
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
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
      Left            =   4530
      TabIndex        =   3
      Top             =   7140
      Width           =   1095
   End
   Begin VB.Frame FrCobro 
      Height          =   6855
      Left            =   60
      TabIndex        =   5
      Top             =   90
      Width           =   6855
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
         ItemData        =   "frmTESParciales.frx":0000
         Left            =   1590
         List            =   "frmTESParciales.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Tag             =   "Tipo de pago|N|N|||formapago|tipforpa|||"
         Top             =   4260
         Width           =   2475
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
         Index           =   1
         Left            =   1680
         TabIndex        =   0
         Text            =   "Text2"
         Top             =   1470
         Width           =   1215
      End
      Begin VB.TextBox txtDescCta 
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
         Index           =   1
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   1470
         Width           =   3495
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
         Index           =   5
         Left            =   4710
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   2910
         Width           =   1755
      End
      Begin VB.TextBox Text2 
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
         Left            =   4710
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   3825
         Width           =   1755
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
         Left            =   4710
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2460
         Width           =   1755
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
         Index           =   3
         Left            =   4710
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1980
         Width           =   1755
      End
      Begin VB.TextBox Text3 
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
         Left            =   1590
         TabIndex        =   1
         Top             =   3825
         Width           =   1305
      End
      Begin MSComctlLib.ListView ListView8 
         Height          =   1455
         Left            =   240
         TabIndex        =   25
         Top             =   5130
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   2566
         View            =   3
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
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase "
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   9
         Left            =   240
         TabIndex        =   27
         Top             =   4290
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cobros realizados: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   8
         Left            =   240
         TabIndex        =   26
         Top             =   4830
         Width           =   1920
      End
      Begin VB.Label Label4 
         Caption         =   "Datos vto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   56
         Left            =   270
         TabIndex        =   19
         Top             =   360
         Width           =   6150
      End
      Begin VB.Label Label4 
         Caption         =   "Datos vto"
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
         Index           =   57
         Left            =   270
         TabIndex        =   18
         Top             =   720
         Width           =   6270
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   6600
         Y1              =   4710
         Y2              =   4710
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   3330
         Width           =   6195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cta banco"
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
         Index           =   7
         Left            =   270
         TabIndex        =   15
         Top             =   1470
         Width           =   1050
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   1380
         Picture         =   "frmTESParciales.frx":0004
         Top             =   1530
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   6
         Left            =   3720
         TabIndex        =   13
         Top             =   3870
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Pagado"
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
         Index           =   5
         Left            =   3780
         TabIndex        =   12
         Top             =   2910
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   6600
         Y1              =   3690
         Y2              =   3690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gastos"
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
         Index           =   4
         Left            =   3840
         TabIndex        =   9
         Top             =   2520
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe TOTAL"
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
         Index           =   2
         Left            =   3060
         TabIndex        =   7
         Top             =   2070
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   3840
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   1230
         Picture         =   "frmTESParciales.frx":6856
         Top             =   3870
         Width           =   240
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox Text1 
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Vencimiento"
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
      Left            =   360
      TabIndex        =   24
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
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
      Index           =   1
      Left            =   360
      TabIndex        =   23
      Top             =   1080
      Width           =   675
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
Dim cad As String
Dim PrimeraVez As Boolean
Dim TipForpa As Integer

Dim LineaCobro As Long

Private Sub PonFoco(ByRef T1 As TextBox)
    T1.SelStart = 0
    T1.SelLength = Len(T1.Text)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub




Private Sub ImprimirRecibo()

    If ImporteFormateado(Text2(0).Text) <= 0 Then
        MsgBox "No se pueden emitir recibos por importes menores o iguales a cero", vbExclamation
        Exit Sub
    End If
    
    frmTESImpRecibo.pImporte = Text2(0).Text
    frmTESImpRecibo.pFechaRec = Text3(0).Text
    frmTESImpRecibo.pFecFactu = RecuperaValor(Vto, 3)
    frmTESImpRecibo.pNumFactu = RecuperaValor(Vto, 2)
    frmTESImpRecibo.pNumSerie = RecuperaValor(Vto, 1)
    frmTESImpRecibo.pNumOrden = RecuperaValor(Vto, 4)
    frmTESImpRecibo.pNumlinea = LineaCobro
    
    frmTESImpRecibo.Show vbModal
    
End Sub

Private Sub Command1_Click(Index As Integer)
Dim B As Boolean
    CadenaDesdeOtroForm = ""
    If Index = 0 Then
        'Comprobamos importes. Y fecha de contabilizacioon
        If Not DatosOK Then Exit Sub
        
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
        
        If chkImprimir.Value = 1 Then ImprimirRecibo
        
    End If
    Unload Me
End Sub

Private Sub CargarListView()
Dim RS As ADODB.Recordset
Dim IT As ListItem
    
    On Error GoTo ECargarlistview
    
    
    ListView8.ColumnHeaders.Clear
    ListView8.ListItems.Clear
    
    
    ListView8.ColumnHeaders.Add , , "Fecha", 1400.2522
    ListView8.ColumnHeaders.Add , , "Usuario", 2000.2522
    ListView8.ColumnHeaders.Add , , "Tipo", 900.2522
    ListView8.ColumnHeaders.Add , , "Importe", 1700.2522, 1
    
    Set RS = New ADODB.Recordset
    
    cad = "select cobros_realizados.fechaent, cobros_realizados.usuariocobro, tipofpago.siglas,  cobros_realizados.impcobro "
    cad = cad & " from cobros_realizados inner join tipofpago on cobros_realizados.tipforpa = tipofpago.tipoformapago "
    cad = cad & " where numserie = " & DBSet(RecuperaValor(Vto, 1), "T")
    cad = cad & " and numfactu = " & DBSet(RecuperaValor(Vto, 2), "N")
    cad = cad & " and fecfactu = " & DBSet(RecuperaValor(Vto, 3), "F")
    cad = cad & " and numorden = " & DBSet(RecuperaValor(Vto, 4), "N")
    cad = cad & " order by numlinea "
    
    RS.Open cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    cad = ""
    While Not RS.EOF
                    
        Set IT = ListView8.ListItems.Add
        
        IT.Text = DBLet(RS.Fields(0))
        IT.SubItems(1) = DBLet(RS.Fields(1))
        IT.SubItems(2) = DBLet(RS.Fields(2))
        IT.SubItems(3) = Format(DBLet(RS.Fields(3)), "###,###,##0.00")
        
        'Siguiente
        RS.MoveNext
    Wend
    NumRegElim = 0
    RS.Close
    Set RS = Nothing
    
    Exit Sub
    
ECargarlistview:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
            
        CargarListView
    
'        TipForpa = DevuelveValor("select tipforpa from formapago where codforpa = " & DBSet(FormaPago, "N"))
        PosicionarCombo Combo1, FormaPago
    
    End If
        
End Sub

Private Sub Form_Load()
        
    Me.Icon = frmPpal.Icon
        
    PrimeraVez = True
        
    If Cobro Then
        Caption = "Cobro"
        Text1(0).Text = RecuperaValor(Vto, 1) & "/" & RecuperaValor(Vto, 2) & "   Fecha: " & RecuperaValor(Vto, 3) & "   Vto. num: " & RecuperaValor(Vto, 4)
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
                
        Label4(57).Caption = Text1(0).Text
        Label4(56).Caption = Text1(2)
        
        
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
    
    CargaCombo
    
    Label4(5).Visible = Cobro
    Text1(5).Visible = Cobro
    Me.Height = Me.FrCobro.Height + 1200 '240 + Me.Command1(0).Height + 240
    'Text2(0).Text = ""
    Text2(0).Text = Format(impo, FormatoImporte)
    Caption = Caption & " de factura"
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
            PonFoco Text2(Index)
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


Private Function DatosOK() As Boolean
Dim Im As Currency

    On Error GoTo EDa
    DatosOK = False
    
    cad = ""
    If Text2(0).Text = "" Then cad = "importe"
    If Text3(0).Text = "" Then cad = cad & " fecha"
    If cad <> "" Then
        MsgBox "Falta: " & cad, vbExclamation
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
    cad = ""
    If impo < 0 Then
        'Importes negativos
        If Im >= 0 Then
            cad = "negativo"
        Else
            If Im < impo Then cad = "X"
        End If
    Else
        If Im <= 0 Then
            cad = "positivo"
        Else
            If Im > impo Then cad = "X"
        End If
    End If
        
    If cad <> "" Then
        
        If cad = "X" Then
            cad = "Importe a pagar mayor que el importe restante.(" & Format(Im, FormatoImporte) & " : " & Format(impo, FormatoImporte) & ")"
        Else
            cad = "El importe debe ser " & cad
        End If
        MsgBox cad, vbExclamation
        Exit Function
    End If
    
        
        
    'Comprobaremos un par de cosillas
    If CuentaBloqeada(RecuperaValor(Cta, 1), CDate(Text3(0).Text), True) Then Exit Function
        
        
        
    DatosOK = True
    Exit Function
EDa:
    MuestraError Err.Number, "Datos Ok"
End Function


Private Function RealizarAnticipo() As Boolean

        Conn.BeginTrans
        If Contabilizar Then
        
        
            Conn.CommitTrans
            RealizarAnticipo = True
            
        Else
            'Conn.RollbackTrans
            TirarAtrasTransaccion
            RealizarAnticipo = False
        End If
End Function


Private Function Contabilizar() As Boolean
Dim Mc As Contadores
Dim FP As Ctipoformapago
Dim Sql As String
Dim Ampliacion As String
Dim Numdocum As String
Dim Conce As Integer
Dim LlevaContr As Boolean
Dim Im As Currency
Dim Debe As Boolean
Dim ElConcepto As Integer
Dim vNumDiari As Integer

    On Error GoTo ECon
    Contabilizar = False
    Set Mc = New Contadores
    If Mc.ConseguirContador("0", CDate(Text3(0).Text) <= vParam.fechafin, True) = 1 Then Exit Function

    Set FP = New Ctipoformapago
    If FP.Leer(Combo1.ListIndex) Then ' antes forma de pago
        Set Mc = Nothing
        Set FP = Nothing
    End If
    
    
    'importe
    impo = ImporteFormateado(Text2(0).Text)
    
    'Inserto cabecera de apunte
    Sql = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari, feccreacion, usucreacion, desdeaplicacion) VALUES ("
    If Cobro Then
        Sql = Sql & FP.diaricli
        vNumDiari = FP.diaricli
    Else
        Sql = Sql & FP.diaripro
        vNumDiari = FP.diaripro
    End If
    Sql = Sql & ",'" & Format(Text3(0).Text, FormatoFecha) & "'," & Mc.Contador
    Sql = Sql & ",'"
    Sql = Sql & "Generado desde Tesorería el " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & DevNombreSQL(vUsu.Nombre)
    If impo < 0 Then Sql = Sql & "  (ABONO)"
    Sql = Sql & "',"
    Sql = Sql & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & ",'ARICONTA 6: Contabilizar Cobros')"
    
    
    Conn.Execute Sql
        
        
        
        
        
    'Inserto en las lineas de apuntes
    Sql = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, "
    Sql = Sql & "codmacta, numdocum, codconce, ampconce,timporteD,"
    Sql = Sql & " timporteH, codccost, ctacontr, idcontab, punteada) VALUES ("
    If Cobro Then
        Sql = Sql & FP.diaricli
    Else
        Sql = Sql & FP.diaripro
    End If
    Sql = Sql & ",'" & Format(Text3(0).Text, FormatoFecha) & "'," & Mc.Contador & ","
    
    
    
    
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
    cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce), "N")
    Ampliacion = cad & " " & Ampliacion
    Ampliacion = Mid(Ampliacion, 1, 35)
    
    
    
    
    'Ahora ponemos linliapu codmacta numdocum codconce ampconce timported timporte codccost ctacontr idcontab punteada
    'Cuenta Cliente/proveedor
    cad = "1,'" & Text1(1).Text & "','" & Numdocum & "'," & Conce & ",'" & DevNombreSQL(Ampliacion) & "',"
    'Importe cobro-pago
    ' nos lo dire "debe"
    If Not Debe Then
        cad = cad & "NULL," & TransformaComasPuntos(CStr(impo))
    Else
        cad = cad & TransformaComasPuntos(CStr(impo)) & ",NULL"
    End If
    'Codccost
    cad = cad & ",NULL,"
    If LlevaContr Then
        cad = cad & "'" & txtCta(1).Text & "'"
    Else
        cad = cad & "NULL"
    End If
    cad = cad & ",'contab',0)"
    cad = Sql & cad
    Conn.Execute cad
    
       
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
    cad = DevuelveDesdeBD("nomconce", "conceptos", "codconce", CStr(Conce), "N")
    Ampliacion = cad & " " & Ampliacion
    Ampliacion = Mid(Ampliacion, 1, 35)
    
    
    
    
    
    
    
    
    
    cad = "2,'" & txtCta(1).Text & "','" & Numdocum & "'," & Conce & ",'" & Ampliacion & "',"
    'Importe cliente
    'Si el cobro/pago va al debe el contrapunte ira al haber
    If Not Debe Then
        'al debe
        cad = cad & TransformaComasPuntos(CStr(impo)) & ",NULL"
    Else
        'al haber
        cad = cad & "NULL," & TransformaComasPuntos(CStr(impo))
    End If
    
    'Codccost
    cad = cad & ",NULL,"
    
    If LlevaContr Then
        cad = cad & "'" & Text1(1).Text & "'"
    Else
        cad = cad & "NULL"
    End If
    cad = cad & ",'idcontab',0)"
    cad = Sql & cad
    Conn.Execute cad
    
    'Insertamos en la temporal para que lo ac
    If Cobro Then
        Sql = FP.diaricli
    Else
        Sql = FP.diaripro
    End If

    InsertaTmpActualizar Mc.Contador, Sql, Text3(0).Text
    
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
        Sql = "cobros"
        Ampliacion = "fecultco"
        Numdocum = "impcobro"
        'El importe es el total. Lo que ya llevaba mas lo de ahora
        If Text1(5).Text <> "" Then impo = impo + ImporteFormateado(Text1(5).Text)
    Else
        
        Sql = "pagos"
        Ampliacion = "fecultpa"
        Numdocum = "imppagad"
        'El importe es el total. Lo que ya llevaba mas lo de ahora
        If Text1(4).Text <> "" Then impo = impo + ImporteFormateado(Text1(4).Text)
    End If
    
    
    '++monica
    If Cobro Then
        Dim NumLin As Long
        
        NumLin = DevuelveValor("select max(numlinea) from cobros_realizados where numserie = " & DBSet(RecuperaValor(Vto, 1), "T") & " AND numfactu=" & DBSet(RecuperaValor(Vto, 2), "N") & " and fecfactu=" & DBSet(RecuperaValor(Vto, 3), "F") & " AND numorden =" & RecuperaValor(Vto, 4))
        NumLin = NumLin + 1
        
        LineaCobro = NumLin
    
        Sql = "insert into cobros_realizados (numserie, numfactu, fecfactu, numorden, numlinea, numdiari, fechaent, "
        Sql = Sql & " numasien, usuariocobro, tipforpa, impcobro, fecrealizado) values (" & DBSet(RecuperaValor(Vto, 1), "T") & ","
        Sql = Sql & DBSet(RecuperaValor(Vto, 2), "N") & "," & DBSet(RecuperaValor(Vto, 3), "F") & ","
        Sql = Sql & DBSet(RecuperaValor(Vto, 4), "N") & "," & DBSet(NumLin, "N") & "," & DBSet(vNumDiari, "N") & ","
        Sql = Sql & DBSet(Text3(0).Text, "F") & "," & DBSet(Mc.Contador, "N") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(Combo1.ItemData(Combo1.ListIndex), "N") & "," & DBSet(Text2(0).Text, "N")
        Sql = Sql & "," & DBSet(Now, "FH") & ")"
    
        Conn.Execute Sql
    
        Sql = "update cobros set impcobro = (select sum(impcobro) from cobros_realizados where numserie = " & DBSet(RecuperaValor(Vto, 1), "T") & " AND numfactu=" & DBSet(RecuperaValor(Vto, 2), "N") & " and fecfactu=" & DBSet(RecuperaValor(Vto, 3), "F") & " AND numorden =" & RecuperaValor(Vto, 4) & ") "
        Sql = Sql & " , fecultco = " & DBSet(Text3(0).Text, "F")
        Sql = Sql & " where numserie = " & DBSet(RecuperaValor(Vto, 1), "T") & " and numfactu = " & DBSet(RecuperaValor(Vto, 2), "N")
        Sql = Sql & " and fecfactu = " & DBSet(RecuperaValor(Vto, 3), "F") & " and numorden = " & DBSet(RecuperaValor(Vto, 4), "N")
    
        Conn.Execute Sql
    
    End If
    
    
    
    
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
        cad = txtCta(Index).Text
        impo = 0
        If cad <> "" Then
            If CuentaCorrectaUltimoNivel(cad, CadenaDesdeOtroForm) Then
                cad = DevuelveDesdeBD("codmacta", "bancos", "codmacta", cad, "T")
                If cad = "" Then
                    CadenaDesdeOtroForm = ""
                    MsgBox "La cuenta contable no esta asociada a ninguna cuenta bancaria", vbExclamation
                End If
            Else
                MsgBox CadenaDesdeOtroForm, vbExclamation
                cad = ""
                CadenaDesdeOtroForm = ""
            End If
            impo = 1
        Else
            CadenaDesdeOtroForm = ""
        End If
        
        
        txtCta(Index).Text = cad
        txtDescCta(Index).Text = CadenaDesdeOtroForm
        If cad = "" And impo <> 0 Then
            PonFoco txtCta(Index)
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
  
      cad = "select nomclien,domclien,pobclien,cpclien,proclien,razosoci,dirdatos,codposta,despobla,desprovi"
      'MAYO 2010
      cad = cad & ",codbanco,codsucur,digcontr,scobro.cuentaba,scobro.codmacta"
      cad = cad & " from scobro,cuentas where scobro.codmacta =cuentas.codmacta and"
      cad = cad & " numserie ='" & RecuperaValor(Me.Vto, 1) & "' and codfaccl=" & RecuperaValor(Me.Vto, 2)
      cad = cad & " and fecfaccl='" & Format(RecuperaValor(Me.Vto, 3), FormatoFecha) & "' and numorden=" & RecuperaValor(Me.Vto, 4)
      miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
      
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
      cad = "1,'" & RecuperaValor(Vto, 1) & "/" & RecuperaValor(Vto, 2) & "'"
      
     
      
      
      'Lugar Vencimiento
      cad = cad & ",'" & Lugar & "'"
      
      'text3 mostrare el codmacta
      'Cad = Cad & ",'" & DevNombreSQL(.SubItems(5)) & "',"
      cad = cad & ",'" & miRsAux!codmacta & "',"
      
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
      cad = cad & Aux & ","
  
      '5 y 6.
      'text5: nomclien
      'texto6:domclien
      If QueDireccionMostrar = 0 Then
          'Cad = Cad & "NULL,NULL"
          'Siempre el nomclien
          'cad = cad & "'" & DevNombreSQL(.SubItems(5)) & "',NULL"
      Else
          If QueDireccionMostrar = 1 Then
              cad = cad & "'" & DevNombreSQL(DBLet(miRsAux!nomclien, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!domclien, "T")) & "'"
          Else
              cad = cad & "'" & DevNombreSQL(DBLet(miRsAux!razosoci, "T")) & "','" & DevNombreSQL(DBLet(miRsAux!dirdatos, "T")) & "'"
          End If
      End If
      
      
      
      'IMPORTES
      '--------------------
      cad = cad & "," & TransformaComasPuntos(CStr(ImporteFormateado(Text2(0).Text)))
      
      'El segundo importe NULL
      cad = cad & ",NULL"
      
      'FECFAS
      '--------------
      'Libramiento o pago     Auqi pone NOW
      cad = cad & ",'" & Format(Text3(0).Text, FormatoFecha) & "'"
      cad = cad & ",'" & Format(RecuperaValor(Vto, 5), FormatoFecha) & "'"
      
      '3era fecha  NULL
      cad = cad & ",NULL"
      
      'OBSERVACIONES
      '------------------
      Aux = EscribeImporteLetra(ImporteFormateado(Text2(0).Text))
      
      Aux = "       ** " & Aux
      cad = cad & ",'" & Aux & "**',"
      
      
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
      cad = cad & Aux
      
      
      
      'OPCION
      '--------------
      cad = cad & ",NULL)"
      
    'El sql completo
    Aux = "INSERT INTO Usuarios.ztesoreriacomun (codusu, codigo, texto1, texto2, texto3, texto4, texto5, "
    Aux = Aux & "texto6, importe1, importe2, fecha1, fecha2, fecha3, observa1, observa2, opcion)"
    Aux = Aux & " VALUES (" & vUsu.Codigo & "," & cad
      
      
    Conn.Execute Aux

    miRsAux.Close
End Sub




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


