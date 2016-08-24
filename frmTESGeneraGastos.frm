VERSION 5.00
Begin VB.Form frmTESGeneraGastos 
   Caption         =   "Nuevos gastos fijos"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameNuevo 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3960
         TabIndex        =   16
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdGenerar 
         Appearance      =   0  'Flat
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   2880
         Width           =   975
      End
      Begin VB.Frame Frameintervalo 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   4695
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   2
            Left            =   3840
            TabIndex        =   25
            Top             =   1635
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   21
            Top             =   1635
            Width           =   1215
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1080
            Width           =   2895
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Index           =   7
            Left            =   4440
            TabIndex        =   24
            Top             =   1680
            Width           =   195
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Increme."
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
            Index           =   6
            Left            =   3120
            TabIndex        =   23
            Top             =   1680
            Width           =   765
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Importe"
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
            Index           =   5
            Left            =   240
            TabIndex        =   22
            Top             =   1680
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ejercicio origen"
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
            Index           =   4
            Left            =   240
            TabIndex        =   19
            Top             =   1080
            Width           =   1290
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ejercicio destino"
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
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   1380
         End
      End
      Begin VB.Frame FramePerio 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1095
         Left            =   720
         TabIndex        =   7
         Top             =   1800
         Width           =   3735
         Begin VB.OptionButton Option1 
            Caption         =   "1 "
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "2"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   11
            Top             =   360
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "3 "
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   10
            Top             =   360
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "4 "
            Height          =   255
            Index           =   4
            Left            =   2280
            TabIndex        =   9
            Top             =   360
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "6 "
            Height          =   255
            Index           =   6
            Left            =   3000
            TabIndex        =   8
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Periodicidad en meses"
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
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   0
            Width           =   1890
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkContabilizado 
         Caption         =   "Contabilizado"
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
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
         Index           =   2
         Left            =   2400
         TabIndex        =   6
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha gasto"
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
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1020
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   600
         Picture         =   "frmTESGeneraGastos.frx":0000
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Generar nuevos gastos"
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
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   120
         Width           =   4410
      End
   End
End
Attribute VB_Name = "frmTESGeneraGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '0.- Un unico pago     SI datosotroform<>"" entonces estoy modificando
    '1.- Copiar de otro año
    '2.- Generar desde 0
    
Public Elemento As Integer
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1



Private Sub chkContabilizado_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdCancelar_Click()
    CadenaDesdeOtroForm = ""
    Unload Me
End Sub

Private Sub cmdGenerar_Click()
Dim impo As Currency
Dim F As Date
Dim FMin As Date
Dim cad As String
Dim B As Boolean
Dim I As Integer
Dim Incre As Boolean
Dim FAux As Date

    'Comprobar
    If Opcion <> 1 Then
        If Not EsFechaOK(Text3(0)) Then
            MsgBox "Fecha incorrecta", vbExclamation
            Exit Sub
        End If
        
        F = CDate(Text3(0).Text)
        If F < vParam.fechaini Then
            MsgBox "Fecha menor que ejercicios", vbExclamation
            Exit Sub
        End If
    End If

    If Opcion = 0 Then
        FormatTextImporte Text1(0)
        
        If Trim(Text1(0).Text) = "" Then
            MsgBox "Ponga el importe", vbExclamation
            Exit Sub
        End If
    
        impo = ImporteFormateado(Text1(0).Text)
    End If
    
    
    
    
    If Opcion = 1 Then
        'Copiar de campaña anterior
        'Tendra dos opciones , o poner
        'un importe fijo o incremental
        FormatTextImporte Text1(1)
        FormatTextImporte Text1(2)
        If Not (Text1(1).Text <> "" Xor Text1(2).Text <> "") Then
            MsgBox "Ponga el importe o el incremento", vbExclamation
            Exit Sub
        End If
        Incre = (Text1(2).Text <> "")
            
    End If
    
    
    
    If Opcion > 0 And Opcion < 3 Then
        'Habra que comprobar mas cosas
        If Opcion = 1 Then
            If Combo1(0).ListIndex = -1 Or Combo1(1).ListIndex = -1 Then
                MsgBox "Seleccione periodos", vbExclamation
                Exit Sub
            End If
            
        Else
            '---------------------------------------
            'Si la opcion es periodico
            
        End If
    End If
    
    
    Select Case Opcion
    Case 0
        
        If Not SQLGasto(F, impo, CadenaDesdeOtroForm <> "") Then Exit Sub
                
            
        

    Case 1
        
        Set miRsAux = New ADODB.Recordset
        F = CDate(Mid(Combo1(1).Text, InStr(1, Combo1(1).Text, ":") + 1))
        cad = "Select * from sgastfijd where codigo =" & Elemento & " AND Fecha >='" & Format(F, FormatoFecha)
        F = DateAdd("yyyy", 1, F)
        '                  ojo, solo menor pq es fechainici, no fin de ejercicio
        cad = cad & "' AND Fecha<'" & Format(F, FormatoFecha) & "'"
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        FMin = CDate(Mid(Combo1(0).Text, InStr(1, Combo1(0).Text, ":") + 1))
                
        If miRsAux.EOF Then
            cad = ""
        Else
            While Not miRsAux.EOF
                    F = miRsAux.Fields!Fecha
                    While F < CDate(FMin)
                        F = DateAdd("yyyy", 1, F)
                    Wend
                    
                    If Incre Then
                        impo = ImporteFormateado(Text1(2).Text)
                        impo = (miRsAux!Importe * impo) / 100
                        impo = Round(impo, 2) + miRsAux!Importe
                    Else
                        impo = ImporteFormateado(Text1(1).Text)
                    End If
                    SQLGasto F, impo, False
                
                
                    miRsAux.MoveNext
            Wend
        End If
    Case 2
        'GEneracion de gastos periodicos, a partir de una fecha
        B = False
        If DiasMes(Month(F), Year(F)) = Day(F) Then
            If Day(F) = 31 Then
                B = True
            Else
                cad = "Es le ultimo dia de mes. Quiere que periodifique con fecha ultimo dia de mes?"
                If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then B = True
            End If
        End If
        
        
        'Veremos la perioricidad
        NumRegElim = DevuelvePeriodoGastos
        
        impo = ImporteFormateado(Text1(0).Text)
        
        If Not B Then
            cad = "Va a generar gastos fijos , a partir de : " & Format(F, "dd/mm/yyyy")
            cad = cad & " con frecuencia cada " & NumRegElim & " mes(es)"
            cad = cad & vbCrLf & vbCrLf & "¿Continuar?"
            If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        
        
        'AHORA GENERAMOS LOS PAGOS. Durante un año
        'Comprobamo si ya existen
        FAux = F
        FMin = DateAdd("yyyy", 1, F)
        
        
        Set miRsAux = New ADODB.Recordset
        cad = "Select * from sgastfijd WHERE codigo =" & Elemento
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not miRsAux.EOF Then
            While F < FMin
            
                cad = " fecha ='" & Format(F, FormatoFecha) & "'"
                miRsAux.Find cad, , adSearchForward, 1
                If Not miRsAux.EOF Then
                    MsgBox "Ya existe gastos en esta fecha: " & F, vbExclamation
                    F = CDate("31/12/2199")
                Else
                    If B Then
                        'Ultimo dia de mes. Sumo un mes
                        F = DateAdd("m", NumRegElim, F)
                        F = CDate(DiasMes(Month(F), Year(F)) & "/" & Month(F) & "/" & Year(F))
                    Else
                        F = DateAdd("m", NumRegElim, F)
                    End If
                End If  'mirsaux.eof
            Wend
        End If
        miRsAux.Close
        Set miRsAux = Nothing
        If F = CDate("31/12/2199") Then Exit Sub
        'AHora inserto
        F = FAux
        While F < FMin
            'Inserto el gasto
            SQLGasto F, impo, False
            If B Then
                'Ultimo dia de mes. Sumo un mes
                F = DateAdd("m", NumRegElim, F)
                F = CDate(DiasMes(Month(F), Year(F)) & "/" & Month(F) & "/" & Year(F))
            Else
                F = DateAdd("m", NumRegElim, F)
            End If
        Wend
        
        
        
    End Select
    CadenaDesdeOtroForm = "OK"
    Unload Me
    
End Sub


Private Function SQLGasto(F As Date, Im As Currency, Modificar As Boolean) As Boolean
Dim cad As String

    If Modificar Then
        cad = "UPDATE sgastfijd SET importe=" & TransformaComasPuntos(CStr(Im))
        cad = cad & " , fecha= '" & Format(F, FormatoFecha)
        cad = cad & "' , contabilizado =" & chkContabilizado.Value
        F = CDate(RecuperaValor(CadenaDesdeOtroForm, 1))
        cad = cad & " WHERE codigo =" & Elemento & " AND fecha ='" & Format(F, FormatoFecha) & "'"
    Else
        cad = "INSERT INTO sgastfijd (codigo, fecha, importe) VALUES (" & Elemento
        cad = cad & ",'" & Format(F, FormatoFecha) & "'," & TransformaComasPuntos(CStr(Im)) & ")"
    End If
    SQLGasto = Ejecuta(cad)
    
End Function

Private Function DevuelvePeriodoGastos()
Dim O As OptionButton
    DevuelvePeriodoGastos = 1
    For Each O In Me.Option1
        If O.Value Then
            DevuelvePeriodoGastos = O.Index
            Exit Function
        End If
    Next
End Function


Private Sub Form_Load()

    
    Me.frameNuevo.Visible = True
    Me.Frameintervalo.Visible = Opcion = 1
    Me.FramePerio.Visible = Opcion = 2
    Label2(0).Caption = "Generar nuevos gastos"
    chkContabilizado.Visible = Opcion = 0
    If Opcion = 0 Then
        
        'Text3(0).Enabled = (CadenaDesdeOtroForm = "")
        If CadenaDesdeOtroForm <> "" Then
            Text3(0).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            Text1(0).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            chkContabilizado.Visible = True
            If RecuperaValor(CadenaDesdeOtroForm, 3) = 1 Then
                chkContabilizado.Value = 1
            Else
                chkContabilizado.Value = 0
            End If
            Label2(0).Caption = "Modificar gastos"
        End If
        
    Else
        If Opcion = 1 Then
            
            CargaCombos
            
        End If
    End If
End Sub


Private Sub CargaCombos()
Dim cad As String
Dim F3 As Date
Dim F2 As Date
         
    cad = "select min(fecha) ,max(fecha) from sgastfijd where codigo =" & Elemento
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        
        If Month(miRsAux.Fields(0)) >= Month(vParam.fechaini) Then
            F2 = CDate(Format(vParam.fechaini, "dd/mm/") & Year(miRsAux.Fields(0)))
        Else
            F2 = CDate(Format(vParam.fechaini, "dd/mm/") & Year(miRsAux.Fields(0)) - 1)
        End If
        
        If Month(miRsAux.Fields(1)) > Month(vParam.fechafin) Then
            F3 = CDate(Format(vParam.fechafin, "dd/mm/") & Year(miRsAux.Fields(1)) + 1)
        Else
            F3 = CDate(Format(vParam.fechafin, "dd/mm/") & Year(miRsAux.Fields(1)))
        End If
            
            
        Do
            cad = "Ejercicio : " & Format(F2, "dd/mm/yyyy")
            Combo1(1).AddItem cad
            F2 = DateAdd("yyyy", 1, F2)
        Loop Until F2 > F3
        
        If F2 >= vParam.fechafin Then
            
        Else
            F2 = DateAdd("d", 1, vParam.fechafin)
        End If
        Combo1(0).AddItem "Ejercicio: " & Format(F2, "dd/mm/yyyy")
        Combo1(0).ListIndex = 0
        Combo1(1).ListIndex = Combo1(1).ListCount - 1
        cmdGenerar.Enabled = True
    Else
        cmdGenerar.Enabled = False
    End If
    miRsAux.Close
    Set miRsAux = Nothing
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




Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
   
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    FormatTextImporte Text1(Index)
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


Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub
