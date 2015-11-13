Attribute VB_Name = "libIBAN"
Option Explicit




'A partir de una cuenta banco formateada y todos los numeros juntos (chr(20))
'  devuelve DOS(2) caracteres del IBAN
'  calculados como dice la formula
'  i=ctabanco_con ES... mod 97
'  i = 98-i
' format(i,"00"             'para que copie                     'es lo que devuelve
'
'Puede NO poner pais. Sera ES
Public Function DevuelveIBAN2(PAIS As String, ByVal CtaBancoFormateada As String, DosCaracteresIBAN As String) As Boolean
Dim AUx As String
Dim N As Long
Dim CadenaPais As String
On Error GoTo EDevuelveIBAN
    DevuelveIBAN2 = False
    DosCaracteresIBAN = ""
    
    
    
    If PAIS = "" Then
        PAIS = "ES"
    Else
        If Len(PAIS) <> 2 Then
            PAIS = "ES"
        Else
            PAIS = UCase(PAIS)
        End If
    End If
    
    
    'Ejemplo mio: 20770294901101867914  IBAN: 41
    'Construir el IBAn:
    'A la derecha de la cuenta se pone
    '   el ES00-->   20770294961101915202ES00 ->92
    'Se transforma las letras ES a numeros.
    ' E=14 S=28
    '           ->>> 20770294961101915202 142800
    If PAIS = "ES" Then
        CadenaPais = "1428"
    Else
        N = Asc(Mid(PAIS, 1, 1))
        If N < 65 Or N > 90 Then Err.Raise 513, , "LEtra incorrecta PAIS: " & PAIS
        N = N - 55
        CadenaPais = CStr(N)
        N = Asc(Mid(PAIS, 2, 1))
        If N < 65 Or N > 90 Then Err.Raise 513, , "LEtra incorrecta PAIS: " & PAIS
        N = N - 55
        CadenaPais = CadenaPais & CStr(N)
    End If
    'Se le añaden 2 ceros al final
    CadenaPais = CadenaPais & "00"
    'Esta es la cadena para ES. SiCadenaPais  fuera otro pais es aqui donde hay que cambiar
    CtaBancoFormateada = CtaBancoFormateada & "142800"
    AUx = ""
    While CtaBancoFormateada <> ""
        If Len(CtaBancoFormateada) >= 6 Then
            AUx = AUx & Mid(CtaBancoFormateada, 1, 6)
            CtaBancoFormateada = Mid(CtaBancoFormateada, 7)
        Else
            AUx = AUx & CtaBancoFormateada
            CtaBancoFormateada = ""
        End If
        
        N = CLng(AUx)
        N = N Mod 97
        
        AUx = CStr(N)
    Wend
        
    N = 98 - N
    
    DosCaracteresIBAN = Format(N, "00")
    DevuelveIBAN2 = True
    Exit Function
EDevuelveIBAN:
    CadenaPais = Err.Description
    CadenaPais = Err.Number & "   " & CadenaPais
    MsgBox "Devuelve IBAN. " & vbCrLf & CadenaPais, vbExclamation
    Err.Clear
End Function




Public Function IBAN_Correcto(IBAN As String) As Boolean
Dim AUx As String
    IBAN_Correcto = False
    AUx = ""
    If Len(IBAN) <> 4 Then
        AUx = "Longitud incorrecta"
    Else
        If IsNumeric(Mid(AUx, 3, 2)) Then
            AUx = "Digitos 3 y 4 deben ser numericos"
        Else
            'Podriamos comprobar lista de paises
    
        End If
    End If
    If AUx <> "" Then
        MsgBox AUx, vbExclamation
    Else
        IBAN_Correcto = True
    End If
End Function
