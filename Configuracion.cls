VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "Configuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'#Generador DAVID(Ariadna)



'-------------------------------------------------------
'-------------------------------------------------------
'Declaracion de las variables
'---- Nombre del DNS
Private mvarODBC As String 'Variable local


'Variables de este modulo
Dim NF as Integer
Dim Mens as String


'------------------------------------------------------
'------------------------------------------------------
'      Propiedades del modulo de configuracion
'------------------------------------------------------
'------------------------------------------------------

'---- 
'---- Nombre del DNS
'---- 
Public Property Let ODBC (ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.ODBC = 5
     mvarODBC =vData
End Property

Public Property Get ODBC() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.ODBC
     ODBC = mvarODBC
End Property







'//////////////////////////////////////////////////////
'             LEER  Y  GUARDAR
'//////////////////////////////////////////////////////
'
'
'
'
'
'
'------------------------------------------------------
'     LEER
'
Public Function Leer() As Byte
    '--------------------------------------------------
    '   Leer:
    '   Busca un fichero de nombre Config.ini y lee de él los valores
    '   de parametrización para la aplicación. El fichero se busca
    '   en el directorio desde el que se ejecuta la aplicación.
    '--------------------------------------------------
    Dim Fichero As String ' Nombre completo del fichero a leer
    Dim LinAnt As String
    Dim Lin As String
    Dim NomParam As String ' Para leer el nombre del parámetro
    Dim ValParam As String ' Para leer el valor del parámetro
    Dim NFile As Integer
    
    On Error GoTo Err1
    Fichero = App.Path & "\Config.ini"
    NFile = FreeFile
    Open Fichero For Input As #NFile   ' Abre el archivo.
    Do While Not EOF(NFile) ' Repite el bucle hasta el final del archivo.
        Line Input #NFile, Lin
        Select Case Lin
           CASE "[ODBC]"
               ValParam = ""
               NomParam = Lin

           Case "<-------->"
               Select Case NomParam
                    Case "[ODBC]": mvarODBC = ValParam
                End Select
            Case Else
                If ValParam = "" Then
                    ValParam = Lin
                Else
                    ValParam = ValParam & vbCrLf & Lin
                End If
        End Select
    Loop
    Close #NFile    ' Cierra el archivo.
    Leer = 0
    Exit Function
Err1:
   Leer = 1
End Function




'------------------------------------------------------
'     GUARDAR

Public Function Grabar() As Byte
    '--------------------------------------------------
    '   GuardaConfig:
    '   Guardamos en el fichero Config.ini los valores de los
    '   parámetros de la aplicación
    '--------------------------------------------------
    Dim Fichero As String ' Nombre completo del fichero a leer
    Dim NFile As Integer

    On Error GoTo Err1
    Fichero = App.Path & "\Config.ini"
    NFile = FreeFile
    Open Fichero For Output As #NFile   ' Abre el archivo.
    Print #NFile, "[ODBC]"
    Print #NFile, mvarODBC
    Print #NFile, "<-------->"


    Close #NFile
    Grabar = 0
    Exit Function
Err1:
    Mens = "Error " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbCritical
    Grabar = 1
End Function
