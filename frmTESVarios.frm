VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTESVarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "1"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   12855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAgregarCuentas 
      Height          =   6015
      Left            =   0
      TabIndex        =   204
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton cmdInsertaCta 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5400
         TabIndex        =   208
         Top             =   1080
         Width           =   315
      End
      Begin VB.TextBox txtDCtaNormal 
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
         Index           =   5
         Left            =   1560
         TabIndex        =   209
         Text            =   "Text9"
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtCtaNormal 
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
         Left            =   120
         TabIndex        =   207
         Text            =   "Text9"
         Top             =   1080
         Width           =   1365
      End
      Begin VB.CommandButton cmdAceptarCtas 
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
         Left            =   3360
         TabIndex        =   210
         Top             =   5400
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
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
         Index           =   21
         Left            =   4680
         TabIndex        =   212
         Top             =   5400
         Width           =   1095
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   206
         Top             =   1560
         Width           =   5655
      End
      Begin VB.Image imgCtaNorma 
         Height          =   240
         Index           =   6
         Left            =   1230
         ToolTipText     =   "Cuentas agrupadas"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Eliminar"
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
         Index           =   23
         Left            =   600
         TabIndex        =   213
         Top             =   5400
         Width           =   1470
      End
      Begin VB.Image imgEliminarCta 
         Height          =   240
         Left            =   240
         Top             =   5400
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta "
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
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   211
         Top             =   720
         Width           =   735
      End
      Begin VB.Image imgCtaNorma 
         Height          =   240
         Index           =   5
         Left            =   870
         ToolTipText     =   "Cuentas individuales"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "AGREGAR CUENTAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   15
         Left            =   480
         TabIndex        =   205
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame FrameModiRemeTal 
      Height          =   3015
      Left            =   30
      TabIndex        =   257
      Top             =   60
      Width           =   6765
      Begin VB.CommandButton cmdModRemTal 
         Caption         =   "Modificar"
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
         Left            =   4080
         TabIndex        =   260
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
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
         Index           =   25
         Left            =   5280
         TabIndex        =   261
         Top             =   2400
         Width           =   1095
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
         Index           =   8
         Left            =   2040
         TabIndex        =   262
         Text            =   "Text2"
         Top             =   1800
         Width           =   4335
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
         Index           =   8
         Left            =   480
         TabIndex        =   259
         Text            =   "Text2"
         Top             =   1800
         Width           =   1485
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
         Height          =   360
         Index           =   27
         Left            =   480
         TabIndex        =   258
         Text            =   "Text1"
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label Label3 
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   510
         TabIndex        =   303
         Top             =   1530
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   18
         Left            =   510
         TabIndex        =   302
         Top             =   780
         Width           =   840
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Modificar remesa"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   263
         Top             =   240
         Width           =   5295
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   8
         Left            =   1710
         Top             =   1530
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   27
         Left            =   1680
         Top             =   810
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6240
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameDevlucionRe 
      Height          =   7245
      Left            =   2970
      TabIndex        =   81
      Top             =   30
      Width           =   5835
      Begin VB.Frame FrameDevDesdeVto 
         Height          =   1215
         Left            =   120
         TabIndex        =   273
         Top             =   600
         Width           =   5655
         Begin VB.TextBox txtNumero 
            Height          =   285
            Left            =   3960
            TabIndex        =   62
            Text            =   "Text10"
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox txtSerie 
            Height          =   285
            Index           =   4
            Left            =   2160
            TabIndex        =   60
            Text            =   "Text10"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtnumfac 
            Height          =   285
            Index           =   4
            Left            =   2760
            TabIndex        =   61
            Text            =   "Text10"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtDCtaNormal 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   2400
            TabIndex        =   274
            Text            =   "Text9"
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtCtaNormal 
            Height          =   285
            Index           =   11
            Left            =   1080
            TabIndex        =   59
            Text            =   "Text9"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Image imgFra 
            Height          =   255
            Left            =   1800
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "Serie / Fra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   38
            Left            =   600
            TabIndex        =   276
            Top             =   720
            Width           =   1065
         End
         Begin VB.Image imgCtaNorma 
            Height          =   240
            Index           =   11
            Left            =   840
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta"
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
            Height          =   195
            Index           =   37
            Left            =   120
            TabIndex        =   275
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   3840
         TabIndex        =   73
         Text            =   "Text4"
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CheckBox chkAgrupadevol2 
         Caption         =   "Agrupa apunte banco"
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   6360
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   1
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Tag             =   "Ampliacion debe/CLIENTES|N|N|0||stipoformapago|ampdecli|||"
         Top             =   6000
         Width           =   2820
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Tag             =   "Ampliacion debe/CLIENTES|N|N|0||stipoformapago|ampdecli|||"
         Top             =   4920
         Width           =   2820
      End
      Begin VB.TextBox txtConcepto 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   76
         Text            =   "Text10"
         Top             =   5520
         Width           =   495
      End
      Begin VB.TextBox txtDConcpeto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2760
         TabIndex        =   199
         Text            =   "Text9"
         Top             =   5520
         Width           =   2895
      End
      Begin VB.TextBox txtConcepto 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   74
         Text            =   "Text10"
         Top             =   4440
         Width           =   495
      End
      Begin VB.TextBox txtDConcpeto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2760
         TabIndex        =   197
         Text            =   "Text9"
         Top             =   4440
         Width           =   2895
      End
      Begin VB.Frame FrameDevDesdeFichero 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1035
         Left            =   120
         TabIndex        =   148
         Top             =   600
         Width           =   5535
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   150
            TabIndex        =   63
            Text            =   "Text8"
            Top             =   420
            Width           =   5295
         End
         Begin VB.Image Image4 
            Height          =   240
            Left            =   900
            Top             =   150
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fichero"
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
            Index           =   10
            Left            =   120
            TabIndex        =   149
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.CheckBox chkDevolRemesa2 
         Caption         =   "Contabilizar gasto remesa"
         Height          =   255
         Left            =   960
         TabIndex        =   72
         Top             =   3510
         Width           =   2295
      End
      Begin VB.OptionButton optDevRem 
         Caption         =   "% x  rec, con MINIMO"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   69
         Top             =   3015
         Width           =   2175
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   4200
         TabIndex        =   71
         Text            =   "Text4"
         Top             =   3000
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton optDevRem 
         Caption         =   "% x Recibo"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   68
         Top             =   2640
         Width           =   1215
      End
      Begin VB.OptionButton optDevRem 
         Caption         =   "� x Recibo"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   67
         Top             =   2280
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   9
         Left            =   4560
         TabIndex        =   80
         Top             =   6720
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   6
         Left            =   4410
         TabIndex        =   65
         Text            =   "Text3"
         Top             =   915
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   5
         Left            =   2400
         TabIndex        =   64
         Text            =   "Text3"
         Top             =   915
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   1230
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   70
         Text            =   "Text4"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdDevolRem 
         Caption         =   "Devolucion"
         Height          =   375
         Left            =   3360
         TabIndex        =   79
         Top             =   6720
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "EUROS"
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
         Index           =   17
         Left            =   5040
         TabIndex        =   272
         Top             =   3570
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Haber"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   202
         Top             =   5520
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Debe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   201
         Top             =   4440
         Width           =   585
      End
      Begin VB.Label Label7 
         Caption         =   "Concepto"
         Height          =   195
         Index           =   10
         Left            =   960
         TabIndex        =   200
         Top             =   5550
         Width           =   750
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   2
         Left            =   1800
         Top             =   5520
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Concepto"
         Height          =   195
         Index           =   9
         Left            =   960
         TabIndex        =   198
         Top             =   4440
         Width           =   750
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   1
         Left            =   1800
         Top             =   4440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Datos contabilizaci�n"
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
         Index           =   12
         Left            =   120
         TabIndex        =   196
         Top             =   4080
         Width           =   1800
      End
      Begin VB.Line Line1 
         X1              =   2160
         X2              =   5640
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "EUROS"
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
         Index           =   9
         Left            =   5040
         TabIndex        =   147
         Top             =   2400
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Minimo (�)"
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
         Index           =   8
         Left            =   3240
         TabIndex        =   145
         Top             =   3045
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "DEVOLUCION REMESA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   87
         Top             =   120
         Width           =   5175
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "A�o"
         Height          =   255
         Index           =   6
         Left            =   3810
         TabIndex        =   86
         Top             =   930
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   5
         Left            =   1770
         TabIndex        =   85
         Top             =   930
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Remesa"
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
         TabIndex        =   84
         Top             =   960
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         TabIndex        =   83
         Top             =   1920
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   11
         Left            =   960
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gastos "
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
         Left            =   3240
         TabIndex        =   82
         Top             =   2400
         Width           =   630
      End
      Begin VB.Image imgRem 
         Height          =   240
         Index           =   1
         Left            =   1080
         Top             =   937
         Width           =   240
      End
   End
   Begin VB.Frame FrameReclamaEmail 
      Height          =   6975
      Left            =   3450
      TabIndex        =   291
      Top             =   0
      Width           =   10335
      Begin VB.CommandButton cmdEliminarReclama 
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   299
         ToolTipText     =   "Eliminar"
         Top             =   6360
         Width           =   375
      End
      Begin VB.CommandButton cmdReclamas 
         Caption         =   "Continuar"
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
         Left            =   7560
         TabIndex        =   298
         Top             =   6360
         Width           =   1215
      End
      Begin VB.OptionButton optReclama 
         Caption         =   "Correctos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   8760
         TabIndex        =   296
         Top             =   450
         Width           =   1365
      End
      Begin VB.OptionButton optReclama 
         Caption         =   "Sin email"
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
         Index           =   0
         Left            =   7230
         TabIndex        =   295
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
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
         Index           =   31
         Left            =   9000
         TabIndex        =   292
         Top             =   6360
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   5295
         Left            =   240
         TabIndex        =   293
         Top             =   840
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   9340
         View            =   3
         Arrange         =   2
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cuenta"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Email"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   5
         Left            =   600
         Top             =   6360
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   4
         Left            =   240
         Top             =   6360
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Email:"
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
         Height          =   195
         Index           =   44
         Left            =   6300
         TabIndex        =   297
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label15 
         Caption         =   "Email cuentas reclamacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   294
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Frame FrameCobros 
      Height          =   6495
      Left            =   0
      TabIndex        =   13
      Top             =   30
      Width           =   5295
      Begin VB.TextBox txtDCtaNormal 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   1800
         TabIndex        =   264
         Text            =   "Text2"
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox txtCtaNormal 
         Height          =   285
         Index           =   9
         Left            =   480
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Frame FrameCobroEfectivo 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   216
         Top             =   3840
         Width           =   5055
         Begin VB.TextBox txtDescCta 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2880
            TabIndex        =   217
            Text            =   "Text2"
            Top             =   120
            Width           =   2175
         End
         Begin VB.TextBox txtCta 
            Height          =   285
            Index           =   2
            Left            =   1680
            TabIndex        =   6
            Text            =   "Text2"
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "LLeva banco:"
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
            Index           =   26
            Left            =   0
            TabIndex        =   218
            Top             =   120
            Width           =   1335
         End
         Begin VB.Image imgCuentas 
            Height          =   240
            Index           =   2
            Left            =   1440
            Top             =   120
            Width           =   240
         End
      End
      Begin VB.Frame FrameCobroTarjeta 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   214
         Top             =   3360
         Width           =   5055
         Begin VB.TextBox txtImporte 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   5
            Text            =   "Text4"
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Gastos (�)"
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
            Index           =   16
            Left            =   0
            TabIndex        =   215
            Top             =   120
            Width           =   1005
         End
      End
      Begin VB.OptionButton optOrdCob 
         Caption         =   "Nombre cliente"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   8
         Top             =   5160
         Width           =   1575
      End
      Begin VB.OptionButton optOrdCob 
         Caption         =   "Fecha vencimiento"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   10
         Top             =   5520
         Width           =   2055
      End
      Begin VB.OptionButton optOrdCob 
         Caption         =   "Fecha factura"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   9
         Top             =   5520
         Width           =   1455
      End
      Begin VB.OptionButton optOrdCob 
         Caption         =   "Cuenta cliente"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   7
         Top             =   5160
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         ForeColor       =   &H80000015&
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   44
         Text            =   "Text2"
         Top             =   4440
         Width           =   3495
      End
      Begin VB.CommandButton cmdCobro 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2760
         TabIndex        =   11
         Top             =   6000
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtDescCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   2520
         Width           =   3255
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   3960
         TabIndex        =   12
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Image imgCtaNorma 
         Height          =   240
         Index           =   9
         Left            =   840
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
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
         Height          =   195
         Index           =   35
         Left            =   120
         TabIndex        =   265
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "Ordenar efectos"
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
         Height          =   195
         Index           =   22
         Left            =   120
         TabIndex        =   195
         Top             =   4920
         Width           =   1620
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   1560
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "ORDENAR COBROS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   2400
         TabIndex        =   20
         Top             =   720
         Width           =   465
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   3120
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   3960
         TabIndex        =   19
         Top             =   720
         Width           =   420
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   4680
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha vencimiento"
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
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1620
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   840
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha cobro"
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
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   3000
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta bancaria"
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
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   1410
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de pago"
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
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   4440
         Width           =   1140
      End
   End
   Begin VB.Frame FrImprimeRecibos 
      Height          =   7215
      Left            =   0
      TabIndex        =   233
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdIMprime 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3480
         TabIndex        =   231
         Top             =   6600
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   24
         Left            =   4680
         TabIndex        =   232
         Top             =   6600
         Width           =   1095
      End
      Begin VB.TextBox txtFPDesc 
         BackColor       =   &H80000018&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   253
         Text            =   "Text10"
         Top             =   5880
         Width           =   2895
      End
      Begin VB.TextBox txtFPDesc 
         BackColor       =   &H80000018&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   252
         Text            =   "Text10"
         Top             =   5520
         Width           =   2895
      End
      Begin VB.TextBox txtFP 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   230
         Text            =   "Text10"
         Top             =   5880
         Width           =   615
      End
      Begin VB.TextBox txtFP 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   229
         Text            =   "Text10"
         Top             =   5520
         Width           =   615
      End
      Begin VB.TextBox txtDCtaNormal 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2760
         TabIndex        =   247
         Text            =   "Text9"
         Top             =   4275
         Width           =   3015
      End
      Begin VB.TextBox txtCtaNormal 
         Height          =   285
         Index           =   7
         Left            =   1560
         TabIndex        =   227
         Text            =   "Text9"
         Top             =   4275
         Width           =   1095
      End
      Begin VB.TextBox txtDCtaNormal 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2760
         TabIndex        =   246
         Text            =   "Text9"
         Top             =   4680
         Width           =   3015
      End
      Begin VB.TextBox txtCtaNormal 
         Height          =   285
         Index           =   8
         Left            =   1560
         TabIndex        =   228
         Text            =   "Text9"
         Top             =   4680
         Width           =   1095
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Index           =   3
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   223
         Text            =   "Text10"
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Index           =   2
         Left            =   4080
         MaxLength       =   3
         TabIndex        =   224
         Text            =   "Text10"
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox txtnumfac 
         Height          =   285
         Index           =   3
         Left            =   2520
         TabIndex        =   225
         Text            =   "Text10"
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtnumfac 
         Height          =   285
         Index           =   2
         Left            =   4080
         TabIndex        =   226
         Text            =   "Text10"
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   25
         Left            =   2520
         TabIndex        =   219
         Text            =   "Text1"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   24
         Left            =   4080
         TabIndex        =   220
         Text            =   "Text1"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   23
         Left            =   4080
         TabIndex        =   222
         Text            =   "Text1"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   20
         Left            =   2520
         TabIndex        =   221
         Text            =   "Text1"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Image imgFP 
         Height          =   255
         Index           =   1
         Left            =   1800
         Top             =   5880
         Width           =   255
      End
      Begin VB.Image imgFP 
         Height          =   255
         Index           =   0
         Left            =   1800
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Forma pago"
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
         Height          =   195
         Index           =   32
         Left            =   240
         TabIndex        =   256
         Top             =   5280
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   35
         Left            =   1200
         TabIndex        =   255
         Top             =   5520
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   34
         Left            =   1200
         TabIndex        =   254
         Top             =   5925
         Width           =   465
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "IMPRIMIR RECIBOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   9
         Left            =   360
         TabIndex        =   251
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta"
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
         Height          =   195
         Index           =   31
         Left            =   240
         TabIndex        =   250
         Top             =   3960
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   33
         Left            =   720
         TabIndex        =   249
         Top             =   4320
         Width           =   465
      End
      Begin VB.Image imgCtaNorma 
         Height          =   240
         Index           =   8
         Left            =   1320
         Top             =   4680
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   32
         Left            =   720
         TabIndex        =   248
         Top             =   4725
         Width           =   465
      End
      Begin VB.Image imgCtaNorma 
         Height          =   240
         Index           =   7
         Left            =   1320
         Top             =   4320
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Serie"
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
         Height          =   195
         Index           =   30
         Left            =   240
         TabIndex        =   245
         Top             =   2400
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   31
         Left            =   4080
         TabIndex        =   244
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   30
         Left            =   2520
         TabIndex        =   243
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Numero factura"
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
         Height          =   195
         Index           =   29
         Left            =   240
         TabIndex        =   242
         Top             =   3240
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   29
         Left            =   4080
         TabIndex        =   241
         Top             =   3240
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   28
         Left            =   2520
         TabIndex        =   240
         Top             =   3240
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   27
         Left            =   2520
         TabIndex        =   239
         Top             =   840
         Width           =   465
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   25
         Left            =   3240
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   26
         Left            =   4080
         TabIndex        =   238
         Top             =   840
         Width           =   420
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   24
         Left            =   4560
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha vencimiento"
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
         Height          =   195
         Index           =   28
         Left            =   240
         TabIndex        =   237
         Top             =   840
         Width           =   1620
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha factura"
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
         Height          =   195
         Index           =   27
         Left            =   240
         TabIndex        =   236
         Top             =   1680
         Width           =   1620
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   23
         Left            =   4560
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   25
         Left            =   4080
         TabIndex        =   235
         Top             =   1680
         Width           =   420
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   20
         Left            =   3240
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   24
         Left            =   2520
         TabIndex        =   234
         Top             =   1680
         Width           =   465
      End
   End
   Begin VB.Frame frameAcercaDE 
      BorderStyle     =   0  'None
      Height          =   3795
      Left            =   0
      TabIndex        =   92
      Top             =   0
      Width           =   5475
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   11
         Left            =   3960
         TabIndex        =   100
         Top             =   3360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "ARIMONEY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   915
         Left            =   240
         TabIndex        =   95
         Top             =   120
         Width           =   4695
      End
      Begin VB.Image Image2 
         Height          =   540
         Left            =   1740
         Top             =   1800
         Width           =   2880
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   99
         Top             =   1200
         Width           =   3795
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C/ Uruguay 11, Despacho 710"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   120
         TabIndex        =   98
         Top             =   2640
         Width           =   2610
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "46007 - VALENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3240
         TabIndex        =   97
         Top             =   2640
         Width           =   1620
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   81.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1725
         Left            =   3960
         TabIndex        =   96
         Top             =   0
         Width           =   1350
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tfno: 96 380 55 79"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   780
         TabIndex        =   94
         Top             =   3000
         Width           =   1650
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax: 96 342 09 38"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3240
         TabIndex        =   93
         Top             =   3000
         Width           =   1560
      End
   End
   Begin VB.Frame FrameImpagados 
      Height          =   3495
      Left            =   0
      TabIndex        =   88
      Top             =   0
      Width           =   5175
      Begin MSComctlLib.ListView ListView1 
         Height          =   2415
         Left            =   1440
         TabIndex        =   91
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Importe"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cerrar"
         Height          =   375
         Index           =   10
         Left            =   3840
         TabIndex        =   90
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Devoluciones"
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
         Left            =   120
         TabIndex        =   89
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame FrameContabilizarGasto 
      Height          =   3855
      Left            =   0
      TabIndex        =   163
      Top             =   0
      Width           =   7935
      Begin VB.CommandButton cmdContabiliGasto 
         Caption         =   "Contabilizar"
         Height          =   375
         Left            =   5280
         TabIndex        =   178
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtDCC 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   187
         Text            =   "Text9"
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox txtCC 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   172
         Text            =   "Text10"
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtDescCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   4680
         TabIndex        =   184
         Text            =   "Text2"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   6
         Left            =   3360
         TabIndex        =   167
         Text            =   "Text2"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtDConcpeto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   4440
         TabIndex        =   182
         Text            =   "Text9"
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox txtConcepto 
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   169
         Text            =   "Text10"
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   4560
         MaxLength       =   35
         TabIndex        =   171
         Text            =   "Text9"
         Top             =   2520
         Width           =   3135
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   19
         Left            =   6600
         TabIndex        =   179
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtDCtaNormal 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   177
         Text            =   "Text9"
         Top             =   2520
         Width           =   3015
      End
      Begin VB.TextBox txtCtaNormal 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   170
         Text            =   "Text9"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1680
         TabIndex        =   166
         Text            =   "Text4"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtDDiario 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   174
         Text            =   "Text9"
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtDiario 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   168
         Text            =   "Text9"
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   120
         TabIndex        =   165
         Text            =   "Text1"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Image imgCtaNorma 
         Height          =   240
         Index           =   0
         Left            =   720
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgCC 
         Height          =   240
         Index           =   0
         Left            =   1440
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Centro de coste"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   186
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Image imgConcepto 
         Height          =   240
         Index           =   0
         Left            =   4560
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label14 
         Caption         =   "Cuenta banco"
         Height          =   255
         Left            =   3360
         TabIndex        =   185
         Top             =   840
         Width           =   1215
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   6
         Left            =   4560
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Importe"
         Height          =   195
         Index           =   7
         Left            =   1680
         TabIndex        =   183
         Top             =   840
         Width           =   750
      End
      Begin VB.Label Label7 
         Caption         =   "Concepto"
         Height          =   195
         Index           =   6
         Left            =   3840
         TabIndex        =   181
         Top             =   1560
         Width           =   750
      End
      Begin VB.Image imgDiario 
         Height          =   240
         Index           =   0
         Left            =   600
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Ampliaci�n"
         Height          =   195
         Index           =   13
         Left            =   4560
         TabIndex        =   180
         Top             =   2280
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   176
         Top             =   2280
         Width           =   585
      End
      Begin VB.Label Label7 
         Caption         =   "Diario"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   175
         Top             =   1560
         Width           =   750
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   19
         Left            =   720
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   173
         Top             =   840
         Width           =   750
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "CONTABILIZAR GASTO FIJO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   8
         Left            =   1320
         TabIndex        =   164
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame FrameeMPRESAS 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      TabIndex        =   188
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   20
         Left            =   4320
         TabIndex        =   192
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmdEmpresa 
         Caption         =   "Regresar"
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   189
         Top             =   4800
         Width           =   975
      End
      Begin MSComctlLib.ListView lwE 
         Height          =   3615
         Left            =   120
         TabIndex        =   190
         Top             =   840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   6376
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "dsdsd"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Empresas en el sistema"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Index           =   8
         Left            =   120
         TabIndex        =   191
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame FrameSelecGastos 
      Height          =   7335
      Left            =   0
      TabIndex        =   157
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   18
         Left            =   3360
         TabIndex        =   162
         Top             =   6840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdListadoGastos 
         Caption         =   "Seguir"
         Height          =   375
         Left            =   4680
         TabIndex        =   161
         Top             =   6840
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   6135
         Left            =   120
         TabIndex        =   159
         Top             =   600
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   10821
         View            =   3
         Arrange         =   2
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Elemento"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Importe"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgcheckall 
         Height          =   240
         Index           =   1
         Left            =   600
         ToolTipText     =   "Quitar seleccion"
         Top             =   6840
         Width           =   240
      End
      Begin VB.Image imgcheckall 
         Height          =   240
         Index           =   0
         Left            =   240
         ToolTipText     =   "Seleccionar todos"
         Top             =   6840
         Width           =   240
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Left            =   1440
         TabIndex        =   160
         Top             =   6840
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "la la la la"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   7
         Left            =   120
         TabIndex        =   158
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame FrameContabilRem2 
      Height          =   4215
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   5535
      Begin VB.CheckBox chkAgrupaCancelacion 
         Caption         =   "Agrupa cancelacion"
         Height          =   255
         Left            =   240
         TabIndex        =   277
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox txtTexto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   269
         Text            =   "Text3"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtTexto 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   268
         Text            =   "Text3"
         Top             =   1440
         Width           =   3135
      End
      Begin VB.CommandButton cmdContabRemesa 
         Caption         =   "Contabilizar"
         Height          =   375
         Left            =   2880
         TabIndex        =   52
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   51
         Text            =   "Text4"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   1320
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   4
         Left            =   3840
         TabIndex        =   49
         Text            =   "Text3"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   48
         Text            =   "Text3"
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   8
         Left            =   4200
         TabIndex        =   54
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Importe"
         Height          =   255
         Index           =   8
         Left            =   1440
         TabIndex        =   271
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Banco"
         Height          =   255
         Index           =   7
         Left            =   1440
         TabIndex        =   270
         Top             =   1440
         Width           =   495
      End
      Begin VB.Image imgRem 
         Height          =   240
         Index           =   0
         Left            =   1080
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gastos (�)"
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
         Left            =   2760
         TabIndex        =   58
         Top             =   2640
         Width           =   885
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   10
         Left            =   840
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         TabIndex        =   57
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Remesa"
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
         Index           =   21
         Left            =   240
         TabIndex        =   56
         Top             =   960
         Width           =   690
      End
      Begin VB.Label Label6 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   55
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "A�o"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   53
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "CONTABILIZAR REMESA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame FrameElimVtos 
      Height          =   4455
      Left            =   0
      TabIndex        =   101
      Top             =   0
      Width           =   12015
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   120
         TabIndex        =   105
         Top             =   840
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   5318
         View            =   3
         Arrange         =   2
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "A�o"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   1305
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tipo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cuenta"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Banco"
            Object.Width           =   4234
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Importe"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Descripci�n"
            Object.Width           =   4586
         EndProperty
      End
      Begin VB.CommandButton cmdEliminaEfectos 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   9600
         TabIndex        =   104
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   12
         Left            =   10680
         TabIndex        =   103
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "ELIMINAR VENCIMIENTOS REMESADOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   4
         Left            =   240
         TabIndex        =   102
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame Framepagos 
      Height          =   7215
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   5295
      Begin VB.Frame FrameDocPorveedor 
         Height          =   1095
         Left            =   120
         TabIndex        =   278
         Top             =   3840
         Width           =   4815
         Begin VB.TextBox txtTexto 
            Height          =   285
            Index           =   3
            Left            =   1080
            MaxLength       =   100
            TabIndex        =   29
            Text            =   "Text2"
            Top             =   690
            Width           =   3615
         End
         Begin VB.TextBox txtTexto 
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   28
            Text            =   "Text2"
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto"
            Height          =   315
            Index           =   37
            Left            =   240
            TabIndex        =   281
            Top             =   720
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "N�. Doc."
            Height          =   195
            Index           =   36
            Left            =   240
            TabIndex        =   280
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label2 
            Caption         =   "Documento"
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
            Height          =   195
            Index           =   39
            Left            =   0
            TabIndex        =   279
            Top             =   0
            Width           =   1500
         End
      End
      Begin VB.TextBox txtCtaNormal 
         Height          =   285
         Index           =   10
         Left            =   720
         TabIndex        =   25
         Text            =   "Text2"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtDCtaNormal 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   2040
         TabIndex        =   266
         Text            =   "Text2"
         Top             =   1920
         Width           =   3015
      End
      Begin VB.OptionButton optOrdPag 
         Caption         =   "Nombre proveedor"
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   31
         Top             =   5760
         Width           =   1815
      End
      Begin VB.OptionButton optOrdPag 
         Caption         =   "Cuenta proveedor"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   30
         Top             =   5760
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optOrdPag 
         Caption         =   "Fecha factura"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   33
         Top             =   6120
         Width           =   1455
      End
      Begin VB.OptionButton optOrdPag 
         Caption         =   "Fecha vencimiento"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   32
         Top             =   6120
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         ForeColor       =   &H80000015&
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   45
         Text            =   "Text2"
         Top             =   5040
         Width           =   3135
      End
      Begin VB.CommandButton cmdOrdenarPago 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2640
         TabIndex        =   34
         Top             =   6600
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   35
         Top             =   6600
         Width           =   1095
      End
      Begin VB.TextBox txtDescCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   36
         Text            =   "Text2"
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   1560
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   2280
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Proveedor"
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
         Height          =   195
         Index           =   36
         Left            =   120
         TabIndex        =   267
         Top             =   1680
         Width           =   900
      End
      Begin VB.Image imgCtaNorma 
         Height          =   240
         Index           =   10
         Left            =   1080
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Ordenar efectos"
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
         Height          =   195
         Index           =   24
         Left            =   120
         TabIndex        =   203
         Top             =   5520
         Width           =   1620
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   1560
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de pago"
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
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   43
         Top             =   5040
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta bancaria"
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
         Index           =   6
         Left            =   120
         TabIndex        =   42
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha pago"
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
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   41
         Top             =   3360
         Width           =   1020
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   5
         Left            =   1200
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha vencimiento"
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
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   1620
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   4
         Left            =   2880
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   2280
         TabIndex        =   39
         Top             =   960
         Width           =   420
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   3
         Left            =   1440
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   38
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "ORDENAR  PAGOS     "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame FrameElimnaHcoReme 
      Height          =   2535
      Left            =   0
      TabIndex        =   150
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton cmdEliminaHco 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   1440
         TabIndex        =   154
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   17
         Left            =   2640
         TabIndex        =   155
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   1710
         TabIndex        =   151
         Text            =   "Text1"
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label10 
         Height          =   255
         Left            =   120
         TabIndex        =   156
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Eliminar hist�rico de remesas"
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
         Height          =   240
         Index           =   6
         Left            =   240
         TabIndex        =   153
         Top             =   480
         Width           =   3105
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   17
         Left            =   1320
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   11
         Left            =   720
         TabIndex        =   152
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Frame FrameRecaudacionEjecutiva 
      Height          =   7815
      Left            =   0
      TabIndex        =   282
      Top             =   0
      Visible         =   0   'False
      Width           =   12735
      Begin VB.CommandButton cmdRecaudaEjec 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   9600
         TabIndex        =   286
         Top             =   7320
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   29
         Left            =   10680
         TabIndex        =   283
         Top             =   7320
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   6495
         Left            =   120
         TabIndex        =   285
         Top             =   720
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   11456
         View            =   3
         Arrange         =   2
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Serie"
            Object.Width           =   1288
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Factura"
            Object.Width           =   1552
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha factura"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "NumOrden"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "F. Vto"
            Object.Width           =   1854
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Cta"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Nombre"
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "NIF"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "CtaBancaria"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Domicilio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Poblacion"
            Object.Width           =   2187
         EndProperty
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Vencimientos recaudaci�n ejecutiva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   11
         Left            =   3120
         TabIndex        =   284
         Top             =   240
         Width           =   5115
      End
   End
   Begin VB.Frame FrameTransfer 
      Height          =   5895
      Left            =   0
      TabIndex        =   126
      Top             =   0
      Width           =   5295
      Begin VB.CheckBox chkPagoDom 
         Caption         =   "Pago en fecha introducida"
         Height          =   255
         Left            =   2280
         TabIndex        =   132
         Top             =   3480
         Width           =   2655
      End
      Begin VB.TextBox txtDCtaNormal 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2160
         TabIndex        =   300
         Text            =   "Text9"
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtCtaNormal 
         Height          =   285
         Index           =   12
         Left            =   960
         TabIndex        =   129
         Text            =   "Text9"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox cboConcepto 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   133
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   960
         TabIndex        =   134
         Text            =   "Text6"
         Top             =   4320
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   2280
         TabIndex        =   128
         Text            =   "Text1"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   960
         TabIndex        =   127
         Text            =   "Text1"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   960
         TabIndex        =   131
         Text            =   "Text1"
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtCta 
         Height          =   285
         Index           =   5
         Left            =   960
         TabIndex        =   130
         Text            =   "Text2"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtDescCta 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2400
         TabIndex        =   137
         Text            =   "Text2"
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   15
         Left            =   3960
         TabIndex        =   136
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton cmdTr 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2790
         TabIndex        =   135
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta "
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
         Index           =   45
         Left            =   120
         TabIndex        =   301
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Image imgCtaNorma 
         Height          =   240
         Index           =   12
         Left            =   1560
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   5
         Left            =   1560
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Concepto trans."
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
         Height          =   195
         Index           =   43
         Left            =   2400
         TabIndex        =   290
         Top             =   3240
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion"
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
         Index           =   15
         Left            =   240
         TabIndex        =   144
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   143
         Top             =   120
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   11
         Left            =   960
         TabIndex        =   142
         Top             =   960
         Width           =   465
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   16
         Left            =   2880
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   10
         Left            =   2280
         TabIndex        =   141
         Top             =   960
         Width           =   420
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   15
         Left            =   1560
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha vencimiento"
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
         Height          =   195
         Index           =   18
         Left            =   120
         TabIndex        =   140
         Top             =   720
         Width           =   1620
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   14
         Left            =   1320
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha pago"
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
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   139
         Top             =   3240
         Width           =   1020
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   4
         Left            =   1560
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta bancaria"
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
         Index           =   16
         Left            =   120
         TabIndex        =   138
         Top             =   2280
         Width           =   1815
      End
   End
   Begin VB.Frame FrameDeuda 
      Height          =   7335
      Left            =   0
      TabIndex        =   106
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton cmdCancelarDeuda 
         Caption         =   "CANCELAR"
         Height          =   375
         Left            =   5040
         TabIndex        =   146
         Top             =   6840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame FrameDH_cta 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   240
         TabIndex        =   122
         Top             =   600
         Width           =   6135
         Begin VB.TextBox txtCtaNormal 
            Height          =   285
            Index           =   2
            Left            =   1680
            TabIndex        =   110
            Text            =   "Text9"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtDCtaNormal 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2880
            TabIndex        =   194
            Text            =   "Text9"
            Top             =   600
            Width           =   3015
         End
         Begin VB.TextBox txtCtaNormal 
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   109
            Text            =   "Text9"
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox txtDCtaNormal 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2880
            TabIndex        =   193
            Text            =   "Text9"
            Top             =   120
            Width           =   3015
         End
         Begin VB.Image imgCtaNorma 
            Height          =   240
            Index           =   2
            Left            =   1320
            Top             =   600
            Width           =   240
         End
         Begin VB.Image imgCtaNorma 
            Height          =   240
            Index           =   1
            Left            =   1320
            Top             =   120
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   9
            Left            =   720
            TabIndex        =   125
            Top             =   600
            Width           =   420
         End
         Begin VB.Label Label1 
            Caption         =   "Desde"
            Height          =   195
            Index           =   8
            Left            =   720
            TabIndex        =   124
            Top             =   120
            Width           =   465
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta "
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
            Index           =   14
            Left            =   0
            TabIndex        =   123
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   3960
         TabIndex        =   112
         Text            =   "Text1"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   2400
         TabIndex        =   111
         Text            =   "Text1"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdPorNIF 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   114
         Top             =   6840
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   13
         Left            =   5040
         TabIndex        =   115
         Top             =   6840
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   116
         Text            =   "Text5"
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   240
         TabIndex        =   108
         Text            =   "Text4"
         Top             =   1080
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1815
         Left            =   2400
         TabIndex        =   113
         Top             =   4800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3201
         View            =   3
         Arrange         =   2
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView lwtipopago 
         Height          =   2295
         Left            =   2400
         TabIndex        =   289
         Top             =   2400
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   4048
         View            =   3
         Arrange         =   2
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   2
         Left            =   1680
         Top             =   4800
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   3
         Left            =   2040
         Top             =   4800
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   1680
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   2040
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Tipos de pago"
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
         Height          =   195
         Index           =   42
         Left            =   240
         TabIndex        =   288
         Top             =   2400
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "Empresas"
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
         Height          =   195
         Index           =   41
         Left            =   360
         TabIndex        =   287
         Top             =   4800
         Width           =   3060
      End
      Begin VB.Label Label2 
         Caption         =   "NIF"
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
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   121
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   195
         Index           =   7
         Left            =   2400
         TabIndex        =   120
         Top             =   1680
         Width           =   465
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   13
         Left            =   4560
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   6
         Left            =   3960
         TabIndex        =   119
         Top             =   1680
         Width           =   420
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   12
         Left            =   3000
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fechas vencimientos"
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
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   118
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   960
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   6960
         Width           =   4095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "DEUDA por NIF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   5
         Left            =   960
         TabIndex        =   107
         Top             =   120
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmTESVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Opcion As Byte
    '
    '0 .- Pedir datos para ordenar cobros
    
    '3.- Reclamaciones
    '4.- Remesas
    
    
    '5.- Pregunta numero TALON pagare
    
    'Cambio situacion remesa
    '----------------------------
    '6.-  De A a B.   Generar banco
            
    '8.- Contabilizar remesa
        
    '9.- Devolucion remesa
        
    '10.- Mostrar vencimientos impagdos

    '11.- ACERCA DE
        
    '12  - Eliminar vtos
    
    '13.- Deuda total consolidada
    '14.-   "         ""      pero desde hasta
        
        
    '15.- Realizar transferencias
        
    '16.- Devolucion remesa desde fichero del banco
    '--------------------------------
    
    
    '17.- Eliminar informacion HCO remesas
    
    '18.- Selecci�n de gastos para el listado de tesoreria
    
    '19.- Contabilizar gastos
    
    '20.- Seleccion de empresas disponibles, para el usuario
    
    
    '21- Listado pagos (cobros donde se indican las cuentas que quiero que apar
    
    
    'Mas sobre remesas.
    '22.- Cancelacion cliente
    '23.- Confirmacion remesa
    
    
    
    '24.- Impresion de todos los tipos de recibos
    
    '25.- Cambiar banco y/o fecha vto para la remesa de talon pagare
    
    '28 .- Devolucion remesa desde un vto
    
    
    '29 .- Recaudacion ejecutiva
    
    
    '31 .- Reclamaciones por email.
            'Tendra los que tienen email
    
    
Public SubTipo As Byte

    'Para la opcion 22
    '   Remesas cancelacion cliente.
    '       1:  Efectos
    '       2: Talones pagares
    
'Febrero 2010
'Cuando pago proveedores con un talon, y le he indicado el numero
Public NumeroDocumento As String
    
    
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBa As frmBanco
Attribute frmBa.VB_VarHelpID = -1

Private WithEvents frmCCtas As frmColCtas
Attribute frmCCtas.VB_VarHelpID = -1
Private WithEvents frmB As frmBasico
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmP As frmFormaPago
Attribute frmP.VB_VarHelpID = -1


Dim RS As ADODB.Recordset
Dim SQL As String
Dim I As Integer
Dim IT As ListItem  'Comun
Dim PrimeraVez As Boolean
Dim Cancelado As Boolean
Dim CuentasCC As String





Private Sub chkAgrupadevol_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboConcepto_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cboTipoRemesa_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub chkComensaAbonos_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkDevolRemesa2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkPagoDom_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmbReferencia_KeyPress(KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub cmbRemesa_Click()
    'Si es talon o pagare pido el banco YA
    'Me.FrameBancoRemesa.Visible = cmbRemesa.ListIndex > 0
End Sub

Private Sub cmbRemesa_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptarCtas_Click()
    If List1.ListCount = 0 Then
        MsgBox "Introduzca cuentas", vbExclamation
        Exit Sub
    End If
    
    'Cargo en CadenaDesdeOtroForm las cuentas empipadas
    CuentasCC = ""
    For I = 0 To List1.ListCount - 1
        CuentasCC = CuentasCC & Mid(List1.List(I), 1, vEmpresa.DigitosUltimoNivel) & "|"
    Next I

    CadenaDesdeOtroForm = CuentasCC
    Unload Me

End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 21 Or Index = 25 Or Index = 31 Then CadenaDesdeOtroForm = "" 'ME garantizo =""
    If Index = 31 Then
        If MsgBox("�Cancelar el proceso?", vbQuestion + vbYesNo) = vbYes Then SubTipo = 0
    End If
    Unload Me
End Sub



Private Sub cmdCancelarDeuda_Click()
    Cancelado = True
End Sub




Private Sub cmdCobro_Click()
Dim cad As String
Dim Importe As Currency

    'Algunas conideraciones
    'Fecha pago tiene k tener valor
    If Text1(2).Text = "" Then
        MsgBox "Fecha de pago debe tener valor", vbExclamation
        PonerFoco Text1(2)
        Exit Sub
    End If
    
    
    
    'VEMOS SI LA FECHA ESTA DENTRO DEL EJERCICIO
    If FechaCorrecta2(CDate(Text1(2).Text), False) > 1 Then
        MsgBox "Fecha cobro fuera de ejercicios", vbExclamation
        PonerFoco Text1(2)
        Exit Sub
    End If
    
    If txtCta(0).Text = "" Then
        MsgBox "Seleccione la cuenta contable asociada al banco", vbExclamation
        PonerFoco txtCta(0)
        Exit Sub
    End If
    
    
    
    Importe = 0
    If SubTipo = 6 Then
        If txtImporte(4).Text <> "" Then
            If InStr(1, txtImporte(4).Text, ",") > 0 Then
                Importe = ImporteFormateado(txtImporte(4).Text)
            Else
                Importe = CCur(TransformaPuntosComas(txtImporte(4).Text))
            End If
        End If
    End If
    If vParamT.IntereseCobrosTarjeta > 0 Then
        If Importe < 0 Or Importe >= 100 Then
            MsgBox "Intereses cobro tarjeta. Valor entre 0..100", vbExclamation
            PonerFoco Me.txtImporte(4)
            Exit Sub
            
        End If
        
        'Solo dejaremos IR cliente a cliente
        If Me.txtCtaNormal(9).Text = "" And Importe > 0 Then
            MsgBox "Seleccione una cuenta cliente", vbExclamation
            PonerFoco Me.txtCtaNormal(9)
            Exit Sub
        End If
    End If
    
    
    If SubTipo = 0 Then
        If txtCta(2).Text <> "" Then Importe = CCur(txtCta(2).Text)
    End If

'
    'Llegados a este punto montaremos el sql
    SQL = ""
    
    If Text1(0).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " scobro.fecvenci >= '" & Format(Text1(0).Text, FormatoFecha) & "'"
    End If
        
        
    If Text1(1).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " scobro.fecvenci <= '" & Format(Text1(1).Text, FormatoFecha) & "'"
    End If
    
    
    
    'Forma de pago
    If SQL <> "" Then SQL = SQL & " AND "
    SQL = SQL & " sforpa.tipforpa = " & SubTipo



    If Me.txtCtaNormal(9).Text <> "" Then
        'Los de un cliente solamente
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " scobro.codmacta = '" & txtCtaNormal(9).Text & "'"
    End If

    'Si son talones o pagares, NO deben estar remesados
    If SubTipo = vbTalon Or SubTipo = vbPagare Then
        SQL = SQL & " AND (codrem is null )"
    End If

    'Para contabilizar transferecias efectuadas por los cobros.
    'NO LAS QUE HAGAMOS COMO ABONOS'    If SubTipo = 1 Then
'        SQL = SQL & " AND impvenci >0 "
'    End If

    
    Screen.MousePointer = vbHourglass
    cad = " FROM scobro,sforpa WHERE scobro.codforpa = sforpa.codforpa AND "
    'Hacemos un conteo
    Set RS = New ADODB.Recordset
    I = 0
    RS.Open "SELECT Count(*) " & cad & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        I = DBLet(RS.Fields(0), "N")
    End If
    RS.Close
    Set RS = Nothing
    
    
    'Si es talon o pagare vere si esta en parametros lo de contabiliza contra cuenta puente
    'Si es asi, avisare sobre la forma correcta de contabilizacion
    If I > 0 Then
        If SubTipo = vbTalon Or SubTipo = vbPagare Then
            If SubTipo = vbTalon Then
                cad = "contatalonpte"
            Else
                cad = "contapagarepte"
            End If
            cad = DevuelveDesdeBD(cad, "paramtesor", "codigo", 1)
            If cad = "" Then cad = "0"
            If Val(cad) > 0 Then
                cad = "La forma de contabilizar pagar�s / talones es mediante la opci�n de remesas (talones,pagar�s)" & vbCrLf
                cad = cad & "�Desea continuar con la contabilizaci�n?"
                If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then I = -1 'Para que no haga nada(ni mostrar el msg de no hay registros
            End If
            
            
            'Si esta recibido el documento NO dejo contabilizarlo
            SQL = SQL & " AND scobro.recedocu = 0"
            
            
        End If
    End If
    
    
    
    
    If I <= 0 Then
        If I = 0 Then MsgBox "Ning�n dato con esos valores.", vbExclamation
    
    Else
        'La ordenacion de los efectos
        If optOrdCob(1).Value Then
            I = 1
        ElseIf optOrdCob(2).Value Then
            I = 2
        ElseIf optOrdCob(3).Value Then
            I = 3
        Else
            I = 0
        End If
        'Hay datos, abriremos el forumalrio para k seleccione
        'los pagos que queremos hacer
        If BloqueoManual(True, "ORDECOBRO", CStr(SubTipo)) Then
            
            With frmTESVerCobrosPagos
                .ImporteGastosTarjeta_ = Importe
                .OrdenacionEfectos = CByte(I)
                .vSql = SQL
                .OrdenarEfecto = True
                .Regresar = False
                .ContabTransfer = False
                .Cobros = True
                .Tipo = SubTipo
                .SegundoParametro = ""
                'Los textos
                .vTextos = Text1(2).Text & "|" & Me.txtCta(0).Text & " - " & Me.txtDescCta(0).Text & "|" & SubTipo & "|"
                
                'Marzo2013   Cobramos un solo cliente
                'Aparecera un boton para traer todos los cobros
                .CodmactaUnica = Trim(txtCtaNormal(9).Text)
                
                .Show vbModal
            End With
            BloqueoManual False, "ORDECOBRO", ""
            'Memorizo la ordenacion
            LeerGuardarOrdenacion False, True, I
        Else
            MsgBox "Proceso bloqueado por otro usuario", vbExclamation
        End If
        
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdContabiliGasto_Click()
    
    'COmprobaciones
    
    If Text1(19).Text = "" Or txtCta(6).Text = "" Or txtImporte(3).Text = "" Or _
        txtDiario(0).Text = "" Or txtCtaNormal(0).Text = "" Or txtConcepto(0).Text = "" Then
            MsgBox "Campos vacios. Todos los campos son obligatorios", vbExclamation
            Exit Sub
    End If
    
    If txtCC(0).Visible Then
        If txtCC(0).Text = "" Then
            MsgBox "Centro de coste obligatorio", vbExclamation
            Exit Sub
        End If
    End If
    
     
    'OK. Contabilizamos
    '---------------------------------------------
    
    'Borro primero la tmp
    SQL = "DELETE FROM tmpactualizar where codusu =" & vUsu.Codigo
    If Not Ejecuta(SQL) Then Exit Sub

    
    Conn.BeginTrans
    
    If ContabilizarGastoFijo Then
        Conn.CommitTrans
        '-----------------------------------------------------------
        'Ahora actualizamos los registros que estan en tmpactualziar
        frmTESActualizar.OpcionActualizar = 20
        frmTESActualizar.Show vbModal
        Unload Me
    Else
        TirarAtrasTransaccion
    End If
    
    
End Sub

Private Sub cmdContabRemesa_Click()
Dim B As Boolean
Dim Importe As Currency
Dim CC As String
Dim Opt As Byte
Dim AgrupaCance As Boolean
Dim ContabilizacionEspecialNorma19 As Boolean


'Dim ImporteEnRecepcion As Currency
'Dim TalonPagareBeneficios As String
    SQL = ""
    If Text3(3).Text = "" Or Text3(4).Text = "" Then
        SQL = "Ponga una remesa."
    Else
        If Not IsNumeric(Text3(3).Text) Or Not IsNumeric(Text3(4).Text) Then SQL = "La remesa debe ser num�rica"
    End If
    
    If Text1(10).Text = "" Then SQL = "Ponga la fecha de abono"
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    'Fecha pertenece a ejercicios contbles
    If FechaCorrecta2(CDate(Text1(10).Text), True) > 1 Then Exit Sub
    
    
    'Ahora miramos la remesa. En que sitaucion , y de que tipo es
    SQL = "Select * from remesas where codigo =" & Text3(3).Text
    SQL = SQL & " AND anyo =" & Text3(4).Text
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If RS.EOF Then
        MsgBox "Ninguna remesa con esos valores", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Sub

    End If
    
    'Tiene valor
    SQL = ""
    B = AdelanteConLaRemesa()
    ContabilizacionEspecialNorma19 = False
    If B Then
        'Si es norma19 y tiene le parametro de contabilizacion por fecha comprobaremos la fecha de los vtos
        If Opcion = 8 Then
        
            'Se podrian agrupar los IFs, pero asi de momento me entero mas
        
            'Para RECIBOS BANCARIOS SOLO
            If DBLet(RS!Tiporem, "N") = 1 Then
                If vParamT.Norma19xFechaVto Then
                    If Not IsNull(RS!Tipo) Then
                        If RS!Tipo = 0 Then
                            'NORMA 19
                            'Contbiliza por fecha VTO
                            'Comprobaremos que toooodos estan en fecha ejercicio
                            SQL = ComprobacionFechasRemesaN19PorVto
                            If SQL <> "" Then SQL = "-Comprobando fechas remesas N19" & vbCrLf & SQL
                            
                            
                            If txtImporte(0).Text <> "" Then SQL = SQL & vbCrLf & "N19 no permite gastos bancario"
                            
                            
                            If SQL <> "" Then
                                B = False
                            Else
                                ContabilizacionEspecialNorma19 = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
    End If

    If Not B Then
        If SQL = "" Then SQL = "Error y punto"
        RS.Close
        Set RS = Nothing
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    SQL = "Select scobro.codmacta,nommacta,fecbloq from scobro,cuentas where scobro.codmacta = cuentas.codmacta"
    SQL = SQL & " and  codrem =" & Text3(3).Text
    SQL = SQL & " AND anyorem =" & Text3(4).Text
    SQL = SQL & " AND fecbloq <='" & Format(Text1(10).Text, FormatoFecha) & "' GROUP BY 1"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not miRsAux.EOF
        SQL = SQL & miRsAux!codmacta & Space(10) & miRsAux!FecBloq & Space(10) & miRsAux!Nommacta & vbCrLf
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If SQL <> "" Then
        CC = "Cuenta          Fec. bloqueo           Nombre" & vbCrLf & String(80, "-") & vbCrLf
        CC = "Cuentas bloqueadas" & vbCrLf & vbCrLf & CC & SQL
        MsgBox CC, vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Sub
    End If
       
       
       
    'Bloqueariamos la opcion de modificar esa remesa
        
        Importe = TextoAimporte(txtImporte(0).Text)
  
        'Tiene gastos. Falta ver si tiene la cuenta de gastos configurada. ASi como
        'si es analitica, el CC asociado
        CC = ""
        If vParam.autocoste Then CC = "codccost"
            
        SQL = DevuelveDesdeBD("ctagastos", "ctabancaria", "codmacta", RS!codmacta, "T", CC)
        If SQL = "" Then
            MsgBox "Falta configurar la cuenta de gastos del banco:" & RS!codmacta, vbExclamation
            Set RS = Nothing
            Exit Sub
        End If
        
        If vParam.autocoste Then
            If CC = "" Then
                MsgBox "Necesita asignar centro de coste a la cuenta de gastos del banco: " & RS!codmacta, vbExclamation
                Set RS = Nothing
                Exit Sub
            End If
        End If
        
        SQL = SQL & "|" & CC & "|"
        
        
        'A�ado, si tiene, la cuenta de ingresos
        CC = DevuelveDesdeBD("ctaingreso", "ctabancaria", "codmacta", RS!codmacta, "T")
        If CC = "" Then
            If Importe > 0 Then
                MsgBox "Falta configurar la cuenta de ingresos del banco:" & RS!codmacta, vbExclamation
                Set RS = Nothing
                Exit Sub
            End If
        End If
        
        SQL = SQL & CC & "|"   'La
        

    SQL = RS!codmacta & "|" & SQL
    
    
    'Contab. remesa. Si es talon/pagare vamos a comprobar si hay diferencias entre el importe del documento
    'y el total de lineas
    B = False    'Si ya se ha hecho la pregunta no la volveremos a repetir
    'TalonPagareBeneficios = ""    'Solo para TAL/PAG y si hay importe beneficios etc

    
    'Pregunta conbilizacion
    If Not B Then   'Si no hemos hecho la pregunta en otro sitio la hacemos ahora
        Select Case Opcion
        Case 8
            CC = "Va a abonar"
        Case 22
            CC = "Procede a realizar la cancelacion del cliente de"
        Case 23
            CC = "Procede a realizar la confirmacion de"
        End Select
        CC = CC & " la remesa: " & RS!Codigo & " / " & RS!Anyo & vbCrLf & vbCrLf
        CC = CC & Space(30) & "�Continuar?"
        If SubTipo = 2 Then
            If Val(RS!Tiporem) = 3 Then
                CC = "Tal�n" & vbCrLf & CC
            Else
                CC = "Pagar�" & vbCrLf & CC
            End If
            CC = "Tipo: " & CC
        End If
    
        If MsgBox(CC, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    'Para llevarlos a hco
    Conn.Execute "DELETE from tmpactualizar  where codusu =" & vUsu.Codigo
    
        
    
    If Opcion = 8 Then
        'CONTABILIZACION    ABONO REMESA
        
        'NORMA 19
        '------------------------------------
        
        'Contabilizaremos la remesa
        Conn.BeginTrans
        
        'mayo 2012
        If ContabilizacionEspecialNorma19 Then
            'Utiliza Morales
            'Es para contabilizar los recibos por fecha de vto
            
            B = ContabNorma19PorFechaVto(RS!Codigo, RS!Anyo, SQL)
        Else
            'Toooodas las demas opciones estan aqui
        
                                    'Efecto(1),pagare(2),talon(3)
            B = ContabilizarRecordsetRemesa(RS!Tiporem, DBLet(RS!Tipo, "N") = 0, RS!Codigo, RS!Anyo, SQL, CDate(Text1(10).Text), Importe)
        
        End If
        
        'si se contabiliza entonces updateo y la pongo en
        'situacion Q. Contabilizada a falta de devueltos ,
        If B Then
            Conn.CommitTrans
            'AQUI updateamos el registro pq es una tabla myisam
            'y no debemos meterla en la transaccion
            SQL = "UPDATE remesas SET"
            SQL = SQL & " situacion= 'Q'"
            SQL = SQL & " WHERE codigo=" & RS!Codigo
            SQL = SQL & " and anyo=" & RS!Anyo

            If Not Ejecuta(SQL) Then MsgBox "Error actualizando tabla remesa.", vbExclamation
            
            
            'Ahora actualizamos los registros que estan en tmpactualziar
            frmTESActualizar.OpcionActualizar = 20
            frmTESActualizar.Show vbModal
            Screen.MousePointer = vbDefault
            'Cerramos
            RS.Close
            Unload Me
            Exit Sub
        Else
            'ANtes
            'Conn.RollbackTrans
            'Ahora
            TirarAtrasTransaccion
        End If
    
    
    Else
        Conn.BeginTrans
      
        'Cancelacion /confirmacion cliente
        If SubTipo = 1 Then
            'EFECTOS
            If Opcion <= 23 Then
            
                'YA NO EXISTE CONFIRMACION REMESA
                Opt = Opcion - 22 '0.Cancelar   1.Confirmar
                AgrupaCance = False
                If Me.chkAgrupaCancelacion.Visible Then
                    If Me.chkAgrupaCancelacion.Value = 1 Then AgrupaCance = True
                End If
                
                'para la 23 NO deberiamos llegar. Ese proceso lo hemos eliminado
                If Opt = 0 Then
                    B = RemesasCancelacionEfectos(RS!Codigo, RS!Anyo, SQL, CDate(Text1(10).Text), Importe, AgrupaCance)
                Else
                    B = False
                    MsgBox " NO deberia haber entrado con confirmacion remesas", vbExclamation
                End If
            Else
                B = False
                MsgBox "Opcion incorrecta (>23)", vbExclamation
            End If
            
        Else
            MsgBox "AHora no deberia estar aqui!!!!!", vbExclamation
            
                                 '
            'B = RemesasCancelacionTALONPAGARE(Val(Rs!tiporem) = 3, Rs!Codigo, Rs!Anyo, SQL, CDate(Text1(10).Text), Importe)
        End If
        If B Then
            Conn.CommitTrans
            
            
            'Ahora actualizamos los registros que estan en tmpactualziar
            frmTESActualizar.OpcionActualizar = 20
            frmTESActualizar.Show vbModal
            Screen.MousePointer = vbDefault
            'Cerramos
            RS.Close
            Unload Me
            Exit Sub
            
        Else
            TirarAtrasTransaccion
        End If
        
    End If
    
    
    
    RS.Close
    Set RS = Nothing
    Screen.MousePointer = vbDefault
End Sub




Private Function AdelanteConLaRemesa() As Boolean
Dim C As String

    AdelanteConLaRemesa = False
    SQL = ""
    
    'Efectos eliminados
    If RS!Situacion = "Z" Or RS!Situacion = "Y" Then SQL = "Efectos eliminados"
    
    'abierta sin llevar a banco. Esto solo es valido para las de efectos
    If SubTipo = 1 Then
        If RS!Situacion = "A" Then SQL = "Remesa abierta. Sin llevar al banco."
    
    End If
    'Ya contabilizada
    If RS!Situacion = "Q" Then SQL = "Remesa abonada."
    
    If SQL <> "" Then Exit Function
    
    
    
    
    If Opcion = 8 Then
        'COntbilizar / abonar remesa
        '---------------------------------------------------------------------------
        If SubTipo = 1 Then
            'Febrero 2009
            'Ahora toooodas las remesas se hace lo mismmo
            ' De llevada a banco a cancelar cliente. De cancelar a abonar y de abonar a eliminar. NO
            'hay distinciones entre remesas. Para podrer abonar una remesa esta tiene que estar cancelada
            
        Else
            If RS!Tiporem = 2 And vParamT.PagaresCtaPuente Then
                If RS!Situacion <> "F" Then SQL = "La remesa NO puede abonarse. Falta cancelaci�n "
            End If
            
            If RS!Tiporem = 3 And vParamT.TalonesCtaPuente Then
                If RS!Situacion <> "F" Then SQL = "La remesa NO puede abonarse. Falta cancelaci�n "
            End If
        End If
        
            
    Else
       'Vamos a proceder al proceso de generacion cancelacion  /* CANCELACION */
       If SubTipo = 1 Then
            'Para los efectos la norma no tiene que ser 19
            'Febrero 2009.  Para tooodas las normas
            'If Rs!Tipo = 0 Then
            '    SQL = "Proceso no v�lido para NORMA 19"
            '    Exit Function
            'End If
        
       End If
       
       'Para elos tipos 1,2
       If Opcion = 22 Then
            'Cancelacion cliente
            'Para los efectos, tiene que estar generado soporte. Para talones/pagares no es obligado
            If SubTipo = 1 Then
                If RS!Situacion <> "B" Then SQL = "Para cancelar la remesa deberia esta en situaci�n 'Soporte generado'"
            Else
                If RS!Situacion = "F" Then SQL = "Remesa YA cancelada"
            End If
        Else
            'Febrero 2009
            'No hay confirmacion
            SQL = "Opci�n de confirmacion NO es v�lida"
            'Confirmacion
            'If Rs!situacion <> "F" Then SQL = "Para confirmar la remesa esta deberia estar 'Cancelacion cliente'"
       End If
       
       
       'Si hasta aqui esta bien:
       'Compruebo que tiene configurado en parametros
       If SQL = "" Then
            'Comprobamos si esta bien configurada
            '
            If SubTipo = 1 Then
                    If Opcion = 22 Then
                        'SQL = "4310"
                        SQL = "RemesaCancelacion"
                    Else
                        SQL = "RemesaConfirmacion"
                    End If
                    SQL = DevuelveDesdeBD(SQL, "paramtesor", "codigo", "1")
                    If SQL = "" Then
                        SQL = "Falta configurar par�metros cuentas confirmaci�n/cancelaci�n remesa. "
                    Else
                        'OK. Esta configurado
                        SQL = ""
                    End If
                    
            Else
                'talones pagares
                'Veremos si esta configurado(y bien configurado) para el proceso
                If RS!Tiporem = 2 Then
                    'Pagare
                    C = "contapagarepte"
                ElseIf RS!Tiporem = 3 Then
                    'Talones
                    C = "contatalonpte"
                Else
                    'NO DEBIA HABERSE METIDO AQUI
                    C = ""
                    
                End If
                If C = "" Then
                    SQL = "Error validando tipo de remesa"
                    
                Else
                    C = DevuelveDesdeBD(C, "paramtesor", "codigo", 1)
                    If C = "" Then C = "0"
                    If Val(C) = 0 Then
                        SQL = "Falta configurar la aplicacion para las remesas de talones / pagares"
                    Else
                        SQL = ""
                    End If
                End If
            End If
       End If
    End If
    AdelanteConLaRemesa = SQL = ""
    
End Function







Private Sub cmdDevolRem_Click()
Dim Importe As Currency
Dim GastoDevolGral As Currency
Dim CadenaVencimiento As String
Dim MultiRemesaDevuelta As String
Dim TipoFicheroDevolucion As Byte

    MultiRemesaDevuelta = ""
    CadenaVencimiento = ""
    If Opcion = 16 Then
        'DESDE FICHERO
        Text8.Text = Trim(Text8.Text)
        If Text8.Text = "" Then Exit Sub
        If Dir(Text8.Text, vbArchive) = "" Then
            MsgBox "El fichero: " & Text8.Text & "    NO existe", vbExclamation
            Exit Sub
        End If
        Text3(5).Text = ""
        Text3(6).Text = ""
        
        'Si que existe el fichero
        TipoFicheroDevolucion = EsFicheroDevolucionSEPA2(Text8.Text)
        If TipoFicheroDevolucion > 0 Then
            If TipoFicheroDevolucion = 2 Then
                'SEPA xml
                ProcesaFicheroDevolucionSEPA_XML Text8, SQL
            Else
                ProcesaCabeceraFicheroDevolucionSEPA Text8, SQL
            End If
        Else
            'Texto normal
            ProcesaCabeceraFicheroDevolucion Text8.Text, SQL
        End If
        If SQL = "" Then Exit Sub
        
        
    
        
        MultiRemesaDevuelta = SQL
        Text3(5).Text = RecuperaValor(SQL, 1)
        Text3(6).Text = RecuperaValor(SQL, 2)
        
    End If
    If Opcion = 28 Then
        
        If txtSerie(4).Text = "" Or txtSerie(4).Text = "" Then
            MsgBox "Indique el numero de factura", vbExclamation
            Exit Sub
        End If
    
        'Desde el Vto
        Set RS = New ADODB.Recordset
        
        SQL = ""
        If Me.txtCtaNormal(11).Text <> "" Then SQL = SQL & " AND codmacta='" & Me.txtCtaNormal(11).Text & "'"
        If txtSerie(4).Text <> "" Then SQL = SQL & " AND numserie = '" & txtSerie(4).Text & "'"
        If txtNumFac(4).Text <> "" Then SQL = SQL & " AND codfaccl = " & txtNumFac(4).Text
        If txtNumero.Text <> "" Then SQL = SQL & " AND numorden = " & txtNumero.Text
        SQL = Mid(SQL, 5)
        
        
        SQL = "Select codrem,anyorem,NUmSerie,codfaccl,numorden from scobro where " & SQL
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RS.EOF Then
            SQL = "Ninguna pertence a ninguna remesa "
            MsgBox SQL, vbExclamation
            RS.Close
            Exit Sub
        End If
        Text3(5).Text = DBLet(RS!CodRem, "T")
        Text3(6).Text = DBLet(RS!AnyoRem, "T")
        CadenaVencimiento = RS!NUmSerie & "|" & RS!codfaccl & "|" & RS!numorden & "|"
        RS.Close
        Set RS = Nothing
    End If
    
    
    SQL = ""
    If Text3(5).Text = "" Or Text3(6).Text = "" Then
        If Opcion = 9 Then
            SQL = "Ponga una remesa."
        Else
            SQL = "ERROR leyendo remesa"
        End If
    Else
        If Not IsNumeric(Text3(5).Text) Or Not IsNumeric(Text3(6).Text) Then SQL = "La remesa debe ser num�rica"
    End If
    
    If Text1(11).Text = "" Then SQL = "Ponga la fecha de abono"
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    'Fecha pertenece a ejercicios contbles
    If FechaCorrecta2(CDate(Text1(11).Text), True) > 1 Then Exit Sub
    
    
    If txtImporte(1).Text = "" Then
        MsgBox "Indique el gasto por recibo", vbExclamation
        Exit Sub
    End If
    '
    If Me.optDevRem(2).Value Then
        If (txtImporte(2).Text = "") Then
            MsgBox "Debe poner valores del  minimo", vbExclamation
            Exit Sub
        End If
        
    End If
    
    If txtImporte(1).Text <> "" Then
        'Hay gravamen por gastos
        'Bloqueariamos la opcion de modificar esa remesa
        Importe = TextoAimporte(txtImporte(1).Text)
        If Me.optDevRem(1).Value Or Me.optDevRem(2).Value Then
            'Porcentual. No puede ser superior al 100%
            If Importe > 100 Then
                MsgBox "Importe no puede ser superior al 100%", vbExclamation
                Exit Sub
            End If
        End If
        
    Else
        Importe = 0
    End If
    
    'Comprobamos los conceptos y ampliaciones
    SQL = ""
    If txtConcepto(1).Text <> "" Then
        If txtDConcpeto(1).Text = "" Then SQL = "Concepto cliente"
    End If
    If txtConcepto(2).Text <> "" Then
        If txtDConcpeto(2).Text = "" Then SQL = "Concepto banco"
    End If
    
    
    If SQL = "" Then
        If Combo2(0).ListIndex = -1 Or Combo2(1).ListIndex = -1 Then SQL = "Ampliacion concepto incorrecta"
    End If
    
    If SQL <> "" Then
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
    
    'Nuevo Noviembre 2009
    GastoDevolGral = 0
    If Me.chkDevolRemesa2.Value = 1 Then
        'Ha puesto gasto devolucion pero NO indica el gasto
        GastoDevolGral = TextoAimporte(txtImporte(5).Text)
        If GastoDevolGral = 0 Then
            MsgBox "Ha marcado contabilizar gasto y no lo ha indicado", vbExclamation
            Exit Sub
        End If
    
    Else
        If txtImporte(5).Text <> "" Then
            MsgBox "Ha indicado el gasto pero no ha marcado contabilizarlo", vbExclamation
            Exit Sub
        End If
    End If
    'Ahora miramos la remesa. En que sitaucion , y de que tipo es
    SQL = "Select * from remesas where codigo =" & Text3(5).Text
    SQL = SQL & " AND anyo =" & Text3(6).Text
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        SQL = "Ninguna remesa con esos valores."
        If Opcion = 16 Then SQL = SQL & "  Remesa: " & Text3(5).Text & " / " & Text3(6).Text
        MsgBox SQL, vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Sub
    End If
    
    
    'Tiene valor
    If RS!Situacion = "A" Then
        MsgBox "Remesa abierta. Sin llevar al banco.", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Sub
    End If
    
    
    
    If Asc(RS!Situacion) < Asc("Q") Then
        MsgBox "Remesa sin contabilizar.", vbExclamation
        RS.Close
        Set RS = Nothing
        Exit Sub
    End If
    
    
    
    
    SQL = RS!Codigo & "|" & RS!Anyo & "|" & RS!codmacta & "|" & Text1(11).Text & "|"
    
    
    Importe = TextoAimporte(txtImporte(1).Text)   ''Levara el gasto por recibo
    If Me.optDevRem(1).Value Or Me.optDevRem(2).Value Then SQL = SQL & "%"
    SQL = SQL & "|"
    If Me.optDevRem(2).Value Then SQL = SQL & TextoAimporte(txtImporte(2).Text)
    SQL = SQL & "|"
    
    
    'SQL llevara hasta ahora
    '        remes    cta ban  fec contb tipo gasto el 1: si tiene valor es el minimo por recibo
    ' Ej:    1|2009|572000005|20/11/2009|%|1|
    
    
    'Si contabilizamos el gasto, o pro contra vendra como factura bancaria desde otro lugar(norma34 p.e.)
    If GastoDevolGral = 0 Then
        'NO HAY GASTO
        SQL = SQL & "0|"
    Else
        SQL = SQL & CStr(GastoDevolGral) & "|"
        If Me.chkDevolRemesa2.Value = 1 Then
            'Voy a contabi�izar los gastos.
            'Vere si tiene CC
            If vParam.autocoste Then
                If DevuelveDesdeBD("codccost", "ctabancaria", "codmacta", RS!codmacta, "T") = "" Then
                    MsgBox "Va a contabilizar los gastos pero no esta configurado el Centro de coste para el banco: " & RS!codmacta, vbExclamation
                    RS.Close
                    Set RS = Nothing
                    Exit Sub
                End If
            End If
        End If
    End If
    
    'Depues del gasto
    'A�adire el fichero, si es autmatico
    If Opcion = 16 Then SQL = SQL & Text8.Text
    SQL = SQL & "|"
    'Nov 2012. En las devoluciones puede ser que el fichero traiga mas de una devolucion
    If Opcion = 16 Then
        If Text8.Text <> "" Then
            'Tengo que subsituir | por #
            MultiRemesaDevuelta = Replace(MultiRemesaDevuelta, "|", "#")
            SQL = SQL & MultiRemesaDevuelta
        End If
    End If
    SQL = SQL & "|"
    

    
    'Cierro aqui
    RS.Close
    
    'Bloqueamos la devolucion
    BloqueoManual True, "Devolrem", vUsu.Codigo
    'Hacemos la devolucion
    frmTESRemesas.Opcion = 2
    frmTESRemesas.vRemesa = SQL
    frmTESRemesas.ImporteRemesa = Importe 'Utilizamos esta variable para indicar el gasto a cargar por recibo
    
    '28Marzo2007
    'Para la contabilizacion de la devolucion
    'Client
    SQL = txtConcepto(1).Text & "|" & Combo2(0).ListIndex & "|"
    'y el banco
    SQL = SQL & txtConcepto(2).Text & "|" & Combo2(1).ListIndex & "|"
    'Noviembre 2009
    'Agrupa el apunte del banco
    SQL = SQL & Abs(chkAgrupadevol2.Value) & "|"
    
    
    
    frmTESRemesas.ValoresDevolucionRemesa = SQL
    'Si es desde el vto, para que lo busque
    frmTESRemesas.vSql = CadenaVencimiento
    
    frmTESRemesas.Show vbModal

    'Desbloqueamos
    BloqueoManual False, "Devolrem", vUsu.Codigo

End Sub

'Private Function EsFicheroSEPA() As Boolean
'
'    On Error GoTo eEsFicheroSEPA
'
'    EsFicheroSEPA = False
'
'
'End Function


Private Sub cmdEliminaEfectos_Click()
Dim Byt As Byte
Dim Forpa As Ctipoformapago
Dim Agrupar As Boolean
Dim Seguir As Boolean
Dim CtaPuente As Boolean

    SQL = ""
    For I = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(I).Checked Then SQL = SQL & "1"
    Next I
    
    If SQL = "" Then
        MsgBox "Seleccione alguna remesa para eliminar los vencimientos", vbExclamation
        Exit Sub
    End If
    
    
    'Comprobar que hay efectos
    Set miRsAux = New ADODB.Recordset
    Byt = 0
    If Not ComprobarEfectosBorrar Then Byt = 1
    Set miRsAux = Nothing
    If Byt = 1 Then
        MsgBox "No se puede borrar ningun efecto", vbExclamation
        Exit Sub
    End If
    
    
    
    
    'Para ofertar los valores por defecto
    For I = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(I).Checked Then
            'Cogere la primera forma de pago
            If ListView2.ListItems(I).Tag = 2 Then
                Byt = vbPagare
            ElseIf ListView2.ListItems(I).Tag = 3 Then
                Byt = vbTalon
            Else
                Byt = vbTipoPagoRemesa
            End If
            Exit For
        End If
    Next I
    Set Forpa = New Ctipoformapago
    Forpa.Leer CInt(Byt)
    
    
    If Byt = vbPagare Then
        'Sobre talones
        CtaPuente = vParamT.PagaresCtaPuente
    ElseIf Byt = vbTalon Then
        CtaPuente = vParamT.TalonesCtaPuente
    Else
        'Efectos. Viene de cancelacion
    End If
    
    
    
    If CtaPuente Then
            SQL = Forpa.diaricli & "|" & Forpa.condecli & "|" & Forpa.conhacli & "|"
            frmTESPedirConceptos.Intercambio = SQL
            frmTESPedirConceptos.Opcion = 0 'Eliminar efectos
            frmTESPedirConceptos.Show vbModal
            
            If CadenaDesdeOtroForm = "" Then Exit Sub
            Forpa.diaricli = RecuperaValor(CadenaDesdeOtroForm, 1)
            Forpa.condecli = RecuperaValor(CadenaDesdeOtroForm, 2)
            Forpa.conhacli = RecuperaValor(CadenaDesdeOtroForm, 3)
            Agrupar = RecuperaValor(CadenaDesdeOtroForm, 4) = "1"
    Else
        'No lleva apunte. Con preguntar sobra
        If MsgBox("Eliminar efectos?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    'Llegados aqui borraremos cada una de las remesas seleccionadas
    For I = ListView2.ListItems.Count To 1 Step -1
        If ListView2.ListItems(I).Checked Then
        

            'Elimino tmpactualizar
            SQL = "DELETE FROM tmpactualizar where codusu =" & vUsu.Codigo
            Ejecuta SQL

        
            Set IT = ListView2.ListItems(I)
                        
                Seguir = True
                If IT.Tag > 1 Then
                    'Comprobamos que si el importe del talon no coincide con la linea

                    If DiferenciaEnImportes(CInt(IT.Index)) Then
                        If Not ComprobarTodosVencidos(CInt(IT.Index)) Then Seguir = False
                    End If
                End If
                    
                'Haremos una apunte de cancelacion
                ' 5208 contra
               
                If Seguir Then
                    Conn.BeginTrans
                    If IT.Tag = 1 Then
                        'RECIBOS
                        Byt = RemesasEliminarVtosRem2(ListView2.ListItems(I).SubItems(1), ListView2.ListItems(I).Text, Now, Forpa, Agrupar)
                    Else
                        'TALONES PAGARES
                        Byt = RemesasEliminarVtosTalonesPagares(IT.Tag, ListView2.ListItems(I).SubItems(1), ListView2.ListItems(I).Text, Now, Forpa, Agrupar)
                    End If
                    If Byt < 2 Then
                        Conn.CommitTrans
                        If Byt = 1 Then
                            frmTESActualizar.OpcionActualizar = 20
                            frmTESActualizar.Show vbModal
                        End If
                    Else
                        TirarAtrasTransaccion
                    End If
            
                    If Byt < 2 Then ListView2.ListItems.Remove I
                End If
        End If
    Next I
    Screen.MousePointer = vbDefault
End Sub








Private Sub cmdEliminaHco_Click()
        
    If Text1(17).Text = "" Then
        MsgBox "Fecha de pago debe tener valor", vbExclamation
        PonerFoco Text1(17)
        Exit Sub
    End If
    
    
    'Comprobaciones
    Set RS = New ADODB.Recordset
    SQL = "Select count(*) from remesas where fecremesa <='" & Format(Text1(17).Text, FormatoFecha) & "' AND tiporem "
    'Tipo remesa
    If SubTipo = vbTipoPagoRemesa Then
        SQL = SQL & " = 1 " 'EFECTOS
    Else
        SQL = SQL & " <> 1 " 'Talones y pagares
    End If
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If Not RS.EOF Then I = DBLet(RS.Fields(0), "N")
    RS.Close
    If I = 0 Then
        MsgBox "Ninguna remesa anterior a la fecha seleccionada", vbExclamation
        Exit Sub
    End If
    
    
    RS.Open SQL & " AND (situacion<'Y' or situacion=NULL)", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If Not RS.EOF Then I = DBLet(RS.Fields(0), "N")
    RS.Close
    
    If I <> 0 Then
        MsgBox "Hay rememesas que no se pueden eliminar", vbExclamation
        Exit Sub
    End If
    
    'Comprobare que hay remesas en situacion Y
    ' y NO tienen vencimientos, y las updateare a Z
    '------------------------------------------------
    SQL = Replace(SQL, "count(*)", "*")
    RS.Open SQL & " AND situacion='Y' ", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CuentasCC = ""
    While Not RS.EOF
        CuentasCC = CuentasCC & "codrem = " & RS!Codigo & " AND anyorem = " & RS!Anyo & "|"
        RS.MoveNext
    Wend
    RS.Close
    
    While CuentasCC <> ""
        I = InStr(1, CuentasCC, "|")
        If I = 0 Then
            CuentasCC = ""
        Else
            SQL = Mid(CuentasCC, 1, I - 1)
            CuentasCC = Mid(CuentasCC, I + 1)
            
            RS.Open "Select count(*) from scobro where " & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            I = 0
            If Not RS.EOF Then I = DBLet(RS.Fields(0), "N")
            RS.Close
            If I = 0 Then
                SQL = Replace(SQL, "codrem", "codigo")
                SQL = Replace(SQL, "anyorem", "anyo")
                SQL = "UPDATE remesas set situacion='Z' WHERE " & SQL
                Conn.Execute SQL
            End If
        End If
    Wend
    
    Screen.MousePointer = vbHourglass
    I = 0
    SQL = "Select * from remesas where fecremesa <='" & Format(Text1(17).Text, FormatoFecha) & "'"
    SQL = SQL & " AND situacion='Z'  AND tiporem "
    If SubTipo = vbTipoPagoRemesa Then
        SQL = SQL & " = 1 " 'EFECTOS
    Else
        SQL = SQL & " <> 1 " 'Talones y pagares
    End If
    SQL = SQL & " order by codigo,anyo"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "Select count(*) from scobro where "
    Set miRsAux = New ADODB.Recordset
    While Not RS.EOF
        Label10.Caption = "Comprobando: " & RS!Codigo & " - " & RS!Anyo
        Label10.Refresh
        
        miRsAux.Open SQL & " codrem =" & RS!Codigo & " AND anyorem =" & RS!Anyo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        I = 0
        If Not miRsAux.EOF Then I = DBLet(miRsAux.Fields(0), "N")
        miRsAux.Close
        If I > 0 Then
            MsgBox "Efectos sin eliminar.  " & Label10.Caption, vbExclamation
            RS.Close
            Label10.Caption = ""
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        RS.MoveNext
    Wend
    SQL = RS.Source
    RS.Close
    
    
    
    
    'Llegados aqui... a borrar
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not RS.EOF
        I = I + 1
        RS.MoveNext
    Wend
    RS.Close
    SQL = "�Seguro que desea eliminar los datos selecionados. (Historico remesas. Total: " & I & ")"
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then
        Screen.MousePointer = vbDefault
        Label10.Caption = ""
        Exit Sub
    End If
    Label10.Caption = "Eliminado datos."
    Me.Refresh
    SQL = " delete from remesas where fecremesa <='" & Format(Text1(17).Text, FormatoFecha) & "'"
    SQL = SQL & " AND situacion='Z'"
    Conn.Execute SQL
    
    
    If SubTipo <> vbTipoPagoRemesa Then
        'Comprobaremos si en la recepcion de documentos tb hay que eliminar los datos
        EliminarEnRecepcionDocumentos
    
    End If
    
    Unload Me
End Sub

Private Sub cmdEliminarReclama_Click()
    SQL = ""
    For I = 1 To Me.ListView6.ListItems.Count
        If Me.ListView6.ListItems(I).Checked Then SQL = SQL & "X"
    Next
    
    If SQL = "" Then Exit Sub
    SQL = "Desea quitar de la reclamacion las cuentas seleccionadas(" & Len(SQL) & ") ?"
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        SQL = "DELETE FROM  tmpentrefechas WHERE codUsu = " & vUsu.Codigo & " AND fechaadq = '"
        For I = Me.ListView6.ListItems.Count To 1 Step -1
            If ListView6.ListItems(I).Checked Then
                CuentasCC = SQL & ListView6.ListItems(I).Text & "'"
                Conn.Execute CuentasCC
                ListView6.ListItems.Remove I
            End If
        Next I
    End If
End Sub

Private Sub cmdEmpresa_Click(Index As Integer)
    CadenaDesdeOtroForm = ""
    If Index = 0 Then
        SQL = ""
        CuentasCC = ""
        For I = 1 To lwE.ListItems.Count
            If Me.lwE.ListItems(I).Checked Then
                SQL = SQL & Me.lwE.ListItems(I).Text & "|"
                CuentasCC = CuentasCC & "1" 'Contador
            End If
        Next I
        CadenaDesdeOtroForm = Len(CuentasCC) & "|" & SQL
        'Vemos las conta
        SQL = ""
        For I = 1 To lwE.ListItems.Count
            If Me.lwE.ListItems(I).Checked Then
                SQL = SQL & Me.lwE.ListItems(I).Tag & "|"
            End If
        Next I
        CadenaDesdeOtroForm = CadenaDesdeOtroForm & SQL
    End If
    Unload Me

End Sub

Private Sub cmdIMprime_Click()
Dim C As String

    C = ""
    
    
    CuentasCC = "fecvenci"
    If Text1(25).Text <> "" Then C = C & " AND " & CuentasCC & " >= '" & Format(Text1(25).Text, FormatoFecha) & "'"
    If Text1(24).Text <> "" Then C = C & " AND " & CuentasCC & " <= '" & Format(Text1(24).Text, FormatoFecha) & "'"
    
    CuentasCC = "fecfaccl"
    If Text1(20).Text <> "" Then C = C & " AND " & CuentasCC & " >= '" & Format(Text1(20).Text, FormatoFecha) & "'"
    If Text1(23).Text <> "" Then C = C & " AND " & CuentasCC & " <= '" & Format(Text1(23).Text, FormatoFecha) & "'"
    
    CuentasCC = "numserie"
    If txtSerie(3).Text <> "" Then C = C & " AND " & CuentasCC & " >= '" & txtSerie(3).Text & "'"
    If txtSerie(2).Text <> "" Then C = C & " AND " & CuentasCC & " <= '" & txtSerie(2).Text & "'"

    
    CuentasCC = "codfaccl"
    If txtNumFac(3).Text <> "" Then C = C & " AND " & CuentasCC & " >= " & txtNumFac(3).Text
    If txtNumFac(2).Text <> "" Then C = C & " AND " & CuentasCC & " <= " & txtNumFac(2).Text
    
    
    CuentasCC = "scobro.codmacta"
    If txtCtaNormal(7).Text <> "" Then C = C & " AND " & CuentasCC & " >= '" & txtCtaNormal(7).Text & "'"
    If txtCtaNormal(8).Text <> "" Then C = C & " AND " & CuentasCC & " <= '" & txtCtaNormal(8).Text & "'"
        
    CuentasCC = "scobro.codforpa"
    If txtFP(0).Text <> "" Then C = C & " AND " & CuentasCC & " >= " & txtFP(0).Text
    If txtFP(1).Text <> "" Then C = C & " AND " & CuentasCC & " <= " & txtFP(1).Text
    
    If C <> "" Then
        SQL = Mid(C, 5)
    Else
        SQL = ""
    End If

    'TROZO FINAL SQL
    C = "SELECT count(*) "
    C = C & " FROM ((scobro INNER JOIN sforpa ON scobro.codforpa = sforpa.codforpa) INNER JOIN stipoformapago ON sforpa.tipforpa = stipoformapago.tipoformapago) INNER JOIN cuentas ON scobro.codmacta = cuentas.codmacta"
    If SQL <> "" Then C = C & " WHERE " & SQL
    
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    NumRegElim = 0
    If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    Set miRsAux = Nothing
    If NumRegElim = 0 Then
        MsgBox "Ning�n dato a mostrar", vbExclamation
        Exit Sub
    End If
'--monica
'    frmVerCobrosImprimir.vSQL = SQL
'    frmVerCobrosImprimir.Show vbModal
End Sub





Private Sub cmdInsertaCta_Click()
    
    txtCtaNormal(5).Text = Trim(txtCtaNormal(5).Text)
    If txtCtaNormal(5).Text = "" Then Exit Sub
    
    If InStr(1, CuentasCC, txtCtaNormal(5).Text & "|") > 0 Then
        MsgBox "Ya la ha a�adido", vbExclamation
    Else
        CuentasCC = CuentasCC & txtCtaNormal(5).Text & "|"
        SQL = txtCtaNormal(5).Text & "      " & txtDCtaNormal(5).Text
        List1.AddItem SQL
        txtCtaNormal(5).Text = ""
        txtDCtaNormal(5).Text = ""
    End If
    PonerFoco Me.txtCtaNormal(5)
    
End Sub

Private Sub cmdListadoGastos_Click()
Dim I1 As Currency
Dim ITot As Currency
Dim C As Long
Dim RC As String
Dim F As Date
    
    RC = RecuperaValor(CadenaDesdeOtroForm, 4)
    
    If RecuperaValor(CadenaDesdeOtroForm, 3) = 0 Then
        'SIN DETALLAR. Va por fechas
        ITot = 0
        F = CDate(ListView4.ListItems(1).SubItems(1))
        ITot = 0
        For I = 1 To ListView4.ListItems.Count
            If ListView4.ListItems(I).Checked Then
                If CDate(ListView4.ListItems(I).SubItems(1)) <> F Then
                    NumRegElim = NumRegElim + 1
                    SQL = "'GASTO'," & NumRegElim & ",'" & Format(F, FormatoFecha) & "','GASTOS PENDIENTES',NULL,"
                    'HAY GASTOS
                    If ITot > 0 Then
                        SQL = SQL & "NULL," & TransformaComasPuntos(CStr(ITot))
                    Else
                        SQL = SQL & TransformaComasPuntos(CStr(Abs(ITot))) & ",NULL"
                    End If
                    SQL = RC & SQL & ")"
                    Conn.Execute SQL
                    'Reasignamos
                    F = CDate(ListView4.ListItems(I).SubItems(1))
                    ITot = ImporteFormateado(ListView4.ListItems(I).SubItems(2))
                              
                Else
                    I1 = ImporteFormateado(ListView4.ListItems(I).SubItems(2))
                    ITot = ITot + I1
                End If
            End If
        Next I
                
        If ITot <> 0 Then
                NumRegElim = NumRegElim + 1
                SQL = "'GASTO'," & NumRegElim & ",'" & Format(F, FormatoFecha) & "','GASTOS PENDIENTES',NULL,"
                'HAY GASTOS
                If ITot > 0 Then
                    SQL = SQL & "NULL," & TransformaComasPuntos(CStr(ITot))
                Else
                    SQL = SQL & TransformaComasPuntos(CStr(Abs(ITot))) & ",NULL"
                End If
                SQL = RC & SQL & ")"
                Conn.Execute SQL
        End If

    Else
         'DETALLAR PAGOS COBROS

            
            'INSERT INTO tmpfaclin (codusu, IVA,codigo, Fecha, Cliente,
            'cta, ImpIVA, Total) VALUES (100,'COBRO',2,'2005-09-28',
            ''A2500565/1','4320001',0,NULL)
            For I = 1 To ListView4.ListItems.Count
                If ListView4.ListItems(I).Checked Then
                    
                    NumRegElim = NumRegElim + 1
                    SQL = "'GASTO'," & NumRegElim & ",'" & Format(ListView4.ListItems(I).SubItems(1), FormatoFecha) & "','"
                    SQL = SQL & DevNombreSQL(ListView4.ListItems(I).Text) & "',NULL,"
                    I1 = ImporteFormateado(ListView4.ListItems(I).SubItems(2))
                    If I1 <> 0 Then
                        If I1 > 0 Then
                            SQL = SQL & "NULL," & TransformaComasPuntos(CStr(I1))
                        Else
                            SQL = SQL & TransformaComasPuntos(CStr(Abs(I1))) & ",NULL"
                        End If
                        SQL = SQL & ")"
                        SQL = RC & SQL
                        Conn.Execute SQL
                    End If
                End If
            Next I
        
   
        
    End If
    
    
    
    
    'Cerramos
    Unload Me
End Sub

Private Sub cmdModRemTal_Click()
    If Text1(27).Text = "" And Me.txtCta(8).Text = "" Then Exit Sub
    SQL = ""
    If Text1(27).Text <> "" Then SQL = SQL & vbCrLf & "Fecha: " & Text1(27).Text
    If txtCta(8).Text <> "" Then SQL = SQL & vbCrLf & "Cuenta: " & txtCta(8).Text & " " & txtDescCta(8).Text
    SQL = "Desea actualizar a los valores indicados?"
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    CadenaDesdeOtroForm = Text1(27).Text & "|" & Me.txtCta(8).Text & "|"
    Unload Me
End Sub

Private Sub cmdOrdenarPago_Click()
Dim cad As String
Dim Forpa As Integer

    'Algunas conideraciones
    'Fecha pago tiene k tener valor
    If Text1(5).Text = "" Then
        MsgBox "Fecha de pago debe tener valor", vbExclamation
        PonerFoco Text1(5)
        Exit Sub
    End If
    
    
    
    'VEMOS SI LA FECHA ESTA DENTRO DEL EJERCICIO
    If FechaCorrecta2(CDate(Text1(5).Text), False) > 1 Then
        MsgBox "Fecha pago fuera de fechas de ejercicios", vbExclamation
        PonerFoco Text1(5)
        Exit Sub
    End If
    
    
    If txtCta(1).Text = "" Then
        MsgBox "Seleccione la cuenta contable asociada al banco", vbExclamation
        PonerFoco txtCta(1)
        Exit Sub
    End If
    
    
    'Llegados a este punto montaremos el sql
    SQL = ""
    
    If Text1(3).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " spagop.fecefect >= '" & Format(Text1(3).Text, FormatoFecha) & "'"
    End If
        
        
    If Text1(4).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " spagop.fecefect <= '" & Format(Text1(4).Text, FormatoFecha) & "'"
    End If
    
    
    If SQL <> "" Then SQL = SQL & " AND "
    SQL = SQL & " sforpa.tipforpa = " & SubTipo

    'Si pone proveedor
    If txtCtaNormal(10).Text <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
        SQL = SQL & " spagop.ctaprove = '" & Me.txtCtaNormal(10).Text & "'"
    End If
    
    
    Screen.MousePointer = vbHourglass
    cad = " FROM spagop,sforpa WHERE spagop.codforpa = sforpa.codforpa AND "
    'Hacemos un conteo
    Set RS = New ADODB.Recordset
    I = 0
    RS.Open "SELECT Count(*) " & cad & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        I = DBLet(RS.Fields(0), "N")
    End If
    RS.Close
    Set RS = Nothing
    
    
    If I = 0 Then
        MsgBox "Ning�n dato con esos valores.", vbExclamation
    Else
        'La ordenacion de los efectos
        If optOrdPag(1).Value Then
            I = 1
        ElseIf optOrdPag(2).Value Then
            I = 2
        ElseIf optOrdPag(3).Value Then
            I = 3
        Else
            I = 0
        End If
        
    
    
    
        If BloqueoManual(True, "ORDEPAGO", CStr(SubTipo)) Then
        
            'El campo Observaciones lo meto en la BD en la tabla
            'Y luego lo leere desde ahi
            If SubTipo = 2 Or SubTipo = 3 Then
                If FrameDocPorveedor.Visible Then GuardaDatosConceptoTalonPagare
            End If
        
            'Hay datos, abriremos el forumalrio para k seleccione
            'los pagos que queremos hacer
            With frmTESVerCobrosPagos
                .vSql = SQL
                .OrdenarEfecto = True
                .Regresar = False
                .Cobros = False
                .NumeroTalonPagere = ""
                If SubTipo = 2 Or SubTipo = 3 Then
                    If FrameDocPorveedor.Visible Then .NumeroTalonPagere = txtTexto(2).Text
                End If
                .OrdenacionEfectos = I
                'Los texots
                .Tipo = SubTipo
                
                'Marzo2013   Cobramos un solo proveedor
                'Aparecera un boton para traer todos los pagos
                .CodmactaUnica = Trim(txtCtaNormal(10).Text)
                
                
                .vTextos = Text1(5).Text & "|" & Me.txtCta(1).Text & " - " & Me.txtDescCta(1).Text & "|" & SubTipo & "|"
                .Show vbModal
            End With
            BloqueoManual False, "ORDEPAGO", ""
            LeerGuardarOrdenacion False, False, I
        Else
            MsgBox "Proceso bloqueado por otro usuario", vbExclamation
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub




Private Sub cmdPorNIF_Click()

    If Opcion = 13 Then
        If Text4.Text = "" Or Text5.Text = "" Then
            MsgBox "Introduzca el NIF", vbExclamation
            Exit Sub
        End If
    End If
    SQL = ""
    For I = 1 To ListView3.ListItems.Count
        If ListView3.ListItems(I).Checked Then
            SQL = "O"
            Exit For
        End If
    Next I
    If SQL = "" Then
        MsgBox "Seleccione al menos una empresa", vbExclamation
        Exit Sub
    End If
    
    'Tipos de pago
    SQL = ""
    For I = 1 To lwtipopago.ListItems.Count
        If lwtipopago.ListItems(I).Checked Then
            SQL = "O"
            Exit For
        End If
    Next I
    If SQL = "" Then
        MsgBox "Seleccione al menos un tipo de pago", vbExclamation
        Exit Sub
    End If
    
    
    
    
    'Preparo
    Me.cmdPorNIF.Enabled = False
    Me.cmdCancelar(13).Enabled = False
    Me.cmdCancelarDeuda.Visible = True
    Me.cmdCancelarDeuda.Cancel = True
    Me.Refresh
    Screen.MousePointer = vbHourglass
    Cancelado = False
'  INSERT INTO tmp347 (codusu, cliprov, cta, nif) VALUES (
    '-----------------------------------------------------------------------------
    Label9.Caption = "Preparando tablas"
    Label9.Refresh
    SQL = "Delete from tmp347 where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    SQL = "Delete from Usuarios.ztesoreriacomun where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    'tmpfaclin  ... sera para cuando es mas de uno
    SQL = "Delete from tmpfaclin where codusu =" & vUsu.Codigo
    Conn.Execute SQL
    
    
    
    
    
    
    
    
    'AHORA INSERTO EN LAS TABLA tmpcta las cuentas que tienen ese NIF , para cada empresa seleccionada
    SQL = ""
    Screen.MousePointer = vbHourglass
    If Opcion = 13 Then
        '------------------------------------------
        'UNO SOLO
        For I = 1 To ListView3.ListItems.Count
            If ListView3.ListItems(I).Checked Then
                If Cancelado Then Exit For
                Label9.Caption = "Obteniendo tabla1: " & ListView3.ListItems(I).Text
                Label9.Refresh
                
                SQL = "Select " & vUsu.Codigo & "," & Mid(ListView3.ListItems(I).Key, 2) & ",codmacta,nifdatos"
                SQL = "INSERT INTO tmp347 (codusu, cliprov, cta, nif) " & SQL
                SQL = SQL & " FROM Conta" & ListView3.ListItems(I).Tag & ".cuentas WHERE nifdatos = '" & Text4.Text & "' ORDER BY codmacta"
                If Not Ejecuta(SQL) Then Exit Sub
                DoEvents
            End If
        Next I
        
        
    Else
        '�Desde Hasta
        'Cargamos
        CargaCtasparaAgruparNIF
        
    End If
        

    
        
        
    If SQL <> "" Then
        If GeneraCobrosPagosNIF Then
        
        
            SQL = ""
            For I = 1 To Me.ListView3.ListItems.Count
                If Me.ListView3.ListItems(I).Checked Then SQL = SQL & "1"
            Next
            If Len(SQL) > 1 Then
                SQL = "0"
            Else
                SQL = "1"
            End If
            SQL = "SoloUnaEmpresa= " & SQL & "|"

        
            With frmImprimir
                
                .OtrosParametros = SQL & "Cuenta= """"|"
                .NumeroParametros = 2
                .FormulaSeleccion = "{ado.codusu}=" & vUsu.Codigo
                .SoloImprimir = False
                .Opcion = 25
                .Show vbModal
            End With
        
        End If
    End If
    
    Me.cmdPorNIF.Enabled = True
    Me.cmdCancelar(13).Enabled = True
    Me.cmdCancelarDeuda.Visible = False
    Me.cmdCancelar(13).Cancel = True
    Label9.Caption = ""
    Screen.MousePointer = vbDefault
End Sub

'
'
'




Private Sub NuevaRemTalPag()
'--monica
'Dim CtaPuente As Boolean
'Dim Forpa As String
'Dim cad As String
'Dim Impor As Currency
'
''Algunas conideraciones
'
'        'Para talones y pagares obligado la cuenta bancaria
'        If txtCta(3).Text = "" Then
'            MsgBox "Indique la cuenta bancaria", vbExclamation
'            Exit Sub
'        End If
'
'
'
'    'Fecha remesa tiene k tener valor
'    If Text1(8).Text = "" Then
'        MsgBox "Fecha de remesa debe tener valor", vbExclamation
'        PonerFoco Text1(8)
'        Exit Sub
'    End If
'
'
'
'    'VEMOS SI LA FECHA ESTA DENTRO DEL EJERCICIO
'    If FechaCorrecta2(CDate(Text1(8).Text), True) > 1 Then Exit Sub
'
'        'NO hago la pregunta. Si no tiene la cuenta puente dejo continuar igual
''        If Me.cmbRemesa.ListIndex = 0 Then
''            SQL = Abs(vParam.PagaresCtaPuente)
''        Else
''            SQL = Abs(vParam.TalonesCtaPuente)
''        End If
''        If SQL = "0" Then
''
''            MsgBox "Falta configurar la opci�n en parametros", vbExclamation
''            Exit Sub
''        End If
'
'    If Me.cmbRemesa.ListIndex = 0 Then
'        CtaPuente = vParamT.PagaresCtaPuente
'    Else
'        CtaPuente = vParamT.TalonesCtaPuente
'    End If
'
'
'
'    'A partir de la fecha generemos leemos k remesa corresponde
'    SQL = "select max(codigo) from remesas where anyo=" & Year(CDate(Text1(8).Text))
'    Set miRsAux = New ADODB.Recordset
'    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    NumRegElim = 0
'    If Not miRsAux.EOF Then
'        NumRegElim = DBLet(miRsAux.Fields(0), "N")
'    End If
'    miRsAux.Close
'
'    NumRegElim = NumRegElim + 1
'    txtRemesa.Text = NumRegElim
'
'
'
'        If Me.cmbRemesa.ListIndex = 0 Then
'            SQL = " talon = 0"
'        Else
'            SQL = " talon = 1"
'        End If
'
'        'Si no lleva cuenta puente, no hace falta que este contabilizada
'        'Es decir. Solo mirare contabilizados si llevo ctapuente
'        If CtaPuente Then SQL = SQL & " AND contabilizada = 1 "
'        SQL = SQL & " AND LlevadoBanco = 0 "
'
'        'de la recepcion de factura
'        If Text1(6).Text <> "" Then SQL = SQL & " AND fechavto >= '" & Format(Text1(6).Text, FormatoFecha) & "'"
'        If Text1(7).Text <> "" Then SQL = SQL & " AND fechavto <= '" & Format(Text1(7).Text, FormatoFecha) & "'"
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'        'Fecha recepcion
'        If Text1(22).Text <> "" Then SQL = SQL & " AND fecharec >= '" & Format(Text1(22).Text, FormatoFecha) & "'"
'        If Text1(21).Text <> "" Then SQL = SQL & " AND fecharec <= '" & Format(Text1(21).Text, FormatoFecha) & "'"
'
'
'
'
'    Screen.MousePointer = vbHourglass
'    Set RS = New ADODB.Recordset
'
'    'Que la cuenta NO este bloqueada
'    I = 0
'    cad = "select cuentas.codmacta,nommacta,FecBloq from "
'    cad = cad & "scarecepdoc,cuentas where scarecepdoc.codmacta=cuentas.codmacta"
'    cad = cad & " AND (not (fecbloq is null) and fecbloq < '" & Format(CDate(Text1(8).Text), FormatoFecha) & "') "
'    cad = cad & " AND " & SQL & " GROUP by 1"
'
'
'
'    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not RS.EOF Then
'        cad = ""
'        I = 1
'        While Not RS.EOF
'            cad = cad & RS!codmacta & " - " & RS!Nommacta & " : " & RS!FecBloq & vbCrLf
'            RS.MoveNext
'        Wend
'    End If
'
'    RS.Close
'
'    If I > 0 Then
'        cad = "Las siguientes cuentas estan bloquedas." & vbCrLf & String(60, "-") & vbCrLf & cad
'        MsgBox cad, vbExclamation
'        Screen.MousePointer = vbDefault
'
'        Exit Sub
'    End If
'
'
'    cad = " FROM scarecepdoc,cuentas where scarecepdoc.codmacta=cuentas.codmacta AND"
'
'    'Hacemos un conteo
'    RS.Open "SELECT * " & cad & SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    I = 0
'    cad = ""
'    While Not RS.EOF
'        I = I + 1
'        cad = cad & " OR ( id = " & RS!Codigo & ") "
'        RS.MoveNext
'    Wend
'    RS.Close
'    If I = 0 Then
'        MsgBox "Ningun dato con esos valores", vbExclamation
'        Exit Sub
'    End If
'    cad = "(" & Mid(cad, 4) & ")"
'    SQL = " from scobro where (numserie,codfaccl,fecfaccl,numorden) in (select numserie ,numfaccl,fecfaccl,numvenci from slirecepdoc where " & cad & ")"
'    SQL = "select sum(impvenci),sum(impcobro),sum(gastos) " & SQL
'
'
'
'
'    'La suma
'    If I > 0 Then
'
'        Impor = 0
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        'If Not Rs.EOF Then Impor = DBLet(Rs.Fields(0), "N") - DBLet(Rs.Fields(1), "N") + DBLet(Rs.Fields(2), "N")
'
'        'Solo el impcobro
'        If Not RS.EOF Then Impor = DBLet(RS.Fields(1), "N")
'        RS.Close
'        If Impor = 0 Then I = 0
'    End If
'
'
'    Set RS = Nothing
'
'    If I = 0 Then
'        MsgBox "Ningun dato a remesar con esos valores(II)", vbExclamation
'    Else
'
'
'        'Preparamos algunas cosillas
'        'Aqui guardaremos cuanto llevamos a cada banco
'        SQL = "Delete from tmpCierre1 where codusu =" & vUsu.Codigo
'        Conn.Execute SQL
'
'        'Si son talones o pagares NO hay reajuste en bancos
'        'Con lo cual cargare la tabla con el banco
'
'        If SubTipo <> vbTipoPagoRemesa Then
'            ' Metermos cta banco, n�remesa . El resto no necesito
'            SQL = "INSERT INTO tmpcierre1 (codusu, cta, nomcta, acumPerD) VALUES ("
'            SQL = SQL & vUsu.Codigo & ",'" & txtCta(3).Text & "','"
'            'ANTES
'            'SQL = SQL & DevNombreSQL(Me.txtDescCta(3).Text) & "'," & TransformaComasPuntos(CStr(Impor)) & ")"
'            'AHora.
'            SQL = SQL & txtRemesa.Text & "',0)"
'            Conn.Execute SQL
'        End If
'
'
'        'Lo qu vamos a hacer es , primero bloquear la opcioin de remesar
'        If BloqueoManual(True, "Remesas", "Remesas") Then
'
'            Me.Visible = False
'
'
'            'Remesas de talones y pagares
'            frmRemeTalPag.vRemesa = "" 'NUEVA
'            frmRemeTalPag.SQL = cad
'            frmRemeTalPag.Talon = cmbRemesa.ListIndex = 1 '0 pagare   1 talon
'            frmRemeTalPag.Text1(0).Text = Me.txtCta(3).Text & " - " & txtDescCta(3).Text
'            frmRemeTalPag.Text1(1).Text = Text1(8).Text
'            frmRemeTalPag.Show vbModal
'
'            'Desbloqueamos
'            BloqueoManual False, "Remesas", ""
'            Unload Me
'        Else
'            MsgBox "Otro usuario esta generando remesas", vbExclamation
'        End If
'
'    End If
'
'    Screen.MousePointer = vbDefault
End Sub




Private Function UpdatearCobrosRemesa() As Boolean
Dim Im As Currency
    On Error GoTo EUpdatearCobrosRemesa
    UpdatearCobrosRemesa = False
    
    SQL = "Select * from scobro WHERE codrem=" & Text3(0).Text
    SQL = SQL & " AND anyorem =" & Text3(1).Text
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
            While Not miRsAux.EOF
                SQL = "UPDATE scobro SET fecultco = '" & Format(Text1(9).Text, FormatoFecha) & "', impcobro = "
                Im = miRsAux!ImpVenci
                If Not IsNull(miRsAux!Gastos) Then Im = Im + miRsAux!Gastos
                SQL = SQL & TransformaComasPuntos(CStr(Im))
                
                SQL = SQL & " ,siturem = 'B'"
                
                
                'WHERE
                SQL = SQL & " WHERE numserie='" & miRsAux!NUmSerie
                SQL = SQL & "' AND  codfaccl =  " & miRsAux!codfaccl
                SQL = SQL & "  AND  fecfaccl =  '" & Format(miRsAux!fecfaccl, FormatoFecha)
                SQL = SQL & "' AND  numorden =  " & miRsAux!numorden
                'Muevo siguiente
                miRsAux.MoveNext
                
                'Ejecuto SQL
                If Not Ejecuta(SQL) Then MsgBox "Error: " & SQL, vbExclamation
            Wend
    End If
    miRsAux.Close
                    
                    
                    
    UpdatearCobrosRemesa = True
    Exit Function
EUpdatearCobrosRemesa:
    
End Function

Private Sub cmdRecaudaEjec_Click()

    'Comprobaciones
    CuentasCC = ""
    SQL = ""
    For I = 1 To Me.ListView5.ListItems.Count
        If Me.ListView5.ListItems(I).Checked Then
            If Me.ListView5.ListItems(I).ForeColor = vbRed Then
                CuentasCC = CuentasCC & "A"
            Else
                SQL = SQL & "A"
            End If
        End If
    Next I
    
    If Len(CuentasCC) > 0 Then
        MsgBox "Hay vencimientos (" & Len(CuentasCC) & ")  seleccionados que tienen errores ", vbExclamation
        Exit Sub
    End If
    
    If Len(SQL) = 0 Then
        MsgBox "Seleccione los vencimientos ", vbExclamation
        Exit Sub
    End If
    
    
    
    'OK vamos con la generacion del fichero
    SQL = ""
    For I = 1 To Me.ListView5.ListItems.Count
        With ListView5.ListItems(I)
            If .Checked Then
                '(numserie,codfaccl,fecfaccl,numorden)
                SQL = SQL & ", ('" & .Text & "',"
                SQL = SQL & .SubItems(1) & ",'" & Format(.SubItems(2), FormatoFecha)
                SQL = SQL & "'," & .SubItems(3) & ")"
                
            End If
        End With
    Next I
    SQL = Mid(SQL, 2) 'quitamos la primera coma
    If GeneraFicheroRecaudacionEjecutiva(SQL) Then Unload Me
    
End Sub

Private Sub cmdReclamas_Click()
    
    'Borraremos los que tienen mail erroneo
    Set RS = New ADODB.Recordset
    SQL = "SELECT fechaadq FROM  tmpentrefechas,cuentas WHERE fechaadq=codmacta  "
    SQL = SQL & " AND codUsu = " & vUsu.Codigo & " AND "
    SQL = SQL & " coalesce(maidatos,'')='' GROUP BY fechaadq  "
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not RS.EOF
        SQL = SQL & ", '" & RS!fechaadq & "'"
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    If SQL <> "" Then
        SQL = "DELETE FROM  tmpentrefechas WHERE codUsu = " & vUsu.Codigo & " AND  fechaadq IN (" & Mid(SQL, 2) & ")"
        Conn.Execute SQL
    End If
        
        
    SQL = DevuelveDesdeBD("count(*)", "tmpentrefechas", "codusu", CStr(vUsu.Codigo))
    If Val(SQL) = 0 Then
        MsgBox "Ninguna reclamacion a enviar", vbExclamation
    Else
        CadenaDesdeOtroForm = "OK"
    End If
    SubTipo = 0
    
    Unload Me
End Sub

Private Sub cmdRemesas_Click()
    
    If SubTipo <> vbTipoPagoRemesa Then
        NuevaRemTalPag
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdTr_Click()
Dim cad As String
'--monica
'
'
'    'Vemos que todo es correcto: fechas, textos , cta banco...
'    If txtCta(5).Text = "" Then
'        MsgBox "Seleccione la cuenta contable del banco", vbExclamation
'        Exit Sub
'    End If
'    If Text1(14).Text = "" Then
'        MsgBox "Ponga fecha transferenica", vbExclamation
'        Exit Sub
'    End If
'    I = EsFechaOK(Text1(14))
'    If I > 1 Then
'        If I = 2 Then
'            MsgBox "Fecha  ejercicios cerrados.", vbExclamation
'            Exit Sub
'        End If
'        cad = "Fecha fuera de  ejercicios . �Desea continuar?"
'        If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
'    End If
'    If Text6.Text = "" Then
'        MsgBox "Debe poner la descripci�n", vbExclamation
'        Exit Sub
'    End If
'
'
'
'    'Comprobamos recibos
'    'Montamos el sql
'    '---------------
'    If SubTipo >= 1 Then
'
'        If SubTipo = 1 Then
'            cad = "1" '1: TRANSFERENCIA
'        Else
'            cad = "5" '3: CONFIRMING O PAGO DOMICILIADO
'        End If
'        cad = " sforpa.tipforpa = " & cad   '1: TRANSFERENCIAS
'        'Al poner el impefect, solo cogemos importes positivos
'        'cad = cad & " AND spagop.transfer is null and spagop.impefect>0"
'        cad = cad & " AND spagop.transfer is null"
'        'El importe NO DEBE SER INFERIOR A 0
'        cad = cad & " AND impefect > 0"
'        SQL = "spagop.fecefect"
'    Else
'        'Transferencias en cobros. SOn abonos
'        cad = " sforpa.tipforpa = " & 1   '1: TRANSFERENCIAS
'        cad = cad & " AND scobro.transfer is null"
'        'Es decir, cojeremos aquellos vencimientos cuyo importe sea
'        'menor a 0 sea cual sea su forma de pago
'        cad = " scobro.impvenci <0 AND scobro.transfer is null"
'        SQL = "scobro.fecvenci"
'    End If
'
'
'
'    'Las fechas desde / hasta
'    'estoy guardando en la variable SQL la columna fecha, para hacerla efectiva
'    'segun sea desde o hasta
'    If Text1(15).Text <> "" Then cad = cad & " AND " & SQL & " >= '" & Format(Text1(15).Text, FormatoFecha) & "'"
'    If Text1(16).Text <> "" Then cad = cad & " AND " & SQL & " <= '" & Format(Text1(16).Text, FormatoFecha) & "'"
'    If Me.txtCtaNormal(12).Text <> "" And Me.txtDCtaNormal(12).Text <> "" Then
'        SQL = "scobro.codmacta"
'        If SubTipo >= 1 Then SQL = "spagop.ctaprove"
'        cad = cad & " AND " & SQL & " = '" & txtCtaNormal(12) & "'"
'    End If
'
'    SQL = ""
'    'Vemos si hay recibos
'    If Not frmTransferencias2.VerHayEfectos(cad) Then
'        MsgBox "Ningun paga a efectuar con esos valores", vbExclamation
'        Exit Sub
'    End If
'
'    SQL = "Transferencias"
'    If SubTipo = 0 Then SQL = SQL & "co"
'    'Bloqueamos el crear transferencias
'    If Not BloqueoManual(True, SQL, CStr(vEmpresa.codempre)) Then
'        MsgBox "El proceso esta bloqueado por otro usuario", vbExclamation
'        Exit Sub
'     End If
'
'    'Obtenemos contador
'    NumRegElim = Val(SugerirCodigoSiguienteTransferencia)
'    I = NumRegElim
'
'
'    'Abrimos la pantalla de seleccionar pagos cobros
'    With frmVerCobrosPagos
'            .vSQL = cad
'            .OrdenarEfecto = True
'            .Regresar = False
'            .Cobros = (SubTipo = 0)
'            .ContabTransfer = False
'            'Los texots
'            .Tipo = 1
'
'            '.vTextos = Text1(5).Text & "|" & Me.txtCta(1).Text & " - " & Me.txtDescCta(1).Text & "|" & SubTipo & "|"
'            .vTextos = Text1(14).Text & "|" & txtCta(5).Text & " - " & Me.txtDescCta(5).Text & "|1|"  '1: transferencia
'            If Me.SubTipo = 2 Then
'                'Es un pago domiciliado
'                If vParam.PagosConfirmingCaixa Then
'                    .vTextos = .vTextos & "|CAIXA confirming|"
'                Else
'                    .vTextos = .vTextos & "|PAGO DOMICILIADO|"
'                End If
'            Else
'                .vTextos = .vTextos & "||"
'            End If
'            .SegundoParametro = NumRegElim
'            NumRegElim = 0
'            Me.Hide
'            .Show vbModal
'    End With
'
'
'    'Si ha seleccionado recibos, marcare para cuando vuelva a
'    'la pantalla de trasnferencias, lance el proceso de generacion de
'    'diskette
'    SQL = "Transferencias"
'    If SubTipo = 0 Then SQL = SQL & "co"
'    BloqueoManual False, SQL, ""
'    If NumRegElim > 0 Then
'
'        'Selec la usma de los recibos
'        If SubTipo = 0 Then
'            SQL = "Select sum(impvenci) from scobro"
'        Else
'            SQL = "Select sum(impefect) from spagop"
'        End If
'        SQL = SQL & " WHERE transfer =" & I
'        Set miRsAux = New ADODB.Recordset
'        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'        SQL = "stransfer"
'        If SubTipo = 0 Then SQL = SQL & "cob"
'
'
'
'
'
'
'        SQL = "INSERT INTO " & SQL & " (codigo, Descripcion, fecha, codmacta, diskette,importe,conceptoTrans"
'        If Me.SubTipo = 2 Then SQL = SQL & ",subtipo" 'para poder meter el UNO aqui
'        SQL = SQL & ") VALUES (" & I & ",'" & DevNombreSQL(Text6.Text) & "','" & Format(Text1(14).Text, FormatoFecha) & "','" & _
'            txtCta(5).Text & "',0,"
'        'LA suma
'        SQL = SQL & TransformaComasPuntos(CStr(Abs(DBLet(miRsAux.Fields(0), "N"))))
'        'Tpo remesa
'        If Me.SubTipo < 2 Then
'            SQL = SQL & "," & Me.cboConcepto.ItemData(cboConcepto.ListIndex)
'        Else
'            'PAGO DOMICILIADO
'            SQL = SQL & "," & Me.chkPagoDom.Value
'            SQL = SQL & ",1"  'para poder meter el UNO aqui de pago domiciiliado
'        End If
'        SQL = SQL & ")"
'        miRsAux.Close
'        Set miRsAux = Nothing
'
'        Conn.Execute SQL
'
'        frmTransferencias2.DatosADevolverBusqueda = I
'        Unload Me
'    Else
'        espera 0.2
'        'Me.Visible = True
'        Me.Show vbModal
'    End If
'    Screen.MousePointer = vbDefault
End Sub

Private Function SugerirCodigoSiguienteTransferencia() As String
    
    SQL = "Select Max(codigo) from stransfer"
    If SubTipo = 0 Then SQL = SQL & "cob"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, , , adCmdText
    SQL = "1"
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            SQL = CStr(RS.Fields(0) + 1)
        End If
    End If
    RS.Close
    Set RS = Nothing
    SugerirCodigoSiguienteTransferencia = SQL
End Function




Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcion
        Case 12
            'Elimiar efectos
            CargaRemesas
            
            If ListView2.ListItems.Count = 0 And SubTipo = 3 Then
                Unload Me
                Exit Sub
            End If
            
        Case 13, 14
            If Opcion = 13 Then
                PonerFoco Text4
            Else
                PonerFoco txtCtaNormal(1)
            End If
        Case 15
            Text1(15).SetFocus
            
        Case 18
            Screen.MousePointer = vbHourglass
            CargaGastos
            cmdListadoGastos.Default = True
            PonerFoco cmdListadoGastos
        Case 19
            Screen.MousePointer = vbHourglass
            CargaDatosContabilizarGastos
            PonerFoco Text1(19)
            
        Case 21
          
            CargalistaCuentas
            PonerFoco txtCtaNormal(5)
            
        Case 25
            PonerFoco Text1(27)
            
        Case 29
            CargarVtosRecaudaEjecutiva
            
        Case 31
            
            ReclamacionGargarList
            If ListView6.ListItems.Count = 0 Then optReclama(1).Value = True
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
    Limpiar Me
    PrimeraVez = True
    Me.Icon = frmPpal.Icon
    
    
    'Cago los iconos
    CargaImagenesAyudas Me.imgCtaNorma, 1, "Seleccionar cuenta"
    CargaImagenesAyudas imgCuentas, 1, "Cuenta contable banco"
    CargaImagenesAyudas imgRem, 1, "Seleccionar remesa"
    CargaImagenesAyudas imgFP, 1, "Seleccionar Forma de pago"
    CargaImagenesAyudas imgConcepto, 1, "Concepto"
    CargaImagenesAyudas Image1, 2


'    CargaImagenesAyudas ImageAyuda, 3

    Me.imgEliminarCta.Picture = frmPpal.ImaListBotoneras32.ListImages(5).Picture
    
    
'    Carga1ImagenAyuda Me.Image4, 1
'    Carga1ImagenAyuda Me.Image3, 1
'    Carga1ImagenAyuda imgFra, 1
'--    CargaImagenesAyudas Me.Image4, 1
'    CargaImagenesAyudas Me.Image3, 1
'    CargaImagenesAyudas imgFra, 1
    
    
    FrameCobros.Visible = False
    Framepagos.Visible = False
    FrameContabilRem2.Visible = False
    FrameDevlucionRe.Visible = False
    FrameImpagados.Visible = False
    frameAcercaDE.Visible = False
    FrameElimVtos.Visible = False
    FrameDeuda.Visible = False
    FrameTransfer.Visible = False
    FrameElimnaHcoReme.Visible = False
    FrameSelecGastos.Visible = False
    FrameContabilizarGasto.Visible = False
    FrameeMPRESAS.Visible = False
    FrameAgregarCuentas.Visible = False
    FrImprimeRecibos.Visible = False
    FrameModiRemeTal.Visible = False
    FrameDevDesdeVto.Visible = False
    FrameRecaudacionEjecutiva.Visible = False
    FrameReclamaEmail.Visible = False
    
    Select Case Opcion
    Case 0
        '
        Caption = "Cobros"
        CargaList
        Text1(0).Text = ""
        Text1(1).Text = Format(Now - 1, "dd/mm/yyyy")
        Text1(2).Text = Format(Now, "dd/mm/yyyy")
        Me.txtCta(0).Text = ""
        Me.txtDescCta(0).Text = ""
        FrameCobros.Visible = True
        H = FrameCobros.Height + 60
        W = FrameCobros.Width
        I = LeerGuardarOrdenacion(True, True, I)
        Me.optOrdCob(I).Value = True
        'En el 0 guardo la opcion por defecto
        Me.optOrdCob(0).Tag = I
        
        
        FrameCobroTarjeta.Visible = SubTipo = 6
        FrameCobroEfectivo.Visible = SubTipo = 0
       
        'Abril 2014
        'NAVARRES
        'Forma de pago tarjeta. Llevar� en lugar de GASTOS, el % de interes(dese parametros)
        'Navarres. Si tiene valor % gastos tarjeta, el dato de gastos pasa a ser
        ' %, ofertando el valor de la columna
        If SubTipo = 6 Then
            If vParamT.IntereseCobrosTarjeta > 0 Then
                Label4(16).Caption = "% Intereses"
                txtImporte(4).Text = Format(vParamT.IntereseCobrosTarjeta, FormatoImporte)
            End If
        End If
       
       
       
       'If SubTipo = 6 Then Me.txtImporte(4).TabIndex = 4
       'If SubTipo = 0 Then txtCta(2).TabIndex = 4
    Case 1
        Caption = "Pagos"
        CargaList
        Text1(3).Text = ""
        Text1(4).Text = Format(Now - 1, "dd/mm/yyyy")
        Text1(5).Text = Format(Now, "dd/mm/yyyy")
        Me.txtCta(1).Text = ""
        Me.txtDescCta(1).Text = ""
        Framepagos.Visible = True
        FrameDocPorveedor.Visible = False
        H = Framepagos.Height
        W = Framepagos.Width
        I = LeerGuardarOrdenacion(True, False, I)
        Me.optOrdPag(I).Value = True
    Case 8, 22, 23
        'Utilizare el mismo FRAM para
        '   8.- Contabilizar / Abono remesa
        '   22- Cancelacion cliente
        '   23- Confirmacion remesa
        '  TANTO DE EFECTOS como de talones pagares
        FrameContabilRem2.Visible = True
        
        Caption = "Remesas"
        If SubTipo = 1 Then
            Caption = Caption & " EFECTOS"
        Else
            Caption = Caption & " talones/pagar�s"
        End If
        chkAgrupaCancelacion.Visible = False
        
        If Opcion = 8 Then
            SQL = "Abono remesa"
            CuentasCC = "Contabilizar"
        Else
        
            If Opcion = 22 Then
            
                SQL = DevuelveDesdeBD("RemesaCancelacion", "paramtesor", "codigo", "1", "N")
                chkAgrupaCancelacion.Visible = Len(SQL) = vEmpresa.DigitosUltimoNivel
                SQL = "Cancelacion cliente"
                CuentasCC = "Can. cliente"
            Else
                SQL = "Confirmacion remesa"
                CuentasCC = "Confirmar"
            End If
            
        End If
        Label5(2).Caption = SQL
        cmdContabRemesa.Caption = CuentasCC
        CuentasCC = ""
        'Los gastos solo van en la contabilizacion
        Label4(2).Visible = Opcion = 8
        txtImporte(0).Visible = Opcion = 8
        
        'noviembre 2009
        'Opcion 8. Contabilizar(ABONO)
        ' tipo  efectos
        ' si tiene cta efectos comerciales descontados y es de ultimo nivel
        ' mostrar el agrupar efectos comerciales descontad
        ' DEBERIA IR AQUI el check visible o no.
        'Veremos si hay que ponerlo o no
        
        
        W = FrameContabilRem2.Width
        H = FrameContabilRem2.Height
    Case 9, 16, 28
        If SubTipo = 1 Then
            Caption = "EFECTOS"
        Else
            Caption = "TALONES / PAGARES"
        End If
        FrameDevlucionRe.Visible = True
        FrameDevDesdeFichero.Visible = Opcion = 16
        Me.FrameDevDesdeVto.Visible = Opcion = 28
        Caption = "Devolucion remesa (" & UCase(Caption) & ")"
        W = FrameDevlucionRe.Width
        H = FrameDevlucionRe.Height
        Text1(11).Text = Format(Now, "dd/mm/yyyy")
        txtImporte(1).Text = 0
    
        'Nuevo 28Marzo2007
        PonerValoresPorDefectoDevilucionRemesa
        
    Case 10
        Me.FrameImpagados.Visible = True
        Caption = "Devoluciones"
        W = FrameImpagados.Width
        H = FrameImpagados.Height
        CargaImpagados
        CargaIconoListview ListView1
        
        
    Case 11
        CargaImagen
        Me.Caption = "Acerca de ....."
        W = Me.frameAcercaDE.Width
        H = Me.frameAcercaDE.Height + 50
        Me.frameAcercaDE.Visible = True
        Label13.Caption = "Versi�n:  " & App.Major & "." & App.Minor & "." & App.Revision & " "
        
    Case 12
        Me.Caption = "Borrar vencimientos"
        W = Me.FrameElimVtos.Width
        H = Me.FrameElimVtos.Height + 200
        Me.FrameElimVtos.Visible = True
        CargaIconoListview ListView2
        
    Case 13, 14
        Caption = "DEUDA x NIF"
        If Opcion = 13 Then
            Label5(5).Caption = "Informe situaci�n por NIF"
        Else
            Label5(5).Caption = "Informe situaci�n por cuenta"
        End If
        
        W = Me.FrameDeuda.Width
        H = Me.FrameDeuda.Height + 200
        Me.FrameDeuda.Visible = True
        Text1(13).Text = Format(Now, "dd/mm/yyyy")
        cargaEmpresasTesor ListView3
        cargaTipoPagos
        FrameDH_cta.Visible = Opcion = 14
        Label9.Caption = ""
        
    Case 15
        'Tansferenicas
        FrameTransfer.Visible = True
        Label4(7).Caption = "Realizar transferencia"
        Label2(45).Caption = "Proveedor"
        If SubTipo = 2 Then
            If vParamT.PagosConfirmingCaixa Then
                Me.Caption = "Caixa confirming"
            Else
                Me.Caption = "Pagos domiciliados"
            End If
            Label4(7).Caption = "Realizar " & LCase(Me.Caption)
            Me.cboConcepto.ListIndex = 1
        Else
            Me.Caption = "Realizar transferencia"
            If SubTipo = 0 Then
                Me.Caption = Me.Caption & " (ABONOS)"
                Label2(45).Caption = "Cliente"
            End If
            Me.cboConcepto.ListIndex = 0
        End If
        W = Me.FrameTransfer.Width
        H = Me.FrameTransfer.Height + 200
        Text1(16).Text = Format(Now, "dd/mm/yyyy")
        Text1(14).Text = Text1(16).Text
        
        Me.cboConcepto.Visible = SubTipo <> 2
        Label2(43).Visible = SubTipo <> 2
        chkPagoDom.Visible = SubTipo = 2
        
    Case 17
    
        FrameElimnaHcoReme.Visible = True
        Me.Caption = "Hco remesas"
        If SubTipo <> vbTipoPagoRemesa Then Me.Caption = Me.Caption & " (Talones-Pagar�s)"
        W = Me.FrameElimnaHcoReme.Width
        H = Me.FrameElimnaHcoReme.Height '+ 200
        Text1(17).Text = Format(DateAdd("m", -2, Now), "dd/mm/yyyy")
    
    Case 18
    
        FrameSelecGastos.Visible = True
        Me.Caption = "Seleccionar gastos"
        W = Me.FrameSelecGastos.Width
        H = Me.FrameSelecGastos.Height '+ 200
        Label5(7).Caption = RecuperaValor(CadenaDesdeOtroForm, 1)
        CargaIconoListview ListView4
        
    Case 19
        'CONTABILIZAR GASTOS FIJOS
        PonerCuentasCC
        Me.Caption = "Contabilizar gastos fijos "
        FrameContabilizarGasto.Visible = True
        W = Me.FrameContabilizarGasto.Width
        H = Me.FrameContabilizarGasto.Height '+ 200
        
    Case 20
        
        Me.Caption = "Empresas disponibles"
        FrameeMPRESAS.Visible = True
        W = Me.FrameeMPRESAS.Width
        H = Me.FrameeMPRESAS.Height '+ 200
        cargaempresas
        
    Case 21
        Caption = "Seleccionar cuentas"
        FrameAgregarCuentas.Visible = True
        W = Me.FrameAgregarCuentas.Width
        H = Me.FrameAgregarCuentas.Height + 200
        
    Case 24
        Caption = "Impresion"
         
        FrImprimeRecibos.Visible = True
        W = Me.FrImprimeRecibos.Width
        H = Me.FrImprimeRecibos.Height + 200
        
    Case 25
        Caption = "Remesas"
        FrameModiRemeTal.Visible = True
        W = Me.FrameModiRemeTal.Width
        H = Me.FrameModiRemeTal.Height + 100
    Case 29
        Caption = "Recaudacion"
        FrameRecaudacionEjecutiva.Visible = True
        W = Me.FrameRecaudacionEjecutiva.Width
        H = Me.FrameRecaudacionEjecutiva.Height + 100
        
    Case 31
        
        Caption = "Reclamacion"
        FrameReclamaEmail.Visible = True
        W = Me.FrameReclamaEmail.Width
        H = Me.FrameReclamaEmail.Height + 100
        SubTipo = 1 'Para que cuando le de al ASPA del forma NO cierre
        
    End Select
    
    
    Me.Height = H + 360
    Me.Width = W + 90
    
    H = Opcion
    If Opcion = 7 Then H = 6
    If Opcion = 14 Then H = 13
    If Opcion = 16 Or Opcion = 28 Then H = 9
    If Opcion = 22 Or Opcion = 23 Then H = 8
    Me.cmdCancelar(H).Cancel = True
    
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    SQL = RecuperaValor(CadenaDevuelta, 1)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    If Opcion = 31 Then
        If SubTipo = 1 Then
            Cancel = 1
            Exit Sub
        End If
    End If

    If Opcion = 4 Then
        'REMESAS BANCARIAS
        If vParamT.RemesasPorEntidad Then
            If txtCta(3).Text <> txtCta(3).Tag Then LeerGuardarBancoDefectoEntidad False
        End If
        
    End If

    Set RS = Nothing
    Set miRsAux = Nothing
        
    
    NumeroDocumento = "" 'Para reestrablecerlo siempre
End Sub


Private Sub frmBa_DatoSeleccionado(CadenaSeleccion As String)
    I = CInt(imgCuentas(0).Tag)
    Me.txtCta(I).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescCta(I).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(Image1(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCCtas_DatoSeleccionado(CadenaSeleccion As String)
    SQL = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub frmP_DatoSeleccionado(CadenaSeleccion As String)
    txtFP(I).Text = RecuperaValor(CadenaSeleccion, 1)
    txtFPDesc(I).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmRe_DatoSeleccionado(CadenaSeleccion As String)
    If I = 0 Then
        Text3(3).Text = RecuperaValor(CadenaSeleccion, 1)
        Text3(4).Text = RecuperaValor(CadenaSeleccion, 2)
        Text1(10).Text = RecuperaValor(CadenaSeleccion, 3)
    Else
        'DEVOLUCIOIN
        Text3(5).Text = RecuperaValor(CadenaSeleccion, 1)
        Text3(6).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
    
End Sub

Private Sub Image1_Click(Index As Integer)
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text1(Index).Text <> "" Then frmC.Fecha = CDate(Text1(Index).Text)
    Image1(0).Tag = Index
    frmC.Show vbModal
    Set frmC = Nothing
    If Text1(Index).Text <> "" Then PonerFoco Text1(Index)
End Sub


Private Sub PonerFoco(ByRef o As Object)
    On Error Resume Next
    o.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub ObtenerFoco(ByRef T As TextBox)
    T.SelStart = 0
    T.SelLength = Len(T.Text)
End Sub

Private Sub KEYpress(ByRef KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Image3_Click()
        Set frmCCtas = New frmColCtas
        SQL = ""
        frmCCtas.DatosADevolverBusqueda = "0"
        frmCCtas.Show vbModal
        Set frmCCtas = Nothing
        If SQL <> "" Then
            'TEngo cuenta contable
            Text5.Text = SQL
            SQL = "nommacta"
            Text4.Text = DevuelveDesdeBD("nifdatos", "cuentas", "codmacta", Text5.Text, "T", SQL)
            If Text4.Text = "" Then
                Text5.Text = ""
                MsgBox "La cuenta no tiene NIF.", vbExclamation
            Else
                Text5.Text = SQL
            End If
        End If

End Sub

Private Sub Image4_Click()
    SQL = ""
    cd1.ShowOpen
    If cd1.FileName <> "" Then SQL = cd1.FileName
    If SQL <> "" Then
        If Dir(SQL, vbArchive) = "" Then
            MsgBox "Fichero NO existe", vbExclamation
            SQL = ""
        End If
    End If
    If SQL <> "" Then Text8.Text = SQL
End Sub

Private Sub imgCC_Click(Index As Integer)
    LanzaBuscaGrid 2
    If SQL <> "" Then
        txtCC(Index).Text = SQL
        txtCC_LostFocus Index
    End If
End Sub

Private Sub imgCheck_Click(Index As Integer)

    If Index < 2 Then
        'Selecciona forma de pago
        For I = 1 To Me.lwtipopago.ListItems.Count
            Me.lwtipopago.ListItems(I).Checked = Index = 1
        Next

    ElseIf Index < 4 Then
        'Empresas
         For I = 1 To Me.ListView3.ListItems.Count
            Me.ListView3.ListItems(I).Checked = Index = 3
        Next
    Else
        'Reclamaciones
        If Me.optReclama(1).Value Then
            'Solo en correctos, los incorrectos se iran tooodos
            For I = 1 To Me.ListView6.ListItems.Count
                Me.ListView6.ListItems(I).Checked = Index = 5
            Next
        End If
    End If
End Sub

Private Sub imgcheckall_Click(Index As Integer)
    Cancelado = (Index = 0)
    For I = 1 To ListView4.ListItems.Count
        ListView4.ListItems(I).Checked = Cancelado
    Next I
    Cancelado = False
End Sub

Private Sub imgConcepto_Click(Index As Integer)
  
    LanzaBuscaGrid 1
    If SQL <> "" Then
        txtConcepto(Index).Text = SQL
        txtConcepto_LostFocus Index
    End If
End Sub

Private Sub imgCtaNorma_Click(Index As Integer)

        If Index <> 6 Then

               Set frmCCtas = New frmColCtas
               SQL = ""
               frmCCtas.DatosADevolverBusqueda = "0"
               frmCCtas.Show vbModal
               
               Set frmCCtas = Nothing
               If SQL <> "" Then
                   txtCtaNormal(Index).Text = SQL
                   txtCtaNormal_LostFocus Index
               End If
            
        Else
        
            'Para las cuentas agrupadas
            SQL = ""
            LanzaBuscaGrid 3
            If SQL <> "" Then
                If MsgBox("Va a insetar las cuentas del grupo de tesoreria: " & SQL & vbCrLf & "�Continuar?", vbQuestion + vbYesNo) = vbYes Then
                    Screen.MousePointer = vbHourglass
                    Set miRsAux = New ADODB.Recordset
                    CargaGrupo
                    Set miRsAux = Nothing
                    Screen.MousePointer = vbDefault
                End If
            End If
        End If
            
            
End Sub

Private Sub imgCuentas_Click(Index As Integer)

    imgCuentas(0).Tag = Index
    Set frmBa = New frmBanco
    frmBa.DatosADevolverBusqueda = "OK"
    frmBa.Show vbModal
    Set frmBa = Nothing
End Sub


Private Sub imgDiario_Click(Index As Integer)
  
    LanzaBuscaGrid 0
    If SQL <> "" Then
        txtDiario(Index).Text = SQL
        txtDiario_LostFocus Index
    End If
End Sub

Private Sub imgEliminarCta_Click()
    If List1.SelCount = 0 Then Exit Sub
    
    SQL = "Desea quitar la(s) cuenta(s): " & vbCrLf
    For I = 0 To List1.ListCount - 1
        If List1.Selected(I) Then SQL = SQL & List1.List(I) & vbCrLf
    Next I
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        For I = List1.ListCount - 1 To 0 Step -1
            If List1.Selected(I) Then
                SQL = Trim(Mid(List1.List(I), 1, vEmpresa.DigitosUltimoNivel + 2))
                NumRegElim = InStr(1, CuentasCC, SQL)
                If NumRegElim > 0 Then CuentasCC = Mid(CuentasCC, 1, NumRegElim - 1) & Mid(CuentasCC, NumRegElim + vEmpresa.DigitosUltimoNivel + 1) 'para que quite el pipe final
                List1.RemoveItem I
            End If
        Next I
    
    End If
    NumRegElim = 0
End Sub

Private Sub imgFP_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Set frmP = New frmFormaPago
    I = Index
    frmP.DatosADevolverBusqueda = "0|1"
    frmP.Show vbModal
    Set frmP = Nothing
End Sub

Private Sub imgFra_Click()
        CadenaDesdeOtroForm = ""
        SQL = ""
        If txtCtaNormal(11).Text <> "" Then SQL = "scobro.codmacta = '" & txtCtaNormal(11).Text & "'"
        frmTESVerCobrosPagos.vSql = SQL
        frmTESVerCobrosPagos.OrdenarEfecto = False
        frmTESVerCobrosPagos.Regresar = True
        frmTESVerCobrosPagos.Cobros = True
        frmTESVerCobrosPagos.Show vbModal
        If CadenaDesdeOtroForm <> "" Then

            txtSerie(4).Text = RecuperaValor(CadenaDesdeOtroForm, 1)
            txtNumFac(4).Text = RecuperaValor(CadenaDesdeOtroForm, 2)
            Me.txtNumero.Text = RecuperaValor(CadenaDesdeOtroForm, 4)
            PonerFoco Text1(11)
        End If
        CadenaDesdeOtroForm = ""
End Sub

Private Sub ListView2_DblClick()
  '  Stop
  '  For NumRegElim = 1 To ListView2.ColumnHeaders.Count: Debug.Print ListView2.ColumnHeaders(NumRegElim).Text & ": " & ListView2.ColumnHeaders(NumRegElim).Width: Next NumRegElim
End Sub

Private Sub ListView4_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Item.Checked = Not Item.Checked
End Sub

Private Sub optDevRem_Click(Index As Integer)
        txtImporte(2).Visible = Index = 2
        Label4(8).Visible = Index = 2
        If Index <> 0 Then
            Label4(9).Caption = "%"
        Else
            Label4(9).Caption = "�uros"
        End If
End Sub

Private Sub optDevRem_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optOrdCob_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optOrdPag_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optReclama_Click(Index As Integer)
    ReclamacionGargarList
    cmdEliminarReclama.Visible = Index = 1
End Sub

Private Sub optSepaXML_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).Text = "" Then Exit Sub
    
    If Not EsFechaOK(Text1(Index)) Then
        MsgBox "Fecha incorrecta: " & Text1(Index).Text, vbExclamation
        Text1(Index).Text = ""
        PonerFoco Text1(Index)
    End If
    
End Sub



Private Sub CargaList()
    


        SQL = DevuelveDesdeBD("descformapago", "stipoformapago", "tipoformapago", CStr(SubTipo), "N")
        Text2(Opcion).Text = SQL
                
        
End Sub


Private Sub Text3_GotFocus(Index As Integer)
    With Text3(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Text3_LostFocus(Index As Integer)
    With Text3(Index)
        .Text = Trim(.Text)
        If .Text = "" Then Exit Sub
        
        If Not IsNumeric(.Text) Then
            MsgBox "Campo debe ser num�rico: " & .Text, vbExclamation
            .Text = ""
            PonerFoco Text3(Index)
        End If
        
        'Para que vaya a la tabal y traiga datos remesa
        If Index = 3 Or Index = 4 Then CamposRemesaAbono
    End With
End Sub


Private Sub Text4_GotFocus()
    ObtenerFoco Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text4_LostFocus()
    Text4.Text = Trim(Text4.Text)
    If Text4.Text = "" Then
        Text5.Text = ""
        Exit Sub
    End If
    
    SQL = DevuelveDesdeBD("nommacta", "cuentas", "nifdatos", Text4.Text, "T")
    If SQL = "" Then
        MsgBox "NIF no encontrado", vbExclamation
        Text5.Text = ""
        PonerFoco Text4
    End If
    
    Text5.Text = SQL
    
End Sub

Private Sub Text6_GotFocus()
    ObtenerFoco Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub









Private Sub Text7_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCC_GotFocus(Index As Integer)
    ObtenerFoco txtCC(Index)
End Sub

Private Sub txtCC_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCC_LostFocus(Index As Integer)
    txtCC(Index).Text = Trim(txtCC(Index).Text)
    SQL = ""
    I = 0
    If txtCC(Index).Text <> "" Then
            
        SQL = DevuelveDesdeBD("nomccost", "cabccost", "codccost", txtCC(Index).Text, "T")
        If SQL = "" Then
            MsgBox "Concepto no existe", vbExclamation
            I = 1
        End If

    End If
    Me.txtDCC(Index).Text = SQL
    If I = 1 Then
        txtCC(Index).Text = ""
        PonerFoco txtCC(Index)
    End If

End Sub

Private Sub txtConcepto_GotFocus(Index As Integer)
    ObtenerFoco txtConcepto(Index)
End Sub

Private Sub txtConcepto_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtConcepto_LostFocus(Index As Integer)
    'Lost focus
    txtConcepto(Index).Text = Trim(txtConcepto(Index).Text)
    SQL = ""
    I = 0
    If txtConcepto(Index).Text <> "" Then
        If Not IsNumeric(txtConcepto(Index).Text) Then
            MsgBox "Campo num�rico", vbExclamation
            I = 1
        Else
            
            SQL = DevuelveDesdeBD("nomconce", "conceptos", "codconce", txtConcepto(Index).Text, "N")
            If SQL = "" Then
                MsgBox "Concepto no existe", vbExclamation
                I = 1
            End If
        End If
    End If
    Me.txtDConcpeto(Index).Text = SQL
    If I = 1 Then
        txtConcepto(Index).Text = ""
        PonerFoco txtConcepto(Index)
    End If
End Sub

Private Sub txtCta_GotFocus(Index As Integer)
    ObtenerFoco txtCta(Index)
End Sub

Private Sub txtCta_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtCta_LostFocus(Index As Integer)
Dim DevfrmCCtas As String

        txtCta(Index).Text = Trim(txtCta(Index).Text)
        DevfrmCCtas = txtCta(Index).Text
        I = 0
        If DevfrmCCtas <> "" Then
            If CuentaCorrectaUltimoNivel(DevfrmCCtas, SQL) Then
                DevfrmCCtas = DevuelveDesdeBD("codmacta", "ctabancaria", "codmacta", DevfrmCCtas, "T")
                If DevfrmCCtas = "" Then
                    SQL = ""
                    MsgBox "La cuenta contable no esta asociada a ninguna cuenta bancaria", vbExclamation
                End If
            Else
                MsgBox SQL, vbExclamation
                DevfrmCCtas = ""
                SQL = ""
            End If
            I = 1
        Else
            SQL = ""
        End If
        
        
        txtCta(Index).Text = DevfrmCCtas
        txtDescCta(Index).Text = SQL
        If DevfrmCCtas = "" And I = 1 Then

            PonerFoco txtCta(Index)
        End If

        
End Sub



Private Function CopiarArchivo() As Boolean
On Error GoTo ECopiarArchivo

    CopiarArchivo = False
    'cd1.CancelError = True
    cd1.FileName = ""
    cd1.ShowSave
    If cd1.FileName <> "" Then
    
        If Dir(cd1.FileName, vbArchive) <> "" Then
            If MsgBox("El archivo " & cd1.FileName & " ya existe" & vbCrLf & vbCrLf & "�Sobreescribir?", vbQuestion + vbYesNo) = vbNo Then Exit Function
            Kill cd1.FileName
        End If
        'Hacemos la copia
        FileCopy SQL, cd1.FileName
        CopiarArchivo = True
    End If
    
    
    Exit Function
ECopiarArchivo:
    MuestraError Err.Number, "Copiar Archivo"
End Function







Private Sub txtCtaNormal_GotFocus(Index As Integer)
    ObtenerFoco txtCtaNormal(Index)
End Sub
    
Private Sub txtCtaNormal_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCtaNormal_LostFocus(Index As Integer)
Dim DevfrmCCtas As String
       
        DevfrmCCtas = Trim(txtCtaNormal(Index).Text)
        I = 0
        If DevfrmCCtas <> "" Then
            If CuentaCorrectaUltimoNivel(DevfrmCCtas, SQL) Then
                
            Else
                MsgBox SQL, vbExclamation
                If Index < 3 Or Index = 9 Or Index = 10 Or Index = 11 Then
                    DevfrmCCtas = ""
                    SQL = ""
                End If
            End If
            I = 1
        Else
            SQL = ""
        End If
        
        
        txtCtaNormal(Index).Text = DevfrmCCtas
        txtDCtaNormal(Index).Text = SQL
        If DevfrmCCtas = "" And I = 1 Then
            PonerFoco txtCtaNormal(Index)
        End If
        VisibleCC
    
        
        If Index = 10 Then
            FrameDocPorveedor.Visible = False
            If SubTipo = 2 Or SubTipo = 3 Then
                FrameDocPorveedor.Visible = SQL <> ""
                If SQL = "" Then
                    txtTexto(2).Text = ""
                    txtTexto(3).Text = ""
                End If
            End If
        
        End If
End Sub

Private Sub txtDiario_GotFocus(Index As Integer)
       ObtenerFoco txtDiario(Index)
End Sub

Private Sub txtDiario_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtDiario_LostFocus(Index As Integer)
    'Lost focus
    txtDiario(Index).Text = Trim(txtDiario(Index).Text)
    SQL = ""
    I = 0
    If txtDiario(Index).Text <> "" Then
        If Not IsNumeric(txtDiario(Index).Text) Then
            MsgBox "Campo num�rico", vbExclamation
            I = 1
        Else
            
            SQL = DevuelveDesdeBD("desdiari", "tiposdiario", "numdiari", txtDiario(Index).Text, "N")
            If SQL = "" Then
                MsgBox "Diario no existe", vbExclamation
                I = 1
            End If
        End If
    End If
    Me.txtDDiario(Index).Text = SQL
    If I = 1 Then
        txtDiario(Index).Text = ""
        PonerFoco txtDiario(Index)
    End If
            
   
End Sub



Private Sub txtFP_GotFocus(Index As Integer)
    ObtenerFoco txtFP(Index)
End Sub

Private Sub txtFP_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtFP_LostFocus(Index As Integer)
    SQL = ""
    txtFP(Index).Text = Trim(txtFP(Index).Text)
    If txtFP(Index).Text <> "" Then
        If Not IsNumeric(txtFP(Index).Text) Then
            MsgBox "Campo debe ser numerico: " & txtFP(Index).Text, vbExclamation
            txtFP(Index).Text = ""
        Else
            SQL = DevuelveDesdeBD("nomforpa", "sforpa", "codforpa", txtFP(Index).Text)
            If SQL = "" Then SQL = "NO existe la forma de pago"
        End If
    End If
    Me.txtFPDesc(Index).Text = SQL
End Sub

Private Sub txtImporte_GotFocus(Index As Integer)
    With txtImporte(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtImporte_LostFocus(Index As Integer)
 Dim Valor
        txtImporte(Index).Text = Trim(txtImporte(Index))
        If txtImporte(Index).Text = "" Then Exit Sub
        

        If Not EsNumerico(txtImporte(Index).Text) Then
            txtImporte(Index).Text = ""
            Exit Sub
        End If
    
        
        If Index = 6 Or Index = 7 Then
           
            If InStr(1, txtImporte(Index).Text, ",") > 0 Then
                Valor = ImporteFormateado(txtImporte(Index).Text)
            Else
                Valor = CCur(TransformaPuntosComas(txtImporte(Index).Text))
            End If
            txtImporte(Index).Text = Format(Valor, FormatoImporte)
        End If
        
End Sub





Private Sub CargaImpagados()

    SQL = "Select fechadev,gastodev from sefecdev  WHERE numserie='" & RecuperaValor(CadenaDesdeOtroForm, 1)
    SQL = SQL & "' AND  codfaccl =  " & RecuperaValor(CadenaDesdeOtroForm, 2)
    SQL = SQL & "  AND  fecfaccl =  '" & Format(RecuperaValor(CadenaDesdeOtroForm, 3), FormatoFecha)
    SQL = SQL & "' AND  numorden =  " & RecuperaValor(CadenaDesdeOtroForm, 4)
    SQL = SQL & " order by fechadev"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set IT = ListView1.ListItems.Add
        IT.Text = Format(RS!fechadev, "dd/mm/yyyy")
        IT.SubItems(1) = Format(RS!gastodev, FormatoImporte)
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub


Private Sub CargaImagen()
On Error Resume Next
    Image2.Picture = LoadPicture(App.Path & "\minilogo.bmp")
    'Image1.Picture = LoadPicture(App.path & "\fondon.gif")
    Err.Clear
End Sub


Private Sub CargaRemesas()
    
    ListView2.ListItems.Clear
    
    If SubTipo > 2 Then
        CargaRemes 3  'Cargamos todo
        CargaRemes 2  'Cargamos todo
    Else
        CargaRemes SubTipo
    End If
    
    
End Sub


Private Sub CargaRemes(SubT As Byte)
Dim F As Date
Dim Dias As Integer

    On Error GoTo EC
    
    
    
    
 
    ' 3 es que esta cargando todo
    If SubT = 1 Or SubT = 3 Then
        'Efectos
        '
        SQL = "Select codigo,anyo, fecremesa,"
        If SubT = 3 Then
            SQL = SQL & " tiporemesa2.descripciont "
        Else
            SQL = SQL & " tiporemesa."
        End If
        SQL = SQL & "descripcion,descsituacion,remesas.codmacta,nommacta,remesadiasmenor, remesadiasmayor, "
        SQL = SQL & "Importe , remesas.descripcion as Desc1, remesas.Tipo,situacion,Tiporem from cuentas,tiposituacionrem,ctabancaria,"
        SQL = SQL & "remesas left join tiporemesa"
        If SubT = 3 Then SQL = SQL & "2" 'Para que carge, en lugar de norma19 norma52 etc que carge efectos, talon, pagare
        SQL = SQL & " on remesas.tipo"
        If SubT = 3 Then SQL = SQL & "rem"
        SQL = SQL & "=tiporemesa"
        If SubT = 3 Then SQL = SQL & "2" 'Para que carge, en lugar de norma19 norma52 etc que carge efectos, talon, pagare
        SQL = SQL & ".tipo where remesas.codmacta=cuentas.codmacta and situacio=situacion and ctabancaria.codmacta=remesas.codmacta"
        SQL = SQL & " AND tiporem = 1 "   'Efectos
        'Solo borrare las contabilizadas o pendientes de eliminar tooodos los efectos
        SQL = SQL & " AND (situacion ='Q' or situacion ='Y')"
                
        
    Else
        'Talones Remesesas
        SQL = "Select codigo,anyo, fecremesa,tiporemesa2.descripciont descripcion,descsituacion,remesas.codmacta,nommacta,talondias,pagaredias, "
        SQL = SQL & "Importe , remesas.descripcion as Desc1, remesas.Tipo,situacion,Tiporem from cuentas,tiposituacionrem,ctabancaria,"
        SQL = SQL & "remesas left join tiporemesa2 on remesas.tiporem=tiporemesa2.tipo "
        SQL = SQL & "where remesas.codmacta=cuentas.codmacta and situacio=situacion and ctabancaria.codmacta=remesas.codmacta"
        SQL = SQL & " AND tiporem > 1 "   'Pagares remesas
       'Solo borrare las contabilizadas o pendientes de eliminar tooodos los efectos
        SQL = SQL & " AND (situacion ='Q' or situacion ='Y')"
    
    End If
    
    SQL = SQL & " ORDER BY anyo,codigo"   'Solo borrare las contabilizadas
    Set RS = New ADODB.Recordset
    
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        'Ver los dias
        If SubT = 1 Or SubT = 3 Then
            'Efectos recibos
            Dias = DBLet(RS!remesadiasmenor, "N")
            I = DBLet(RS!remesadiasmayor, "N")
            If I < Dias And I > 0 Then Dias = I
        Else
            If RS!Tiporem = 2 Then
                'Pagare
                Dias = DBLet(RS!pagaredias, "N")
            Else
                'talon
                Dias = DBLet(RS!talondias, "N")
            End If
            
        End If
        F = RS!fecremesa
        
        If SubT = 2 Then
            'If RS!Codigo > 159 Then Stop
            SQL = "anyorem=" & RS!Anyo & " AND codrem "
            SQL = DevuelveDesdeBD("min(fecvenci)", "scobro", SQL, RS!Codigo, "N")
            If SQL <> "" Then
                If CDate(SQL) > F Then F = SQL
            End If
        End If
        
        F = DateAdd("d", Dias, F)
        If F < Now Then
            Set IT = ListView2.ListItems.Add
            IT.Text = RS!Anyo
            IT.SubItems(1) = RS!Codigo
            IT.SubItems(2) = RS!Descripcion
            IT.SubItems(3) = RS!fecremesa
            IT.SubItems(4) = RS!codmacta
            IT.SubItems(5) = RS!Nommacta
            IT.SubItems(6) = Format(RS!Importe, FormatoImporte)
            IT.SubItems(7) = RS!Desc1
            IT.Tag = RS!Tiporem
        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    Exit Sub
EC:
    MuestraError Err.Number, "Cargando vencimientos"
End Sub



'
Public Function GeneraCobrosPagosNIF() As Boolean
Dim cad As String
Dim L As Long
Dim Empre As String
Dim Importe  As Currency

Dim QueTipoPago As String

    'Guardaremos en la variable QueTipoPago que tipos de pago ha seleccionado
    'Si selecciona todos los tipos de pago NO pondremos el IN en el select
    QueTipoPago = ""
    cad = "" 'para saber si ha selccionado todos
    For L = 1 To Me.lwtipopago.ListItems.Count
        If lwtipopago.ListItems(L).Checked Then
            QueTipoPago = QueTipoPago & ", " & Me.lwtipopago.ListItems(L).Tag
        Else
            cad = "NO" 'No estan todos seleccionados
        End If
    Next
    If cad = "" Then
        'Estan todos. No tiene sentido hacer el Select in
        QueTipoPago = ""
    Else
        QueTipoPago = Mid(QueTipoPago, 2)
    End If
    
    
    
'En la tabla  INSERT INTO tmp347 (codusu, cliprov, cta, nif) VALUES ((
' Tendremos codccos: la empresa
'                  : cta, cada uno de los valores
'INSERT INTO ztesoreriacomun (codusu, codigo, texto1, texto2, texto3, texto4,
'texto5, texto6, importe1, importe2, fecha1, fecha2, fecha3, observa1,
'observa2, opcion) VALUES
    GeneraCobrosPagosNIF = False
    L = 1
    SQL = "Select * from tmp347 where codusu =" & vUsu.Codigo & " ORDER BY cliprov,cta"
    Set RS = New ADODB.Recordset
    Set miRsAux = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not RS.EOF
        If Cancelado Then
            RS.Close
            Exit Function
        End If
        'Los labels
        Label9.Caption = "Nif: " & RS!NIF & " - " & RS!Cta
        Label9.Refresh
        
        'SQL insert
        SQL = "INSERT INTO Usuarios.ztesoreriacomun (codusu,texto1, codigo,texto2,  texto3,texto4, texto5,fecha1,fecha2,"   'texto5, texto6,
        SQL = SQL & " importe1, importe2,opcion"
        SQL = SQL & ") VALUES ("
        'NIF      Nombre
        SQL = SQL & vUsu.Codigo & ",'" & RS!NIF & "',"
        
        
        '-------
        Empre = DameEmpresa(CStr(RS!cliprov))
        
        'COBROS
        cad = "Select fecfaccl,numserie,codfaccl, numorden,impvenci,impcobro,gastos,fecvenci,nommacta from conta" & RS!cliprov & ".scobro as c1,"
        cad = cad & "conta" & RS!cliprov & ".cuentas as c2 "
        If QueTipoPago <> "" Then cad = cad & ", conta" & RS!cliprov & ".sforpa as sforpa"
        cad = cad & " where c1.codmacta = c2.codmacta AND c1.codmacta='" & RS!Cta & "'"
        If QueTipoPago <> "" Then cad = cad & " AND c1.codforpa=sforpa.codforpa AND sforpa.tipforpa in (" & QueTipoPago & ")"
        'Fechas
        If Text1(12).Text <> "" Then cad = cad & " AND fecvenci >='" & Format(Text1(12).Text, FormatoFecha) & "'"
        If Text1(13).Text <> "" Then cad = cad & " AND fecvenci <='" & Format(Text1(13).Text, FormatoFecha) & "'"
        
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            'Los label
            If Cancelado Then
                miRsAux.Close
                Exit Function
            End If
            
            'Insetamos codigo,  texto3
            '                    empresa
            cad = L & ",'" & Empre & "','"
            cad = cad & miRsAux!NUmSerie & "/" & Format(miRsAux!codfaccl, "0000000000") & " : " & miRsAux!numorden & "','"
            cad = cad & RS!Cta & "','"
            cad = cad & DevNombreSQL(miRsAux!Nommacta) & "','"
            'texto4: fecha
            cad = cad & Format(miRsAux!fecfaccl, FormatoFecha) & "','"
            cad = cad & Format(miRsAux!FecVenci, FormatoFecha) & "',"
            
            
            'En importe1 estara el importe del cobro. En el 2 tb
'            Importe = DBLet(miRsAux!Gastos, "N") - DBLet(miRsAux!impcobro, "N")
'            Importe = Importe + miRsAux!impvenci
'             Cad = Cad & TransformaComasPuntos(CStr(Importe)) & "," & TransformaComasPuntos(CStr(Importe))


            Importe = DBLet(miRsAux!Gastos, "N")
            cad = cad & TransformaComasPuntos(CStr(Importe))
            Importe = miRsAux!ImpVenci - DBLet(miRsAux!impcobro, "N")
            cad = cad & "," & TransformaComasPuntos(CStr(Importe))
           
            
            
            'un cero para importe 2  y un cero para la opcion
            cad = cad & ",0)"
            
            'Ejecutamos
            cad = SQL & cad
            Ejecuta cad
            
            L = L + 1
            miRsAux.MoveNext
            DoEvents
        Wend
        miRsAux.Close
        
        'PAGOS
        cad = "Select numfactu,numorden,fecfactu,imppagad,fecefect,impefect,nommacta from conta" & RS!cliprov & ".spagop ,conta" & RS!cliprov & ".cuentas "
        If QueTipoPago <> "" Then cad = cad & ", conta" & RS!cliprov & ".sforpa as sforpa"
        cad = cad & " where ctaprove = codmacta AND ctaprove='" & RS!Cta & "'"
        If QueTipoPago <> "" Then cad = cad & " AND spagop.codforpa=sforpa.codforpa AND sforpa.tipforpa in (" & QueTipoPago & ")"
        
        
        'Fechas
        If Text1(12).Text <> "" Then cad = cad & " AND fecefect >='" & Format(Text1(12).Text, FormatoFecha) & "'"
        If Text1(13).Text <> "" Then cad = cad & " AND fecefect <='" & Format(Text1(13).Text, FormatoFecha) & "'"
        miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRsAux.EOF
            'Los label
            If Cancelado Then
                miRsAux.Close
                Exit Function
            End If
            
            'Insetamos codigo,  texto3,t5
            '                    empresa
            cad = L & ",'" & Empre & "','"
            cad = cad & DevNombreSQL(miRsAux!NumFactu) & " : " & miRsAux!numorden & "','"
            cad = cad & RS!Cta & "','"
            cad = cad & DevNombreSQL(miRsAux!Nommacta) & "','"
            ' fecha1 y 2
            cad = cad & Format(miRsAux!FecFactu, FormatoFecha) & "','"
            cad = cad & Format(miRsAux!fecefect, FormatoFecha) & "',"
            
            
            'En importe1 estara el importe del cobro
            Importe = DBLet(miRsAux!imppagad, "N")

            Importe = miRsAux!ImpEfect - Importe
            cad = cad & TransformaComasPuntos(CStr(0)) & "," & TransformaComasPuntos(CStr(-1 * Importe))
            
            cad = cad & ",1)" '1: pago
            
            
            
            
            'Ejecutamos
            cad = SQL & cad
            Ejecuta cad
            
            L = L + 1
            miRsAux.MoveNext
            
            DoEvents
        Wend
        miRsAux.Close
        
        
        'SIGUIENTE CUENTA
        RS.MoveNext
    Wend
    RS.Close
    
    cad = "DELETE FROM usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo & " AND importe1+importe2=0"
    Conn.Execute cad
    
    cad = "select count(*) from usuarios.ztesoreriacomun where codusu = " & vUsu.Codigo
    RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    L = 0
    If Not RS.EOF Then
        L = DBLet(RS.Fields(0), "N")
    End If
    RS.Close
    
    Set RS = Nothing
    Set miRsAux = Nothing
    
    If L = 0 Then
        'ERROR. MO HAY DATOS
        MsgBox "Sin datos.", vbExclamation
    Else
        GeneraCobrosPagosNIF = True
    End If
        
End Function



Private Function DameEmpresa(ByVal S As String) As String
    DameEmpresa = "NO ENCONTRADA"
    For I = 1 To ListView3.ListItems.Count
        If ListView3.ListItems(I).Tag = S Then
            DameEmpresa = DevNombreSQL(ListView3.ListItems(I).Text)
            Exit For
        End If
    Next I
    
End Function






Private Sub cargaTipoPagos()
    'FALTARA VER LO DE QUITAR EMPRESAS NO PERMITIDAS
 
    lwtipopago.ListItems.Clear
    SQL = "select tipoformapago,descformapago,siglas from stipoformapago order by tipoformapago"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set IT = lwtipopago.ListItems.Add
        IT.Key = "C" & miRsAux!tipoformapago
        IT.Text = miRsAux!descformapago
      '  IT.SubItems(1) = miRsAux!siglas
        IT.Tag = miRsAux!tipoformapago
        
        If miRsAux!tipoformapago > 0 Then IT.Checked = True  'menos el efectivo  todas
         
        miRsAux.MoveNext
        
    Wend
    miRsAux.Close
    Set miRsAux = New ADODB.Recordset
End Sub



Private Sub CargaCtasparaAgruparNIF()
    I = 0
    SQL = "select cuentas.codmacta,nifdatos from scobro,cuentas where scobro.codmacta=cuentas.codmacta"
    SQL = SQL & " and not (nifdatos is null)  "
    If txtCtaNormal(1).Text <> "" Then SQL = SQL & " and cuentas.codmacta >= '" & txtCtaNormal(1).Text & "'"
    If txtCtaNormal(2).Text <> "" Then SQL = SQL & " and cuentas.codmacta <= '" & txtCtaNormal(2).Text & "'"
    SQL = SQL & " group by  codmacta,nifdatos"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF
        If Cancelado Then
            miRsAux.Close
            Exit Sub
        End If
        SQL = "INSERT INTO tmpfaclin (codusu, codigo, NIF) VALUES (" & vUsu.Codigo & "," & I & ",'" & miRsAux!nifdatos & "')"
        Ejecuta SQL
        miRsAux.MoveNext
        DoEvents
        I = I + 1
    Wend
    miRsAux.Close
    If Cancelado Then Exit Sub
    'AHora los nifs en los pagos
    SQL = "select cuentas.codmacta,nifdatos from spagop,cuentas where ctaprove=cuentas.codmacta"
    SQL = SQL & " and not (nifdatos is null) "
    If txtCtaNormal(1).Text <> "" Then SQL = SQL & " and cuentas.codmacta >= '" & txtCtaNormal(1).Text & "'"
    If txtCtaNormal(2).Text <> "" Then SQL = SQL & " and cuentas.codmacta <= '" & txtCtaNormal(2).Text & "'"
    
    SQL = SQL & " group by  codmacta,nifdatos"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF
        If Cancelado Then
            miRsAux.Close
            Exit Sub
        End If
        SQL = "INSERT INTO tmpfaclin (codusu, codigo, NIF) VALUES (" & vUsu.Codigo & "," & I & ",'" & miRsAux!nifdatos & "')"
        Ejecuta SQL
        miRsAux.MoveNext
        I = I + 1
        DoEvents
    Wend
    
    miRsAux.Close
    If Cancelado Then Exit Sub
    
    'Ahora cargaremos la tabla tmp347 que tendra las cuentas
    'Para cada NIF generaremos sus datos, con las empresas
    SQL = "Select nif from tmpfaclin where codusu =" & vUsu.Codigo & " group by nif"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Label9.Caption = "Nif: " & miRsAux!NIF
        Label9.Refresh

        For I = 1 To ListView3.ListItems.Count
            If ListView3.ListItems(I).Checked Then
                If Cancelado Then
                    miRsAux.Close
                    Exit Sub
                End If
                SQL = "Select " & vUsu.Codigo & "," & Mid(ListView3.ListItems(I).Key, 2) & ",codmacta,'" & miRsAux!NIF & "'"
                SQL = "INSERT INTO tmp347 (codusu, cliprov, cta, nif) " & SQL
                SQL = SQL & " FROM Conta" & ListView3.ListItems(I).Tag & ".cuentas WHERE nifdatos = '" & miRsAux!NIF & "' ORDER BY codmacta"
                If Not Ejecuta(SQL) Then Exit Sub
            
                DoEvents
            
            End If
        Next I
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    Label9.Caption = "Cuentas obtenidas. Leyendo BD"
    Me.Refresh
    espera 0.5
    
End Sub




Private Sub CargaGastos()
Dim Importe As Currency
    Label11.Caption = "Cargando datos"
    Label11.Refresh


    'ESTO ES UN POCO MARCIANO
    '-------------------------------------------------
    '
    ' El recodset mirsaux  viene cargado desde la fase anterior
    ' De ese modo, con una unica .open . Si no es EOF lanzamos esta pantalla
    ' si es EOF ni nos molestamos en abrirla

    While Not miRsAux.EOF
        Set IT = ListView4.ListItems.Add()
        IT.Text = miRsAux!Descripcion
        IT.SubItems(1) = Format(miRsAux!Fecha, "dd/mm/yyyy")
        IT.SubItems(2) = Format(miRsAux!Importe, FormatoImporte)
        IT.Checked = True
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    Label11.Caption = ""
    
    
    
End Sub

Private Sub CargaDatosContabilizarGastos()
    txtCta(6).Text = RecuperaValor(CadenaDesdeOtroForm, 3)
    txtDescCta(6).Text = RecuperaValor(CadenaDesdeOtroForm, 4)
    txtCtaNormal(0).Text = RecuperaValor(CadenaDesdeOtroForm, 5)
    txtDCtaNormal(0).Text = RecuperaValor(CadenaDesdeOtroForm, 6)
    Text9.Text = RecuperaValor(CadenaDesdeOtroForm, 2)
    'Fecha e Importe
    SQL = RecuperaValor(CadenaDesdeOtroForm, 7)
    I = InStr(8, SQL, " ")
    Text1(19).Text = Trim(Mid(SQL, 1, I))
    txtImporte(3).Text = Trim(Mid(SQL, I))
    'ASignaremos cadenadesdeotroform el valor para hacer el UPDATE del registro SI se contabiliza
    SQL = RecuperaValor(CadenaDesdeOtroForm, 1) & "|"
    CadenaDesdeOtroForm = SQL & Text1(19).Text & "|" & Text9.Text & "|"
    
    VisibleCC
End Sub

Private Sub PonerCuentasCC()

    CuentasCC = ""
    If vParam.autocoste Then
        SQL = "Select * from parametros"
        Set miRsAux = New ADODB.Recordset
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        'NO PUEDE SER EOF
        CuentasCC = "|" & miRsAux!grupogto & "|" & miRsAux!grupovta & "|"
        miRsAux.Close
        Set miRsAux = Nothing
    End If
End Sub


Private Sub VisibleCC()
Dim B As Boolean

    B = False
    If vParam.autocoste Then
        If txtCtaNormal(0).Text <> "" Then
                SQL = "|" & Mid(txtCtaNormal(0).Text, 1, 1) & "|"
                If InStr(1, CuentasCC, SQL) > 0 Then B = True
        End If
    End If
    Label1(14).Visible = B
    txtCC(0).Visible = B
    txtDCC(0).Visible = B
    imgCC(0).Visible = B
End Sub



Private Sub LanzaBuscaGrid(Opcion As Integer)

'No tocar variable SQL
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String


'--monica
'
'
'    SQL = ""
'    Screen.MousePointer = vbHourglass
'    Set frmB = New frmBuscaGrid
'    frmB.vSQL = ""
'
'    '###A mano
'    frmB.vDevuelve = "0|"   'Siempre el 0
'
'    frmB.vSelElem = 0
'
'    'Ejemplo
'        'Cod Diag.|idDiag|N|10�
'        Select Case Opcion
'        Case 0
'            'DIARIO
'            cad = "Codigo|numdiari|N|15�"
'            cad = cad & "Descripcion|desdiari|T|60�"
'            frmB.vTabla = "tiposdiario"
'            frmB.vTitulo = "Diario"
'        Case 1
'            'CONCEPTO
'            cad = "Codigo|codconce|N|15�"
'            cad = cad & "Descripcion|nomconce|T|60�"
'            frmB.vTabla = "Conceptos"
'            frmB.vTitulo = "Conceptos"
'
'            frmB.vSQL = " codconce <900"
'
'        Case 2
'            'CC
'            cad = "Codigo|codccost|N|15�"
'            cad = cad & "Descripcion|nomccost|T|60�"
'            frmB.vTabla = "cabccost"
'            frmB.vTitulo = "Centros de coste"
'
'        Case 3
'            'Cuentas agrupadas bajo el concepto: grupotesoreria
'            cad = "Grupo tesoreria|grupotesoreria|T|60�"
'            frmB.vTabla = "cuentas"
'            frmB.vSQL = " grupotesoreria <> '' GROUP BY 1"
'            frmB.vTitulo = "Cuentas grupos tesoreria"
'        End Select
'
'
'        frmB.vCampos = cad
'
'
'
'
'
''        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'
'
'    Screen.MousePointer = vbDefault
End Sub




Private Function ContabilizarGastoFijo() As Boolean
Dim Mc As Contadores
Dim FechaAbono As Date
Dim Importe As Currency
    On Error GoTo EContabilizarGastoFijo
    ContabilizarGastoFijo = False
    Set Mc = New Contadores
    
    FechaAbono = CDate(Text1(19).Text)
    If Mc.ConseguirContador("0", FechaAbono <= vParam.fechafin, True) = 1 Then Exit Function
   
    
    
    'Insertamos la cabera
    SQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, bloqactu, numaspre, obsdiari) VALUES ("
    SQL = SQL & txtDiario(0).Text & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador
    SQL = SQL & ", 1, NULL, '"
    SQL = SQL & "Gasto fijo : " & RecuperaValor(CadenaDesdeOtroForm, 1) & " - " & DevNombreSQL(RecuperaValor(CadenaDesdeOtroForm, 3)) & vbCrLf
    SQL = SQL & "Generado desde Tesoreria el " & Format(Now, "dd/mm/yyyy") & " por " & DevNombreSQL(vUsu.Nombre) & "');"
    If Not Ejecuta(SQL) Then Exit Function
    
    If InStr(1, txtImporte(3).Text, ",") > 0 Then
        'Texto formateado
        Importe = ImporteFormateado(txtImporte(3).Text)
    Else
        Importe = CCur(TransformaPuntosComas(txtImporte(3).Text))
    End If
    I = 1
    Do
        'Lineas de apuntes .
         SQL = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, "
         SQL = SQL & "codmacta, numdocum, codconce, ampconce,timporteD,"
         SQL = SQL & " timporteH, ctacontr, codccost,idcontab, punteada) "
         SQL = SQL & "VALUES (" & txtDiario(0).Text & ",'" & Format(FechaAbono, FormatoFecha) & "'," & Mc.Contador & "," & I & ",'"
         
         'Cuenta
         If I = 1 Then
            SQL = SQL & txtCtaNormal(0).Text
         Else
            SQL = SQL & txtCta(6).Text
        End If
        SQL = SQL & "','" & Format(Val(RecuperaValor(CadenaDesdeOtroForm, 1)), "000000000") & "'," & txtConcepto(0).Text & ",'"
        
        'Ampliacion
        SQL = SQL & DevNombreSQL(Mid(txtDConcpeto(0).Text & " " & Text9.Text, 1, 30)) & "',"
                        
        If I = 1 Then
            SQL = SQL & TransformaComasPuntos(CStr(Importe)) & ",NULL,'"
            'Contrapar
            SQL = SQL & txtCta(6).Text
        Else
            SQL = SQL & "NULL," & TransformaComasPuntos(CStr(Importe)) & ",'"
            'Contrpar
            SQL = SQL & txtCtaNormal(0).Text
        End If
        
        'Solo para la line NO banco
        If I = 1 And txtCC(0).Visible Then
            SQL = SQL & "','" & txtCC(0).Text & "'"
        Else
            SQL = SQL & "',NULL"
        End If
        SQL = SQL & ",'CONTAB',0)"
        
        If Not Ejecuta(SQL) Then Exit Function
        I = I + 1
    Loop Until I > 2  'Una para el banoc, otra para la cuenta
   
    
    'Insertamos para pasar a hco
    InsertaTmpActualizar Mc.Contador, txtDiario(0).Text, FechaAbono
    
    
    
    

    'AHora actualizamos el gasto
    FechaAbono = RecuperaValor(CadenaDesdeOtroForm, 2)
    SQL = "UPDATE sgastfijd SET"
    SQL = SQL & " contabilizado=1"
    SQL = SQL & " WHERE codigo=" & RecuperaValor(CadenaDesdeOtroForm, 1)
    SQL = SQL & " and fecha='" & Format(FechaAbono, FormatoFecha) & "'"
    Conn.Execute SQL


    
    
    ContabilizarGastoFijo = True
    Exit Function
EContabilizarGastoFijo:
    MuestraError Err.Number, "Contabilizar Gasto Fijo"
End Function




'------------------------------------------------------------
'Empresas prohibidas
Private Sub cargaempresas()
Dim Prohibidas As String

On Error GoTo Ecargaempresas

    VerEmresasProhibidas Prohibidas
    
    SQL = "Select * from Usuarios.Empresas where tesor=1 order by codempre"
    Set lwE.SmallIcons = frmMantenusu.ImageList1
    lwE.ListItems.Clear
    Set RS = New ADODB.Recordset
    I = -1
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        SQL = "|" & RS!codempre & "|"
        If InStr(1, Prohibidas, SQL) = 0 Then
            Set IT = lwE.ListItems.Add(, , RS!nomempre, , 3)
            IT.Tag = RS!codempre
            If IT.Tag = vEmpresa.codempre Then
                IT.Checked = True
                I = IT.Index
            End If
            IT.ToolTipText = RS!CONTA
        End If
        RS.MoveNext
    Wend
    RS.Close
    If I > 0 Then Set lwE.SelectedItem = lwE.ListItems(I)
Ecargaempresas:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos empresas"
    Set RS = Nothing
End Sub

Private Sub VerEmresasProhibidas(ByRef VarProhibidas As String)

On Error GoTo EVerEmresasProhibidas
    VarProhibidas = "|"
    SQL = "Select codempre from Usuarios.usuarioempresasariconta WHERE codusu = " & (vUsu.Codigo Mod 1000)
    SQL = SQL & " order by codempre"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
          VarProhibidas = VarProhibidas & RS!codempre & "|"
          RS.MoveNext
    Wend
    RS.Close
    Exit Sub
EVerEmresasProhibidas:
    MuestraError Err.Number, Err.Description & vbCrLf & " Consulte soporte t�cnico"
    Set RS = Nothing
End Sub



Private Sub txtNumFac_GotFocus(Index As Integer)
    ObtenerFoco txtNumFac(Index)
End Sub

Private Sub txtNumFac_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtNumFac_LostFocus(Index As Integer)
    txtNumFac(Index).Text = Trim(txtNumFac(Index).Text)
    If txtNumFac(Index).Text = "" Then Exit Sub
    If Not IsNumeric(txtNumFac(Index).Text) Then
        MsgBox "Campo numerico.", vbExclamation
        If Index = 4 Then txtNumFac(Index).Text = ""
        PonerFoco txtNumFac(Index)
    End If
End Sub

Private Sub txtSerie_GotFocus(Index As Integer)
    ObtenerFoco txtSerie(Index)
End Sub

Private Sub txtSerie_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtSerie_LostFocus(Index As Integer)
    If txtSerie(Index).Text = "" Then Exit Sub
    txtSerie(Index).Text = UCase(txtSerie(Index).Text)
    If Index = 4 Then txtSerie(Index).Text = Mid(txtSerie(Index).Text, 1, 1)
End Sub


Private Function LeerGuardarOrdenacion(Leer As Boolean, Cobros As Boolean, Valor As Integer) As Integer
Dim C As String
Dim NF As Integer
Dim Fichero As String

On Error GoTo ELeerGuardarOrdenacion
    LeerGuardarOrdenacion = 0
    
    NF = FreeFile
    If Cobros Then
        Fichero = App.Path & "\OrdenCob.xdf"
    Else
        Fichero = App.Path & "\OrdenPag.xdf"
    End If
    If Leer Then
        
        If Dir(Fichero, vbArchive) <> "" Then
            
            Open Fichero For Input As #NF
            Line Input #NF, C
            Close #NF
            
            LeerGuardarOrdenacion = Val(C)
    
        End If
    Else
        
            Open Fichero For Output As #NF
            Print #NF, Valor
            Close #NF
    
    End If
    Exit Function
ELeerGuardarOrdenacion:
    Err.Clear
End Function



Private Sub PonerValoresPorDefectoDevilucionRemesa()
Dim FP As Ctipoformapago

    On Error GoTo EPonerValoresPorDefectoDevilucionRemesa
    
    
    Set FP = New Ctipoformapago
    FP.Leer vbTipoPagoRemesa
    Me.txtConcepto(1).Text = FP.condecli
    Me.txtConcepto(2).Text = FP.conhapro
    'Ampliaciones
    Combo2(0).ListIndex = FP.ampdecli
    Combo2(1).ListIndex = FP.amphapro
    
    'Que carge el concepto
    txtConcepto_LostFocus 1
    txtConcepto_LostFocus 2
    Set FP = Nothing
    Exit Sub
EPonerValoresPorDefectoDevilucionRemesa:
    MuestraError Err.Number, "PonerValoresPorDefectoDevilucionRemesa"
    Set FP = Nothing
End Sub


Private Sub CargalistaCuentas()
    List1.Clear
    If CadenaDesdeOtroForm <> "" Then
        Do
            I = InStr(1, CadenaDesdeOtroForm, "|")
            If I > 0 Then
                SQL = Mid(CadenaDesdeOtroForm, 1, I - 1)
                CadenaDesdeOtroForm = Mid(CadenaDesdeOtroForm, I + 1)
                CuentaCorrectaUltimoNivel SQL, CuentasCC
                SQL = SQL & "      " & CuentasCC
                List1.AddItem SQL
            End If
        Loop Until I = 0
        CadenaDesdeOtroForm = ""
        
        'Genero Cuentas CC  (por no declarar mas variables vamos)
        CuentasCC = ""
        For I = 0 To List1.ListCount - 1
            SQL = Mid(List1.List(I), 1, vEmpresa.DigitosUltimoNivel)
            CuentasCC = CuentasCC & SQL & "|"
        Next I
    Else
        CuentasCC = ""
    End If
    
End Sub

Private Sub CargaGrupo()

    On Error GoTo ECargaGrupo
    
    SQL = "Select codmacta,nommacta FROM cuentas where grupotesoreria ='" & DevNombreSQL(SQL) & "'"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not miRsAux.EOF
        SQL = miRsAux!codmacta
        If InStr(1, CuentasCC, SQL & "|") > 0 Then
            I = 1
        Else
            CuentasCC = CuentasCC & SQL & "|"
            SQL = SQL & "      " & miRsAux!Nommacta
            List1.AddItem SQL
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If I > 0 Then MsgBox "Algunas cuentas YA habian sido insertadas", vbExclamation
    Exit Sub
ECargaGrupo:
    MuestraError Err.Number, "CargaGrupo"
End Sub

Private Function ComprobarEfectosBorrar() As Boolean
Dim J As Integer
Dim Dias As Integer
Dim Tipopago As Byte
    ComprobarEfectosBorrar = False
    
    For J = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(J).Checked Then

                If ListView2.ListItems(J).Tag = 2 Then
                    'Tipopago = vbPagare
                    Tipopago = 2
                ElseIf ListView2.ListItems(J).Tag = 3 Then
                    'Tipopago = vbTalon
                    Tipopago = 3
                Else
                    'Tipopago = vbTipoPagoRemesa
                    Tipopago = 1
                End If
        
                    
                'Datos bancos. Importe maximo para dias 1, dias2 si no llega
                If Tipopago = 3 Then
                    'Sobre talones
                    'SQL = "100000000,talondias,talondias"
                    SQL = "talondias"
                ElseIf Tipopago = 2 Then
                    'SQL = "100000000,pagaredias,pagaredias"
                    SQL = "pagaredias"
                Else
                    'Efectos.
                    'SQL = "remesariesgo,remesadiasmenor,remesadiasmayor"
                    SQL = "remesadiasmenor"
                End If
   
                    
                'ANTES   Marzo 2011
                'Datos bancos. Importe maximo para dias 1, dias2 si no llega
''                If SubTipo = 3 Then
''                    'Sobre talones
''                    'SQL = "100000000,talondias,talondias"
''                    SQL = "talondias"
''                ElseIf SubTipo = 2 Then
''                    'SQL = "100000000,pagaredias,pagaredias"
''                    SQL = "pagaredias"
''                Else
''                    'Efectos.
''                    'SQL = "remesariesgo,remesadiasmenor,remesadiasmayor"
''                    SQL = "remesadiasmenor"
''                End If
                    
                SQL = "select " & SQL & " from remesas r,ctabancaria b where r.codmacta=b.codmacta and codigo=" & ListView2.ListItems(J).SubItems(1) & " AND anyo = " & ListView2.ListItems(J).Text
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If miRsAux.EOF Then
                    SQL = "Error grave datos banco" & vbCrLf & SQL
                Else
                    SQL = ""
                    Dias = DBLet(miRsAux.Fields(0), "N")
                End If
                
                miRsAux.Close
                
                If SQL <> "" Then
                    MsgBox SQL, vbExclamation
                    Exit Function
                End If
                
                SQL = "Select fecvenci from scobro WHERE codrem=" & ListView2.ListItems(J).SubItems(1)
                SQL = SQL & " AND anyorem = " & ListView2.ListItems(J).Text
                miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                SQL = ""
                If miRsAux.EOF Then
                    'NO hay ningun vencimiento menor.
                    SQL = "UPDATE remesas Set situacion=""Z"" where codigo =" & ListView2.ListItems(J).SubItems(1)
                    SQL = SQL & " AND anyo= " & ListView2.ListItems(J).Text
                    EjecutarSQL SQL
                    
                    
                    
                Else
                    While Not miRsAux.EOF
                        NumRegElim = DateDiff("d", miRsAux!FecVenci, Now)
                        
                        If NumRegElim > Dias Then SQL = "OK"
                        miRsAux.MoveNext
                    Wend
                    
                End If
                
                'Cierro el RS
                miRsAux.Close
                
                            
                
                
                
                
                If SQL = "OK" Then
                    ComprobarEfectosBorrar = True
                    Exit Function
                End If
                    
        End If 'De checked
    Next J


End Function


'Podria darse el caso que el importe del talon/pagare
'Se distinto a la suma de los vencimientos que lo comoponen
'con lo cual el apunte de abono debera contemplar esa diferencia
'y llevarlo a una cuenta 6-7
Private Function ComprobarImportesRemTalonPagare(ImporteRemesa As Currency, ByRef ImporteDocumentos As Currency) As Boolean
Dim DocumentosRecibido As Long

    On Error GoTo EComprobarImportesRemTalonPagare
    

    ComprobarImportesRemTalonPagare = False


    

    CuentasCC = "select l.id from   slirecepdoc l left join  scobro  on l.numserie=scobro.numserie and"
    CuentasCC = CuentasCC & " l.numfaccl=scobro.codfaccl and   l.fecfaccl=scobro.fecfaccl and l.numvenci=scobro.numorden"
    CuentasCC = CuentasCC & " WHERE codrem=" & Text3(3).Text & " AND anyorem=" & Text3(4).Text
    CuentasCC = CuentasCC & " group by id"
    
    
    Set miRsAux = New ADODB.Recordset
    
    miRsAux.Open CuentasCC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ImporteDocumentos = 0
    DocumentosRecibido = 0
    CuentasCC = ""
    While Not miRsAux.EOF
        If IsNull(miRsAux!Id) Then
            CuentasCC = "Hay vencimientos asociados a la remesa sin estar en la recepcion de documentos."
        Else
        
            If DocumentosRecibido <> miRsAux!Id Then
                
                If DocumentosRecibido > 0 Then ImporteDocumentos = ImporteDocumentos + CCur(DBLet(DevuelveDesdeBD("importe", "scarecepdoc", "codigo", CStr(DocumentosRecibido))))
                DocumentosRecibido = miRsAux!Id
        
            End If
            
            
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    If DocumentosRecibido > 0 Then ImporteDocumentos = ImporteDocumentos + CCur(DBLet(DevuelveDesdeBD("importe", "scarecepdoc", "codigo", CStr(DocumentosRecibido))))
    
    Set miRsAux = Nothing
    
    If CuentasCC <> "" Then MsgBox CuentasCC, vbExclamation
    
    
    
    
    ComprobarImportesRemTalonPagare = True
    
    
    
    Exit Function
EComprobarImportesRemTalonPagare:
    MuestraError Err.Number
End Function



Private Function DiferenciaEnImportes(Indice As Integer) As Boolean
Dim RB As ADODB.Recordset
Dim C As String
Dim Impor As Currency
Dim Codigo As Integer

    C = "select scobro.impvenci,l.importe,id from slirecepdoc l left join  scobro  on l.numserie=scobro.numserie and"
    C = C & " l.numfaccl=scobro.codfaccl and   l.fecfaccl=scobro.fecfaccl and l.numvenci=scobro.numorden"
    C = C & " WHERE anyorem = " & ListView2.ListItems(Indice).Text
    C = C & " AND codrem = " & ListView2.ListItems(Indice).SubItems(1) & " ORDER BY ID"
    
    Set RB = New ADODB.Recordset
    RB.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    DiferenciaEnImportes = False
    Codigo = 0
    While Not RB.EOF
        If RB!Id <> Codigo Then
            'Ha cambiado de documento
            If Codigo > 0 Then
                C = DevuelveDesdeBD("importe", "scarecepdoc", "codigo", CStr(Codigo))
                If CCur(C) <> Impor Then
                    'Ya esta clara la diferencia. Nos piramos
                    DiferenciaEnImportes = True
                    RB.Close
                    Exit Function
                End If
            End If
            'Reestablecemos
            Codigo = RB!Id
            Impor = 0
        End If
        'El importe
        Impor = Impor + RB!Importe
        'Siguiente
        RB.MoveNext
    Wend
    RB.Close
        
    If Codigo > 0 Then
        C = DevuelveDesdeBD("importe", "scarecepdoc", "codigo", CStr(Codigo))
        If CCur(C) <> Impor Then
            'Ya esta clara la diferencia. Nos piramos
            DiferenciaEnImportes = True
        End If
    End If
    Set RB = Nothing
End Function


'Cuando eliminamos un pagare/talon en los cuales el importe del talon
'no se corresponde con el de los vencimientos, entonces el program
'debe intentar que se eliminen todos a la vez
Private Function ComprobarTodosVencidos(Indice As Integer) As Boolean
Dim RV As ADODB.Recordset
Dim C As String
Dim Dias As Integer
        
        Set RV = New ADODB.Recordset
        If SubTipo = 3 Then
            C = "talondias"
        Else
            'SQL = "100000000,pagaredias,pagaredias"
            C = "pagaredias"
        End If
              
                    
        C = "select " & C & " from remesas r,ctabancaria b where r.codmacta=b.codmacta and codigo="
        C = C & ListView2.ListItems(Indice).SubItems(1) & " AND anyo = " & ListView2.ListItems(Indice).Text
        RV.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Dias = DBLet(RV.Fields(0), "N")
        RV.Close
    

        C = "select fecvenci from slirecepdoc l left join  scobro  on l.numserie=scobro.numserie and"
        C = C & " l.numfaccl=scobro.codfaccl and   l.fecfaccl=scobro.fecfaccl and l.numvenci=scobro.numorden"
        C = C & " WHERE anyorem= " & ListView2.ListItems(Indice).Text
        C = C & " AND codrem = " & ListView2.ListItems(Indice).SubItems(1)
        
        RV.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        C = ""
        While Not RV.EOF
            NumRegElim = DateDiff("d", RV!FecVenci, Now)
            If NumRegElim < Dias Then C = C & "#"
            RV.MoveNext
        Wend
        RV.Close
        Set RV = Nothing
        If C <> "" Then
            C = "Existen " & Len(C) & " vencimiento(s)  que no han vencido todavia."
            C = C & vbCrLf & "�Continuar?"
            If MsgBox(C, vbQuestion + vbYesNo) = vbNo Then Exit Function
        End If
        
        ComprobarTodosVencidos = True
End Function


Private Sub CamposRemesaAbono()
       
   Me.txtTexto(0).Text = ""
   Me.txtTexto(1).Text = ""
   
   
   If Text3(3) <> "" And Text3(4).Text <> "" Then
        
        Set RS = New ADODB.Recordset
        SQL = "select importe,nommacta from remesas,cuentas where remesas.codmacta=cuentas.codmacta "
        SQL = SQL & " and anyo=" & Text3(4).Text & " and codigo=" & Text3(3).Text
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            Me.txtTexto(0).Text = RS!Nommacta
            Me.txtTexto(1).Text = Format(RS!Importe, FormatoImporte)
        End If
        RS.Close
        Set RS = Nothing
    End If
    
End Sub



Private Sub EliminarEnRecepcionDocumentos()
Dim CtaPte As Boolean
Dim J As Integer
Dim CualesEliminar As String
On Error GoTo EEliminarEnRecepcionDocumentos

    'Comprobaremos si hay datos
    
        'Si no lleva cuenta puente, no hace falta que este contabilizada
        'Es decir. Solo mirare contabilizados si llevo ctapuente
        CuentasCC = ""
        CualesEliminar = ""
        J = 0
        For I = 0 To 1
            ' contatalonpte
            SQL = "pagarecta"
            If I = 1 Then SQL = "contatalonpte"
            CtaPte = (DevuelveDesdeBD(SQL, "paramtesor", "codigo", "1") = "1")
            
            'Repetiremos el proceso dos veces
            SQL = "Select * from scarecepdoc where fechavto<='" & Format(Text1(17).Text, FormatoFecha) & "'"
            SQL = SQL & " AND   talon = " & I
            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RS.EOF
                    'Si lleva cta puente habra que ver si esta contbilizada
                    J = 0
                    If CtaPte Then
                        If Val(RS!Contabilizada) = 0 Then
                            'Veo si tiene lineas. S
                            SQL = DevuelveDesdeBD("count(*)", "slirecepdoc", "id", CStr(RS!Codigo))
                            If SQL = "" Then SQL = "0"
                            If Val(SQL) > 0 Then
                                CuentasCC = CuentasCC & RS!Codigo & " - No contabilizada" & vbCrLf
                                J = 1
                            End If
                        End If
                    End If
                    If J = 0 Then
                        'Si va benee
                        If Val(DBLet(RS!llevadobanco, "N")) = 0 Then
                            SQL = DevuelveDesdeBD("count(*)", "slirecepdoc", "id", CStr(RS!Codigo))
                            If SQL = "" Then SQL = "0"
                            If Val(SQL) > 0 Then
                                CuentasCC = CuentasCC & RS!Codigo & " - Sin llevar a banco" & vbCrLf
                                J = 1
                            End If
                    
                        End If
                    End If
                    'Esta la borraremos
                    If J = 0 Then CualesEliminar = CualesEliminar & ", " & RS!Codigo
                    
                    RS.MoveNext
            Wend
            RS.Close
            
            
            
        Next I
        
        

        
        If CualesEliminar = "" Then
            'No borraremos ninguna
            If CuentasCC <> "" Then
                CuentasCC = "No se puede eliminar de la recepcion de documentos los siguientes registros: " & vbCrLf & vbCrLf & CuentasCC
                MsgBox CuentasCC, vbExclamation
                
            End If
            Exit Sub
        End If
            
        
        
        'Si k hay para borrar
        CualesEliminar = Mid(CualesEliminar, 2)
        J = 1
        SQL = "X"
        Do
            I = InStr(J, CualesEliminar, ",")
            If I > 0 Then
                J = I + 1
                SQL = SQL & "X"
            End If
        Loop Until I = 0
        
        SQL = "Va a eliminar " & Len(SQL) & " registros de la recepcion de documentos." & vbCrLf & vbCrLf & vbCrLf
        If CuentasCC <> "" Then CuentasCC = "No se puede eliminar de la recepcion de documentos los siguientes registros: " & vbCrLf & vbCrLf & CuentasCC
        SQL = SQL & vbCrLf & CuentasCC
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
            SQL = "DELETE from slirecepdoc where id in (" & CualesEliminar & ")"
            Conn.Execute SQL
            
            SQL = "DELETE from scarecepdoc where codigo in (" & CualesEliminar & ")"
            Conn.Execute SQL
    
        End If

    Exit Sub
EEliminarEnRecepcionDocumentos:
    MuestraError Err.Number, Err.Description
End Sub



Private Sub txtTexto_GotFocus(Index As Integer)
    ObtenerFoco txtTexto(Index)
End Sub

Private Sub txtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
        KEYpress KeyAscii
End Sub

Private Sub GuardaDatosConceptoTalonPagare()
    CuentasCC = "DELETE FROM tmpimpbalance WHERE codusu = " & vUsu.Codigo
    Conn.Execute CuentasCC
  
    If txtTexto(3).Text <> "" Then
        CuentasCC = "Insert into `tmpimpbalance` (`codusu`,`Pasivo`,`codigo`,`QueCuentas`) VALUES (" & vUsu.Codigo
        CuentasCC = CuentasCC & ",'Z',1,'" & DevNombreSQL(txtTexto(3).Text) & "')"
        Ejecuta CuentasCC
        
    End If
    CuentasCC = ""
End Sub





Private Function ComprobacionFechasRemesaN19PorVto() As String
Dim Aux As String

    ComprobacionFechasRemesaN19PorVto = ""
    Aux = "anyorem = " & RS!Anyo & " AND codrem "
    Aux = DevuelveDesdeBD("min(fecvenci)", "scobro", Aux, RS!Codigo)
    If Aux = "" Then
        ComprobacionFechasRemesaN19PorVto = "Error fechas vto"
    Else
        If CDate(Aux) < vParam.fechaini Then
            ComprobacionFechasRemesaN19PorVto = "Vtos con fecha menor que inicio de ejercicio"
        End If
    End If
    If ComprobacionFechasRemesaN19PorVto <> "" Then Exit Function
    
    ComprobacionFechasRemesaN19PorVto = ""
    Aux = "anyorem = " & RS!Anyo & " AND codrem "
    Aux = DevuelveDesdeBD("max(fecvenci)", "scobro", Aux, RS!Codigo)
    If Aux = "" Then
        ComprobacionFechasRemesaN19PorVto = "Error fechas vto"
        Exit Function
    End If
    If CDate(Aux) > DateAdd("yyyy", 1, vParam.fechafin) Then ComprobacionFechasRemesaN19PorVto = "Vtos con fecha mayor que fin de ejercicio"
    
    
    
End Function


Private Sub CargarVtosRecaudaEjecutiva()
Dim LineaOK As Boolean
Dim Importe As Currency


    On Error GoTo eCargarVtosRecaudaEjecutiva
    SQL = "Select numserie,codfaccl,fecfaccl,numorden,fecvenci,impvenci,gastos,impcobro,scobro.codmacta,nommacta,nifdatos"
    SQL = SQL & ",dirdatos,codposta,despobla,desprovi,codbanco ,codsucur,digcontr,scobro.cuentaba"
    SQL = SQL & NumeroDocumento
    SQL = SQL & " ORDER BY numserie,codfaccl,numorden"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Me.ListView5.ListItems.Clear
    
    While Not RS.EOF
        
        
        'If RS!codfaccl = 13188 Then Stop
        
        Set IT = ListView5.ListItems.Add
        IT.Text = RS!NUmSerie
        IT.SubItems(1) = Format(RS!codfaccl, "000000")
        IT.SubItems(2) = Format(RS!fecfaccl, "dd/mm/yyyy")
        IT.SubItems(3) = Format(RS!numorden, "00")
        IT.SubItems(4) = Format(RS!FecVenci, "dd/mm/yyyy")
        
        Importe = DBLet(RS!Gastos, "N")
        Importe = Importe - DBLet(RS!impcobro, "N")
         
        
        IT.SubItems(5) = Format(RS!ImpVenci - Importe, FormatoImporte)
        If Importe <> 0 Then IT.ListSubItems(5).ForeColor = vbBlue   'marcamos con Azul el lw wn importe que tienen gastos y/o parcial
     
    
        IT.SubItems(6) = RS!codmacta
        IT.SubItems(7) = Trim(RS!Nommacta)   'NOMBRE OBLIGADO
        
        'direc
        IT.SubItems(8) = Trim(DBLet(RS!nifdatos, "N"))
        IT.SubItems(10) = Trim(DBLet(RS!dirdatos, "N"))
        IT.SubItems(11) = Right("     " & DBLet(RS!codposta), 5) & " " & Trim(DBLet(RS!desPobla, "N"))
        
        
        
        'codbanco ,codsucur,digcontr,cuentaba
        If DBLet(RS!codbanco, "N") = 0 Then
            SQL = "----"
        Else
            SQL = Format(RS!codbanco, "0000")
        End If
        CuentasCC = SQL & " "
        If DBLet(RS!codsucur, "N") = 0 Then
            SQL = "----"
        Else
            SQL = Format(RS!codsucur, "0000")
        End If
        CuentasCC = CuentasCC & " " & SQL
        'CC,cuentaba
        If Trim(DBLet(RS!digcontr, "T")) = "" Then
            SQL = "--"
        Else
            If Not IsNumeric(RS!digcontr) Then
                SQL = "--"
            Else
                SQL = Right("--" & RS!digcontr, 2)
            End If
        End If
        CuentasCC = CuentasCC & " " & SQL
        If DBLet(RS!Cuentaba, "N") = 0 Then
            SQL = "----------"
        Else
            SQL = Format(RS!Cuentaba, "0000000000")
        End If
        CuentasCC = CuentasCC & " " & SQL
                
        IT.SubItems(9) = CuentasCC
        IT.ToolTipText = IT.SubItems(7)
        
        'Validaciones
        LineaOK = True
        
        
        'No pueden estar vacios ni NOMBRE, NIF,CTABANCO,direccion y boblacion
        'Ademas NIF y ctabanco tendras comprobaciones especiales
        For I = 7 To 11
            If IT.SubItems(I) = "" Then
                LineaOK = False
                IT.ListSubItems(I).ForeColor = vbRed
            End If
        Next
        'NIF
        If IT.SubItems(8) <> "" Then
            If Not Comprobar_NIF(IT.SubItems(8)) Then
                LineaOK = False
                IT.ListSubItems(8).ForeColor = vbRed
            End If
        End If
        
        'Cta banco
        If InStr(1, IT.SubItems(9), "-") > 0 Then
                'EROR tiene un -  que he puesto al formatearla
                LineaOK = False
                IT.ListSubItems(9).ForeColor = vbRed
        End If
        
        If Not LineaOK Then
            IT.Bold = True
            IT.ForeColor = vbRed
        End If
        RS.MoveNext
        
    Wend
    RS.Close
    
    
    
eCargarVtosRecaudaEjecutiva:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        cmdRecaudaEjec.Enabled = False
    End If
    Set RS = Nothing
End Sub






Private Sub ReclamacionGargarList()
    ListView6.ListItems.Clear
    
    SQL = "SELECT fechaadq,maidatos,razosoci,nommacta FROM  tmpentrefechas,cuentas WHERE fechaadq=codmacta  "
    SQL = SQL & " AND codUsu = " & vUsu.Codigo & " AND "
    If Me.optReclama(0).Value Then
        'Sin email
        SQL = SQL & " coalesce(maidatos,'')='' "
        ListView6.Checkboxes = False
    Else
        SQL = SQL & " maidatos<>'' "
        ListView6.Checkboxes = True
    End If
    SQL = SQL & " GROUP BY fechaadq  "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set IT = ListView6.ListItems.Add
        IT.Text = RS!fechaadq
        IT.SubItems(1) = RS!Nommacta
        IT.SubItems(2) = DBLet(RS!maidatos, "T")
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing

End Sub





Private Sub LeerGuardarBancoDefectoEntidad(Leer As Boolean)
On Error GoTo eLeerGuardarBancoDefectoEntidad

    I = -1
    SQL = App.Path & "\BancRemEn.xdf"
    If Leer Then
        txtCta(3).Text = ""
        If Dir(SQL, vbArchive) <> "" Then
            I = FreeFile
            Open SQL For Input As #I
            If Not EOF(I) Then
                Line Input #I, SQL
                txtCta(3).Text = SQL
                txtCta(3).Tag = SQL
            End If
        End If
    
    Else
        'Guardar
        If Me.txtCta(3).Text = "" Then
            If Dir(SQL, vbArchive) <> "" Then Kill SQL
        Else
            I = FreeFile
            Open SQL For Output As #I
            Print #I, txtCta(3).Text
            
        End If
        
        
    End If
    
    If I >= 0 Then Close #I
    Exit Sub
eLeerGuardarBancoDefectoEntidad:
    Err.Clear
End Sub
